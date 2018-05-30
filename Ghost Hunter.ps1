#Requires -version 3.0
<#
    Check that all disconnected sessions still exist and if not report the gorey details

    Guy Leech, 2018

    Modification History:

    29/05/18  GL  Added help and remoting of Citrix cmdlets if not available locally
#>

<#
.SYNOPSIS

Search for sessions that Citrix report as being disconnected where that session no longer exists on the specified server. These are referred to as "ghost" sessions.
Optionally, set the state of that session to "hidden" which disables session sharing for that specific session allowing subsequently launched applications for that user to launch ok.

.DESCRIPTION

Ghost sessions should not occur but it has been observed to happen on at least 7.13.
The script will also attempt to find the user's session logoff in the User Profile Service event log on the server reporting the ghost session and also if they have sessions on any other server.

.PARAMETER ddc

The Delivery Controller to query for disconnected sessions.

.PARAMETER hide

Any disconnected session found not to exist will have its Citrix "hidden" flag set so affected users launching new applications won't get an error due to session sharing failure.

.PARAMETER mailServer

Name of the SMTP server to use to send email notifying of ghost sessions. If not specified then the results will be displayed on screen in a grid view

.PARAMETER proxyMailServer

If the mail server only allows SMTP connections from specific machines, use this option to proxy the email via that machine

.PARAMETER from

The email address which the email will be sent from

.PARAMETER subject

The subject of the email. This can include PowerShell expressions which will be evaluated

.PARAMETER forceIt

Suppresses prompting to confirm that a session should be hidden

.PARAMETER historyFile

A file which keeps track of which sessions are already hidden so newly discovered ghost sessions can be highlighted and an email sent

.PARAMETER subject

The subject of the email sent to inform of new ghost sessions

.EXAMPLE

& '.\ghost hunter.ps1' -ddc ctxddc001 -hide -mailserver smtp001 -recipients guy.leech@mars.com,sarah.kennedy@venus.com -historyfile c:\scripts\ghosties.csv

Query disconnected sessions from the Delivery Controller ctxddc001 and any ghost sessions will be set to hidden. If any of these are new since the script was last run,
since they will have been recorded in the c:\scripts\ghosties.csv file, send an email to the listed recipients via SMTP mail server smtp001

.EXAMPLE

& '.\ghost hunter.ps1' -ddc ctxddc001 

Query disconnected sessions from the Delivery Controller ctxddc001 and any ghost sessions will be displayed in an on screen grid view.
Any sessions selected when the "OK" button is clicked in the grid view will be placed in the clipboard.

.NOTES

If running the script as a scheduled task and using -hide then specify the -forceIt parameter otherwise the script will hang since it will prompt to confirm the hidden setting is to be set.
The user running the script must have sufficient privileges to query the sessions via Get-BrokerSession, set the session state to "hidden" and to query remote event logs.

#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [string]$ddc = 'localhost' ,
    [switch]$hide ,
    [string]$proxyMailServer = 'localhost' ,
    [string]$mailserver ,
    [string[]]$recipients ,
    [switch]$forceIt ,
    [string]$historyFile ,
    [string]$subject = "Ghost XenApp sessions detected" ,
    [string]$from  = "$($env:COMPUTERNAME)@$($env:USERDNSDOMAIN)" ,
    [int]$maxRecordCount = 10000 ,
    [string[]]$snapins = @( 'Citrix.Broker.Admin.*'  )
)

ForEach( $snapin in $snapins )
{
    Add-PSSnapin $snapin -ErrorAction Continue ## if this fails then we will try to import the snapin from the DDC specified
}

if( $forceIt )
{
     $ConfirmPreference = 'None'
}

## if no local Citrix PoSH cmdlets then try and pull in from DDC
$DDCsession = $null
if( ( Get-Command -Name Get-BrokerSession -ErrorAction SilentlyContinue ) -eq $null )
{
    if( $ddc -eq 'localhost' )
    {
        Write-Error "Unable to find required Citrix cmdlet Get-BrokerSession - aborting" ## we have already tried a direct Add-PSSnapin
        Exit 1
    }
    else
    {
        $DDCsession = New-PSSession -ComputerName $ddc
        if( $DDCsession )
        {
            $null = Invoke-Command -Session $DDCsession -ScriptBlock { Add-PSSnapin $using:snapins }
            $null = Import-PSSession -Session $DDCsession -Module $snapins
        }
        else
        {
            Write-Error "Unable to remote to $ddc to add required Citrix cmdlet Get-BrokerSession - aborting"
            Exit 1
        }
    }
}

[int]$count = 0

[hashtable]$sessions = @{}

[datetime]$startTime = Get-Date

[array]$results = @( Get-BrokerSession -AdminAddress $ddc -MaxRecordCount $maxRecordCount -SessionState Disconnected | ForEach-Object `
{ 
    $session = $_
    [string]$domainname,$username = $session.UserName -split '\\'
    if( [string]::IsNullOrEmpty( $username ) )
    {
        ## don't know why, it just does this occasionally
        $domainname,$username = $session.UntrustedUserName -split '\\' 
    }
    $count++
    Write-Verbose "$count : checking $UserName on $($session.HostedMachineName)"
    if( [string]::IsNullOrEmpty( $username ) )
    {
        Write-Warning "No user name found for session on $($session.HostedMachineName) via client $($session.ClientName)"
    }
    else
    {
        [array]$serverSessions = $sessions[ $session.HostedMachineName ]
        if( ! $serverSessions )
        {
            ## Get users from machine - if we just run quser then get error for no users so this method make it squeaky clean
            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = "quser.exe"
            $pinfo.Arguments = "/server:$($session.HostedMachineName)"
            $pinfo.RedirectStandardError = $true
            $pinfo.RedirectStandardOutput = $true
            $pinfo.UseShellExecute = $false
            $pinfo.WindowStyle = 'Hidden'
            $pinfo.CreateNoWindow = $true
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $pinfo
            $null = $process.Start()
            $process.WaitForExit()
            ## Output of quser is fixed width but can't do simple parse as SESSIONNAME is empty when session is disconnected so we break it up based on header positions
            [string[]]$fieldNames = @( 'USERNAME','SESSIONNAME','ID','STATE','IDLE TIME','LOGON TIME' )
            [string[]]$allOutput = $process.StandardOutput.ReadToEnd() -split "`n"
            [string]$header = $allOutput[0]
            $serverSessions = @( $allOutput | Select -Skip 1 | ForEach-Object `
            {
                [string]$line = $_
                if( ! [string]::IsNullOrEmpty( $line ) )
                {
                    $result = New-Object -TypeName PSCustomObject
                    For( [int]$index = 0 ; $index -lt $fieldNames.Count ; $index++ )
                    {
                        [int]$startColumn = $header.IndexOf($fieldNames[$index])
                        ## if last column then can't look at start of next field so use overall line length
                        [int]$endColumn = if( $index -eq $fieldNames.Count - 1 ) { $line.Length } else { $header.IndexOf( $fieldNames[ $index + 1 ] ) }
                        try
                        {
                            Add-Member -InputObject $result -MemberType NoteProperty -Name $fieldNames[ $index ] -Value ( $line.Substring( $startColumn , $endColumn - $startColumn ).Trim() )
                        }
                        catch
                        {
                            throw $_
                        }
                    }
                    $result
                }      
            }) 
            $sessions.Add( $session.HostedMachineName , $serverSessions )
        }
        $usersActualSession = $null
        if( $serverSessions )
        {
            ForEach( $serverSession in $serverSessions )
            {
                if( $serverSession.Username -eq $UserName )
                {
                    $usersActualSession = $serverSession
                    break
                }
            }
        }
        if( ! $usersActualSession )
        {
            $otherSessions = @( Get-BrokerSession -AdminAddress ywcxp2003 -UserName "$domainname\$UserName" | ?{ $_.SessionKey -ne $session.SessionKey } )
            Add-Member -InputObject $session -MemberType NoteProperty -Name OtherSessions -Value ( ( $otherSessions | Select -ExpandProperty HostedMachineName ) -join ',' )
            Write-Warning "No session found on server $($session.HostedMachineName) for user $username, has $($otherSessions.Count) other sessions"
            if( $hide -and ! $session.Hidden -and $PSCmdlet.ShouldProcess( $username ,  "Hide session on $($session.HostedMachineName)" )  )
            {
                Write-Verbose "Setting hidden property"
                Set-BrokerSession -InputObject $session -Hidden $true
            }
            ## Get user logon and logoff events from that server to add to output
            $sid = (New-Object System.Security.Principal.NTAccount($domainname + '\' + $username)).Translate([System.Security.Principal.SecurityIdentifier]).value
            $events = @( Get-WinEvent -ComputerName $session.HostedMachineName -FilterHashtable @{ LogName = 'Microsoft-Windows-User Profile Service/Operational' ; id = 1,4 ; UserId = $sid } )
            if( $events -and $events.Count )
            {
                Add-Member -InputObject $session -MemberType NoteProperty -Name ActualLogonTime  -Value ( $events | Where-Object { $_.Id -eq 1 } | Select -First 1 -ExpandProperty TimeCreated )
                Add-Member -InputObject $session -MemberType NoteProperty -Name ActualLogoffTime -Value ( $events | Where-Object { $_.Id -eq 4 } | Select -First 1 -ExpandProperty TimeCreated )
            }
            else
            {
                Write-Warning "Unable to find logon and logoff events for user $username (sid $sid) on $($session.HostedMachineName)"
            }
            $session
        }
    }
})

if( $DDCsession )
{
    $null = Remove-PSSession -Session $DDCsession
    $DDCsession = $null
}

[string]$status = "Found $($results.Count) ghost sessions out of $count disconnected across $($sessions.Count) servers at $(Get-Date $startTime -Format G)"

Write-Verbose $status

if( $results -and $results.Count )
{
    if( [string]::IsNullOrEmpty( $mailserver ) )
    {
        $selected = @( $results | Out-GridView -Title $status -PassThru )
        if( $selected -and $selected.Count )
        {
            $selected | Clip.exe
        }
    }
    else
    {
        [int]$alreadyAlerted = 0

        if( ! [string]::IsNullOrEmpty( $historyFile )  )
        {
            [array]$existing = $null
            if( Test-Path -Path $historyFile -PathType Leaf -ErrorAction SilentlyContinue )
            {
                $existing = Import-Csv -Path $historyFile
            }
            ForEach( $result in $results )
            {
                ForEach( $ghost in $existing )
                {
                    if( $ghost.SessionId -eq $result.SessionId -and $ghost.hostedmachinename -eq $result.HostedMachineName )
                    {
                        $alreadyAlerted++
                        Write-Verbose "Already alerted on $($result.Username) ($($result.untrustedusername)) on $($result.HostedMachineName)"
                        break
                    }            
                }
            }
         
            $results | Export-Csv -Path $historyFile
        }

        if( $alreadyAlerted -ne $results.Count )
        {
            if( $recipients[0].IndexOf( ',' ) -ge 0 )
            {
                $recipients = $recipients -split ','
            }
        
            [string]$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
            $style += "TABLE{border: 1px solid black; border-collapse: collapse;}"
            $style += "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
            $style += "TD{border: 1px solid black; padding: 5px; }"
            $style += "</style>"

            [string]$body = ($results | Select UserName,UntrustedUserName,HostedMachineName,ActualLogonTime,ActualLogoffTime,OtherSessions,Hidden) |  ConvertTo-Html -Head $style
            if( [string]::IsNullOrEmpty( $subject ) )
            {
                $subject = Invoke-Expression $status ## expands any cmdlets such as Get-Date
            }
            Invoke-Command -ComputerName $proxyMailServer -ScriptBlock { Send-MailMessage -SmtpServer $using:mailserver -To $using:recipients -From $using:from -Subject $using:subject -Body $using:body -BodyAsHtml }
        }
        else
        {
            Write-Verbose "Not emailing as have already alerted on all $($results.Count) ghost sessions"
        }
    }
}
