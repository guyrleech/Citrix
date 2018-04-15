#requires -version 3.0
<#
    Script to find disconnected sessions and end them if they have been disconnected over specified period.
    Will aslo atttempt to terminate any processes running for users specified on the command line

    Guy Leech, 2018

    Use this script at your own risk - no warranty provided
#>


<#
.SYNOPSIS

Find disconnected XenApp sessions disconnected over a specified threshold and terminate them. Can also terminate specified processes in case they are preventing logoff.

.PARAMETER threshold

Disconnection time in hours and minutes, e.g. 6:30, over which the session will be disconnected

.PARAMETER ddc

The Desktop Delivery Controller to query to get disconnected session information

.PARAMETER forceIt

Do not prompt for confirmation before terminating stuck processes and disconnecting processes

.PARAMETER logfile

A csv file that will be appended to with details of the sessions terminated

.PARAMETER killProcesses

A comma separated list of processes to terminate if they are running in the users session

.EXAMPLE

& '.\End Disconnected sessions.ps1' -threshold 8:30 -ddc ctxddc01 -killProcesses stuckprocess,anotherstuckprocess -logfile c:\support\disconnected.terminations.csv

End all disconnected sessions found via delivery controller ctxddc01 which have been disconnected for more than 8 hours 30 minutes.
If there are processes still running, before the session is ternminated, called stuckprocessor or anotherstuckprocess they will be terminated.
Results will be appended to the CSV file c:\support\disconnected.terminations.csv

#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
	[Parameter(Mandatory=$false,HelpMessage='Disconnected threshold in -hours:minutes')]
	[string]$threshold ,
    [string]$ddc = 'localhost' ,
    [switch]$forceIt ,
    [string]$logfile ,
    [string[]]$killProcesses
)

[string[]]$snapins = @( 'Citrix.Broker.Admin.*'  )

if( $killProcesses -and $killProcesses -contains 'csrss' )
{
    Write-Error 'Kiling csrss causes BSoDs so not continuing'
    return
}

[int]$hours,[int]$minutes = $threshold -split ':'

if( ( ! $hours -and ! $minutes ) -or $minutes -lt 0 -or $minutes -gt 59 )
{
    Write-Error "Bad threshold of $threshold specified"
    return
} 

## Reform in case minutes wasn't specified
$threshold = "-$([math]::Abs( $hours )):$minutes"

ForEach( $snapin in $snapins )
{
    Add-PSSnapin $snapin -ErrorAction Stop
}

if( $forceIt )
{
     $ConfirmPreference = 'None'
}

# reform incase flattened by scheduled task engine
if( $killProcesses -and $killProcesses.Count -eq 1 -and $killProcesses.IndexOf(',') -ge 0 )
{
    $killProcesses = $killProcesses -split ','
}

$disconnected = @( Get-BrokerSession -AdminAddress $ddc -SessionState 'Disconnected' -Filter { SessionStateChangeTime -lt $threshold } )

Write-Verbose "Got $($disconnected.Count)  disconnected sessions over $threshold`n$($disconnected|select username,UntrustedUserName,HostedMachineName,StartTime,SessionStateChangeTime|Format-Table -AutoSize|Out-String)"

if( $disconnected -and $disconnected.Count -gt 0 )
{
    [array]$processes = @()
    if( $killProcesses -and $killProcesses.Count )
    {
        ForEach( $session in $disconnected )
        {
            [string]$username = $session.Username
            if( [string]::IsNullOrEmpty( $username ) )
            {
                $username = $session.UntrustedUsername
            }
            $username = ($username -split '\\')[-1] ## strip domain name off
            if( ! [string]::IsNullOrEmpty( $session.HostedMachineName ) -and ! [string]::IsNullOrEmpty( $username ) )
            {
                if( (quser /server:$($session.HostedMachineName)|select -skip 1| Where-Object{ $_ -match "[^a-z0-9_]$username\s+(\d+)\s" }) )
                {
                    [int]$sessionId = $Matches[1].Trim()
                    if( $sessionId -gt 0 )
                    {
                        ## have to remote it as doesn't return session ids if run via -ComputerName. Can't check username as may be system processes in that session
                        $processes = @( Invoke-Command -ComputerName $session.HostedMachineName -ScriptBlock { Get-Process -IncludeUserName -Name $using:killProcesses | Where-Object { $_.SessionId -eq $using:sessionId } } )
                        if( $processes -and $processes.Count )
                        {
                            Add-Member -InputObject $session -MemberType NoteProperty -Name ProcessesKilled -Value ( ( $processes | select -ExpandProperty Name ) -join ',' )
                            if( $PSCmdlet.ShouldProcess( "Session $sessionId for $username on $($session.HostedMachineName)" , "Kill processes $(($processes|Select -ExpandProperty Name) -join ',')" ) )
                            {
                                Invoke-Command -ComputerName $session.HostedMachineName -ScriptBlock { $using:processes | Stop-Process -Force -PassThru }
                            }
                        }
                        else
                        {
                            Write-Warning "Found no $($killProcesses -join ',') processes to kill in session $sessionId for $username on $($session.HostedMachineName)"
                        }
                    }
                    else
                    {
                        Write-Warning "Failed to get session id via quser for $username on $($session.HostedMachineName)"
                    }
                }
                else
                {
                    Write-Warning "Failed to get session via quser for $username on $($session.HostedMachineName)"
                }
            }
            else
            {
                Write-Warning "Couldn't get username `"$username`" or host `"$($session.HostedMachineName)`""
            }
        }
    }
    if( $PSCmdlet.ShouldProcess( "$($disconnected.Count) disconnected sessions" , 'Log off' ) )
    {
        if( ! [String]::IsNullOrEmpty($logfile) )
        {    
            $disconnected | Select-Object -Property @{n='Sampled';e={Get-Date}},Username,StartTime,UntrustedUserName,SessionStateChangeTime,HostedMachineName,ClientName,ClientAddress,CatalogName,DesktopGroupName,ControllerDNSName,HostingServerName,ProcessesKilled | Export-Csv -NoTypeInformation -Append -Path $logfile
        }

        $disconnected | Stop-BrokerSession
    }
}
