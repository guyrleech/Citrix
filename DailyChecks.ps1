#requires -version 3.0

<#
    Email quick health report of Citrix XenApp/XenDesktop 7.x environment

    Guy Leech, 2018

    Modification history:

    02/05/18  GL   Guard around Get-BrokerApplciationGroup so doesn't error on older 7.x, only do highest used machines for XenApp, not VDI
                   Added timeout when getting boot time from machine via function in Guys.Common.Functions.psm1

    08/05/18  GL   Added Ghost session detection to long time disconnected user table

    09/05/18  GL   Added tags column to most heavily used machines table

    10/05/18  GL   Added option to exclude machine names via regular expression

    26/05/18  GL   Added restart schedule statistics for XenApp delivery groups

    27/05/18  GL   Added tags to overdue reboot 

    30/05/18  GL   Added LoadIndex and LoadIndexes columns

    31/05/18  GL   Filter highest LoadIndex machines where not in maintenance mode

    08/06/18  GL   Added VMware support for use where machines aren't power managed by Citrix

    19/03/20  GL   Added support for Citrix cloud and email server, vCenter & PVS credentials
#>

<#
.SYNOPSIS

Send an HTML email report of some Citrix health checkpoints such as machines not rebooted recently, machines not powered up, not registered or in maintenance mode, users disconnected for too long, file share capacities
Also includes application groups and desktops with tag restrictions

.DESCRIPTION

.PARAMETER profileName

The name of a previosuly saved profile created using the secure client file created in the Cloud portal

.PARAMETER ddcs

Comma separated list of Desktop Delivery Controllers to extract information from although the cmdlets must be available from where the script runs from, e.g. where Studio is installed
Only specify one DDC if you have more than one but they share the same SQL database

.PARAMETER pvss

Comma separated list of Provisioning Services servers to extract information from although the cmdlets must be available from where the script runs from, e.g. where the PVS console is installed
Only specify one PVS server if you have more than one but they share the same SQL database

.PARAMETER pvsCredential

Credentials to access PVS. Use when the account under which the script runs does not have PVS access.

.PARAMETER vcenterCredential

Credentials to access VMware vCenter. Use when the account under which the script runs does not have vCenter access.

.PARAMETER vCentres

Comma separated list of VMware vCentres to connect to. Use this if VMs in Citrix are not power managed

.PARAMETER UNCs

Comma separated list of file shares to report on capacity for

.PARAMETER mailServer

Name of the SMTP server to use to send summary email

.PARAMETER proxyMailServer

If the mail server only allows SMTP connections from specific machines, use this option to proxy the email via that machine

.PARAMETER from

The email address which the email will be sent from

.PARAMETER emailCredential

Credential to authenticate on mail server

.PARAMETER subject

The subject of the email

.PARAMETER qualifier

Text to prepend to the subject of the email. Use the default subject but use this option to specify an environment or customer name

.PARAMETER recipients

Comma separated list of email addresses to send the report to

.PARAMETER noCheckVDAs

Do not check properties on VDAs. Use when they are not accessible from where the script is run or the account running the script does not have access

.PARAMETER maxRecords

Some of the Citrix cmdlets by default onkly return 250 records although this will produce a warning

.PARAMETER disconnectedMinutes

Sessions which have been disconnected longer than this will be reported. Specify a value of 0 to not report on these

.PARAMETER lastRebootedDaysAgo

Machines last rebooted more than this number of days ago will be reported on.

.PARAMETER excludedTags

Comma separated list of tag names to exlude from the report

.PARAMETER excludedMachines

Reegular expression to match against machine names to exlude from the report

.PARAMETER excludedTags

Comma separated list of tag names to exlude from the report

.PARAMETER topCount

Specified how many items will be included where the top n items are displayed, e.g. servers with the most number of sessions

.PARAMETER jobTimeout

How long in seconds to wait for a remote machine to return its boot time before aborting the command

.PARAMETER logFile

Append to a log file at the specified location

.EXAMPLE

& '.\Daily checks.ps1' -ddcs ctxddc01 -pvss -ctxpvs01 -mailserver smtpserver -proxymailserver msscom01 -qualifier "Constoso" -recipients support@somehwere.com,bob@contoso.com

Extract data from Delivery Controller ctxddc01 and PVS server ctxpvs01 and email the results via msscom01 as realying is not allowed via SMTP server smtpserver

.EXAMPLE

& '.\Daily checks.ps1' -profileName CloudAdmin -pvss -ctxpvs01 -mailserver smtpserver -proxymailserver msscom01 -qualifier "Constoso" -recipients support@somehwere.com,bob@contoso.com

Extract data from the Citrix Cloud with credentials & details as stored in the previously created profile "CloudAdmin" and PVS server ctxpvs01 and email the results via msscom01 as realying is not allowed via SMTP server smtpserver

.NOTES

Uses local PowerShell cmdlets for PVS, DDCs and VMware, as well as Active Directory, so run from a machine where both PVS and Studio consoles and the VMware PowerCLI are installed.
Also uses an additional module Guys.Common.Functions.psm1 which should be placed into the same folder as the main script itself.

To store credentials for Citrix Cloud, download the secrets file via this article https://whatisavpnconnection.blogspot.com/2014/08/xenapp-xendesktop-remote-powershell-sdk.html and run the following with that csv file downloaded as an argument

Set-XDCredentials -CustomerId "YourCustomerId" -ProfileType CloudApi -StoreAs CloudAdmin -SecureClientFile "C:\secureclient.csv" 

Where "CloudAdmin" is what is then passed to the -ProfileName argument
#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='OnPrem')]
    [string[]]$ddcs = @( $env:Computername ) ,
    [string[]]$pvss = @( ) ,
    [string[]]$UNCs = @() ,
    [string[]]$vCentres = @() ,
    [Parameter(ParameterSetName='Cloud')]
    [string]$profileName ,
    [string]$mailserver ,
    [string]$proxyMailServer = 'localhost' ,
    [System.Management.Automation.PSCredential]$emailCredential ,
    [System.Management.Automation.PSCredential]$pvsCredential ,
    [System.Management.Automation.PSCredential]$vCenterCredential ,
    [string]$from = "$env:Computername@$env:userdnsdomain" ,
    [string]$subject = "Daily checks $(Get-Date -Format F)" ,
    [string]$qualifier ,
    [string[]]$recipients ,
    [string]$excludedMachines = '^$' ,
    [int]$disconnectedMinutes = 480 ,
    [int]$lastRebootedDaysAgo = 7 ,
    [int]$topCount = 5 ,
    [string[]]$excludedTags ,
    [string]$logfile ,
    [int]$maxRecords = 2000 ,
    [int]$jobTimeout = 30 ,
    [switch]$noCheckVDAs ,
    [string[]]$snapins = @( 'Citrix.Broker.Admin.*'  ) ,
    [string[]]$modules = @( "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll" , 'Guys.Common.Functions.psm1' ) ,
    [string]$vmwareModule = 'VMware.VimAutomation.Core'
)

if( ! [string]::IsNullOrEmpty( $logfile ) )
{
    Start-Transcript -Append $logfile
}

## Can't use WMI/CIM since servers could be non-Windows
Add-Type @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
    public static class Disk
    {
        // Thanks to https://www.pinvoke.net/default.aspx/kernel32.getdiskfreespaceex
        [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool GetDiskFreeSpaceEx(
                string lpDirectoryName, 
                out ulong lpFreeBytesAvailable, 
                out ulong lpTotalNumberOfBytes, 
                out ulong lpTotalNumberOfFreeBytes);
    }
}
'@

ForEach( $snapin in $snapins )
{
    Add-PSSnapin $snapin -ErrorAction Continue
}

ForEach( $module in $modules )
{
    Import-Module $module -ErrorAction SilentlyContinue
    [bool]$loaded = $?
    if( ! $loaded -and $module -notmatch '^[a-z]:\\' -and  $module -notmatch '^\\\\' ) ## only check script folder if not an absolute or UNC path
    {
        ## try same folder as the script if there is no path in the module name
        Import-Module (Join-Path ( & { Split-Path -Path $myInvocation.ScriptName -Parent } ) $module ) -ErrorAction Continue
        $loaded = $?
    }
    if( ! $loaded )
    {
        Write-Warning "Unable to load module `"$module`" so functionality may be limited"
    }
}

[string]$body = ''
$deliveryGroupStats = New-Object -TypeName System.Collections.Generic.List[psobject]
$possiblyOverdueReboot = New-Object -TypeName System.Collections.Generic.List[psobject]
$notPoweredOn = New-Object -TypeName System.Collections.Generic.List[psobject]
$poweredOnUnregistered = New-Object -TypeName System.Collections.Generic.List[psobject]
$longDisconnectedUsers = New-Object -TypeName System.Collections.Generic.List[psobject]
$highestUsedMachines = New-Object -TypeName System.Collections.Generic.List[psobject]
$highestLoadIndexes = New-Object -TypeName System.Collections.Generic.List[psobject]
$sites = New-Object -TypeName System.Collections.Generic.List[psobject]
$pvsRetries = New-Object -TypeName System.Collections.Generic.List[psobject]
$fileShares = New-Object -TypeName System.Collections.Generic.List[psobject]
$deliveryGroupStatsVDI = New-Object -TypeName System.Collections.Generic.List[psobject]
$deliveryGroupStatsXenApp = New-Object -TypeName System.Collections.Generic.List[psobject]
$taggedApplicationGroups = New-Object System.Collections.ArrayList
$taggedDesktops = New-Object System.Collections.ArrayList
$failedToGetBootTime = New-Object System.Collections.ArrayList

## Fix issue where scheduled task doesn't pass as an array
if( $ddcs.Count -eq 1 -and $ddcs[0].IndexOf(',') -ge 0 )
{
    $ddcs = $ddcs[0] -split ','
}

if( $pvss.Count -eq 1 -and $pvss[0].IndexOf(',') -ge 0 )
{
    $pvss = $pvss[0] -split ','
}

if( $UNCs -and $UNCs.Count -eq 1 -and $UNCs[0].IndexOf(',') -ge 0 )
{
    $UNCs = $UNCs[0] -split ','
}

if( $excludedTags -and $excludedTags.Count -eq 1 -and $excludedTags[0].IndexOf(',') -ge 0 )
{
    $excludedTags = $excludedTags[0] -split ','
}

if( $vCentres -and $vCentres.Count -eq 1 -and $vCentres[0].IndexOf(',') -ge 0 )
{
    $vCentres = $vCentres[0] -split ','
}

$vic = $null
if( $vCentres -and $vCentres.Count -gt 0 )
{
    Import-Module $vmwareModule -ErrorAction Stop

    ## Disable certificate and deprecation warnings
    $null = Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -DisplayDeprecationWarnings:$false 
    [hashtable]$vCenterParams = @{ 'Server' = $vCentres }
    if( $PSBoundParameters[ 'vCenterCredential' ] )
    {
        $vCenterParams.Add( 'credential' , $vCenterCredential )
    }
	if( ! ( $vic = Connect-VIServer @vCenterParams ) -or ! $? )
	{
  	    Throw "Failed to connect to vCenters $vCentres"
	}
}

$fileShares = ForEach( $UNC in $UNCs )
{
    [uint64]$userFreeSpace = 0
    [uint64]$totalSize = 0
    [uint64]$totalFreeSpace = 0
    if( [PInvoke.Win32.Disk]::GetDiskFreeSpaceEx( $UNC , [ref]$userFreeSpace , [ref]$totalSize , [ref]$totalFreeSpace ) )
    {
        [pscustomobject]@{ 'UNC' = $UNC ; 'Free Space (GB)' = [math]::Round( $totalFreeSpace / 1GB ) ; 'Total Size (GB)' = [math]::Round( $totalSize / 1GB ) ; 'Percentage Free Space' = [math]::Round( ( $totalFreeSpace / $totalSize ) * 100 , 1 ) }
    }
    else
    {
        $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Error "Failed to get details for $UNC : $LastError"
    }
}

ForEach( $pvs in $pvss )
{
    if( Get-Command -Name Set-PvsConnection -ErrorAction SilentlyContinue )
    {
        [hashtable]$pvsParameters = @{ 'Server' = $pvs }
        if( $PSBoundParameters[ 'pvsCredential' ] )
        {
            $pvsParameters += @{
                'User' = ($pvsCredential.UserName -split '\\')[-1]
                'Domain' = ($pvsCredential.UserName -split '\\')[0]
                'Password' = $pvsCredential.GetNetworkCredential().Password
            }
        }
        Set-PvsConnection @pvsParameters
        $pvsParameters[ 'Password' ] = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
        $pvsParameters = $null

        ## Status is comma separated value where first field is the number of retries
        $pvsRetries += Get-PvsDeviceInfo| Select -Property Name,@{n='PVS Server';e={$_.ServerName}},SiteName,CollectionName,DiskLocatorName,
            @{n='Retries';e={($_.status -split ',')[0] -as [int]}},DiskVersion
    }
    else
    {
        Write-Warning "PVS cmdlets not detected so unable to report on PVS server $pvs"
    }
}

ForEach( $ddc in $ddcs )
{
    [hashtable]$params = @{}
    if( $PSBoundParameters[ 'profileName' ] )
    {
        Get-XDAuthentication -ProfileName $profileName -ErrorAction Stop
        $ddc = 'Cloud'
    }
    else
    {
        $params.Add( 'AdminAddress' , $ddc )
    }
    [array]$machines = @( Get-BrokerMachine @params -MaxRecordCount $maxRecords | Where-Object { $_.MachineName -notmatch $excludedMachines } )
    [array]$users = @( Get-BrokerSession @params -MaxRecordCount $maxRecords  | Where-Object { $_.MachineName -notmatch $excludedMachines } )
    [array]$XenAppDeliveryGroups = @( Get-BrokerDesktopGroup @params -SessionSupport MultiSession )
    [int]$registeredMachines = $machines | Where-Object { $_.RegistrationState -eq 'Registered' } | Measure-Object | Select -ExpandProperty Count

    $body += "Got $($machines.Count) machines from $ddc with $(($users | Where-Object { $_.SessionState -eq 'Active' }).Count) users active and $(($users | Where-Object { $_.SessionState -eq 'Disconnected' }).Count) disconnected`n"
   
    ## See what if any app groups are tag restricted and then get number of available tagged machines
    if( ( Get-Command -Name Get-BrokerTag -ErrorAction SilentlyContinue ) `
        -and ( Get-Command -Name Get-BrokerApplicationGroup -ErrorAction SilentlyContinue ) )## came later on in 7.x so not necessarily present
    {
        [array]$allApplicationGroups = @( Get-BrokerApplicationGroup @params )
        Get-BrokerTag @params | ForEach-Object `
        {
            $tag = $_
            if( ! $excludedTags -or $excludedTags -notcontains $tag.Name )
            {
                ## Now find all app groups restricted by this tag
                $allApplicationGroups | Where-Object { $_.RestrictToTag -eq $tag.Name } | ForEach-Object `
                {
                    $applicationGroup = $_
                    ## now find workers with this tag
                    [int]$taggedMachinesAvailable = $machines | Where-Object { $_.Tags -contains $tag.Name -and $_.InmaintenanceMode -eq $false -and $_.RegistrationState -eq 'Registered' -and $_.WindowsConnectionSetting -eq 'LogonEnabled' } | Measure-Object | Select -ExpandProperty Count
                    [int]$taggedMachinesTotal = $machines | Where-Object { $_.Tags -contains $tag.Name } | Measure-Object | Select -ExpandProperty Count
                    $null = $taggedApplicationGroups.Add( [pscustomobject]@{'Application Group' = $applicationGroup.Name ; 'Tag' = $tag.Name ; 'Tag Description' = $tag.Description ;
                        'Machines available' = $taggedMachinesAvailable ; 'Total machines tagged' = $taggedMachinesTotal ; 'Percentage Available' = [math]::Round( ( $taggedMachinesAvailable / $taggedMachinesTotal ) * 100 )} )
                }
            }
            ## Now check if any delivery groups have desktops which are tag restricted
            $XenAppDeliveryGroups | ForEach-Object `
            {
                $deliveryGroup = $_
                Get-BrokerEntitlementPolicyRule -DesktopGroupUid $deliveryGroup.uid @params -RestrictToTag $tag.Name | ForEach-Object `
                {
                    $desktop = $_
                    [int]$taggedMachinesAvailable = $machines | Where-Object { $_.DesktopGroupName -eq $deliveryGroup.Name -and $_.Tags -contains $tag.Name -and $_.InmaintenanceMode -eq $false -and $_.RegistrationState -eq 'Registered' -and $_.WindowsConnectionSetting -eq 'LogonEnabled' } | Measure-Object | Select -ExpandProperty Count
                    [int]$taggedMachinesTotal = $machines | Where-Object { $_.DesktopGroupName -eq $deliveryGroup.Name -and $_.Tags -contains $tag.Name } | Measure-Object | Select -ExpandProperty Count              
                    $null = $taggedDesktops.Add( [pscustomobject]@{ 'Delivery Group' = $deliveryGroup.Name ; 'Published Desktop' = $desktop.PublishedName ; 'Description' = $desktop.Description ; 'Enabled' = $desktop.Enabled ; 'Tag' = $tag.Name ; 'Tag Description' = $tag.Description ;
                        'Machines available' = $taggedMachinesAvailable ; 'Total machines tagged' = $taggedMachinesTotal ; 'Percentage Available' = [math]::Round( ( $taggedMachinesAvailable / $taggedMachinesTotal ) * 100 ) } )
                }
            }
        }
    }

    <#
    $poweredOnUnregistered += @( $machines | Where-Object { $_.PowerState -eq 'On' -and $_.RegistrationState -eq 'Unregistered' -and ! $_.InMaintenanceMode } | Select MachineName,DesktopGroupName,CatalogName,SessionCount)
    
    $notPoweredOn += @( $machines | Where-Object { $_.PowerState -eq 'Off' } | Select @{n='Machine Name';e={($_.MachineName -split '\\')[-1]}},DesktopGroupName,CatalogName,InMaintenanceMode )
    #>
    
    ## add VMware VM details for each machine from Citrix in a hash table on the unqualified name so we can look up efficiently when enumerating over machines from Citrix
    [hashtable]$vmwareMachines = @{}
    if( $vic )
    {
        $machines | ForEach-Object `
        {
            $name = ($_.MachineName -split '\\')[-1]
            $vm = Get-VM -Name $name -ErrorAction SilentlyContinue
            if( ! $vm )
            {
                $vm = Get-VM -Name ($name + '*') ## lest it isn't actually named as per the NetBIOS name. Slightly risk, perhaps define character after name as parameter like _
            }
            if( $vm )
            {
                $vmwareMachines.Add( $name , $vm )
            }
            else
            {
                Write-Warning "Failed to find VM $name"
            }
        }
    }
    $poweredOnUnregistered += @( $machines | Where-Object { $(if( $vic ) { ($vmwareMachines[ ($_.MachineName -split '\\')[-1] ]|Select-Object -ExpandProperty PowerState) -eq 'PoweredOn' } else { $_.PowerState -eq 'On' } ) -and $_.RegistrationState -eq 'Unregistered' -and ! $_.InMaintenanceMode } | Select @{n='Machine Name';e={($_.MachineName -split '\\')[-1]}},DesktopGroupName,CatalogName,InMaintenanceMode )
    
    $notPoweredOn += @( $machines | Where-Object { $(if( $vic ) { ($vmwareMachines[ ($_.MachineName -split '\\')[-1] ]|Select-Object -ExpandProperty PowerState) -eq 'PoweredOff' } else { $_.PowerState -eq 'Off'  } ) } | Select @{n='Machine Name';e={($_.MachineName -split '\\')[-1]}},DesktopGroupName,CatalogName,InMaintenanceMode )
   
    $possiblyOverdueReboot += if( $lastRebootedDaysAgo -and ! $noCheckVDAs )
    {
        [decimal]$slowestRemoteJob = 0
        [decimal]$fastestRemoteJob = [int]::MaxValue
        [datetime]$lastRebootedThreshold = (Get-Date).AddDays( -$lastRebootedDaysAgo )

        $machines | Where-Object { if( $vic ) { ($vmwareMachines[($_.MachineName -split '\\')[-1]]|Select-Object -ExpandProperty PowerState) -eq 'PoweredOn' } else { $_.PowerState -eq 'On'  }} | ForEach-Object `
        {
            $machine = $_
            [string]$machineName = ($machine.MachineName -split '\\')[-1]
            Write-Verbose "Checking if last reboot of $machineName was before $lastRebootedThreshold"
            [scriptblock]$work = 
            {
                ## don't use CIM as may not have PowerShell v3.0
                [Management.ManagementDateTimeConverter]::ToDateTime( ( Get-WmiObject -Class Win32_OperatingSystem -ErrorAction SilentlyContinue | Select -ExpandProperty LastBootUpTime ) )
            }
            
            $timer = [Diagnostics.Stopwatch]::StartNew()
            ## This function has a timeout as some remote commands can take a long time to timeout which massively slows the script

            $lastBootTime = Get-RemoteInfo -computer $machineName -jobTimeout $jobTimeout -work $work

            $timer.Stop()
            Write-Verbose "Remote call to $machineName took $($timer.Elapsed.TotalSeconds) seconds"
            $fastestRemoteJob = [math]::Min( $fastestRemoteJob , $timer.Elapsed.TotalSeconds )
            $slowestRemoteJob = [math]::Max( $slowestRemoteJob , $timer.Elapsed.TotalSeconds )

            if( $lastBootTime -and [datetime]$lastBootTime -lt $lastRebootedThreshold )
            {
                [pscustomobject]@{ 'Machine' = $machineName ; 'Last Rebooted' = $lastBootTime ; 'Delivery Group' = $machine.DesktopGroupName ; 'Machine Catalogue' = $machine.CatalogName ; 'Maintenance Mode' = $machine.InMaintenanceMode ; 'Registration State' = $machine.RegistrationState ; 'User Sessions' = $machine.SessionCount ; 'Tags' = $machine.Tags -join ' ' }
            }
            elseif( ! $lastBootTime )
            {
                Write-Warning "Failed to get boot time for $machineName"
                $null = $failedToGetBootTime.Add( [pscustomobject]@{ 'Machine' = $machineName ; 'Delivery Group' = $machine.DesktopGroupName ; 'Machine Catalogue' = $machine.CatalogName ; 'Maintenance Mode' = $machine.InMaintenanceMode ; 'Registration State' = $machine.RegistrationState ; 'User Sessions' = $machine.SessionCount ; 'Tags' = $machine.Tags -join ' ' } )
            }
        }
        Write-Verbose "Fatest remote job was $fastestRemoteJob seconds, slowest $slowestRemoteJob seconds"
    }
        
    [int]$inMaintenanceModeAndOn = $machines | Where-Object { $_.InMaintenanceMode -eq $true -and $(if( $vic ) { ($vmwareMachines[($_.MachineName -split '\\')[-1]]|Select-Object -ExpandProperty PowerState) -eq 'PoweredOn' } else { $_.PowerState -eq 'On'  }) } | Measure-Object | Select -ExpandProperty Count

    $body += "`t$inMaintenanceModeAndOn powered on machines are in maintenance mode ($([math]::round(( $inMaintenanceModeAndOn / $machines.Count) * 100))%)`n"
    $body += "`t$registeredMachines machines are registered ($([math]::round(( $registeredMachines / $machines.Count) * 100))%)`n"
    $body += "`t$($notPoweredOn.Count) machines are not powered on ($([math]::round(( $notPoweredOn.Count / $machines.Count) * 100))%)`n"
    if( $lastRebootedDaysAgo -and ! $noCheckVDAs )
    {
        $body += "`t$($possiblyOverdueReboot.Count) machines have not been rebooted in last $lastRebootedDaysAgo days ($([math]::round(( $possiblyOverdueReboot.Count / $machines.Count) * 100))%)`n"
        $body += "`t$($failedToGetBootTime.Count) powered on machines failed to return boot time ($([math]::round(( $failedToGetBootTime.Count / $machines.Count) * 100))%)`n"
    }

    ## Find sessions and users disconnected more than certain number of minutes
    if( $disconnectedMinutes )
    {
        $longDisconnectedUsers += @( $users | Where-Object { $_.SessionState -eq 'Disconnected' -and $_.SessionStateChangeTime -lt (Get-Date).AddMinutes( -$disconnectedMinutes ) } | Select UserName,UntrustedUserName,@{n='Machine Name';e={($_.MachineName -split '\\')[-1]}},StartTime,SessionStateChangeTime,IdleDuration,DesktopGroupName )
        $body += "`t$($longDisconnectedUsers.Count) users have been disconnected over $disconnectedMinutes minutes`n"
    }

    ## Retrieve delivery group stats - separate for VDI and XenApp as we are interested in subtly different things
    $deliveryGroupStatsVDI += Get-BrokerDesktopGroup @params -SessionSupport SingleSession | Sort PublishedName | Select @{'n'='Delivery Controller';'e'={$ddc}},PublishedName,Enabled,InMaintenanceMode,DesktopsAvailable,DesktopsDisconnected,DesktopsInUse,@{n='% available';e={[math]::Round( $_.DesktopsAvailable / ($_.DesktopsAvailable + $_.DesktopsDisconnected + $_.DesktopsInUse) * 100 )}},DesktopsPreparing,DesktopsUnregistered 
    $deliveryGroupStatsXenApp += $XenAppDeliveryGroups | ForEach-Object `
    {
        $deliveryGroup = $_.Name
        [string]$rebootState = $null
        [string]$lastRebootsEnded = $null
        Get-BrokerRebootCycle -DesktopGroupName $deliveryGroup @params | Sort -Property StartTime -Descending | Select -First 1 | ForEach-Object `
        {
            if( ! [string]::IsNullOrEmpty( $rebootState ) )
            {
                $rebootState += ','
            }
            $rebootState += $_.State.ToString()
            if( ! [string]::IsNullOrEmpty( $lastRebootsEnded ) )
            {
                $lastRebootsEnded += ','
            }
            if( $_.EndTime )
            {
                $lastRebootsEnded += (Get-Date $_.EndTime -Format G).ToString()
            }
        }
        if( [string]::IsNullOrEmpty( $rebootState ) )
        {
            $rebootState = 'No schedule'
        }
        [int]$availableServers = ($machines | Where-Object { $_.DesktopGroupName -eq $deliveryGroup -and $_.RegistrationState -eq 'Registered' -and $_.InMaintenanceMode -eq $false -and $_.WindowsConnectionSetting -eq 'LogonEnabled' } | Measure-Object).Count
        Select-Object -InputObject $_ -Property @{'n'='Delivery Controller';'e'={$ddc}},PublishedName,Description,Enabled,InMaintenanceMode,
            @{n='Available Servers';e={$availableServers}},
            @{n='Total Servers';e={$_.TotalDesktops }},
            @{n='% machines available';e={[math]::Round( ( $availableServers / $_.TotalDesktops ) * 100 )}},
            @{n='Total Sessions';e={$_.Sessions}},
            @{n='Disconnected Sessions';e={$_.DesktopsDisconnected}},
            TotalApplications,TotalApplicationGroups ,
            @{n='Restart State';e={$rebootState}} ,
            @{n='Restarts Ended';e={$lastRebootsEnded}}
    }

    ## only do this for XenApp as doesn't make sense for single user OS in VDI
    if( $XenAppDeliveryGroups -and $XenAppDeliveryGroups.Count )
    {
        [array]$highestUserCounts = @( $machines | Sort SessionCount -Descending | Select -First $topCount -Property @{n='Machine Name';e={($_.MachineName -split '\\')[-1]}},SessionCount,DesktopGroupName,@{n='Tags';e={$_.Tags -join ', '}} )
        if( $highestUserCounts.Count )
        {
            $highestUsedMachines += $highestUserCounts
            $body += "`tHighest number of concurrent users is $($highestUserCounts[0].SessionCount) on $($highestUserCounts[0].'Machine Name')`n"
        }
        [array]$highestLoadIndices = @( $machines | Where-Object { $_.InMaintenanceMode -eq $false } | Sort LoadIndex -Descending | Select -First $topCount -Property @{n='Machine Name';e={($_.MachineName -split '\\')[-1]}},SessionCount,LoadIndex,@{n='Load Indexes';e={$_.LoadIndexes -join ','}},DesktopGroupName,@{n='Tags';e={$_.Tags -join ', '}} )
        if( $highestLoadIndices.Count )
        {
            $highestLoadIndexes += $highestLoadIndices
            $body += "`tHighest load index is $($highestLoadIndices[0].LoadIndex) on $($highestLoadIndices[0].'Machine Name') with $($highestLoadIndices[0].SessionCount) sessions`n"
        }
    }
    
    $sites += Get-BrokerSite @params | Select Name,@{'n'='Delivery Controller';'e'={$ddc}},PeakConcurrentLicenseUsers,TotalUniqueLicenseUsers,LicensingGracePeriodActive,LicensingOutOfBoxGracePeriodActive,LicensingGraceHoursLeft,LicensedSessionsActiv
}

if( $recipients -and $recipients.Count -and ! [string]::IsNullOrEmpty( $mailserver ) )
{
    if( $recipients.Count -eq 1 -and $recipients[0].IndexOf(',') -ge 0 )
    {
        $recipients = $recipients[0] -split ','
    }
    
    if( ! [string]::IsNullOrEmpty( $qualifier ) )
    {
        $subject = $qualifier + ' ' + $subject
    }

    [string]$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
    $style += "TABLE{border: 1px solid black; border-collapse: collapse;}"
    $style += "TH{border: 1px solid black; background: #dddddd; padding: 5px;}"
    $style += "TD{border: 1px solid black; padding: 5px;}"
    $style += "</style>"

    ## ConvertTo-Html only works for objects, not raw text
    [string]$htmlBody = "<h2>Summary</h2>`n" + $body -split "`n" | ForEach-Object { "<p>$($_ -replace '\t' , '&nbsp;&nbsp;&nbsp;&nbsp;')</p>`n" }

    $htmlBody += $sites | ConvertTo-Html -Fragment -PreContent '<h2>Site Information<h2>'| Out-String

    if( $deliveryGroupStatsVDI -and $deliveryGroupStatsVDI.Count )
    {
        $htmlBody += $deliveryGroupStatsVDI | ConvertTo-Html -Fragment -PreContent "<h2>Summary of $($deliveryGroupStatsVDI.Count) VDI Delivery Groups<h2>" | Out-String
    }

    if( $deliveryGroupStatsXenApp -and $deliveryGroupStatsXenApp.Count )
   {
        $htmlBody += $deliveryGroupStatsXenApp | Sort '% machines available' -Descending | ConvertTo-Html -Fragment -PreContent "<h2>Summary of $($deliveryGroupStatsXenApp.Count) XenApp Delivery Groups<h2>"| Out-String
    }
    
    if( $possiblyOverdueReboot -and $possiblyOverdueReboot.Count )
    {
        $htmlBody += $possiblyOverdueReboot | Sort 'Delivery Group' | ConvertTo-Html -Fragment -PreContent "<h2>$($possiblyOverdueReboot.Count) machines possibly overdue reboot<h2>" | Out-String
    }

    if( $notPoweredOn -and $notPoweredOn.Count )
    {
        $htmlBody += $notPoweredOn | Sort DesktopGroupName | ConvertTo-Html -Fragment -PreContent "<h2>$($notPoweredOn.Count) machines not powered on<h2>" | Out-String
    }
    if( $taggedApplicationGroups -and $taggedApplicationGroups.Count )
    {
        $htmlBody += $taggedApplicationGroups | Sort Tag  | ConvertTo-Html -Fragment -PreContent "<h2>$($taggedApplicationGroups.Count) tag restricted application groups<h2>" | Out-String
    }
    if( $taggedDesktops -and $taggedDesktops.Count )
    {
        $htmlBody += $taggedDesktops | Sort Tag  | ConvertTo-Html -Fragment -PreContent "<h2>$($taggedDesktops.Count) tag restricted published desktops<h2>" | Out-String
    }
    if( $poweredOnUnregistered -and $poweredOnUnregistered.Count -gt 0 )
    {
        $htmlBody += $poweredOnUnregistered | sort DesktopGroupName| ConvertTo-Html -Fragment -PreContent "<h2>$($poweredOnUnregistered.Count) machines powered on and unregistered but not in maintenance mode<h2>" | Out-String
    }
    if( $highestUsedMachines -and $highestUsedMachines.Count -gt 0 )
    {
        $htmlBody += $highestUsedMachines | sort SessionCount -Descending| ConvertTo-Html -Fragment -PreContent "<h2>$($highestUsedMachines.Count) machines with highest number of users<h2>" | Out-String
    }
    if( $highestLoadIndexes -and $highestLoadIndexes.Count -gt 0 )
    {
        $htmlBody += $highestLoadIndexes | sort LoadIndex -Descending| ConvertTo-Html -Fragment -PreContent "<h2>$($highestLoadIndexes.Count) machines with highest load indexes (not in maintenance mode)<h2>" | Out-String
    }
    if( $failedToGetBootTime -and $failedToGetBootTime.Count )
    {
        $htmlBody += $failedToGetBootTime | sort 'Delivery Group' | ConvertTo-Html -Fragment -PreContent "<h2>$($failedToGetBootTime.Count) powered on machines failed to return boot time<h2>" | Out-String
    }
    if( $pvsRetries -and $pvsRetries.Count -gt 0 )
    {
        ## only add boot time now so that we don't do it for all devices
        ## TODO should we wrap these into runspaces lest they timeout and take an age to do so?
        $htmlBody += $pvsRetries | sort Retries -Descending | Select -First $topCount -Property *,@{n='Boot Time';e={Get-Date -Date (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $_.Name -ErrorAction SilentlyContinue | Select -ExpandProperty LastBootupTime) -Format G}} | Where-Object { $_.Retries } | ConvertTo-Html -Fragment -PreContent "<h2>Machines with highest number of PVS retries<h2>" | Out-String
    }
    if( $fileShares -and $fileShares.Count )
    {
        $htmlBody += $fileShares  | sort 'Percentage Free Space' | ConvertTo-Html -Fragment -PreContent "<h2>File share capacities<h2>" | Out-String
    }
    [hashtable]$sessions = @{}
    if( $longDisconnectedUsers -and $longDisconnectedUsers.Count -gt 0 )
    {
        ## see if any of these are ghosts, as in there isn't a session on the server that Citrix thinks there is
        ForEach( $disconnectedUser in $longDisconnectedUsers )
        {
            [array]$serverSessions = $sessions[ $disconnectedUser.'Machine Name' ]
            if( ! $serverSessions -and ! $noCheckVDAs )
            {
                ## Get users from machine - if we just run quser then get error for no users so this method make it squeaky clean
                $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                $pinfo.FileName = "quser.exe"
                $pinfo.Arguments = "/server:$($disconnectedUser.'Machine Name')"
                $pinfo.RedirectStandardError = $true
                $pinfo.RedirectStandardOutput = $true
                $pinfo.UseShellExecute = $false
                $pinfo.WindowStyle = 'Hidden'
                $pinfo.CreateNoWindow = $true
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $pinfo
                if( $process.Start() )
                {
                    if( $process.WaitForExit( $jobTimeout * 1000 ) )
                    {
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
                        } )
                        $sessions.Add( $disconnectedUser.'Machine Name' , $serverSessions )
                    }
                    else
                    {
                        Write-Warning ( "Timeout of {0} seconds waiting for process to exit {1} {2}" -f $jobTimeout , $pinfo.FileName , $pinfo.Arguments )
                    }
                }
                else
                {
                    Write-Warning ( "Failed to start process {0} {1}" -f $pinfo.FileName , $pinfo.Arguments )
                }
            }
            $usersActualSession = $null
            if( $serverSessions )
            {
                [string]$domainname,$username = $disconnectedUser.UserName -split '\\'
                if( [string]::IsNullOrEmpty( $username ) )
                {
                    $username = ($disconnectedUser.UntrustedUserName -split '\\')[-1]
                }
                ForEach( $serverSession in $serverSessions )
                {
                    if( $Username -eq $serverSession.UserName )
                    {
                        $usersActualSession = $serverSession
                        break
                    }
                }
            }
            if( ! $usersActualSession )
            {
                Add-Member -InputObject $disconnectedUser -MemberType NoteProperty -Name 'Ghost Session' -Value 'Yes'
            }
        }
        $htmlBody += $longDisconnectedUsers | sort SessionStateChangeTime | ConvertTo-Html -Fragment -PreContent "<h2>$($longDisconnectedUsers.Count) users disconnected more than $disconnectedMinutes minutes<h2>"| Out-String
    }
    $htmlBody = ConvertTo-Html -PostContent $htmlBody -Head $style

    [hashtable]$mailParams = @{
        'Subject' = $subject 
        'BodyAsHtml' = $true
        'Body' = $htmlBody
        'From' = $from
        'To' = $recipients
        'SmtpServer' = $mailserver
    }
    if( $PSBoundParameters[ 'emailCredential' ] )
    {
        $mailParams.Add( 'credential' , $emailCredential )
    }
    if( $PSBoundParameters[ 'proxyMailServer' ] )
    {
        Invoke-Command -ComputerName $proxyMailServer -ScriptBlock { Send-MailMessage @Using:mailParams }
    }
    else
    {
        Send-MailMessage @mailParams
    }
}

if( $vic )
{
    Disconnect-VIServer -Server $vCentres -Confirm:$false
}

if( ! [string]::IsNullOrEmpty( $logfile ) )
{
    Stop-Transcript
}
