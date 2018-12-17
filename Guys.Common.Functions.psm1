<#
    Common functions used by multiple scripts

    Guy Leech, 2018

    Modification history:

    20/06/18    GL  Added function to get Citrix PVS devices

    12/12/18    GL  Added -quiet switch to Get-RemoteInfo
#>

## internal use only
Function Extract-RemoteInfo( [array]$remoteInfo )
{
    [hashtable]$fields = @{}
    if( $remoteInfo -and $remoteInfo.Count )
    {
        $osinfo,$logicalDisks,$cpu,$domainMembership = $remoteInfo
        $fields += `
            @{
                'Boot_Time' = [Management.ManagementDateTimeConverter]::ToDateTime( $osinfo.LastBootUpTime )
                'Available Memory (GB)' = [Math]::Round( $osinfo.FreePhysicalMemory / 1MB , 1 )
                'Committed Memory %' = 100 - [Math]::Round( ( $osinfo.FreeVirtualMemory / $osinfo.TotalVirtualMemorySize ) * 100 , 1 )
                'CPU Usage %' = $cpu
                'Free disk space %' = ( $logicalDisks | Sort DeviceID | ForEach-Object { [Math]::Round( ( $_.FreeSpace / $_.Size ) * 100 , 1 ) }) -join ' '
                'Domain Membership' = $domainMembership
            }
    }
    else
    {
        Write-Warning "Get-RemoteInfo failed: $remoteError"
    }
    $fields
}

Function Get-RemoteInfo
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$computer , 
        [int]$jobTimeout = 60 , 
        [Parameter(Mandatory=$true)]
        [scriptblock]$work ,
        [switch]$quiet ,
        [int]$miscparameter1 ## passed through to the script block
    )

    $results = $null

    [scriptblock]$code = `
    {
        Param([string]$computer,[scriptblock]$work,[int]$miscparameter1)
        Invoke-Command -ComputerName $computer -ScriptBlock $work
    }

    try
    {
        ## use a runspace so we can have a timeout  
        $runspace = [RunspaceFactory]::CreateRunspace()
        $runspace.Open()
        $command = [PowerShell]::Create().AddScript($code)
        $command.Runspace = $runspace
        $null = $command.AddParameters( @( $computer , $work , $miscparameter1 ) )
        $job = $command.BeginInvoke()

        ## wait for command to finish
        $wait = $job.AsyncWaitHandle.WaitOne( $jobTimeout * 1000 , $false )

        if( $wait -or $job.IsCompleted )
        {
            if( $command.HadErrors )
            {
                if( ! $quiet )
                {
                    Write-Warning "Errors occurred in remote command on $computer :`n$($command.Streams.Error)"
                }
            }
            else
            {
                $results = $command.EndInvoke($job)
                if( ! $results -and ! $quiet )
                {
                    Write-Warning "No data returned from remote command on $computer"
                }
            }   
            ## if we do these after timeout too then takes an age to return which defeats the point of running via a runspace
            $command.Dispose() 
            $runSpace.Dispose()
        }
        else
        {
            if( ! $quiet )
            {
                Write-Warning "Job to retrieve info from $computer is still running after $jobTimeout seconds so aborting"
            }
            $null = $command.BeginStop($null,$null)
            ## leaking command and runspace but if we dispose it hangs
        }
    }   
    catch
    {
        if( ! $quiet )
        {
            Write-Error "Failed to get remote info from $computer : $($_.ToString())"
        }
    }
    $results
}

Function Get-PVSDevices
{
    Param
        (
        [Parameter(ParameterSetName='Manual',mandatory=$true,HelpMessage='Comma separated list of PVS servers')]
        [string[]]$pvsServers ,
        [Parameter(ParameterSetName='Manual',mandatory=$true,HelpMessage='Comma separated list of Delivery controllers')]
        [string[]]$ddcs ,
        [string[]]$hypervisors ,
        [switch]$dns ,
        [string]$name ,
        [switch]$tags ,
        [string]$ADgroups ,
        [switch]$noRemoting ,
        [switch]$noOrphans ,
        [switch]$noProgress ,
        [ValidateSet('PVS','MCS','Manual','Any')]
        [string]$provisioningType = 'PVS' ,
        [int]$maxRecordCount = 2000 ,
        [int]$jobTimeout = 120 ,
        [int]$timeout = 60 ,
        [int]$cpuSamples = 2 ,
        [int]$maxThreads = 10 ,
        [string]$splitVM ,
        [string]$pvsShare ,
        [string[]]$modules ## need to import them into runspaces even though they are already loaded
    )
    
    if( $noProgress )
    {
        $ProgressPreference = 'SilentlyContinue'
    }

    Write-Progress -Activity "Caching information" -PercentComplete 0

    ## Get all information from DDCs so we can lookup locally
    [hashtable]$machines = @{}

    ForEach( $ddc in $ddcs )
    {
        $machines.Add( $ddc , [System.Collections.ArrayList] ( Get-BrokerMachine -AdminAddress $ddc -MaxRecordCount $maxRecordCount -ErrorAction SilentlyContinue ) )
    }

    ## Make a hashtable so we can index quicker when cross referencing to DDC & VMware
    [hashtable]$devices = @{}

    [hashtable]$vms = @{}

    if( $hypervisors -and $hypervisors.Count )
    {
        ## Cache all VMs for efficiency
        Get-VM | Where-Object { $_.Name -match $name } | ForEach-Object `
        {
            $vms.Add( $_.Name , $_ )
        }
        Write-Verbose "Got $($vms.Count) vms matching `"$name`" from $($hypervisors -split ' ')"
    }

    [scriptblock]$remoteWork = if( ! $noRemoting )
    {
        {
            $osinfo = Get-WmiObject -Class Win32_OperatingSystem
            $logicalDisks = Get-WmiObject -Class Win32_logicaldisk -Filter 'DriveType = 3'
            $cpu = $(if( $using:miscparameter1 -gt 0 ) { [math]::Round( ( Get-Counter -Counter '\Processor(*)\% Processor Time' -SampleInterval 1 -MaxSamples $using:miscparameter1 |select -ExpandProperty CounterSamples| Where-Object { $_.InstanceName -eq '_total' } | select -ExpandProperty CookedValue  | Measure-Object -Average ).Average , 1 ) }) -as [int]
            $domainMembership = Test-ComputerSecureChannel
            $osinfo,$logicalDisks,$cpu,$domainMembership
        }
    }

    [int]$pvsServerCount = 0

    [hashtable]$adparams = @{ 'Properties' = @( 'Created' , 'LastLogonDate' , 'Description' , 'Modified' )  }
    if( ! [string]::IsNullOrEmpty( $ADgroups ) )
    {
        $adparams[ 'Properties' ] +=  'MemberOf' 
    }

    ## https://blogs.technet.microsoft.com/heyscriptingguy/2015/11/28/beginning-use-of-powershell-runspaces-part-3/
    $SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    ## one downside of runspaces is that we don't get modules passed through
    ForEach( $module in $modules )
    {
        [void]$sessionstate.ImportPSModule( $module )
    }

    [void]$sessionstate.ImportPSModule( 'Citrix.PVS.SnapIn' )

    ## also need to import the functions we need from this module
    @( 'Get-ADMachineInfo' , 'Get-RemoteInfo' ,'Extract-RemoteInfo' ) | ForEach-Object `
    {
        $function = $_
        $Definition = Get-Content Function:\$function -ErrorAction Continue
        $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $function , $Definition
        $sessionState.Commands.Add($SessionStateFunction)
    }

    $RunspacePool = [runspacefactory]::CreateRunspacePool(
        1, #Min Runspaces
        $maxThreads ,
        $sessionstate ,
        $host
    )

    $PowerShell = [PowerShell]::Create()

    $PowerShell.RunspacePool = $RunspacePool
    $RunspacePool.Open()

    $jobs = New-Object System.Collections.ArrayList

    ForEach( $pvsServer in $pvsServers )
    {
        $pvsServerCount++
        Set-PvsConnection -Server $pvsServer 

        if( ! $? )
        {
            Write-Output "Cannot connect to PVS server $pvsServer - aborting"
            continue
        }

        ## Get Device info in one go as quite slow
        [hashtable]$deviceInfos = @{}
        Get-PvsDeviceInfo | ForEach-Object `
        {
            $deviceInfos.Add( $_.DeviceId , $_ )
        }

        ## Cache store locations so we can look up vdisk sizes
        [hashtable]$stores = @{}
        Get-PvsStore | ForEach-Object `
        {
            $stores.Add( $_.StoreName , $_.Path )
        }

        ## Get all devices so we can do progress
        $pvsDevices = @( Get-PvsDevice | Where-Object { $_.Name -match $name })
        [decimal]$eachDevicePercent = 100 / [Math]::Max( $pvsDevices.Count , 1 ) ## avoid divide by zero if no devices found
        [int]$counter = 0
    
        # Cache all disk and version info
        [hashtable]$diskVersions = @{}

        Get-PvsSite | ForEach-Object `
        {
            Get-PvsDiskInfo -SiteId $_.SiteId | ForEach-Object `
            {
                $diskVersions.Add( $_.DiskLocatorId , @( Get-PvsDiskVersion -DiskLocatorId $_.DiskLocatorId ) )
            }
        }

        # Get all sites that we can see on this server and find all devices and cross ref to Citrix for catalogues and delivery groups
        $pvsDevices | ForEach-Object `
        {
            $counter++
            $device = $_
            Write-Verbose "$counter / $($pvsDevices.Count) : $($device.Name)"
            [decimal]$percentComplete = $counter * $eachDevicePercent
            Write-Progress -Activity "Processing $($pvsDevices.Count) devices from PVS server $pvsServer" -Status "$($device.name)" -PercentComplete $percentComplete
            $vm = if( $vms -and $vms.count )
            {
                $vms[ $device.Name ]
            }

            [hashtable]$parameters = @{
                'Device' = $device
                'vm' = $vm
                'NoRemoting' = $noRemoting
                'CPUSamples' = $cpuSamples
                'DiskVersions' = $diskVersions
                'deviceInfos' = $deviceInfos
                'stores' = $stores
                'pvsserver' = $pvsServer
                'adparams' = $adparams
                'ADGroups' = $ADgroups
                'ddcs' = $ddcs
                'machines' = $machines
                'dns' = $dns
                'tags' = $tags
                'jobTimeout' = $jobTimeout
                'RemoteWork' = $remoteWork
            }

            $PowerShell = [PowerShell]::Create() 
            $PowerShell.RunspacePool = $RunspacePool   
            [void]$PowerShell.AddScript({
            Param (
                $device,
                $vm ,
                $noRemoting ,
                $cpuSamples ,
                $diskVersions ,
                $deviceInfos ,
                $stores ,
                $pvsserver ,
                $adparams ,
                $ADgroups ,
                $ddcs ,
                $machines ,
                $dns ,
                $tags ,
                $jobTimeout ,
                $remoteWork
            )
        
            [string[]]$cacheTypes = 
            @(
                'Standard Image' ,
                'Cache on Server', 
                'Standard Image' ,
                'Cache in Device RAM', 
                'Cache on Device Hard Disk', 
                'Standard Image' ,
                'Device RAM Disk', 
                'Cache on Server, Persistent',
                'Standard Image' ,
                'Cache in Device RAM with Overflow on Hard Disk' 
            )

            [string[]]$accessTypes = 
            @(
                'Production', 
                'Maintenance', 
                'Maintenance Highest Version', 
                'Override', 
                'Merge', 
                'MergeMaintenance', 
                'MergeTest'
                'Test'
            )
            
            [int]$bootVersion = -1

            ## Can't easily cache this since needs each device's deviceid
            $vDisk = Get-PvsDiskInfo -DeviceId $device.DeviceId

            [hashtable]$fields = @{}

            if( $vm )
            {
                $error.Clear()
                $fields += @{
                    'CPUs' = $vm.NumCpu 
                    'Memory (GB)' = $vm.MemoryGB
                    'Hard Drives (GB)' = $( ( Get-HardDisk -VM $vm -ErrorAction SilentlyContinue | sort CapacityGB | select -ExpandProperty CapacityGB ) -join ' ' )
                    'NICS' = $( ( Get-NetworkAdapter -VM $vm -ErrorAction SilentlyContinue | Sort Type | Select -ExpandProperty Type ) -join ' ' )
                    'Hypervisor' = $vm.VMHost
                }
                if( $error[0] )
                {
                    Write-Warning "VMware error: $($error[0])"
                }
            }
        
            if( $device.Active -and ! $noRemoting )
            {
                $remoteInfo = Get-RemoteInfo -computer $device.Name -miscparameter1 $cpuSamples -jobTimeout $jobTimeout -work $remoteWork -ErrorVariable RemoteError
                $fields += Extract-RemoteInfo $remoteInfo
            }

            $fields.Add( 'PVS Server' , $pvsServer )
            $versions = $null
            if( $vdisk )
            {
                $fields += @{
                    'Disk Name' = $vdisk.Name
                    'Store Name' = $vdisk.StoreName
                    'Disk Description' = $vdisk.Description
                    'Cache Type' = $cacheTypes[$vdisk.WriteCacheType]
                    'Disk Size (GB)' = ([math]::Round( $vdisk.DiskSize / 1GB , 2 ))
                    'Write Cache Size (MB)' = $vdisk.WriteCacheSize }

                $versions = $diskVersions[ $vdisk.DiskLocatorId ] ## Get-PvsDiskVersion -DiskLocatorId $vdisk.DiskLocatorId ## 
            
                if( $versions )
                {
                    ## Now get latest production version of this vdisk
                    $override = $versions | Where-Object { $_.Access -eq 3 } 
                    $vdiskFile = $null
                    $latestProduction = $versions | Where-Object { $_.Access -eq 0 } | Sort Version -Descending | Select -First 1 
                    if( $latestProduction )
                    {
                        $vdiskFile = $latestProduction.DiskFileName
                        $latestProductionVersion = $latestProduction.Version
                    }
                    else
                    {
                        $latestProductionVersion = $null
                    }
                    if( $override )
                    {
                        $bootVersion = $override.Version
                        $vdiskFile = $override.DiskFileName
                    }
                    else
                    {
                        ## Access: Read-only access of the Disk Version. Values are: 0 (Production), 1 (Maintenance), 2 (MaintenanceHighestVersion), 3 (Override), 4 (Merge), 5 (MergeMaintenance), 6 (MergeTest), and 7 (Test) Min=0, Max=7, Default=0
                        $bootVersion = $latestProductionVersion
                    }
                    if( $vdiskFile)
                    {
                        ## Need to see if Store path is local to the PVS server and if so convert to a share so we can get vdisk file info
                        if( $stores[ $vdisk.StoreName ] -match '^([A-z]):(.*$)' )
                        {
                            if( [string]::IsNullOrEmpty( $pvsShare ) )
                            {
                                $vdiskfile = Join-Path ( Join-Path ( '\\' + $pvsServer + '\' + "$($Matches[1])`$"  ) $Matches[2] ) $vdiskFile ## assume regular admin share
                            }
                            else
                            {
                                $vdiskfile = Join-Path ( Join-Path ( '\\' + $pvsServer + '\' + $pvsShare ) ) $vdiskFile
                            }
                        }
                        else
                        {
                            $vdiskFile = Join-Path $stores[ $vdisk.StoreName ] $vdiskFile
                        }
                        if( ( Test-Path $vdiskFile -ErrorAction SilentlyContinue ) )
                        {
                            $fields += @{ 'vDisk Size (GB)' = [math]::Round( (Get-ItemProperty -Path $vdiskFile).Length / 1GB ) }
                        }
                        else
                        {
                            Write-Warning "Could not find disk `"$vdiskFile`" for $($device.name)"
                        }
                    }
                    if( $latestProductionVersion -eq $null -and $override )
                    {
                        ## No production version, only an override so this must be the latest production version
                        $latestProductionVersion = $override.Version
                    }
                    $fields += @{
                        'Override Version' = $( if( $override ) { $bootVersion } else { $null } ) 
                        'Vdisk Latest Version' = $latestProductionVersion 
                        'Latest Version Description' = $( $versions | Where-Object { $_.Version -eq $latestProductionVersion } | Select -ExpandProperty Description )  
                    }      
                }
                else
                {
                    Write-Output "Failed to get vdisk versions for id $($vdisk.DiskLocatorId) for $($device.Name):$($error[0])"
                }
                $fields.Add( 'Vdisk Production Version' ,$bootVersion )
            }
            else
            {
                Write-Output "Failed to get vdisk for device id $($device.DeviceId) device $($device.Name)"
            }
        
            $deviceInfo = $deviceInfos[ $device.DeviceId ]
            if( $deviceInfo )
            {
                $fields.Add( 'Disk Version Access' , $accessTypes[ $deviceInfo.DiskVersionAccess ] )
                $fields.Add( 'Booted Off' , $deviceInfo.ServerName )
                $fields.Add( 'Device IP' , $deviceInfo.IP )
                if( ! [string]::IsNullOrEmpty( $deviceInfo.Status ) )
                {
                    $fields.Add( 'Retries' , ($deviceInfo.Status -split ',')[0] -as [int] ) ## second value is supposedly RAM cache used percent but I've not seen it set
                }
                if( $device.Active )
                {
                    ## Check if booting off the disk we should be as previous info is what is assigned, not what is necessarily being used (e.g. vdisk changed for device whilst it is booted)
                    $bootedDiskName = (( $diskVersions[ $deviceInfo.DiskLocatorId ] | Select -First 1 | Select -ExpandProperty Name ) -split '\.')[0]
                    $fields.Add( 'Booted Disk Version' , $deviceInfo.DiskVersion )
                    if( $bootVersion -ge 0 )
                    {
                        Write-Verbose "$($device.Name) booted off $bootedDiskName, disk configured $($vDisk.Name)"
                        $fields.Add( 'Booted off latest' , ( $bootVersion -eq $deviceInfo.DiskVersion -and $bootedDiskName -eq $vdisk.Name ) )
                        $fields.Add( 'Booted off vdisk' , $bootedDiskName )
                    }
                }
                if( $versions )
                {
                    try
                    {
                        $fields.Add( 'Disk Version Created' ,( $versions | Where-Object { $_.Version -eq $deviceInfo.DiskVersion } | select -ExpandProperty CreateDate ) )
                    }
                    catch
                    {
                        $_
                    }
                }
            }
            else
            {
                Write-Warning "Failed to get PVS device info for id $($device.DeviceId) device $($device.Name)"
            }
        
            $fields += Get-ADMachineInfo -machineName $device.Name -adparams $adparams -adGroups $ADgroups

            if( $device.Active -and $dns )
            {
                [array]$ipv4Address = @( Resolve-DnsName -Name $device.Name -Type A )
                $fields.Add( 'IPv4 address' , ( ( $ipv4Address | Select -ExpandProperty IPAddress ) -join ' ' ) )
            }
        
            if( $machines -and $machines.Count )
            {
                ## Need to find a ddc that will return us information on this device
                ForEach( $ddc in $ddcs )
                {
                    ## can't use HostedMachineName as only populated if registered
                    $machine = $machines[ $ddc ] | Where-Object { $_.MachineName -eq  ( ($device.DomainName -split '\.')[0] + '\' + $device.Name ) }
                    if( $machine )
                    {
                        $fields += @{
                            'Machine Catalogue' = $machine.CatalogName
                            'Delivery Group' = $machine.DesktopGroupName
                            'Registration State' = $machine.RegistrationState
                            'User_Sessions' = $machine.SessionCount
                            'Load Index' = $machine.LoadIndex
                            'Load Indexes' = $machine.LoadIndexes -join ','
                            'Maintenance_Mode' = $( if( $machine.InMaintenanceMode ) { 'On' } else { 'Off' } )
                            'DDC' = $ddc
                        }
                        if( $tags )
                        {
                            $fields.Add( 'Tags' , ( $machine.Tags -join ',' ) )
                        }
                        break
                    }
                }
            }

            Add-Member -InputObject $device -NotePropertyMembers $fields
            $device ## return
        })
    
        [void]$PowerShell.AddParameters( $Parameters )
        $Handle = $PowerShell.BeginInvoke()
        $temp = '' | Select PowerShell,Handle
        $temp.PowerShell = $PowerShell
        $temp.handle = $Handle
        [void]$jobs.Add($Temp)
        }
    }

    [array]$results = @( $jobs | ForEach-Object `
    {
        $_.powershell.EndInvoke($_.handle)
        $_.PowerShell.Dispose()
    })
    $jobs.clear()

    $results | ForEach-Object `
    {
        if( Get-Member -InputObject $_ -Name 'Name' -ErrorAction SilentlyContinue )
        {
            $devices.Add( $_.Name , ( $_ | Select * ) )
        }
        else
        {
            Write-Warning $_
        }
    }

    ## See if we have any devices from DDC machine list which are marked as being in PVS catalogues but not in our devices list so are orphans
    if( ! $noOrphans )
    {
        $machines.GetEnumerator() | ForEach-Object `
        {
            $ddc = $_.Key
            Write-Progress -Activity "Checking for orphans on DDC $ddc" -PercentComplete 98

            ## Cache machine catalogues so we can check provisioning type
            [hashtable]$catalogues = @{}
            Get-BrokerCatalog -AdminAddress $ddc | ForEach-Object { $catalogues.Add( $_.Name , $_ ) }

            ## Add to devices so we can display as much detail as possible if PVS provisioned
            $_.Value | ForEach-Object `
            {
                $machine = $_
                $domainName,$machineName = $machine.MachineName -split '\\'
                if( [string]::IsNullOrEmpty( $machineName ) )
                {
                    $machineName = $domainName
                    $domainName = $null
                }
                if( [string]::IsNullOrEmpty( $name ) -or $machineName -match $name )
                {
                    ## Now see if have this in devices in which case we ignore it - domain name in device record may be FQDN but domain from catalogue will not be (may also be missing in device)
                    #$device = $devices | Where-Object { $_.Name -eq $machineName -and ( ! $domainName -or ! $_.DomainName -or ( $domainName -eq ( $_.DomainName -split '\.' )[0] ) ) }
                    $device = $devices[ $machineName ]
                    if( $device ) ## check same domain
                    {
                        if( $domainName -and $device.DomainName -and $domainName -ne ( $device.DomainName -split '\.' )[0] )
                        {
                            $device = $null ## doesn't quite match
                        }
                    }
                    if( ! $device )
                    {
                        ## Now check machine catalogues so if ProvisioningType = PVS then we will look to see if it an orphan
                        $catalogue = $catalogues[ $machine.CatalogName  ]
                        if( ! $catalogue -or $provisioningType -eq 'Any' -or $catalogue.ProvisioningType -match $provisioningType )
                        {
                            $newItem = [pscustomobject]@{ 
                                'Name' = ( $machine.MachineName -split '\\' )[-1] 
                                'DomainName' = if( $machine.MachineName.IndexOf( '\' ) -gt 0 )
                                {
                                    ($machine.MachineName -split '\\')[0]
                                }
                                else
                                {
                                    $null
                                }
                                'DDC' = $ddc ;
                                'Machine Catalogue' = $machine.CatalogName
                                'Delivery Group' = $machine.DesktopGroupName
                                'Registration State' = $machine.RegistrationState
                                'Maintenance_Mode' = $( if( $machine.InMaintenanceMode ) { 'On' } else { 'Off' } )
                                'User_Sessions' = $machine.SessionCount ; }

                            [hashtable]$adFields = Get-ADMachineInfo -machineName $newItem.Name -adparams $adparams -adGroups $ADgroups
                            if( $adFields -and $adFields.Count )
                            {
                                Add-Member -InputObject $newItem -NotePropertyMembers $adfields
                            }
                            if( ! $noRemoting )
                            {
                                [hashtable]$fields = Extract-RemoteInfo (Get-RemoteInfo -computer $newItem.Name -miscparameter1 $cpuSamples -jobTimeout $jobTimeout -work $remoteWork)
                                if( $fields -and $fields.Count )
                                {
                                    Add-Member -InputObject $newItem -NotePropertyMembers $fields
                                }
                            }
                            if( $tags )
                            {
                                Add-Member -InputObject $newItem  -MemberType NoteProperty -Name 'Tags' -Value ( $machine.Tags -join ',' )
                            }
                            if( $dns )
                            {
                                [array]$ipv4Address = @( Resolve-DnsName -Name $newItem.Name -Type A )
                                Add-Member -InputObject $newItem  -MemberType NoteProperty -Name 'IPv4 address' -Value ( ( $ipv4Address | Select -ExpandProperty IPAddress ) -join ' ' )
                            }

                            $devices.Add( $newItem.Name , $newItem )
                        }
                    }
                }
            }
        }
        ## if we have VMware details then get those VMs and add if not present here
        if( $hypervisors -and $hypervisors.Count )
        {
            ## will already be connected as have already grabbed VMs
            Write-Progress -Activity "Checking for orphans on hypervisor $($hypervisors -split ' ')" -PercentComplete 99

            [int]$vmCount = 0
            $vms.GetEnumerator() | ForEach-Object `
            {
                $vmwareVM = $_.Value
                $vmCount++
                [string]$vmName = $vmwareVM.Name
                ## name may not be just netbios but may have other infor after a separator character like _
                if( ! [string]::IsNullOrEmpty( $splitVM ) )
                {
                    $vmName =  ($vmName -split $splitVM)[0]
                }
                $existingDevice = $devices[ $vmname ]
                ## Now have to see if we have restricted the PVS device retrieval via -name making $devices a subset of all PVS devices
                if( ! $existingDevice -and ! [string]::IsNullOrEmpty( $name ) )
                {
                    $existingDevice = $vmName -notmatch $name
                }
                if( ! $existingDevice )
                {
                    $newItem = [pscustomobject]@{ 
                        'Name' = $vmName
                        'Description' = $vmwareVM.Notes
                        'CPUs' = $vmwareVM.NumCpu 
                        'Memory (GB)' = $vmwareVM.MemoryGB
                        'Hard Drives (GB)' = $( ( Get-HardDisk -VM $vmwareVM -ErrorAction SilentlyContinue | sort CapacityGB | select -ExpandProperty CapacityGB ) -join ' ' )
                        'NICS' = $( ( Get-NetworkAdapter -VM $vmwareVM -ErrorAction SilentlyContinue | Sort Type | Select -ExpandProperty Type ) -join ' ' )
                        'Hypervisor' = $vmwareVM.VMHost
                        'Active' = $( if($vmwareVM.PowerState -eq 'PoweredOn') { $true } else { $false } )
                    }
                    
                    [hashtable]$adFields = Get-ADMachineInfo -machineName $newItem.Name -adparams $adparams -adGroups $ADgroups
                    if( $adFields -and $adFields.Count )
                    {
                        Add-Member -InputObject $newItem -NotePropertyMembers $adfields
                    }

                    if( $vmwareVM.PowerState -eq 'PoweredOn' )
                    {
                        if( ! $noRemoting )
                        { 
                            [hashtable]$fields = Extract-RemoteInfo (Get-RemoteInfo -computer $newItem.Name -miscparameter1 $cpuSamples -jobTimeout $jobTimeout -work $remoteWork)
                            if( $fields -and $fields.Count )
                            {
                                Add-Member -InputObject $newItem -NotePropertyMembers $fields
                            }
                        }
                        if( $dns )
                        {
                            [array]$ipv4Address = @( Resolve-DnsName -Name $newItem.Name -Type A )
                            Add-Member -InputObject $newItem  -MemberType NoteProperty -Name 'IPv4 address' -Value ( ( $ipv4Address | Select -ExpandProperty IPAddress ) -join ' ' )
                        }
                    }

                    $devices.Add( $newItem.Name , $newItem )
                }
                if( ! $vmCount )
                {
                    Write-Warning "Found no VMs on $($hypervisors -split ',') matching regex `"$name`""
                }
            }
        }
    }
    
    Write-Progress -Activity 'Finished' -Completed -PercentComplete 100

    $devices # return
}

Function Get-ADMachineInfo
{
    Param
    (
        [hashtable]$adparams ,
        [string]$adGroups ,
        [string]$machineName
    )
    
    [hashtable]$fields = @{}

    try
    {
        $adaccount = Get-ADComputer $machineName -ErrorAction SilentlyContinue @adparams
        if( $adaccount )
        {
            [string]$groups = $null
            if( ! [string]::IsNullOrEmpty( $ADgroups ) )
            {
                $groups = ( ( $adAccount | select -ExpandProperty MemberOf | ForEach-Object { (( $_ -split '^CN=')[1] -split '\,')[0] } | Where-Object { $_ -match $ADgroups } ) -join ' ' )
            }
            $fields += `
            @{
                'AD Account Created' = $adAccount.Created
                'AD Last Logon' = $adAccount.LastLogonDate
                'AD Account Modified' = $adaccount.Modified
                'AD Description' = $adAccount.Description
                'AD Groups' = $groups
            }
         } 
    }
    catch {}
    $fields # return
}

Export-ModuleMember -Function Get-RemoteInfo,Get-PVSDevices,Get-ADMachineInfo
