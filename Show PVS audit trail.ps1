#requires -version 3.0
<#
    Show or export PVS audit logs

    Guy Leech, 2018

    Modification history:

    21/05/18  GL  Initial release
#>

<#
.SYNOPSIS

Produce grid view or csv report of Citrix Provisioning Services audit logs

.DESCRIPTION

Logs can be exported using the Export-PvsAuditTrail cmdlet but this only exports in XML format so may not be as easy to read/filter as csv or grid view format.
Auditing must be enabled for the farm which can be done in the Options tab for the farm properties in the PVS console.

.PARAMETER pvsServers

Comma separated list of PVS servers to contact. Do not specify multiple servers if they use the same SQL database

.PARAMETER csv

Path to a csv file that will have the results written to it.

.PARAMETER gridview

Show the results in an on screen grid view.

.PARAMETER startDate

The start date for audit events to be retrieved. If not specified then the Citrix cmdlet defaults to one week prior to the current date/time

.PARAMETER endDate

The end date for audit events to be retrieved. If not specified then the Citrix cmdlet defaults to the current date/time

.NOTES

If the PVS cmdlets are not available where the script is run from, e.g. the PVS console is not installed, then the script will attempt to load the required module remotely from the PVS server(s) specified.

Details on the cmdlets used can be found here - https://docs.citrix.com/content/dam/docs/en-us/provisioning-services/7-15/PvsSnapInCommands.pdf

#>

[CmdletBinding()]

Param
(
    [string[]]$pvsServers = @( 'localhost' ) , 
    [string]$csv ,
    [switch]$gridView ,
    [string]$startDate ,
    [string]$endDate ,
    [string]$pvsModule = "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll"
)

[int]$ERROR_INVALID_PARAMETER = 87
[int]$defaultDaysBack = 7

[string[]]$audittypes = @(
    'Many' , 
    'AuthGroup' , 
    'Collection' , 
    'Device' , 
    'Disk' , 
    'DiskLocator' , 
    'Farm' , 
    'FarmView' , 
    'Server' , 
    'Site' , 
    'SiteView' , 
    'Store' ,
    'System' , 
    'UserGroup'
)

[hashtable]$auditActions = @{
 1 = 'AddAuthGroup'
 2 = 'AddCollection'
 3 = 'AddDevice'
 4 = 'AddDiskLocator'
 5 = 'AddFarmView'
 6 = 'AddServer'
 7 = 'AddSite'
 8 = 'AddSiteView'
 9 = 'AddStore'
 10 = 'AddUserGroup'
 11 = 'AddVirtualHostingPool'
 12 = 'AddUpdateTask'
 13 = 'AddDiskUpdateDevice'
 1001 = 'DeleteAuthGroup'
 1002 = 'DeleteCollection'
 1003 = 'DeleteDevice'
 1004 = 'DeleteDeviceDiskCacheFile'
 1005 = 'DeleteDiskLocator'
 1006 = 'DeleteFarmView'
 1007 = 'DeleteServer'
 1008 = 'DeleteServerStore'
 1009 = 'DeleteSite'
 1010 = 'DeleteSiteView'
 1011 = 'DeleteStore'
 1012 = 'DeleteUserGroup'
 1013 = 'DeleteVirtualHostingPool'
 1014 = 'DeleteUpdateTask'
 1015 = 'DeleteDiskUpdateDevice'
 1016 = 'DeleteDiskVersion'
 2001 = 'RunAddDeviceToDomain'
 2002 = 'RunApplyAutoUpdate'
 2003 = 'RunApplyIncrementalUpdate'
 2004 = 'RunArchiveAuditTrail'
 2005 = 'RunAssignAuthGroup'
 2006 = 'RunAssignDevice'
 2007 = 'RunAssignDiskLocator'
 2008 = 'RunAssignServer'
 2009 = 'RunWithReturnBoot'
 2010 = 'RunCopyPasteDevice'
 2011 = 'RunCopyPasteDisk'
 2012 = 'RunCopyPasteServer'
 2013 = 'RunCreateDirectory'
 2014 = 'RunCreateDiskCancel'
 2015 = 'RunDisableCollection'
 2016 = 'RunDisableDevice'
 2017 = 'RunDisableDeviceDiskLocator'
 2018 = 'RunDisableDiskLocator'
 2019 = 'RunDisableUserGroup'
 2020 = 'RunDisableUserGroupDiskLocator'
 2021 = 'RunWithReturnDisplayMessage'
 2022 = 'RunEnableCollection'
 2023 = 'RunEnableDevice'
 2024 = 'RunEnableDeviceDiskLocator'
 2025 = 'RunEnableDiskLocator'
 2026 = 'RunEnableUserGroup'
 2027 = 'RunEnableUserGroupDiskLocator'
 2028 = 'RunExportOemLicenses'
 2029 = 'RunImportDatabase'
 2030 = 'RunImportDevices'
 2031 = 'RunImportOemLicenses'
 2032 = 'RunMarkDown'
 2033 = 'RunWithReturnReboot'
 2034 = 'RunRemoveAuthGroup'
 2035 = 'RunRemoveDevice'
 2036 = 'RunRemoveDeviceFromDomain'
 2037 = 'RunRemoveDirectory'
 2038 = 'RunRemoveDiskLocator'
 2039 = 'RunResetDeviceForDomain'
 2040 = 'RunResetDatabaseConnection'
 2041 = 'RunRestartStreamingService'
 2042 = 'RunWithReturnShutdown'
 2043 = 'RunStartStreamingService'
 2044 = 'RunStopStreamingService'
 2045 = 'RunUnlockAllDisk'
 2046 = 'RunUnlockDisk'
 2047 = 'RunServerStoreVolumeAccess'
 2048 = 'RunServerStoreVolumeMode'
 2049 = 'RunMergeDisk'
 2050 = 'RunRevertDiskVersion'
 2051 = 'RunPromoteDiskVersion'
 2052 = 'RunCancelDiskMaintenance'
 2053 = 'RunActivateDevice'
 2054 = 'RunAddDiskVersion'
 2055 = 'RunExportDisk'
 2056 = 'RunAssignDisk'
 2057 = 'RunRemoveDisk'
 2058 = 'RunDiskUpdateStart'
 2059 = 'RunDiskUpdateCancel'
 2060 = 'RunSetOverrideVersion'
 2061 = 'RunCancelTask'
 2062 = 'RunClearTask'
 2063 = 'RunForceInventory'
 2064 = 'RunUpdateBDM'
 2065 = 'RunStartDeviceDiskTempVersionMode'
 2066 = 'RunStopDeviceDiskTempVersionMode'
 3001 = 'RunWithReturnCreateDisk'
 3002 = 'RunWithReturnCreateDiskStatus'
 3003 = 'RunWithReturnMapDisk'
 3004 = 'RunWithReturnRebalanceDevices'
 3005 = 'RunWithReturnCreateMaintenanceVersion'
 3006 = 'RunWithReturnImportDisk'
 4001 = 'RunByteArrayInputImportDevices'
 4002 = 'RunByteArrayInputImportOemLicenses'
 5001 = 'RunByteArrayOutputArchiveAuditTrail'
 5002 = 'RunByteArrayOutputExportOemLicenses'
 6001 = 'SetAuthGroup'
 6002 = 'SetCollection'
 6003 = 'SetDevice'
 6004 = 'SetDisk'
 6005 = 'SetDiskLocator'
 6006 = 'SetFarm'
 6007 = 'SetFarmView'
 6008 = 'SetServer'
 6009 = 'SetServerBiosBootstrap'
 6010 = 'SetServerBootstrap'
 6011 = 'SetServerStore'
 6012 = 'SetSite'
 6013 = 'SetSiteView'
 6014 = 'SetStore'
 6015 = 'SetUserGroup'
 6016 = 'SetVirtualHostingPool'
 6017 = 'SetUpdateTask'
 6018 = 'SetDiskUpdateDevice'
 7001 = 'SetListDeviceBootstraps'
 7002 = 'SetListDeviceBootstrapsDelete'
 7003 = 'SetListDeviceBootstrapsAdd'
 7004 = 'SetListDeviceCustomProperty'
 7005 = 'SetListDeviceCustomPropertyDelete'
 7006 = 'SetListDeviceCustomPropertyAdd'
 7007 = 'SetListDeviceDiskPrinters'
 7008 = 'SetListDeviceDiskPrintersDelete'
 7009 = 'SetListDeviceDiskPrintersAdd'
 7010 = 'SetListDevicePersonality'
 7011 = 'SetListDevicePersonalityDelete'
 7012 = 'SetListDevicePersonalityAdd'
 7013 = 'SetListDiskLocatorCustomProperty'
 7014 = 'SetListDiskLocatorCustomPropertyDelete'
 7015 = 'SetListDiskLocatorCustomPropertyAdd'
 7016 = 'SetListServerCustomProperty'
 7017 = 'SetListServerCustomPropertyDelete'
 7018 = 'SetListServerCustomPropertyAdd'
 7019 = 'SetListUserGroupCustomProperty'
 7020 = 'SetListUserGroupCustomPropertyDelete'
 7021 = 'SetListUserGroupCustomPropertyAdd'
}

[hashtable]$auditParams = @{}

if( [string]::IsNullOrEmpty( $csv ) -and ! $gridView )
{
    Write-Warning "Neither -csv nor -gridview specified so there will be no output produced"
}

if( ! [string]::IsNullOrEmpty( $startDate ) )
{
    $auditParams.Add( 'BeginDate' , [datetime]::Parse( $startDate ) )
}

if( ! [string]::IsNullOrEmpty( $endDate ) )
{
    $auditParams.Add( 'EndDate' , [datetime]::Parse( $endDate ) )
    if( ! [string]::IsNullOrEmpty( $startDate ) )
    {
        if( $auditParams[ 'EndDate' ] -lt $auditParams[ 'BeginDate' ] )
        {
            Write-Error "End date $endDate earlier than start date $startDate"
            Exit $ERROR_INVALID_PARAMETER
        }
    }
    elseif( $auditParams[ 'EndDate' ] -lt (Get-Date).AddDays( -$defaultDaysBack ) )
    {
        Write-Error "End date $endDate earlier than default start date $((Get-Date).AddDays( -$defaultDaysBack ))"
        Exit $ERROR_INVALID_PARAMETER
    }
}

if( ! [string]::IsNullOrEmpty( $pvsModule ) )
{
    Import-Module $pvsModule -ErrorAction SilentlyContinue
}

$PVSSession = $null

[hashtable]$sites = @{}
[hashtable]$stores = @{}
[hashtable]$collections = @{}
[bool]$localPVScmdlets =  ( Get-Command -Name Set-PvsConnection -ErrorAction SilentlyContinue ) -ne $null

[array]$auditevents = @( ForEach( $pvsServer in $pvsServers )
{
    ## See if we have cmdlets we need and if not try and get them from the PVS server
    if( ! $localPVScmdlets )
    {
        $PVSSession = New-PSSession -ComputerName $pvsServer
        if( $PVSSession )
        {
            $null = Invoke-Command -Session $PVSSession -ScriptBlock { Import-Module $using:pvsModule }
            $null = Import-PSSession -Session $PVSSession -Module 'Citrix.PVS.SnapIn'
        }
    }
    else
    {
        $PVSSession = $null
    }

    Set-PvsConnection -Server $pvsServer 

    if( ! $? )
    {
        Write-Output "Cannot connect to PVS server $pvsServer - aborting"
        continue
    }

    ## Lookup table for site id to name
    Get-PvsSite | ForEach-Object `
    {
        $sites.Add( $_.SiteId , $_.SiteName )
    }
    Get-PvsCollection | ForEach-Object `
    {
        $collections.Add( $_.CollectionId , $_.CollectionName )
    }
    Get-PvsStore | ForEach-Object `
    {
        $stores.Add( $_.StoreId , $_.StoreName )
    }
    Get-PvsAuditTrail @auditParams | ForEach-Object `
    {
        $auditItem = $_
        [string]$subItem = $null
        if( ! [string]::IsNullOrEmpty( $auditItem.SubId ) ) ## GUID of the Collection or Store of the action
        {
            $subItem = $collections[ $auditItem.SubId ]
            if( [string]::IsNullOrEmpty( $subItem ) )
            {
                $subItem = $stores[ $auditItem.SubId ]
            }
        }
        [string]$parameters = $null
        [string]$properties = $null
        if( $auditItem.Attachments -band 0x4 ) ## parameters
        {
            $parameters = ( Get-PvsAuditActionParameter -AuditActionId $auditItem.AuditActionId | ForEach-Object `
            {
                "$($_.name)=$($_.value) "
            } )
        }
        if( $auditItem.Attachments -band 0x8 ) ## properties
        {
            $properties = ( Get-PvsAuditActionProperty -AuditActionId $auditItem.AuditActionId | ForEach-Object `
            {
                "$($_.name):$($_.OldValue)=>$($_.NewValue) "
            } )
        }
        [PSCustomObject]@{ 
            'Time' = $auditItem.Time
            'PVS Server' = $pvsServer
            'Domain' = $auditItem.Domain
            'User' = $auditItem.UserName
            'Type' = $audittypes[ $auditItem.Type ]
            'Action' = $auditActions[ $auditItem.Action -as [int] ]
            'Object Name' = $auditItem.ObjectName
            'Sub Item' = $subItem
            'Path' = $auditItem.Path
            'Site' = $sites[ $auditItem.SiteId ] 
            'Properties' = $properties
            'Parameters' = $parameters }
    }
    if( $PVSSession )
    {
        $null = Remove-PSSession -Session $PVSSession
        $PVSSession = $null
    }
} ) | Sort Time

[string]$title = "Got $($auditevents.Count) audit events from $($pvsServers -join ' ')"

Write-Verbose $title

if( ! [string]::IsNullOrEmpty( $csv ) )
{
    $auditevents | Export-Csv -Path $csv -NoClobber -NoTypeInformation
}

if( $gridView ) 
{
    $selected = $auditevents | Out-GridView -Title $title -PassThru
    if( $selected -and $selected.Count )
    {
        $selected | Set-Clipboard
    }
}
