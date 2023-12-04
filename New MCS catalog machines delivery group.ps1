 <#
.SYNOPSIS
    Create new MCS machine catalog, machines and delivery group

.DESCRIPTION
    Can also create other items as required such as provisioning and naming schemes

.PARAMETER todo!!

.EXAMPLE
    & '.\New MCS catalog machines delivery group.ps1'  -catalog "MCS Server 2022" -deliveryGroup "MCS Server 2022" -ddc grl-xaddc02 -numberOfMachines 2 -sessionsupport MultiSession -machineNamePattern 'gls22mcsdem##' -sourceVM GLCTXMCSMAST22 -CleanOnBoot

.NOTES
    Modification History:

    2021/12/05 @guyrleech  Script Born
    2023/12/04 @guyrleech  Script finally pushed to GitHub. Parameters and Examples still require documenting.
#>

<#
Copyright © 2023 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]

Param
(
    [string]$catalog ,
    [string]$deliveryGroup ,
    [string]$ddc ,
    [string]$profileName ,
    [string]$machineNamePattern ,
    [int]$numberOfMachines ,
    [string]$sourceVM ,
    [string]$sourceSnapshot ,
    [string]$zone ,
    [switch]$CleanOnBoot ,
    [string[]]$hypervisorAddress ,
    [ValidateSet('XenServer', 'SCVMM', 'VCenter', 'Custom', 'AWS', 'CloudPlatform', 'WakeOnLAN')]
    [string]$hypervisorType = 'vcenter' ,
    [PSCredential]$hypervisorCredentials ,
    [ValidateSet('Numeric','Alphabetic')]
    [string]$namingSchemeType = 'numeric' ,
    [ValidateSet('Random', 'Static', 'Permanent')]
    [string]$alllocationType = 'Random' ,
    [ValidateSet('MultiSession', 'SingleSession')]
    [string]$sessionsupport ,
    [ValidateSet('Discard', 'OnLocal' , 'OnPvd')]
    [string]$PersistUserChanges = 'Discard' ,
    [ValidateSet('Private', 'Shared')]
    [string]$desktopKind = 'Shared' ,
    [ValidateSet('DesktopsOnly', 'AppsOnly', 'DesktopsAndApps')]
    [string]$deliveryType = 'DesktopsAndApps' ,
    [ValidateSet('L5', 'L7', 'L7_6', 'L7_7', 'L7_8', 'L7_9', 'L7_20', 'L7_25')]
    [string]$MinimumFunctionalLevel = 'L7_25' ,
    [string]$TenantId ,
    [string]$scope ,
    [switch]$AllowReconnectInMaintenanceMode ,
    [switch]$AppProtectionKeyLoggingRequired ,
    [switch]$AppProtectionScreenCaptureRequired ,
    [switch]$AutoscalingEnabled ,
    [ValidateSet('FourBit', 'EightBit', 'SixteenBit', 'TwentyFourBit')]
    [Alias('colourDepth')]
    [string]$colorDepth = 'TwentyFourBit' ,
    [switch]$SecureIcaRequired ,
    [string]$desktopName ,
    [string]$tagName ,
    [string]$tagDescription ,
    [string]$restrictDesktopToTag ,
    [switch]$disabled ,
    [string[]]$allowedAccounts ,
    [string[]]$allowedDesktopAccounts ,
    [switch]$maintenanceMode ,
    [int]$maxDesktopsPerUser = 1 ,
    [int]$WriteBackCacheDiskSizeGB = 40 ,
    [int]$WriteBackCacheMemorySizeMB = 256 ,
    [string]$functionalLevel ,
    [string]$OU ,
    [string]$description = "Created by script by $env:username from $env:computerName" ,
    [string]$domain = $env:USERDNSDOMAIN ,
    [int]$startCount = 1 ,
    [string]$hostingConnection ,
    [string[]]$hostingNetwork ,
    [string[]]$hostingStorage ,
    [string[]]$hostingTemporaryStorage ,
    [string]$hostingUnit ,
    [switch]$resolve ,
    [string]$saveToJson ,
    [switch]$overwriteJson ,
    [string]$readFromJson
)

if( -Not [string]::IsNullOrEmpty( $saveToJson ) )
{
    if( ( Test-Path -Path $saveToJson -ErrorAction SilentlyContinue ) -and -Not $overwriteJson )
    {
        Throw "JSON file `"$saveToJson`" already exists - use -overwriteJson to overwrite"
    }
    $jsonobject = New-Object -TypeName psobject
    $PSBoundParameters.GetEnumerator() | ForEach-Object `
    {
        $value = $_.value
        if( $_.Value.GetType().Name -eq 'PSCredential' )
        {
            ## by default does not output securestring so we do that
            $value = @{
                Username = $_.Value.Username
                Password = $_.value.Password | ConvertFrom-SecureString
            }
        }
        Add-Member -InputObject $jsonobject -MemberType NoteProperty -Name $_.Key -Value $value
    }
    $jsonobject | ConvertTo-Json | Out-File -FilePath $saveToJson
}

if( -Not [string]::IsNullOrEmpty( $readFromJson ) )
{
    if( -Not ( Test-Path -Path $readFromJson -ErrorAction SilentlyContinue ) )
    {
        Throw "JSON file `"$readFromJson`" not found"
    }
    if( $jsonobject = ( Get-Content -Path $readFromJson | ConvertFrom-Json ) )
    {
        ForEach( $property in $jsonobject.PSObject.Properties )
        {
            if( $PSBoundParameters[ $property.Name ] )
            {
                Write-Warning -Message "-$($property.Name) present on command line so ignoring json value `"$($property.Value)`""
            }
            else
            {
                $value = $property.value
                
                ## need to deal with [switch] types as will be custom objects
                if( $property.Value.GetType().Name -eq 'PSCustomObject' )
                {
                    if( $property.Value.PSObject.Properties[ 'IsPresent' ] -and $property.Value.IsPresent -is [bool] )
                    {
                        $value = $property.Value.IsPresent
                    }
                    elseif( $property.Name -match 'credentials$' -and $property.Value.PSobject.Properties[ 'username' ] -and $property.Value.PSobject.Properties[ 'password' ] )
                    {
                        $value = New-Object System.Management.Automation.PSCredential( $property.Value.username , ( ConvertTo-SecureString -String $property.Value.password ))
                    }
                    else
                    {
                        Throw "Don't know how to process json file item `"$($property.Name)`""
                    }
                }

                try
                {
                    Write-Verbose -Message "Setting `"$($property.Name)`" to `"$Value`""
                    Set-Variable -Name $property.Name -Value $value
                }
                catch
                {
                    Throw "Fatal error with json `"$($property.Name)`" `"$value`" : $_"
                }
            }
        }
    }
    if( -Not $jsonobject )
    {
        Throw "Failed to parse json from `"$readFromJson`""
    }
}

if( -Not [string]::IsNullOrEmpty( $OU ) )
{
    ## see if canonical name (e.g. copied from AD Users & computers) and convert to distinguished
    if( $OU.IndexOf('/') -gt 0 )
    {
        [string]$translatedOU = $null
        try
        {
            ## http://www.itadmintools.com/2011/09/translate-active-directory-name-formats.html
            if( $NameTranslate = New-Object -ComObject NameTranslate )
            {
                [System.__ComObject].InvokeMember( 'Init' , 'InvokeMethod', $null , $NameTranslate , @( 3 , $null ) ) ## ADS_NAME_INITTYPE_GC
                [System.__ComObject].InvokeMember( 'Set' , 'InvokeMethod', $null , $NameTranslate , (2 ,$OU)) ## CANONICALNAME
                $translatedOU = [System.__ComObject].InvokeMember('Get' , 'InvokeMethod' , $null , $NameTranslate , 1) ## DISTINGUISHEDNAME
            }
            else
            {
                Write-Warning -Message "Failed to create ComObject NameTranslate"
            }
        }
        catch
        {
            Throw "Failed to translate OU `"$OU`" from canonical name to distinguished`n$_"
        }
        Write-Verbose -Message "Translated OU `"$OU`" to `"$translatedOU`""
        $OU = $translatedOU
    }

    [bool]$ouexists = $false
    try
    {
        $ouexists = [adsi]::Exists( "LDAP://$OU" )
    }
    catch
    {
        Throw "Badly formatted OU $OU : $_"
    }
    if( -Not $ouexists )
    {
        Throw "OU `"$OU`" not found"
    }
}

[hashtable]$citrixParameters = @{}
$secureclient = $null

[string[]]$citrixModules = @( 'Citrix.Broker.*' , 'Citrix.Host.*' )
[bool]$isCloud = $false

if( ! [string]::IsNullOrEmpty( $ddc ) )
{
    $citrixParameters.Add( 'AdminAddress' , $ddc )
}

if( ! [string]::IsNullOrEmpty( $profileName ) )
{
    $isCloud = $true
}

ForEach( $citrixModule in $citrixModules )
{
    if( ! (  Import-Module -Name $citrixModule -ErrorAction SilentlyContinue -PassThru -Verbose:$false) `
        -and ! ( Add-PSSnapin -Name $citrixModule -ErrorAction SilentlyContinue -PassThru -Verbose:$false) )
    {
        ## No point doing implicit remoting since wouldn't get the XDHyp: PSDrive which we need
       Throw "Failed to load Citrix PowerShell cmdlets from $citrixModule - is this a Delivery Controller or have Studio or the PowerShell SDK installed ?"
    }
}

$thisZoneUid = $null

if( $isCloud )
{
    $authenticationError = $null
    Get-XDAuthentication -ProfileName $profileName -ErrorAction Stop -ErrorVariable authenticationError
   
    if( $null -ne $authenticationError -and $authenticationError.Count )
    {
        Throw "Failed to authenticate with profile `"$profileName`" : $authenticationError"
    }
    if( -not [string]::IsNullOrEmpty( $zone ) )
    {
        if( $null -eq ( $thisZone = Get-ConfigZone @citrixParameters | Where-Object Name -match $zone ) )
        {
            Throw "Unable to find zone matching $zone"
        }
        elseif( $thisZone -is [array] -and $thisZone.Count -ne 1 )
        {
            Throw "Found $($thisZone.Count) zones ($($thisZone.Name -join ',')) matching $zone"
        }
        else
        {
            $thisZoneUid = $thisZone.Uid
        }
    }
}

[array]$hypervisorContents = @()
$tag = $null

if( -Not [string]::IsNullOrEmpty( $tagName ) )
{
    [hashtable]$tagParameters = @{ 'Name' = $tagName }
    if( $null -ne $tagDescription )
    {
        $tagParameters.Add( 'Description' , $tagDescription )
    }
    $tagParameters += $citrixParameters

    if( -Not ( $tag = Get-BrokerTag -Name $tagName -ErrorAction SilentlyContinue @citrixParameters ) )
    {
        if( ! ( $tag = New-BrokerTag @tagParameters ) )
        {
            Write-Warning -Message "Failed to create tag `"$tagName`""
        }
    }
    elseif( $null -ne $tagDescription -and $tagDescription -ne $tag.Description )
    {
        Set-BrokerTag @tagParameters ## update tag description if changed
    }
}

if( -Not [string]::IsNullOrEmpty( $hostingConnection ) )
{
    if( -Not ( $existingHostingConnection = Get-BrokerHypervisorConnection @citrixParameters -Name $hostingConnection -ErrorAction SilentlyContinue ) )
    {
        [hashtable]$newitemParameters = $citrixParameters.Clone()
        if( -not [string]::IsNullOrEmpty( $thisZoneUid ) )
        {
            $newitemParameters.Add( 'ZoneUid' , $thisZoneUid )
        }
        Set-HypAdminConnection @citrixParameters
        if( $null -eq $hypervisorCredentials )
        {
            Throw "Hosting connection `"$hostingConnection`" not found and no credentials specified to create a new one"
        }
        elseif( $null -eq $hypervisorAddress )
        {
            Throw "Hosting connection `"$hostingConnection`" not found and address specified to create a new one"
        }
        elseif(  $PsCmdlet.ShouldProcess( "Hosting connection `"$hostingConnection`"" , "Create" ) )
        {
            $existingHostingConnection = $null
            $hostingConnectionError = $null
            try
            {
                $existingHostingConnection = New-Item -ErrorVariable hostingConnectionError -Verbose:$false -connectiontype $hypervisorType -hypervisoraddress $hypervisorAddress -Path @( "XDHyp:\Connections\$hostingConnection") -Scope @() -Password $hypervisorCredentials.GetNetworkCredential().password -username $hypervisorCredentials.UserName -persist @newitemParameters
            }
            catch
            {
            }

            if( -Not $existingHostingConnection )
            {
                [string]$message = "Failed to create new $hypervisorType hosting connection to $hypervisorAddress"
                if( [string]::IsNullOrEmpty( $zone ) )
                {
                    $message = "$message. Do you need to specify a zone ?"
                }
                Throw "$message"
            }
            else
            {
                Write-Verbose -Message "Hosting connection created ok"
                [string]$hostingConnectionPath = $existingHostingConnection.FullPath

                if( $resolve )
                {
                    Write-Verbose -Message "$(Get-Date -Format G): Getting storage and network via new hosting connection ..."
                    [datetime]$startTime = [datetime]::Now
                    $hypervisorContents = @( Get-ChildItem -Path $existingHostingConnection.FullPath -Recurse -Force -ErrorAction SilentlyContinue -Verbose:$false | Where-Object { $_.ObjectTypeName -match '^(network|storage)$' } )
                    [datetime]$endTime = [datetime]::Now
                    Write-Verbose -Message "$(Get-Date -Format G): Got $($hypervisorContents.Count) items from $($existingHostingConnection.FullPath) in $(($endTime - $startTime).TotalSeconds) seconds"
                    $matchingNetworks = New-Object -TypeName System.Collections.Generic.List[string]
                    $matchingStorage  = New-Object -TypeName System.Collections.Generic.List[string]
                    $matchingTemporaryStorage  = New-Object -TypeName System.Collections.Generic.List[string]
                    [int]$networkCount = 0
                    [int]$storageCount = 0
                    [hashtable]$rootPaths = @{}

                    ForEach( $item in $hypervisorContents )
                    {
                        $itemAdded = $null
                        if( $item.ObjectTypeName -eq 'network' )
                        {
                            $networkCount++
                            ForEach( $network in $hostingNetwork )
                            {
                                if( $item.FullPath -match "$network\.network$" )
                                {
                                    $matchingNetworks.Add( ( $itemAdded = $item.FullPath ) )
                                }
                            }               
                        }
                        elseif( $item.ObjectTypeName -eq 'storage' )
                        {
                            $storageCount++
                            ForEach( $storage in $hostingStorage )
                            {
                                if( $item.FullPath -match "$storage\.storage$" )
                                {
                                    $matchingStorage.Add( ( $itemAdded = $item.FullPath ) )
                                }
                            }
                            ForEach( $storage in $hostingTemporaryStorage )
                            {
                                if( $item.FullPath -match "$storage\.storage$" )
                                {
                                    $matchingTemporaryStorage.Add( ( $itemAdded = $item.FullPath ) )
                                }
                            }
                        }
                        ## need to get the datacenter and cluster for the rootpath
                        if( $itemAdded )
                        {
                            try
                            {
                                ## isolate datacenter and cluster from path "XDHyp:\Connections\Wakefield vCenter\Wakefield.datacenter\Garage.cluster\Local Extra SSD.storage" 
                                if( $null -ne ( [string[]]$split = ( $itemAdded -split '\\') ) -and $split.Count -ge 5 )
                                {
                                    $rootPaths.Add( ( $split[0..4] -join '\' ) , $true )
                                }
                            }
                            catch
                            {
                                ## already got it which is fine
                            }
                        }

                    }
                    if( $rootPaths.Count -ne 1 )
                    {
                        Write-Warning -Message "Got $($rootPaths.Count) root paths for resources - only expected one"
                    }
                    [hashtable]$hostingUnitParameters = @{
                        Verbose = $false
                        HypervisorConnectionName = $hostingConnection
                        Path = ( Join-Path -Path 'XDHyp:\HostingUnits' -ChildPath $hostingUnit)
                        RootPath = $rootPaths.GetEnumerator() | Select-Object -ExpandProperty Key
                    }
                    if( $matchingNetworks.Count -eq 0 )
                    {
                        Write-Warning -Message "Found no networks matching $hostingNetwork out of $networkCount networks"
                    }
                    else
                    {
                        $hostingUnitParameters.Add( 'NetworkPath' , $matchingNetworks )
                    }
                
                    if( $matchingStorage.Count -eq 0 )
                    {
                        Write-Warning -Message "Found no storage matching $hostingStorage out of $storageCount networks"
                    }
                    else
                    {
                        $hostingUnitParameters.Add( 'StoragePath' , $matchingStorage )
                    }

                    if( $matchingTemporaryStorage.Count -eq 0 )
                    {
                        Write-Warning -Message "Found no storage matching $hostingTemporaryStorage out of $storageCount networks"
                    }
                }
                else ## not resolving entities
                {           
                    $matchingNetworks = $hostingNetwork
                    $matchingStorage  = $hostingStorage
                    $matchingTemporaryStorage  = $hostingTemporaryStorage
                    ## TODO need to set rootpath
                }

                $jobgroup = $null

                ForEach( $temporaryStorage in $matchingTemporaryStorage )
                {
                    if( -Not $jobgroup )
                    {
                        $jobgroup = [Guid]::NewGuid()
                        $hostingUnitParameters.Add( 'JobGroup' , $jobgroup )
                    }
                    New-HypStorage -StoragePath $temporaryStorage -StorageType 'TemporaryStorage' -JobGroup $jobgroup -Verbose:$false
                }
                ## PersonalvDiskStoragePath deprecated so no point implementing
                if( -Not ( $hostingUnit = New-Item @hostingUnitParameters ) )
                {
                    ## TODO delete the hosting connection as incomplete
                }
                else
                {
                    ## add hosting unit to broker service
                    if( -Not ( New-BrokerHypervisorConnection -HypHypervisorConnectionUid $existingHostingConnection.HypervisorConnectionUid -Verbose:$false ) )
                    {
                        Write-Warning -Message "Failed to creaate hypervisor connection to hosting connection"
                    }
                }
            }
        }
    }
    else
    {
        Write-Verbose -Message "Hosting connection `"$hostingConnection`" already exists"
        $existingHostingConnection | Set-BrokerHypervisorConnection
    }
}
else ## not specified a hosting connection
{
    $existingHostingConnection = $null
    $existingHostingConnection = Get-BrokerHypervisorConnection @citrixParameters
    if( -Not $existingHostingConnection )
    {
        Throw "No hosting connections found - if you want to create one specify -hostingConnection"
    }
    if( $existingHostingConnection -is [array] -and $existingHostingConnection.Count -ne 1 )
    {
        ## TODO find existing hosting connections in use and if only one, use that
        Throw "Got $($existingHostingConnection.Count) existing hosting connections - please specify one via -hostingConnection ($(($existingHostingConnection | Select-Object -ExpandProperty Name) -join ','))"
    }
    $existingHostingConnection | Set-BrokerHypervisorConnection
}

$baseVM = $null
$snapshot = $null
$existingCatalogue = $null

if( $catalog -match '[\/;:#.*?=<>|\[\]()"'']' )
{
    Throw "Illegal characters in catalog name $catalog"
}

## TODO if no delivery group specified but want to add toi delivery group, see if all catalog machines are in a single delivery group and use that, if not throw an error

## TODO get prov scheme and identity pool via uid not name

if( -Not [string]::IsNullOrEmpty( $catalog ) )
{
    $existingCatalogue = $null
    $existingCatalogue = Get-BrokerCatalog -Name $catalog @citrixParameters -ErrorAction SilentlyContinue
    if( -Not $existingCatalogue )
    {
        if( [string]::IsNullOrEmpty( $sessionsupport ) )
        {
            Throw "Must specify session support via -sessionsupport"
        }
        if( $PsCmdlet.ShouldProcess( "Machine catalog `"$catalog`"" , "Create" ) )
        {
            Write-Verbose -Message "Catalogue `"$catalog`" not found so will create"
 
            [string]$identityPoolName = $catalog -replace '[\/;:#.*?=<>|\[\]()"'']'
            if( $existingIdentityPool = Get-AcctIdentityPool @citrixParameters -IdentityPoolName $identityPoolName )
            {
                Write-Verbose -Message "Already have identity pool `"$($existingIdentityPool.IdentityPoolName)`" with naming scheme $($existingIdentityPool.NamingScheme)"
                if( $null -ne $machineNamePattern -and $existingIdentityPool.NamingScheme -ne $machineNamePattern )
                {
                    Throw "Existing machine naming pattern $($existingIdentityPool.NamingScheme) is different to requested $machineNamePattern"
                }
            }
            elseif( [string]::IsNullOrEmpty( $machineNamePattern ) )
            {
                Write-Warning -Message "Must specify a machine name pattern via -machineNamePattern with # for numbers"
            }
            elseif( $machineNamePattern.IndexOf( '#' ) -lt 0  )
            {
                Write-Warning -Message "No # characters specified in machine name pattern"
            }
            elseif( $machineNamePattern.Length -gt 15  )
            {
                Write-Warning -Message "Machine name pattern is too long - 15 characters maximum"
            }
            else
            {
                ## see if pattern name already in a pool since we will need to use that
                if( $existingIdentityPoolForName = Get-AcctIdentityPool @citrixParameters | Where-Object { $_.NamingScheme -eq $machineNamePattern -and $_.NamingSchemeType -eq $namingSchemeType } )
                {
                    Write-Warning -Message "Already have identity pool `"$($existingIdentityPoolForName.IdentityPoolName)`" for naming pattern $($existingIdentityPoolForName.NamingScheme)"
                }

                [hashtable]$identityPoolParameters = $citrixParameters.Clone()
                $identityPoolParameters += @{
                    'identityPoolName' = $identityPoolName
                    'domain' = $domain
                    'startCount' = $startCount
                    'NamingSchemeType' = $namingSchemeType
                    'NamingScheme' = $machineNamePattern
                }
                if( $thisZoneUid )
                {
                    $identityPoolParameters.Add( 'ZoneUid' , $thisZoneUid )
                }
                if( $OU )
                {
                    $identityPoolParameters.Add( 'OU' , $OU )
                }
                if( -Not ( $existingIdentityPool = New-AcctIdentityPool @identityPoolParameters ) )
                {
                    Write-Warning -Message "Problem creating identity pool"
                }
            }

            [string]$provisioningSchemeName = $catalog -replace  '[\/;:#.*?=<>|\[\]()"'']'
            if( $existingIdentityPool )
            {
                if( $null -eq ( $existingProvisioningScheme = Get-ProvScheme -ProvisioningSchemeName $provisioningSchemeName @citrixParameters -ErrorAction SilentlyContinue ))
                {
                    ## find the hosting unit as need it to create provisioning scheme
                    ## seems to vary which *name property it resides in so check multiple
                    [string]$HypervisorConnectionName = $existingHostingConnection | Select-Object Name,HypervisorConnectionName,PSChildName -ErrorAction SilentlyContinue|ForEach-Object { $_.PSObject.Properties.Value } | Sort-Object -Unique
                    if( [string]::IsNullOrEmpty( $HypervisorConnectionName ))
                    {
                        Write-Warning -Message "Unable to get name from hosting connection to retrieve its hosting connection"
                    }
                    else
                    {
                        Write-Verbose -Message "Searching XDHyp:\HostingUnits for hypervisor connection `"$HypervisorConnectionName`""
                    }
                    $hostingUnitPath = $null
                    $hostingUnitPath = Get-ChildItem -Path XDHyp:\HostingUnits -Verbose:$false | Where-Object { $_.HypervisorConnection.HypervisorConnectionName -eq $HypervisorConnectionName }
                    if( $null -eq $hostingUnitPath )
                    {
                        Throw "Failed to find hosting unit for hypervisor connection `"$HypervisorConnectionName`""
                    }
                    if( $hostingUnitPath -is [array] -and $hostingUnitPath.Count -ne 1 )
                    {
                        Throw "Failed to find single hosting unit matching $escapedHostingPath, found $($hostingUnitPath.Count)"
                    }
                    Write-Verbose -Message "$(Get-Date -Format G): Getting storage and network via existing hosting unit `"$($hostingUnitPath.HostingUnitName)`" $($hostingUnitPath.RootPath) ..."
                    [datetime]$startTime = [datetime]::Now
                    ## don't filter by type as we will need VMs and snapshots later
                    $hypervisorContents = @( Get-ChildItem -Path ($hostingUnitPath.PSPath -split '::')[-1] -Recurse -Force -ErrorAction SilentlyContinue -Verbose:$false )
                    [datetime]$endTime = [datetime]::Now
                    Write-Verbose -Message "$(Get-Date -Format G): Got $($hypervisorContents.Count) items from $($existingHostingConnection.Name) in $(($endTime - $startTime).TotalSeconds) seconds"

                    if( -Not [string]::IsNullOrEmpty( $sourceVM ) )
                    {
                        if( $null -eq ( $baseVM = $hypervisorContents.Where( { $_.ObjectTypeName -eq 'vm' -and $_.Name -eq $sourceVM } ) ) -or $baseVM.Count -eq 0 )
                        {
                            Throw "Unable to find VM $sourceVM"
                        }

                        if( [string]::IsNullOrEmpty( $sourceSnapshot ) )
                        {
                            if( $null -eq ( $snapshot = $hypervisorContents | Where-Object { $_.ObjectTypeName -eq 'snapshot' -and $_.PSParentPath.StartsWith( $baseVM.PSPath ) } | Select-Object -Last 1) )
                            {
                                Throw "Unable to find any snapshots in VM $($baseVM.Name)"
                            }
                        }
                        elseif( $null -eq ( $snapshot = $hypervisorContents | Where-Object { $_.ObjectTypeName -eq 'snapshot' -and $_.PSParentPath.StartsWith( $baseVM.PSPath ) -and $_.Name -match $sourceSnapshot } ) )
                        {
                            Throw "Unable to find snapshot $sourceSnapshot"
                        }
                        elseif( $snapshot -is [array] )
                        {
                            Throw "Found $($snapshot.Count) snapshots ($(($snapshot | Select-Object -ExpandProperty Name) -join ',') matching $sourceSnapshot"
                        }
                    }
                    else
                    {
                        Throw "Must specify a source VM"
                    }
                    [hashtable]$provisioningSchemeParameters = $citrixParameters.Clone()
                    $provisioningSchemeParameters += @{
                        'RunAsynchronously' = $true ## so we can output progress
                        'ProvisioningSchemeName' =  $provisioningSchemeName
                        'HostingUnitName' = $hostingunitpath.HostingUnitName
                        'IdentityPoolName' = $existingIdentityPool.IdentityPoolName
                        'MasterImageVM' = $snapshot.FullPath
                        'CleanOnBoot' = $CleanOnBoot
                    }

                    if( $WriteBackCacheDiskSizeGB -gt 0 )
                    {
                        $provisioningSchemeParameters += @{
                            'UseWriteBackCache' = $true
                            'WriteBackCacheDiskSize' = $WriteBackCacheDiskSizeGB
                        }
                        if( $WriteBackCacheMemorySizeMB -gt 0 )
                        {
                            $provisioningSchemeParameters.Add( 'WriteBackCacheMemorySize' , $WriteBackCacheMemorySizeMB )
                        }
                    }
                    if( -Not [string]::IsNullOrEmpty( $functionalLevel ) )
                    {
                        $provisioningSchemeParameters.Add( 'FunctionalLevel' , $functionalLevel )
                    }
                
                    if( $WriteBackCacheDiskSizeGB -gt 0 -and -Not $CleanOnBoot )
                    {
                        ## otherwise will get an error when creating the provisioning scheme
                        Throw "Cannot use writeback cache with clean on boot disabled"
                    }

                    try
                    {
                        if( ! ( $newprovisioningSchemeTaskId = New-ProvScheme @provisioningSchemeParameters ))
                        {
                            Write-Warning -Message "Error creating provisioning scheme `"$($provisioningSchemeParameters.ProvisioningSchemeName)`""
                        }
                        else
                        {
                            [string]$activity = "Creating provisioning scheme `"$($provisioningSchemeParameters.ProvisioningSchemeName)`""
                            [int]$lastProgress = -1
                            Do
                            {
                                if( $provisioningTask = Get-ProvTask -TaskId $newprovisioningSchemeTaskId )
                                {
                                    if( $provisioningTask.PSObject.Properties[ 'TaskProgress' ] -and $null -ne $provisioningTask.TaskProgress )
                                    {
                                        if( $provisioningTask.TaskProgress -ne $lastProgress )
                                        {
                                            Write-Progress -Activity $activity -Status "$($provisioningTask.TaskProgress)% complete" -PercentComplete $provisioningTask.TaskProgress
                                            $lastProgress = $provisioningTask.TaskProgress
                                        }
                                    }
                                    if( $provisioningTask.Active )
                                    {
                                        Start-Sleep -Milliseconds 5000
                                    }
                                }
                            } While( $provisioningTask -and $provisioningTask.Active )

                            Write-Progress -Completed -Activity $activity

                            if( ! $provisioningTask -or $provisioningTask.TaskState -match 'Error' -or $provisioningTask.TerminatingError )
                            {
                                Write-Warning -Message "Error $($provisioningTask.TerminatingError) creating provisioning scheme `"$($provisioningSchemeParameters.ProvisioningSchemeName)`" at $(Get-Date -Date $provisioningTask.DateFinished -Format G), master image $($provisioningTask.MasterImage)"
                            }
                        
                            $provisioningSchemeId = $provisioningTask.ProvisioningSchemeUid.Guid
                            Add-ProvSchemeControllerAddress -ProvisioningSchemeName $provisioningTask.ProvisioningSchemeName @citrixParameters -ControllerAddress $(if( [string]::IsNullOrEmpty( $ddc ) -or $ddc -eq 'localhost' ) { ([System.Net.Dns]::GetHostByName($env:computerName)).hostname } else { $ddc })
                            $existingProvisioningScheme = Get-ProvScheme -ProvisioningSchemeUid $provisioningSchemeId
                            if( $provisioningTask )
                            {
                                $null = $provisioningTask | Remove-ProvTask
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Error creating provisioning scheme `"$($provisioningSchemeParameters.ProvisioningSchemeName)`" - $_"
                    }
                }
                else
                {
                    Write-Verbose -Message "Provisioning scheme already exists"
                }
            }
            else
            {
                Write-Warning -Message "Unable to create provisioning scheme as have no identity pool"
            }

            [hashtable]$catalogParameters = $citrixParameters.Clone()
            if( -Not [string]::IsNullOrEmpty( $TenantId ) )
            {
                $catalogParameters.Add( 'TenantId' , $TenantId )
            }
            if( -Not [string]::IsNullOrEmpty( $scope ) )
            {
                $catalogParameters.Add( 'Scope' , $Scope )
            }

            $newCatalogue = $null
            $newCatalogue = New-BrokerCatalog @catalogParameters -AllocationType $alllocationType -Description $description -ProvisioningType MCS -ProvisioningSchemeId $existingProvisioningScheme.ProvisioningSchemeUid -Name $catalog -PersistUserChanges $PersistUserChanges -sessionsupport $sessionsupport -MinimumFunctionalLevel $MinimumFunctionalLevel
            if( -Not $newCatalogue )
            {
                Throw "Failed to create machine catalogue `"$catalog`""
            }
            else
            {
                if( $tag )
                {
                    Add-BrokerTag -InputObject $tag -Catalog $newCatalogue @citrixParameters
                }
            }
            $existingCatalogue = $newCatalogue
            Write-Verbose -Message "Created catalog `"$catalog`" ok"
        }
    }
    else
    {
        Write-Verbose -Message "Catalogue `"$catalog`" already exists"
        if( -Not ( $existingIdentityPool = Get-AcctIdentityPool @citrixParameters -IdentityPoolName $catalog ) )
        {
            Write-Warning -Message "Found no identity pool for catalog $catalog"
        }
        if( -Not ( $existingProvisioningScheme = Get-ProvScheme -ProvisioningSchemeName $catalog @citrixParameters ) )
        {
            Write-Warning -Message "Found no provisioning scheme for catalog $catalog"
        }
    }
}

$newMachines = New-Object -TypeName System.Collections.Generic.List[object]

if( $numberOfMachines -gt 0 )
{
    if( $PsCmdlet.ShouldProcess( "Machine catalog `"$catalog`"" , "Create $numberOfMachines machines" ) )
    {
        ## Get AD Accounts for new machines
        [array]$ADaccounts = @( New-AcctADAccount @citrixParameters -Count $numberOfMachines -IdentityPoolName $existingIdentityPool.IdentityPoolName)
        if( $null -eq $ADaccounts -or $ADaccounts.SuccessfulAccountsCount -ne $numberOfMachines )
        {
            Write-Warning -Message "Failed to create $numberOfMachines AD computer accounts"
        }
        elseif( ! ( $newprovisioningTaskId = New-ProvVM -ADAccountName $ADaccounts.SuccessfulAccounts -ProvisioningSchemeUid $existingProvisioningScheme.ProvisioningSchemeUid -RunAsynchronously ))
        {
            Write-Warning -Message "Error provisioning machines"
        }
        else
        {
            [string]$activity = "Provisioning $numberOfMachines VMs"
            [int]$lastProgress = -1
            Do
            {
                if( $provisioningTask = Get-ProvTask -TaskId $newprovisioningTaskId )
                {
                    if( $provisioningTask.PSObject.Properties[ 'TaskProgress' ] -and $null -ne $provisioningTask.TaskProgress )
                    {
                        if( $provisioningTask.TaskProgress -ne $lastProgress )
                        {
                            Write-Progress -Activity $activity -Status "$($provisioningTask.TaskProgress)% complete" -PercentComplete $provisioningTask.TaskProgress
                            $lastProgress = $provisioningTask.TaskProgress
                        }
                    }
                    if( $provisioningTask.Active )
                    {
                        Start-Sleep -Milliseconds 2000
                    }
                }
            } While( $provisioningTask -and $provisioningTask.Active )

            Write-Progress -Completed -Activity $activity

            if( ! $provisioningTask -or $provisioningTask.TaskState -match 'Error' -or $provisioningTask.TerminatingError )
            {
                Write-Warning -Message "Error $($provisioningTask.TerminatingError) provisioning $numberOfMachines VMs at $(Get-Date -Date $provisioningTask.DateFinished -Format G), master image $($provisioningTask.MasterImage)"
            }
            Write-Verbose -Message "Creating $($provisioningTask.CreatedVirtualMachines.Count) new broker machines"
            ForEach( $newMachine in $provisioningTask.CreatedVirtualMachines )
            {
                Write-Verbose -Message "Creating new machine $($newMachine.VMName)"
                $newBrokerMachine = $null
                try
                {
                    $newBrokerMachine = New-BrokerMachine @citrixParameters -CatalogUid $existingCatalogue.Uid -MachineName $newMachine.ADAccountSid -InMaintenanceMode $maintenanceMode.IsPresent
                }
                catch
                {
                }
                
                if( -Not $newBrokerMachine )
                {
                    Write-Warning -Message "Failed to create new machine for $($newMachineName.VMName)"
                }
                else
                {
                    if( $tag )
                    {
                        Add-BrokerTag -InputObject $tag -Machine $newBrokerMachine @citrixParameters
                    }
                    $newMachines.Add( $newBrokerMachine )
                }
            }
            
            if( $provisioningTask )
            {
                $null = $provisioningTask | Remove-ProvTask @citrixParameters
            }
        }
    }
}

if( $deliveryGroup -match '[\/;:#.*?=<>|\[\]()"'']' )
{
    Throw "Illegal characters in delivery group name $deliveryGroup"
}

if( -Not [string]::IsNullOrEmpty( $deliveryGroup ) -and -Not ( $existingDeliveryGroup = Get-BrokerDesktopGroup -Name $deliveryGroup @citrixParameters -ErrorAction SilentlyContinue ) )
{
    if( $PsCmdlet.ShouldProcess( "Delivery group `"$deliveryGroup`"" , "Create" ) )
    {
        Write-Verbose -Message "Delivery group `"$deliveryGroup`" not found so will create"
        [hashtable]$deliveryGroupParameters = $citrixParameters.Clone()

        if( -Not [string]::IsNullOrEmpty( $TenantId ) )
        {
            $deliveryGroupParameters.Add( 'TenantId' , $TenantId )
        }
        if( -Not [string]::IsNullOrEmpty( $scope ) )
        {
            $deliveryGroupParameters.Add( 'Scope' , $Scope )
        }
        <#
        if( -Not [string]::IsNullOrEmpty( $AppProtectionScreenCaptureRequired ) )
        {
            $deliveryGroupParameters.Add( 'AppProtectionScreenCaptureRequired' , $true )
        }
        if( -Not [string]::IsNullOrEmpty( $AppProtectionKeyLoggingRequired ) )
        {
            $deliveryGroupParameters.Add( 'AppProtectionKeyLoggingRequired' , $true )
        }
        #>
        if( -Not ( $existingDeliveryGroup = New-BrokerDesktopGroup @deliveryGroupParameters -Name $deliveryGroup -Description $description -Enabled (-Not $disabled) -SessionSupport $sessionsupport -DesktopKind $desktopKind `
            -DeliveryType $deliveryType -InMaintenanceMode $maintenanceMode.IsPresent -SecureIcaRequired $SecureIcaRequired.IsPresent -ColorDepth $colorDepth `
            -AllowReconnectInMaintenanceMode $AllowReconnectInMaintenanceMode.IsPresent -MinimumFunctionalLevel $MinimumFunctionalLevel ) )
        {
            Write-Warning -Message "Failed to create delivery group `"$deliveryGroup`""
        }
        else
        {
            Write-Verbose -Message "Delivery group `"$deliveryGroup`" created ok"
        
            if( $tag )
            {
                Add-BrokerTag -InputObject $tag -DesktopGroup $existingDeliveryGroup @citrixParameters
            }

            [hashtable]$ruleParameters = $citrixParameters.Clone()

            if( $null -ne $allowedAccounts -and $allowedAccounts.Count -gt 0 )
            {
                $ruleParameters += @{
                    'AllowedUsers' = 'Filtered'
                    'IncludedUsers' = $allowedAccounts
                }
            }
            else
            {
                ## need to create default access policies otherwise get "The User configuration has been manually modified and cannot be changed by Studio" in Studio and doesn't show to users
                $ruleParameters += @{
                    'AllowedUsers' = 'AnyAuthenticated'
                }
            }

            if( ! ( $viaAG = New-BrokerAccessPolicyRule @ruleParameters -AllowedConnections ViaAG -Name "$($deliveryGroup)_AG" -DesktopGroupUid $existingDeliveryGroup.Uid -AllowedProtocols HDX,RDP -AllowRestart $true -IncludedSmartAccessFilterEnabled $true -IncludedUserFilterEnabled $true ) )
            {
                Write-Warning -Message "Failed to create AG policy rule for new desktop group `"$deliveryGroup`""
            }
            if( ! ( $direct = New-BrokerAccessPolicyRule @ruleParameters -AllowedConnections NotViaAG -Name "$($deliveryGroup)_Direct" -DesktopGroupUid $existingDeliveryGroup.Uid -AllowedProtocols HDX,RDP -AllowRestart $true -IncludedSmartAccessFilterEnabled $true -IncludedUserFilterEnabled $true ) )
            {
                Write-Warning -Message "Failed to create Non AG policy rule for new desktop group `"$deliveryGroup`""
            }
        }
    }
}

if( $newMachines -and $newMachines.Count -gt 0 )
{
    if( $PsCmdlet.ShouldProcess( "Delivery group `"$deliveryGroup`"" , "Add $($newMachines.Count) machines" ) )
    {
        Write-Verbose -Message "Adding $($newMachines.Count) new machines to delivery group `"$deliveryGroup`""
        Add-BrokerMachine @citrixParameters -DesktopGroup $existingDeliveryGroup.Uid -InputObject $newMachines
        if( -Not $? )
        {
            Write-Warning -Message "Problem adding $($newMachines.Count) new machines to delivery group `"$deliveryGroup`""
        }
    }
}

if( -Not [string]::IsNullOrEmpty( $desktopName ) )
{
    ## TODO Single User OS: New-BrokerAssignmentPolicyRule -Name "MCS Win 10 Better" -Description "Created via PS New-BrokerAssignmentPolicyRule" -DesktopGroupUid 1666 -MaxDesktops 1 -PublishedName "MCS Win 10 Better"  -IncludedUserFilterEnabled $false 
    
    if( $sessionsupport -ieq 'SingleSession' )
    {
        $existingDesktop = Get-BrokerAssignmentPolicyRule  -DesktopGroupUid $existingDeliveryGroup.Uid -Name $desktopName -ErrorAction SilentlyContinue @citrixParameters
    }
    else
    {
        $existingDesktop = Get-BrokerEntitlementPolicyRule -DesktopGroupUid $existingDeliveryGroup.Uid -Name $desktopName -ErrorAction SilentlyContinue @citrixParameters
    }

    if( -Not $existingDesktop )
    {
        if( $PsCmdlet.ShouldProcess( $desktopName , 'Create Published Desktop' ) )
        {
            ## publish a desktop
            [hashtable]$desktopParameters = $citrixParameters.Clone()
            if( -Not [string]::IsNullOrEmpty( $RestrictDesktopToTag ))
            {
                if( -Not ( $restrictTag = Get-BrokerTag -Name $restrictDesktopToTag -ErrorAction SilentlyContinue @citrixParameters ) )
                {
                    if( ! ( $restrictTag = New-BrokerTag -Name $restrictDesktopToTag -Description "Auto created by script" @citrixParameters ) )
                    {
                        Write-Warning -Message "Failed to create tag `"$tagName`""
                    }
                }
                if( $restrictTag )
                {
                    $desktopParameters.Add( 'RestrictToTag' , $restrictDesktopToTag )
                }
                else
                {
                    Write-Warning -Message "Unable to restrict desktop `"$desktopName`" to tag $restrictDesktopToTag"
                }
            }
            if( $null -ne $allowedDesktopAccounts -and $allowedDesktopAccounts.Count -gt 0 )
            {
                $desktopParameters.Add( 'IncludedUserFilterEnabled' , $true )
                $desktopParameters.Add( 'IncludedUsers' , $allowedDesktopAccounts )
            }
            else
            {
                $desktopParameters.Add( 'IncludedUserFilterEnabled' , $false )
            }

            if( $sessionsupport -ieq 'SingleSession' )
            {
                $newPublishedDesktop = New-BrokerAssignmentPolicyRule -Description $description -Name $desktopName -PublishedName $desktopName -DesktopGroupUid $existingDeliveryGroup.Uid -Enabled $true -MaxDesktops $maxDesktopsPerUser @desktopParameters
            }
            else
            {
                $newPublishedDesktop = New-BrokerEntitlementPolicyRule -Description $description -Name $desktopName -PublishedName $desktopName -DesktopGroupUid $existingDeliveryGroup.Uid -Enabled $true @desktopParameters
            }

            if( -Not $newPublishedDesktop )
            {
                Write-Warning -Message "Failed to create published desktop `"$desktopName`""
            }
        }
    }
    else
    {
        Write-Warning -Message "Desktop `"$desktopName`" already exists"
    }
}
