#Requires -version 3.0

<#
    Change or show StoreFront tracing settings on one or more SF servers.
    Note that this will restart various services so a brief loss of service is possible when seting the trace level

    Use of this script is entirely at your own risk - the author cannot be held responsible for any undesired effects deemed to have been caused by this script.

    Guy Leech, 2018
#>

<#
.SYNOPSIS

Change the logging level of StoreFront servers or show the current trace level(s). Will warn when the settings are not consistent on a server.

.DESCRIPTION

Setting the trace level will cause some of the StoreFront services to be restarted which could result in a short loss of service.
Remember to set the trace level back to what it was when debugging has been completed.

.PARAMETER servers

A comma separated list of StoreFront servers to operate on.
If StoreFront clusters are in use then only one server in each cluster needs to be specified if the -cluster option is specified.

.PARAMETER cluster

Uses StoreFront cmdlets to get a list of the cluster members and adds them to the servers list.

.PARAMETER traceLevel

The trace level to set on the given list of servers. If not specified then the current logging state(s) are reported.

.PARAMETER grid

Output any inconsistent results to an on screen grid view and copy all selected rows into the clipboard when OK is pressed.

.PARAMETER webConfig

The name of the config file to interrogate. There should ne no need to change this.

.PARAMETER installDirKey 

The registry key where the StoreFront installation directory is stored. Only used when -cluster specified and there should ne no need to change this.

.PARAMETER installDirValue

The registry value in which the StoreFront installation directory is stored. Only used when -cluster specified and there should ne no need to change this.

.PARAMETER moduleInstaller

The StoreFront script, relative to the installation directory, which loads the StoreFront modules containing the Get-DSClusterMembersName cmdlet.
Only used when -cluster specified and there should ne no need to change this.

.PARAMETER diagnosticsNode

The XML node in the web.config files which contains the trace settings. Only used when -cluster specified and there should ne no need to change this.

.EXAMPLE

& '.\StoreFront Logging' -servers storefront01,storefront02 -grid

Report the state of logging on the listed StoreFront servers and display in a sortbable and filterable grid view.

.EXAMPLE 

& '.\StoreFront Logging' -servers storefront01 -cluster -traceLevel Error

Set the state of logging to 'Error' on the listed server and any cluster members.
This will cause StoreFront services to be restarted regardless of whether the trace level is changing or not.

.LINK

https://docs.citrix.com/en-us/storefront/3/sf-troubleshoot.html

.NOTES

There doesn't seem to be a cmdlet for reading trace settings so the web.config files are parsed in order to retrieve the settings.

Can be runfrom any server, not necessarily a StoreFront one.

Use -verbose to get more detailed information.
#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]

Param
(
    [switch]$cluster ,
    [string[]]$servers = @( 'locahost' ) ,
    [ValidateSet('Off', 'Error','Warning','Info','Verbose')]
    [string]$traceLevel ,
    [switch]$grid ,
    ## Parameters below here generally should not need to be changed. Only really here as I detest hard coding of strings but I like scripts to be adaptable without code changes where possible
    [string]$webConfig = 'web.config' ,
    [string]$installDirKey = 'SOFTWARE\Citrix\DeliveryServices' ,
    [string]$installDirValue = 'InstallDir' ,
    [string]$moduleInstaller = 'Scripts\ImportModules.ps1' ,
    [string]$diagnosticsNode = 'configuration/system.diagnostics/switches/add' ,
    [string]$logfileNode = 'configuration/system.diagnostics/sharedListeners/add'
)

## The name of the attribute we add to the web.config XML to store the name of that file. Must not exist already in the XML. Does not change the web.config file itself.
Set-Variable -Name 'fileNameAttribute' -Value '__FileName' -Option Constant

if( [string]::IsNullOrEmpty( $traceLevel ) )
{
    [string]$snapin = 'Citrix.DeliveryServices.Web.Commands'
}
else
{
    [string]$snapin = 'Citrix.DeliveryServices.Framework.Commands' 
}

## retrieve versions so store for second pass
[hashtable]$StoreFrontVersions = @{}

## We will retrieve all the cluster members - although multiple SF servers may have been specified, they may be members of different clusters so we get all unique names
if( $cluster )
{
    $newServers = New-Object -TypeName System.Collections.ArrayList
    ForEach( $server in $servers )
    {
        ## Read install dir from registry so we don't have to load all SF cmdlets
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$server)
        if( $reg )
        {
            $RegSubKey = $Reg.OpenSubKey($installDirKey)

            if( $RegSubKey )
            {
                $installDir = $RegSubKey.GetValue($installDirValue) 
                if( ! [string]::IsNullOrEmpty( $installDir ) )
                {
                    $script = Join-Path $installDir $moduleInstaller
                    [string]$version,[string[]]$clusterMembers = Invoke-Command -ComputerName $server -ScriptBlock `
                    {
                        & $using:script
                        (Get-DSVersion).StoreFrontVersion
                        @( (Get-DSClusterMembersName).Hostnames )
                    } 
                    if( ! [string]::IsNullOrEmpty( $version ) )
                    {
                        $StoreFrontVersions.Add( $server , $version )
                    }
                    if( $clusterMembers -and $clusterMembers.Count )
                    {
                        ## now iterate through and add to servers array if not present already although indirectly since we are already iterating over this array so cannot change it
                        $clusterMembers | ForEach-Object `
                        {
                            if( $servers -notcontains $_ -and $newServers -notcontains $_ )
                            {
                                $null = $newServers.Add( $_ )
                            }
                        }
                    }
                    else
                    {
                        Write-Warning "No cluster members found via $server"
                    }
                }
                else
                {
                    Write-Error "Failed to read value `"$installDirValue`" from key HKLM\$installDirKey on $server"
                }
                $RegSubKey.Close()
            }
            else
            {
                Write-Error "Failed to open key HKLM\$installDirKey on $server"
            }
            $reg.Close()
        }
        else
        {
            Write-Error "Failed to open key HKLM on $server"
        }
    }
    if( $newServers -and $newServers.Count )
    {
        Write-Verbose "Adding $($newServers.Count) servers to action list: $($newServers -join ',')"
        $servers += $newServers
    }
}

[int]$badServers = 0 
                              
[array]$results = @( ForEach( $server in $servers )
{
    if( [string]::IsNullOrEmpty( $traceLevel ) )
    {
        ## keyed on file name with value as the XML from that file
        [hashtable]$configFiles = Invoke-Command -ComputerName $server  -ScriptBlock `
        {
            Add-PSSnapin $using:snapin
            [hashtable]$files = @{}
            $dsWebSite = Get-DSWebSite
            $dsWebSite.Applications | ForEach-Object `
            {
                $app = $_
                [string]$configFile = Join-Path ( Join-Path $dsWebSite.PhysicalPath $app.VirtualPath ) $using:webConfig
                [xml]$content = $null

                if( Test-Path $configFile -ErrorAction SilentlyContinue )
                {
                    $content = (Get-Content $configFile)
                }
                $files.Add( $configFile , $content )
            }
            $files
        }
        Write-Verbose "Got $($configFiles.Count) $webConfig files from $server"
        [hashtable]$states = @{}
        $configFiles.GetEnumerator() | ForEach-Object `
        {
            [xml]$node = $_.Value
            [string]$fileName = $_.Key
            $diags = $null
            try
            {
                $diags = @( $node.SelectNodes( "//$diagnosticsNode" ) )
                $logFile = $node.SelectSingleNode( "//$logfileNode" )  ## should only be one
            }
            catch { }
            if( $diags )
            {
                $diags | ForEach-Object `
                {
                    $thisSwitch = $_
                    [string]$module = ($thisSwitch.Name -split '\.')[-1]
                    $info = $null 
                    try
                    {
                        $info = [pscustomobject]@{ 'Server' = $server ; 'Trace Level' = $thisSwitch.Value ; 'Config File' = $fileName ; 'Module' = $module }
                        [string]$version = $StoreFrontVersions[ $server ]
                        if( ! [string]::IsNullOrEmpty( $version ) )
                        {
                            Add-Member -InputObject $info -MemberType NoteProperty -Name 'StoreFront Version' -Value $version
                        }
                        if( $logFile )
                        {
                           Add-Member -InputObject $info -NotePropertyMembers @{ 'Log File' = $logFile.initializeData ; 'Max Size (KB)' = $logFile.maxFileSizeKB  }
                           ## Seems this isn't present for all SF versions
                           if( Get-Member -InputObject $logFile -Name fileCount -ErrorAction SilentlyContinue )
                           {
                               Add-Member -InputObject $info -MemberType NoteProperty -Name 'Log File Count' -Value $logfile.fileCount
                           }
                        }
                        $states.Add( $thisSwitch.Value , [System.Collections.ArrayList]( @( $info ) ) )
                    }
                    catch
                    {
                        if( ! [string]::IsNullOrEmpty( $thisSwitch.Name ) -and $info )
                        {
                            $null = $states[ $thisSwitch.Value ].Add( $info )
                        }
                    }
                }
            }
        }
        $states.GetEnumerator() | select -ExpandProperty Value ## push into results array
        if( $states.Count -gt 1 )
        {
            Write-Warning "Trace levels are inconsistent on $server - $(($states.GetEnumerator()|Select -ExpandProperty Name) -join ',')"
            $states.GetEnumerator() | ForEach-Object { Write-Verbose "$($_.Name) :`n`t$(($_.Value|select -ExpandProperty 'Config File') -join ""`n`t"" )" }
            $badServers++
        }
        elseif( ! $states.Count )
        {
            Write-Warning "No trace levels found on $server"
        }
        Write-Host "$server : logging level is $(($states.GetEnumerator()|select -ExpandProperty Name ) -join ',')" ## Can't be Write-Output otherwise will be captured into the results array
    }
    elseif( $PSCmdlet.ShouldProcess( $server , "Set trace level to $tracelevel & restart services" ) )
    {
        Write-Verbose "Setting trace level to $traceLevel on $server"
        Invoke-Command -ComputerName $server -ScriptBlock { Add-PSSnapin $using:snapin ; Set-DSTraceLevel -All -TraceLevel $using:traceLevel }
    }
} )

if( $grid -and $results.Count )
{
    [string]$title = $( if( $badServers )
    {
        "Inconsistent settings found on $badServers out of"
    }
    else
    {
        ## Now check it's the same consistent setting across all servers
        [string]$lastLevel = $null
        [bool]$matching = $true
        [int]$different = 0
        ForEach( $server in $servers )
        {
            [string]$thisLevel = $results | Where-Object { $_.Server -eq $server } | Select -First 1 -ExpandProperty 'Trace Level'
            if( $lastLevel -and $thisLevel -ne $lastLevel)
            {
                $matching = $false
                $different++
            }
            $lastLevel = $thisLevel
        }
        if( $matching )
        {
            "Consistent settings found on all"
        }
        else
        {
            "Different settings found on $different out of"
        }
    } ) + " $($servers.Count) StoreFront servers"

    $selected = $results | Select * | Out-GridView -Title $title -PassThru
    if( $selected )
    {
        $selected | Clip.exe
    }
}
