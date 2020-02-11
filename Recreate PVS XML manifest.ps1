<#
.SYNOPSIS

Create a new XML manifest file for the given PVS vdisk, fetching version details from SQL and/or the vhdx/avhdx files on disk. This can then be used to import the disk into the PVS console.
For use when there are multiple versions of the vdisk rather than a single, monolithic vhdx file.

.DESCRIPTION

A disk can be absent in the PVS console but still exist in the PVS SQL database. If not in SQL then the versions can be pieced togther from the vhdx/avhdx files albeit without descriptions and some assumptions

.PARAMETER sqlServer

The SQL server (and instance name) where the PVS database resides

.PARAMETER database

The PVS SQL database name

.PARAMETER diskPath

Full path to the vhdx or avhdx which is the first disk version to be added

.PARAMETER credential

Use this and pass a PSCredential object containing the credentials if SQL authentication is required rather than integrated Windows auth

.PARAMETER production

Make the highest disk version production rather than maintenance mode

.PARAMETER startingVersion

The lowest version of the disk to be included in the XML manifest file

.PARAMETER pvsServer

The PVS server to query if the PVS snapins are availble on the machine running the script. It will default to the server the script is running on.

.PARAMETER pvsDll

The full path to the Citrix dll providing the PVS snapin

.PARAMETER secondsBeforeLastWrite

If greater than or equal to zero will set the creation time of the vhdx/avhdx file to this number of seconds before the last write time of the disk

.EXAMPLE

& '.\Recreate PVS XML manifest.ps1' -sqlserver GRL-SQL01\Instance2 -database CitrixProvisioning -diskPath S:\Store\xaserver2016.vhdx -production

Locate details for the disk xaserver2016 in the given SQL database, creating an importable XML file where the highest version of the disk is set to production mode

& '.\Recreate PVS XML manifest.ps1' -diskPath S:\Store\xaserver2016.vhdx

Find all versions of the base disk xaserver2016.vhdx in S:\Store creating an importable XML file where the highest version of the disk is set to maintenance mode

.NOTES

This script is used entirely at your own risk. The author accepts no responsibility for loss or damage caused by its use. Always backup existing XML files before running although the script will fail if one already exists for the given disk - it will not be overwritten

Modification History:

@guyrleech   11/02/20   Initial publice release

#>

[CmdletBinding()]

Param
(
    [string]$sqlServer ,
    [string]$database ,
    [Parameter(mandatory=$true,HelpMessage='Path to base vdisk')]
    [string]$diskPath ,
    [System.Management.Automation.PSCredential]$credential , 
    [switch]$production ,
    [string]$pvsServer ,
    [int]$startingVersion = 0 ,
    [string]$pvsDll = "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll" ,
    [int]$secondsBeforeLastWrite = -1
)

Function Convert-XMLSpecialCharacters
{
    Param
    (
        [string]$string
    )

    $string -replace '&' , '&amp;' -replace '\<' , '&lt;' -replace '\>' , '&gt;' 
}

[string]$queryDiskDetails = @'
SELECT *
  FROM [DiskVersion]
  WHERE diskFileName = @vdiskname + '.vhdx' or diskFileName like @vdiskname + '.%.vhdx' or diskFileName like @vdiskname + '.%.avhdx' or @vdiskname + '.vhd' or diskFileName like @vdiskname + '.%.vhd' or diskFileName like @vdiskname + '.%.avhd'
  ORDER BY version DESC , diskId
'@

[string]$diskName = [IO.Path]::GetFileNameWithoutExtension( $diskPath )
[string]$diskFolder = Split-Path -Path $diskPath -Parent
[string]$XMLManifest = Join-Path -Path $diskFolder -ChildPath ( $diskName + '.xml' )

if( Test-Path -Path $XMLManifest -ErrorAction SilentlyContinue )
{
    Throw "XML manifest file `"$XMLManifest`" already exists"
}

if( ( Test-Path $pvsDll -ErrorAction SilentlyContinue ) -and (Import-Module -Name "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll" -PassThru -Verbose:$false) )
{
    if( $PSBoundParameters[ 'pvsServer' ] )
    {
        Set-PvsConnection -Server $pvsServer
    }
    if( $existingDisk = Get-PvsSite | Get-PvsDiskInfo | Where-Object { $_.DiskLocatorName -eq $diskName } )
    {
        Throw "Disk $diskname already exists in PVS in site $($existingDisk.SiteName), store $($existingDisk.StoreName) created by $($existingDisk.Author) on $(Get-Date -Date ([datetime]$existingDisk.Date) -Format G)"
    }
}

if( $PSBoundParameters[ 'sqlServer' ] -or $PSBoundParameters[ 'database' ] )
{
    $connectionString = "Data Source=$sqlServer;Initial Catalog=$database;"

    if( $PSBoundParameters[ 'credential' ] )
    {
        ## will only work for SQL auth, Windows must be done via RunAs
        $connectionString += "Integrated Security=no;"
        $connectionString += "uid=$($credential.UserName);"
        $connectionString += "pwd=$($credential.GetNetworkCredential().Password);"
    }
    else
    {
        $connectionString += "Integrated Security=SSPI;"
    }

    $dbConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
    $dbConnection.ConnectionString = $connectionString

    try
    {
        $dbConnection.open()
    }
    catch
    {
        Write-Error "Failed to connect with `"$connectionString`" : $($_.Exception.Message)"  
        $connectionString = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
        Exit 1
    }

    $connectionString = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'

    $cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
    $cmd.connection = $dbConnection

    $cmd.CommandText = $queryDiskDetails

    $null = $cmd.Parameters.AddWithValue( "@vdiskname" , $diskName )

    if( $sqlreader = $cmd.ExecuteReader() )
    {
        $datatable = New-Object System.Data.DataTable
        $datatable.Load( $sqlreader )

        $sqlreader.Close()
        $dbConnection.Close()

        ## First pass to see if we have duplicate disk names
        [hashtable]$disks = @{}
        [int]$duplicates = 0
        $duplicateId = $null

        if( ! $datatable.Rows -or ! $datatable.Rows.Count )
        {
            Throw "Found no results in SQL for `"$diskName`""
        }

        ForEach( $version in $datatable.Rows )
        {
            Try
            {
                $disks.Add($version.diskFileName , $version.diskId )
            }
            Catch
            {
                $duplicates++
                $duplicateId = $version.diskId
            }
        }

        if( $duplicates )
        {
            Throw "Disk `"$diskPath`" exists more than once in SQL"
        }

        [string[]]$versions = @( ForEach( $version in $datatable.Rows )
        {
            [string]$diskFile = Join-Path -Path $diskFolder -ChildPath $version.diskFileName
            $diskProperties = Get-ItemProperty -Path $diskFile -ErrorAction SilentlyContinue

            if( $secondsBeforeLastWrite -ge 0 -and $diskProperties -and $diskProperties.CreationTime -gt $diskProperties.LastWriteTime )
            {
                Set-ItemProperty -Path $diskFile -Name CreationTime -Value ( $diskProperties.LastWriteTime.AddSeconds( -$secondsBeforeLastWrite ))
            }

            Write-Verbose -Message "Version $($version.version), description `"$($version.description)`""

            @"
            <version>
                <versionNumber>$($version.version)</versionNumber>
                <description>$(Convert-XMLSpecialCharacters -string $version.description)</description>
                <type>$($version.type)</type>
                <access>$($version.access)</access>
                <createDate>$((Get-Date -Date ($version.createDate.ToUniversalTime() -as [datetime]) -Format s) -replace 'T' , ' ')</createDate>
                <scheduledDate>$($version.scheduledDate)</scheduledDate>
                <deleteWhenFree>$(if( $version.deleteWhenFree -eq 'True' ) { 1 } else { 0 } )</deleteWhenFree>
                <diskFileName>$($version.diskFileName)</diskFileName>
                <size>$(if( $diskProperties ) { $diskProperties.Length })</size>
                <lastWriteTime>$(if( $diskProperties ) { Get-Date -Date $diskProperties.LastWriteTime.ToUniversalTime() -Format s })</lastWriteTime>
              </version>
"@
        })
    }
}
else ## no SQL so go off files found only
{
    if( Test-Path -Path $diskPath -ErrorAction SilentlyContinue )
    {
        [string]$extension = $null
        [System.Collections.Generic.List[string]]$disks = $null
        if( $diskPath -match '\.([\d])+.vhdx?$' -or $diskPath -match '\.([\d]+).avhdx?$' )
        {
            $startingVersion = $Matches[ 1 ]
        }
        else ## no number in file name so version zero
        {
            $startingVersion = 0
            $disks += (Split-Path -Path $diskPath -Leaf)
        }
        
        [int]$highestVesion = 0
        [string]$baseDiskName = $diskName -replace "\.$startingVersion$"
        $disks += @( Get-ChildItem -Path $diskFolder -Name "*$baseDiskName.*" | ForEach-Object `
        {
            $disk = $_
            $extension = [IO.Path]::GetExtension( $disk )
            if( $extension -eq '.vhdx' -or $extension -eq '.avhdx' -or $extension -eq '.vhd' -or $extension -eq '.avhd')
            {
                [int]$thisVersion = -1
                if( $disk -match "$baseDiskName\.([\d])+$([regex]::Escape( $extension))`$" -and ($thisVersion = ($matches[1] -as [int])) -ge $startingVersion )
                {
                    if( $thisVersion -gt $highestVesion )
                    {
                        $highestVesion = $thisVersion
                    }
                    $disk
                }
            }
        })
        if( $disks -and $disks.Count )
        {
            Write-Verbose -Message "Highest version is $highestVesion"

            [string[]]$versions = @( ForEach( $disk in $disks )
            {            
                [string]$diskFile = Join-Path -Path $diskFolder -ChildPath $disk
                $diskProperties = Get-ItemProperty -Path $diskFile -ErrorAction SilentlyContinue

                if( $diskProperties -and $diskProperties.CreationTime -gt $diskProperties.LastWriteTime )
                {
                    Set-ItemProperty -Path $diskFile -Name CreationTime -Value ( $diskProperties.LastWriteTime.AddSeconds( -$secondsBeforeLastWrite ))
                }
                
                $extension = [IO.Path]::GetExtension( $disk )
                [int]$version = 0
                if( $disk -match "\.([\d])+$([regex]::Escape( $extension))`$" )
                {
                    $version = $Matches[1]
                }
                [int]$type = 1 ## Manual
                if( $extension -match '\.vhd' )
                {
                    if( $version -eq 0 )
                    {
                        $type =  0 ## Base
                    }
                    else
                    {
                        $type = 4 ## MergeBase
                    }
                }
                
                ## can only have one maintenance version so make all production except the last one unless -production specified
                [int]$access = 0
                if( $version -eq $highestVesion -and ! $production )
                {
                    $access = 1
                } 
                @"
                <version>
                    <versionNumber>$version</versionNumber>
                    <description></description>
                    <type>$type</type>
                    <access>$access</access>
                    <createDate>$(if( $diskProperties ) { Get-Date -Date $diskProperties.CreationTimeUtc -Format s })</createDate>
                    <scheduledDate></scheduledDate>
                    <deleteWhenFree>0</deleteWhenFree>
                    <diskFileName>$disk</diskFileName>
                    <size>$(if( $diskProperties ) { $diskProperties.Length })</size>
                    <lastWriteTime>$(if( $diskProperties ) { Get-Date -Date $diskProperties.LastWriteTimeUtc -Format s })</lastWriteTime>
                  </version>
"@
            })
        }
        else
        {
            Throw "No disks found"
        }
    }
    else
    {
        Throw "Unable to find disk `"$diskPath`""
    }
}

if( $versions -and $versions.Count )
{
    &{
        '<?xml version="1.0" encoding="utf-8"?>'
        '<versionManifest>'
        "<startingVersion>$startingVersion</startingVersion>"
        $versions
        '</versionManifest>'
    } | Out-File -Filepath $XMLManifest -Encoding utf8
}
