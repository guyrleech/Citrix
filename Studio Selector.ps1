
<#
.SYNOPSIS

Allow user to select what CVAD Delivery Controller to connect Citrix Studio to, either interactively or via script arguments

.DESCRIPTION

Citrix Studio only prompts once for the delivery controller to connect to & stores it in %AppData%\Microsoft\MMC\Studio.
This script allows copies of this file to be made which store different delivery controllers which can be picked interactively at run time or via -server argument.

.PARAMETER server

The delivery controller to connect to or to produce a config file for if -save is specified (and it is assumed that this is the server last connected to).
If not specified a grid view will be presented with all previously configured delivery controllers to chose one from.

.PARAMETER save

Save the existing file produced by running Studio to a file with the name specified by the -server argument.
Keeps the existing file too so that Studio can be run outside of this script and it will connect to the same server as is stored in the file.

.PARAMETER move

Move the existing file produced by running Studio to a file with the name specified by the -server argument.
Removes the existing file too so that Studio when run outside of this script will prompt for the delivery controller to connect to.

.PARAMETER delete

Delete the default Studio config file or for a specific delivery controller if -server is specified.

.PARAMETER console

The path to the Studio executable. If not specified will be retrieved from the registry.

.EXAMPLE

& '.\Studio Selector.ps1'

Prompt the user for which delivery controller to connect to. If no previously saved Studio.server files are found Studio will be launched normally.

.EXAMPLE

& '.\Studio Selector.ps1' -save -server grl-xaddc01

Save the current Studio configuration file to Studio.grl-xaddc01 so it can be picked later.
Assumes the server that Studio was last connected to is the one mentioned - it does not check (yet)

.EXAMPLE

& '.\Studio Selector.ps1' -server grl-xaddc01

Launch Studio to connect to delivery controller grl-xaddc01, as long as Studio was previously connected to this at some juncture and a Studio.grl-xaddc01 file produced via -save or -move

.EXAMPLE

& '.\Studio Selector.ps1' -delete -server grl-xaddc01

Delete the saved Studio file for delivery controller grl-xaddc01, prompting for confirmation

.EXAMPLE

& '.\Studio Selector.ps1' -delete -server grl-xaddc01 -confirm:$false

Delete the saved Studio file for delivery controller grl-xaddc01 without prompting for confirmation

.NOTES

The delivery controller name is stored in a base64 encoded XML text string in the Studio file which is itself an XML file
If a Studio file already exists, it will be renamed to GUID.Studio in the same folder

Modification History:

12/01/2021  @guyrleech  Initial release

#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High',DefaultParameterSetName='None')]

Param
(
    [string]$server ,
    [string]$configFileName = 'Studio' ,
    [Parameter(ParameterSetName='Save')]
    [switch]$save ,
    [Parameter(ParameterSetName='Save')]
    [switch]$move ,
    [Parameter(ParameterSetName='Delete')]
    [switch]$delete ,
    [Parameter(ParameterSetName='None')]
    [string]$console
)

if( ! $PSBoundParameters[ 'console' ] )
{
    if( [string]::IsNullOrEmpty( ( $console = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Citrix\DesktopStudio' -Name LaunchPath | Select-Object -ExpandProperty LaunchPath ) ) )
    {
        Throw "Unable to find console registry value LaunchPath in HKLM\SOFTWARE\Citrix\DesktopStudio"
    }
}

if( ! ( Test-Path -Path $console -PathType Leaf -ErrorAction SilentlyContinue ) )
{
    Throw "Unable to find console launcher `"$console`""
}

[string]$configFileFolder = [System.IO.Path]::Combine( ([Environment]::GetFolderPath( [Environment+SpecialFolder]::ApplicationData )) , 'Microsoft' , 'MMC' )

if( ! ( Test-Path -Path $configFileFolder -PathType Container -ErrorAction SilentlyContinue ) )
{
    Throw "Folder `"$configFileFolder`" does not exist - has MMC ever been run for this user?"
}

[string]$originalConfigFilePath = Join-Path -Path $configFileFolder -ChildPath $configFileName

if( $save -or $move )
{
    if( ! $PSBoundParameters[ 'server' ] )
    {
        Throw "Must specify server name via -server when saving file"
    }
    if( ! ( Test-Path -Path $originalConfigFilePath -PathType Leaf -ErrorAction SilentlyContinue ) )
    {
        Throw "File `"$originalConfigFilePath`" does not exist - has Studio ever been run for this user?"
    }

    [string]$newStudioFile = Join-Path -Path $configFileFolder -ChildPath ( $configFileName + '.' + $server)
    if( ! ( Test-Path -Path $newStudioFile -ErrorAction SilentlyContinue ) -or $PSCmdlet.ShouldProcess( $newStudioFile , 'Overwrite' ) )
    {
        if( $move )
        {
            Move-Item -Path $originalConfigFilePath -Destination $newStudioFile -Force
        }
        else
        {
            Copy-Item -Path $originalConfigFilePath -Destination $newStudioFile -Force
        }
    }
}
elseif( $delete )
{
    [string]$fileName = $originalConfigFilePath
    if( $PSBoundParameters[ 'server' ] )
    {
        $fileName = Join-Path -Path $configFileFolder -ChildPath ($configFileName + ".$server")
    }

    if( ! ( Test-Path -Path $fileName -PathType Leaf -ErrorAction SilentlyContinue ) )
    {
        Write-Warning -Message "$fileName does not exist so cannot delete"
    }
    elseif( $PSCmdlet.ShouldProcess( $fileName , 'Delete' ) )
    {
        Remove-Item -Path $fileName -Force
    }
}
else
{
    [array]$studioFiles = @( ( Get-ChildItem -Path "$originalConfigFilePath.*" | Select-Object -ExpandProperty Name ) -replace "^$configFileName\." )
    if( ! $studioFiles -or ! $studioFiles.Count )
    {
        Write-Warning -Message "No saved $configFileName files found `"$originalConfigFilePath.*`" - have you run $configFileName and then this script with -save?"

        if( ! ( $launched = Start-Process -FilePath $console -WorkingDirectory (Split-Path -Path $console) -PassThru ) )
        {
            Throw "Failed to launch $console"
        }
        Write-Verbose -Message "$console launched as pid $($launched.Id) at $(Get-Date -Format G)"
    }
    else
    {
        Write-Verbose -Message "Found $($studioFiles.Count) studio files ($($studioFiles -join ' , '))"
        $chosen = $null
        if( $PSBoundParameters[ 'server' ] )
        {
            $chosen = $server
        }
        else
        {
            if( $chosenItem = $studioFiles | Select-Object @{n='Server';e={$_}} | Out-GridView -Title "Choose server to connect to with $configFileName" -PassThru  )
            {
                if( $chosenItem -is [array] )
                {
                    Throw "Only 1 server should be selected - $($chosenItem.Count) were selected"
                }
                else
                {
                    $chosen = $chosenItem.Server
                }
            }
        }

        if( ! [string]::IsNullOrEmpty( $chosen ) )
        {
            Write-Verbose -Message "$chosen chosen"
            [string]$studioFileChosen = Join-Path -Path $configFileFolder -ChildPath ($configFileName + ".$chosen")
            if( ! ( Test-Path -Path $studioFileChosen -ErrorAction SilentlyContinue ) )
            {
                Throw "Unable to find chosen file `"$studioFileChosen`""
            }
            [bool]$continue = $true
            [string]$renamedStudioFile = $null
            if( Test-Path -Path $originalConfigFilePath -ErrorAction SilentlyContinue )
            {
                $renamedStudioFile = Join-Path -Path $configFileFolder -ChildPath ((New-Guid).Guid + ".$configFileName" )
                Write-Verbose -Message "Renamed original is $renamedStudioFile"
                Move-Item -Path $originalConfigFilePath -Destination $renamedStudioFile -Force
                $continue = $? -and ( Test-Path -Path $renamedStudioFile -ErrorAction SilentlyContinue )
            }
            if( $continue )
            {
                Copy-Item -Path $studioFileChosen -Destination $originalConfigFilePath -Force
                if( $? -and ( Test-Path -Path $originalConfigFilePath -ErrorAction SilentlyContinue ) )
                {
                    if( ! ( $launched = Start-Process -FilePath $console -WorkingDirectory (Split-Path -Path $console) -PassThru ) )
                    {
                        Throw "Failed to launch $console"
                    }
                    Write-Verbose -Message "$console launched as pid $($launched.Id) at $(Get-Date -Format G)"
                }
                else
                {
                    Throw "Failed to copy `"$studioFileChosen`" to `"$originalConfigFilePath`""
                }
            }
            else
            {
                Throw "Failed to copy original file `"$originalConfigFilePath`" to `"$renamedStudioFile`""
            }
        }
        else
        {
            Write-Verbose -Message "Cancel pressed in grid view"
        }
    }
}