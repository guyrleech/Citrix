<#
.SYNOPSIS

Make given changes to given file and then launch it via FTA. Can also create a shortcut in the sendto folder to a dynamically created wrapper script

.DESCRIPTION

Useful for ICA files from Citrix Virtual Apps & Desktops (XenApp/XenDesktop) when a published app doesn't launch in the resolution you need or across more monitors than you want.

.PARAMETER path

The path to the file to modify

.PARAMETER replacements

A comma separated list of strings to replace of the form sourcetext=newtext where the = delimiter can be changed by the -delimiter argument

.PARAMETER install

Install an explorer sendto menu shortcut with the name of the shortcut being the string passed to this parameter. A new script will be created with this name in the same folder as this script

.PARAMETER uninstall

Remove the given shortcut name from the user's sendto menu

.PARAMETER description

Description put in the shortcut

.PARAMETER deleter

The prefix to use to specify that if the following string is found in the input file that it will be deleted

.PARAMETER encoding

The encoding of the source file. If not specified the source file will be examined to determine the coding

.PARAMETER deleteOriginal

Delete the file specified by -path

.PARAMETER failOnFail

If no changes are made to the source file then do not launch the file

.PARAMETER force

Overwrite the temporary file that will be created from the source file by applying the specified replacements if it exists already

.EXAMPLE

& '.\Modify and launch file.ps1' -path c:\temp\fwerf.ica -replacements DesiredVRES=1360,DesiredHRES=2500,DesiredColor=24,DesiredColor=No -deleteOriginal

Look for strings "DesiredVRES", "DesiredHRES" , "DesiredColor" and "DesiredColor" in the file c:\temp\fwerf.ica, create a new temporary file with the specified replacements for these strings and launch it via
File Type Association, deleting the original source file

.EXAMPLE

& '.\Modify and launch file.ps1' -Install "Send ICA file to WQHD" -replacements DesiredVRES=1360,DesiredHRES=2500,DesiredColor=24,DesiredColor=No -deleteOriginal

Create a shortcut called "Send ICA file to WQHD" in the calling user's explorer send to menu which runs a dynamically created script of the same name which calls this script with the specified arguments

.NOTES

Modification History:

    @guyrleech  21/02/20  Initial release

#>

[CmdletBinding()]

Param
(
    [string]$path ,
    [string[]]$replacements ,
    [string]$delimiter = '=' ,
    [string]$install ,
    [string]$uninstall ,
    [string]$description = "Modify and launch ICA file" ,
    [string]$deleter ,
    [string]$encoding ,
    [switch]$deleteOriginal ,
    [switch]$failOnFail ,
    [switch]$force
)

## https://docs.microsoft.com/en-gb/archive/blogs/samdrey/determine-the-file-encoding-of-a-file-csv-file-with-french-accents-or-other-exotic-characters-that-youre-trying-to-import-in-powershell
Function Get-FileEncoding
{

    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [string]$Path
    )

    [byte[]]$byte = Get-Content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $Path

    if( $byte -and $byte.Count -eq 4 )
    {
        if ( $byte[0] -eq 0xef -and $byte[1] -eq 0xbb -and $byte[2] -eq 0xbf )
        { 'UTF8' }
        elseif ($byte[0] -eq 0xfe -and $byte[1] -eq 0xff)
        { 'Unicode' }
        elseif ($byte[0] -eq 0 -and $byte[1] -eq 0 -and $byte[2] -eq 0xfe -and $byte[3] -eq 0xff)
        { 'UTF32' }
        elseif ($byte[0] -eq 0x2b -and $byte[1] -eq 0x2f -and $byte[2] -eq 0x76)
        { 'UTF7'}
        else
        { 'ASCII' }
    }
}

if( $PSBoundParameters[ 'install' ] -or $PSBoundParameters[ 'uninstall' ] )
{
    [string]$launcherScriptContents = @'
        <#
            Pass file list via explorer and send to to another script

            @guyrleech 2020
        #>

        [string]$otherScript = Join-Path -Path ( Split-Path -Path (& { $myInvocation.ScriptName }) -Parent) -ChildPath '##ScriptName##'

        ForEach( $file in $args )
        {
            & $otherScript ##Arguments## -path $file
        }
'@
    [string]$sendToFolder = [Environment]::GetFolderPath( [Environment+SpecialFolder]::SendTo )
    [string]$lnkFile = Join-Path -Path $sendToFolder -ChildPath ( $(if( $PSBoundParameters[ 'install' ] ) { $install } else { $uninstall }) + '.lnk' )

    if( $PSBoundParameters[ 'install' ] )
    {
        if( Test-Path -Path $lnkFile -ErrorAction SilentlyContinue )
        {
            Throw "Shortcut `"$lnkFile`" already exists"
        }
        [string]$launcherScript = Join-Path -Path (Split-Path -Path (& { $myInvocation.ScriptName }) -Parent) -ChildPath ($install + '.ps1')
        if( Test-Path -Path $launcherScript -ErrorAction SilentlyContinue )
        {
            Throw "Launcher script `"$launcherScript`" already exists and not overwriting"
        }
        $self = Get-WmiObject -Class Win32_Process -Filter "ProcessId = '$pid'"
        [string]$powershellExecutable = $self.executablePath
        if( $powershellExecutable -match '^(.*)_ise(\.exe)$' )
        {
            $powershellExecutable = $Matches[1] + $Matches[2]
        }

        ## build command line for worker script
        ## -replacements DesiredVRES=1360,DesiredHRES=2500,DesiredColor=24,RemoveICAFile=No -deleteOriginal
        [string]$commandLine = "-replacements `"$($replacements -join ',')`" -delimiter $delimiter"
        if( $PSBoundParameters[ 'deleter' ] )
        {
            $commandLine += " -deleter $deleter"
        }
        if( $PSBoundParameters[ 'encoding' ] )
        {
            $commandLine += " -encoding $encoding"
        }
        if( $deleteOriginal )
        {
            $commandLine += " -deleteOriginal"
        }
        if( $failOnFail )
        {
            $commandLine += " -failOnFail"
        }
        if( $force )
        {
            $commandLine += " -force"
        }

        $launcherScriptContents -replace '##ScriptName##' , $(Split-Path -Path (& { $myInvocation.ScriptName }) -Leaf) -replace '##Arguments##' , $commandLine | Set-Content -Path $launcherScript
        $shellObject = New-Object -ComObject Wscript.Shell
        $shortcut = $shellObject.CreateShortcut( $lnkFile )
        $shortcut.TargetPath = $powershellExecutable
        $shortcut.WorkingDirectory = Split-Path -Path $powershellExecutable -Parent
        $shortcut.Arguments = "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -file `"$launcherScript`""
        $shortcut.Description = $description
        $shortcut.WindowStyle = 7 ## minimised
        $shortcut.Save()
    }
    elseif( $PSBoundParameters[ 'uninstall' ] )
    {
        if( ! ( Test-Path -Path $shortcut -ErrorAction SilentlyContinue ) )
        {
            Throw "Shortcut `"$shortcut`" does not exist"
        }
        Remove-Item -Path $shortcut
    }
}
else
{
    [string]$extension = [System.IO.Path]::GetExtension( $path )
    [string]$tempFile = [System.IO.Path]::GetTempFileName()  + $extension

    if( ! ( Test-Path -Path $path -ErrorAction SilentlyContinue ) )
    {
        Throw "Unable to open source file `"$path`""
    }

    if( ( Test-Path -Path $tempFile -ErrorAction SilentlyContinue ) -and ! $force )
    {
        Throw "Temp file `"$tempFile`" already exists and -force not specified"
    }

    if( $PSBoundParameters[ 'deleter' ] -and $delimiter.Length -ne 1 )
    {
        Throw "Deleter must be a single character"
    }

    if( ! $PSBoundParameters[ 'encoding' ] )
    {
        if( ! ( $encoding = Get-FileEncoding -Path $path ) )
        {
            Throw "Unable to determine encoding of `"$path`" so specify with -encoding"
        }
    }

    ## may not be passed as an array so spit it out again although this will fail if there are commas, in quotes, in the array items
    if( $replacements[0].IndexOf( ',' ) -ge 0 )
    {
        $replacements = $replacements -split ','
    }

    ## One time splitting of replacements so we can match quickly when processing the file
    [hashtable]$itemsToMatch = @{}
    ForEach( $replacement in $replacements )
    {
        Try
        {
            ## if the replacement starts with the deleter character then we delete the whole line
            if( $PSBoundParameters[ 'deleter' ] -and $replacement[0] -eq $deleter )
            {
                $itemsToMatch.Add( $replacement.SubString( 1 ) , [int]-1 )
            }
            else ## regular replacment/removal
            {
                [string[]]$split = $replacement -split $delimiter , 2
                ## if nothing after delimiter then it will replace the supplied value with the empty string
                $itemsToMatch.Add( $split[ 0 ] , [string]$(if( $split.Count -lt 2 -or [string]::IsNullOrEmpty( $split[1]) ) { "" } else { [Environment]::ExpandEnvironmentVariables( $split[ 1 ] ) } ) )
            }
        }
        Catch
        {
            Throw "Duplicated element $replacement"
        }
    }

    Try
    {
        [int]$changes = 0
        $newContent = [IO.File]::ReadAllLines( $path ) | ForEach-Object `
        {
            [string]$line = $_
            $split = $line -split $delimiter , 2
            if( $split -and $split.Count -eq 2 )
            {
                $matchedItem = $itemsToMatch[ $split[ 0 ] ]
                if( $matchedItem -ne $null )
                {
                    if( $matchedItem -is [int] )
                    {
                        $line = $null
                        $changes++
                    }
                    else
                    {
                        $line = $split[ 0 ] + $delimiter + $matchedItem
                        if( $matchedItem -eq $split[1] )
                        {
                            Write-Warning "Not changed $($split[0]) as already `"$matchedItem`""
                        }
                        else
                        {
                            $changes++
                        }
                    }
                }
            }

            if( $line )
            {
                $line
            }
        }

        if( ! $changes )
        {
            if( $failOnFail )
            {
                Throw "No changes made to file `"$path`""
            }
            else
            {
                Write-Warning -Message "No changes made to file `"$path`""
            }
        }

        ## Need to maintain encoding
        $newContent | Set-Content -Encoding $encoding -Path $tempFile

        if( Test-Path -Path $tempFile -ErrorAction SilentlyContinue )
        {
            Write-Verbose -Message "Launching `"$newContent`""
            $launched = Start-Process -FilePath $tempFile -Verb Open -PassThru
            if( ! $launched )
            {
                Throw "Failed to launch `"$newContent`""
            }
            elseif( $deleteOriginal )
            {
                Remove-Item -Path $path -Force
            }
        }

    }
    Catch
    {
        Throw $_
    }

    ## don't remove file as process using it may not have processed it yet
}
