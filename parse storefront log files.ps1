#requires -version 3
<#
    Take Citrix (almost) XML format StoreFront logs and convert to csv or grid view

    Guy Leech, 2018
#>

<#
.SYNOPSIS

Produce grid view or csv report of Citrix StoreFront logs

.DESCRIPTION

Allows filtering and/or searching of StoreFront logs. See https://support.citrix.com/article/CTX139592 for changing the log level on StoreFront servers.

.PARAMETER folder

The folder containing the StoreFront logs. Only use this option if the logs are not in the default folder.

.PARAMETER inputFile

A specific StoreFront log file to parse rather than parsing all files in the folder specified via -folder

.PARAMETER computers

A comma separated list of StoreFront servers to retrieve logs from via the C$ share.

.PARAMETER outputFile

The csv file to write results to.

.PARAMETER last

Include log entries in the preceding period where 's' is seconds, 'm' is minutes, 'h' is hours, 'd' is days, 'w' is weeks and 'y' is years so 12h will show all entries written in the last 12 hours.

.PARAMETER sinceBoot

Include log entries since the boot time of the computer being processed.

.PARAMETER start

Include log entries from the specified start date/time such as "02:00:00 01/01/18"

.PARAMETER end

Include log entries up to the specified end date/time such as "23:10:00 02/02/18"

.PARAMETER subtypes

A comma separated list of the types of log entry to include such as "error" or "error,warning"

.EXAMPLE

& '.\parse storefront log files.ps1' -computers storefront01,storefront02 -last 10h

Retrieve all log entries from the two specified StoreFront servers produced in the last 10 hours and output to an on-screen grid view

.EXAMPLE

& '.\parse storefront log files.ps1' -computers storefront01,storefront02 -sinceBoot -subtypes error -outputFile storefront.errors.csv

Retrieve all error log entries from the two specified StoreFront servers produced since each was booted and write to the specified file

#>

[cmdletBinding()]

Param
(
    [string]$folder = "$($env:ProgramFiles)\Citrix\Receiver StoreFront\Admin\trace" ,
    [string[]]$computers = @( 'localhost' ) ,
    [string]$inputFile ,
    [string]$outputFile ,
    [Parameter(ParameterSetName='Last')]
    [string]$last , 
    [Parameter(ParameterSetName='SinceBoot')]
    [switch]$sinceBoot ,
    [Parameter(ParameterSetName='StartTime')]
    [string]$start ,
    [string]$end ,
    [string[]]$subtypes
)

Function Process-LogFile
{
    Param( [string]$inputFile , [string[]]$subtypes , [datetime]$start , [datetime]$end )

    [string]$parent = 'GuyLeech' ## doesn't matter what it is since it only ever lives in memory

    Write-Verbose "Processing $inputFile ..." 
    ## Missing parent context so add
    [xml]$wellformatted = "<$parent>" + ( Get-Content $inputFile ) + "</$parent>"

    <#
    <EventID>0</EventID>
		    <Type>3</Type>
		    <SubType Name="Information">0</SubType>
		    <Level>8</Level>
		    <TimeCreated SystemTime="2016-05-25T14:59:12.8077640Z" />
		    <Source Name="Citrix.DeliveryServices.WebApplication" />
		    <Correlation ActivityID="{00000000-0000-0000-0000-000000000000}" />
		    <Execution ProcessName="w3wp" ProcessID="52168" ThreadID="243" />
    #>

    try
    {
        $wellformatted.$parent.ChildNodes | Where-Object { ( [string]::IsNullOrEmpty( $subtypes ) -or $subtypes -contains $_.System.SubType.Name ) -and [datetime]$_.System.TimeCreated.SystemTime -ge $start -and [datetime]$_.System.TimeCreated.SystemTime -le $end } | ForEach-Object `
        {
            $result = New-Object -TypeName PSObject -Property `
                (@{'Date'=[datetime]$_.System.TimeCreated.SystemTime; 
                'File' = (Split-Path $inputFile -Leaf); 
                'Type'=$_.System.Type;
                'Subtype'=$_.System.Subtype.Name;
                'Level'=$_.System.Level;
                'SourceName'=$_.System.Source.Name;
                'Computer'=$_.System.Computer;
                'ApplicationData'=$_.ApplicationData}) 
            ## Not all objects have all properties
            if( ( Get-Member -InputObject $_.system -Name 'Process' -Membertype Properties -ErrorAction SilentlyContinue ) )
            {
                $result | Add-Member -MemberType NoteProperty -Name 'Process' -Value $_.System.Process.ProcessName
                $result | Add-Member -MemberType NoteProperty -Name 'PID' -Value $_.System.Process.ProcessID;`
                $result | Add-Member -MemberType NoteProperty -Name 'TID' -Value $_.System.Process.ThreadID;`
            }
            $result
        } | Select Date,File,SourceName,Type,SubType,Level,Computer,Process,PID,TID,ApplicationData ## This is the order they will be displayed in
    }
    catch {}
}

if( [string]::IsNullOrEmpty( $inputFile ) -and [string]::IsNullOrEmpty( $folder ) )
{
    Write-Error "Must specify a folder (-folder) or specific log file (-inputFile) to process"
    return
}

if( ! [string]::IsNullOrEmpty( $last ) )
{
    ## see what last character is as will tell us what units to work with
    [int]$multiplier = 0
    switch( $last[-1] )
    {
        "s" { $multiplier = 1 }
        "m" { $multiplier = 60 }
        "h" { $multiplier = 3600 }
        "d" { $multiplier = 86400 }
        "w" { $multiplier = 86400 * 7 }
        "y" { $multiplier = 86400 * 365 }
        default { Write-Error "Unknown multiplier `"$($last[-1])`"" ; return }
    }
    $endDate = Get-Date
    if( $last.Length -le 1 )
    {
        $startDate = $endDate.AddSeconds( -$multiplier )
    }
    else
    {
        $startDate = $endDate.AddSeconds( - ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [int] ) * $multiplier ) )
    }
}
else
{
    if( $sinceBoot )
    {
        if( Get-Command Get-CimInstance -ErrorAction SilentlyContinue )
        {
            ## if doing since boot then we'll do this per server but set $startDate now anyway
            $startDate = Get-Date
        }
        else
        {
            Write-Warning ( "Cannot use -sinceBoot with PowerShell version {0}, requires minimum 3.0" -f $PSVersionTable.PSVersion.Major )
            return
        }
    }
    elseif( ! $start )
    {
        Write-Error "Must specify at least a start date/time with -start or -last"
        return
    }
    else
    {
        $startDate = [datetime]::Parse( $start )
    }
    if( ! $end )
    {
        $endDate = Get-Date
    }
    else
    {
        $endDate = [datetime]::Parse( $end )
    }
}

[datetime]$earliest = $startDate

[int]$counter = 1

$results = ForEach( $computer in $computers )
{
    if( $sinceBoot )
    {
        $startDate = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computer | Select -ExpandProperty LastBootupTime
    }
    
    if( $startDate -lt $earliest )
    {
        $earliest = $startDate
    }

    Write-Verbose "Processing $counter / $($computers.Count) : $computer between $(Get-Date $startDate -Format U) and $(Get-Date $endDate -Format U)"
    
    if( $inputFile )
    {
        [string]$thisInputFile = $inputFile -replace '^([A-Z]):' , "\\$computer\`$1`$"
        Process-LogFile -inputFile $inputFile -levels $subtypes -start $startDate -end $endDate
    }
    elseif( ! [string]::IsNullOrEmpty( $folder ) )
    {
        [string]$thisFolder= $Folder -replace '^([A-Z]):' , "\\$computer\`$1`$"
        Get-ChildItem -Path $thisFolder | Where-Object { $_.LastWriteTime -ge $startDate } | ForEach-Object { Process-LogFile -inputFile $_.FullName -subtypes $subtypes -start $startDate -end $endDate} 
    }
    $counter++
}

if( $results )
{
    [string]$message = "$($results.Count) log entries found between $(Get-Date $earliest -Format U) and $(Get-Date $endDate -Format U) on $($computers -join ' ')"
    Write-Verbose $message

    if( ! [string]::IsNullOrEmpty( $outputFile ) )
    {
        $results | Sort Date | Export-Csv -Path $outputFile -NoTypeInformation -NoClobber
    }
    else
    {
        $chosen = @( $results | Sort Date | Out-GridView -Title $message -PassThru )
        if( $chosen -and $chosen.Count )
        {
            $chosen | Clip.exe
        }
    }
}
else
{
    Write-Warning "No results found between $(Get-Date $earliest -Format U) and $(Get-Date $endDate -Format U)"
}
