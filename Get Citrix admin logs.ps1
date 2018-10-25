#requires -version 3
<#
    Retrieve Citrix Studio logs

    @guyrleech 2018
#>

<#
.SYNOPSIS

Produce grid view or csv report of Citrix XenApp/XenDesktop admin logs such as from actions in Studio or Director

.PARAMETER ddc

The Delivery Controller to connect to, defaults to the local machine which must have the Citrix Studio PowerShell modules available

.PARAMETER username

Only return records for the specified user

.PARAMETER operation

Only return records which match the specified operation such as "log off" or "shadow"

.PARAMETER start

Only return records created on or after the given date/time

.PARAMETER end

Only return records created on or before the given date/time

.PARAMETER last

Only return records created in the last x seconds/minutes/hours/days/weeks years, e.g. 1d for 1 day or 12h for 12 hours

.PARAMETER outputfile

Write the returned records to the csv file named

.PARAMETER gridview

Display the returned results in an on screen filterable and sortable grid view

.PARAMETER configChange

Only return records which are for configuration changes

.PARAMETER adminActions

Only return records for administrative actions like shadowing or logging off

.PARAMETER studioOnly

Only return records for operations performed via Studio

.PARAMETER directorOnly

Only return records for operations performed via Director

.PARAMETER maxRecordCount

Returns at most this number of records. If more records are available than have been returned then a warning message will be displayed.

.EXAMPLE

& '.\Get Citrix admin logs.ps1' -username manuel -gridview -last 14d -operation "Shadow"

Show all shadow operations performed by the user manuel in the last 14 days and display in a grid view

.EXAMPLE

& '.\Get Citrix admin logs.ps1' -start "01/01/2018" -end "31/01/2018" -configChange -outputfile c:\temp\citrix.changes.csv

Show all configuration changes made between the 1st and 31st of January 2018 and write the results to c:\temp\citrix.changes.csv

#>

[CmdletBinding()]

Param
(
    [string]$ddc = 'localhost' ,
    [string]$username ,
    [switch]$configChange ,
    [switch]$adminAction ,
    [switch]$studioOnly ,
    [switch]$directorOnly ,
    [string]$operation ,
    [Parameter(Mandatory=$true, ParameterSetName = "TimeSpan")]
    [datetime]$start ,
    [Parameter(Mandatory=$false, ParameterSetName = "TimeSpan")]
    [datetime]$end = [datetime]::Now ,
    [Parameter(Mandatory=$true, ParameterSetName = "Last")]
    [string]$last ,
    [string]$outputFile ,
    [int]$maxRecordCount = 5000 ,
    [switch]$gridview
)

if( $studioOnly -and $directorOnly )
{
    Throw "Cannot specify both -studioOnly and -directorOnly"
}

if( ! [string]::IsNullOrEmpty( $last ) )
{
    [long]$multiplier = 0
    switch( $last[-1] )
    {
        "s" { $multiplier = 1 }
        "m" { $multiplier = 60 }
        "h" { $multiplier = 3600 }
        "d" { $multiplier = 86400 }
        "w" { $multiplier = 86400 * 7 }
        "y" { $multiplier = 86400 * 365 }
        default { Throw "Unknown multiplier `"$($last[-1])`"" }
    }
    if( $last.Length -le 1 )
    {
        $start = $end.AddHours( -$multiplier )
    }
    else
    {
        $start = $end.AddSeconds( - ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [long] ) * $multiplier ) )
    }
}
elseif( ! $PSBoundParameters[ 'start' ] )
{
    $start = (Get-Date).AddDays( -7 )
}

Add-PSSnapin -Name 'Citrix.ConfigurationLogging.Admin.*'

if( ! ( Get-Command -Name 'Get-LogHighLevelOperation' -ErrorAction SilentlyContinue ) )
{
    Throw "Unable to find the Citrix Get-LogHighLevelOperation cmdlet required"
}

[hashtable]$queryparams = @{
    'AdminAddress' = $ddc
    'SortBy' = '-StartTime'
    'MaxRecordCount' = $maxRecordCount
    'ReturnTotalRecordCount' = $true
}
if( $configChange -and ! $adminAction )
{
    $queryparams.Add( 'OperationType' , 'ConfigurationChange' )
}
elseif( ! $configChange -and $adminAction )
{
    $queryparams.Add( 'OperationType' , 'AdminActivity' )
}
if( ! [string]::IsNullOrEmpty( $username ) )
{
    if( $username.IndexOf( '\' ) -lt 0 )
    {
        $username = $env:USERDOMAIN + '\' + $username
    }
    $queryparams.Add( 'User' , $username )
}
if( $directorOnly )
{
    $queryparams.Add( 'Source' , 'Citrix Director' )
}
if( $studioOnly )
{
    $queryparams.Add( 'Source' , 'Studio' )
}

$recordCount = $null

[array]$results = @( Get-LogHighLevelOperation -Filter { StartTime -ge $start -and EndTime -le $end }  @queryparams -ErrorAction SilentlyContinue -ErrorVariable RecordCount | ForEach-Object -Process `
{
    if( [string]::IsNullOrEmpty( $operation ) -or $_.Text -match $operation )
    {
        $result = [pscustomobject][ordered]@{
            'Started' = $_.StartTime
            'Duration (s)' = [math]::Round( (New-TimeSpan -Start $_.StartTime -End $_.EndTime).TotalSeconds , 2 )
            'User' = $_.User
            'From' = $_.AdminMachineIP
            'Operation' = $_.text
            'Source' = $_.Source
            'Type' = $_.OperationType
            'Targets' = $_.TargetTypes -join ','
            'Successful' = $_.IsSuccessful
        }
        if( ! $configChange )
        {
            Add-Member -InputObject $result -NotePropertyMembers @{
                'Target Process' = $_.Parameters[ 'ProcessName' ]
                'Target Machine' = $_.Parameters[ 'MachineName' ]
                'Target User' = $_.Parameters[ 'UserName' ]
            }
        }
        $result
    }
} )

if( $recordCount -and $recordCount.Count )
{
    if( $recordCount[0] -match 'Returned (\d+) of (\d+) items' )
    {
        if( [int]$matches[1] -lt [int]$matches[2] )
        {
            Write-Warning "Only retrieved $($matches[1]) of a total of $($matches[2]) items, use -maxRecordCount to return more"
        }
        ## else we got all the records
    }
    else
    {
        Write-Error $recordCount[0]
    }
}

if( ! $results -or ! $results.Count )
{
    Write-Warning "No log entries found between $(Get-Date $start -Format G) and $(Get-Date $end -Format G)"
}
else
{
    if( ! [string]::IsNullOrEmpty( $outputFile ) )
    {
        $results | Export-Csv -Path $outputFile -NoClobber -NoTypeInformation
    }
    elseif( $gridview )
    {
        $selected = $results | Out-GridView -Title "$($results.Count) events from $(Get-Date $start -Format G) and $(Get-Date $end -Format G)" -PassThru
        if( $selected )
        {
            $selected | clip.exe
        }
    }
    else
    {
        $results
    }
}
