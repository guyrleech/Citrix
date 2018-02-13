#requires -version 3.0

<#
    Get Citrix PVS target boot time events from event log and convert to CSV for reporting or alerting purposes

    Ensure that each PVS server's stream service has event logging enabled

    Guy Leech, 2017

    Modification history:

    13/02/18   GL   Added chart view option
#>

<#
.SYNOPSIS

Get Citrix PVS target boot time events from event logs and output to CSV or a chart for reporting or alerting, by email, purposes

.DESCRIPTION

.PARAMETER computers

A comma separated list of PVS servers to query. Defaults to the computer running the script if none is given.

.PARAMETER last

Show boot times in the preceding period where 's' is seconds, 'm' is minutes, 'h' is hours, 'd' is days, 'w' is weeks and 'y' is years so 7d will show all events in the last 7 days. The default is all events.

.PARAMETER output

CSV file to create with the results.

.PARAMETER meanAbove 

If the mean (average) time (in seconds) exceeds this value then an email alert is sent.

.PARAMETER medianAbove
 
If the median time (in seconds) exceeds this value then an email alert is sent.

.PARAMETER modeAbove 

If the mode time (in seconds) exceeds this value then an email alert is sent.

.PARAMETER slowestAbove 

If the slowest time (in seconds) exceeds this value then an email alert is sent.

.PARAMETER gridView

Output the results to a graphical grid view where they can be sorted/filtered

.PARAMETER chartView

Display the results in a chart

.PARAMETER mailserver

The SMTP email server to use for sending the email.

.PARAMETER recipients

Comma separated list of email addresses to send the email to.

.PARAMETER subject 

The subject of the email. A default one is provided if none is specified.

.PARAMETER alertSubject 

The subject of the alert email, if one is sent. A default one is provided if none is specified.

.PARAMETER from  

The email address of the sender. A default one is provided if none is specified.

.PARAMETER useSSL

If specified, will communicate with the email server using SSL.

.PARAMETER search

Regex pattern to use when searching and replacing server names when they need sanitising for security purposes.

.PARAMETER replace

Regex pattern to replace in server names when they need sanitising for security purposes.

.EXAMPLE

& '.\Get PVS boot time stats.ps1' -last 7d -output c:\boot.times.csv

Will show PVS boot times in the last 7 days for the PVS server running the script and also output them to the file c:\boot.times.csv

.EXAMPLE

& '.\Get PVS boot time stats.ps1' -last 90d -chartview -computers pvsserver1,pvsserver2

Will show PVS boot times in a chart in the last 90 days for the PVS servers pvsserver1 and pvsserver2

.EXAMPLE

& '.\Get PVS boot time stats.ps1' -last 7d -output c:\boot.times.csv -mailserver mailserver -recipients someone@somewhere.com -meanAbove 180 -computers pvsserver1,pvsserver2

Will show PVS boot times in the last 7 days for the PVS servers pvsserver1 and pvsserver2 and if the average boot time is longer than 3 minutes it will send an email with the details to someone@somewhere.com.

.NOTES

Ensure that each PVS server's stream service has event logging enabled in order for the events this script looks for are generated.

#>

[CmdletBinding()]

Param
(
    [string[]]$computers = @( 'localhost' ) ,
    [string]$last ,
    [string]$output ,
    [switch]$gridView ,
    [switch]$chartView ,
    [int]$chartType = -1 ,
    [int]$meanAbove = -1 ,
    [int]$medianAbove = -1 ,
    [int]$modeAbove = -1 ,
    [int]$slowestAbove = -1 ,
    [string]$mailserver ,
    [string[]]$recipients ,
    [string]$subject = "Citrix PVS boot times from $env:COMPUTERNAME" ,
    [string]$alertSubject = "Citrix PVS boot times alert from $env:COMPUTERNAME" ,
    [string]$from  = "$($env:COMPUTERNAME)@$($env:USERDNSDOMAIN)" ,
    [switch]$useSSL ,
    [string]$providerName = 'StreamProcess' ,
    [int]$eventId = 10 ,
    [string]$eventLog = 'Application' ,
    [string]$search ,
    [string]$replace
)

[array]$events = @()
[int]$slowest = 0
[int]$fastest = [int]::MaxValue
[long]$totalTime = 0
[int]$count = 1
[hashtable]$modes = @{}
[dateTime]$startDate = (Get-Date).AddYears( -20 ) ## Should be long enough ago!

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

$events = ForEach( $computer in $computers )
{
    Write-Verbose "$count / $($computers.Count ) : processing $computer from $startDate"
    @( Get-WinEvent -ComputerName $computer -FilterHashtable @{Logname=$eventLog;ID=$eventId;ProviderName=$providerName;StartTime=$startDate} | Where-Object { $_.Message -match 'boot time'}|select TimeCreated,Message | ForEach-Object `
    {
        ## Message will be "Device xxxxx boot time: 2 minutes 50 seconds."
        if( $_.Message -match '^Device (?<Target>[^\s]+) boot time: (?<minutes>\d+) minutes (?<seconds>\d+) seconds\.$' )
        {
            [int]$boottime = ( $matches[ 'minutes' ] -as [int] ) * 60 + ( $matches[ 'seconds' ] -as [int] )
            New-Object -TypeName PSCustomObject -Property (@{ 'TimeCreated' = $_.TimeCreated ; 'Server' = $computer ; 'Target' = $matches[ 'Target' ] ; 'BootTime' = $boottime })
            $totalTime += $boottime
            if( $boottime -gt $slowest )
            {
                $slowest = $boottime
            }
            if( $boottime -lt $fastest )
            {
                $fastest = $boottime
            }
            ## Add to hash table for mode calculation
            try
            {
                $modes.Add( $boottime , 1 )
            }
            catch
            {
                $modes.Set_Item( $boottime , $modes[ $boottime ] + 1 )
            }
        }
    })
    $count++
}

if( $events.Count -gt 0 )
{
    ## See if we need to transmogrify names to protect sensitive information
    if( ! [string]::IsNullOrEmpty( $search ) )
    {
        $events | ForEach-Object `
        {
            $_.Server = $_.Server -replace $search , $replace
        }
        $computers = $computers -replace $search , $replace
        $subject = $subject  -replace $search , $replace
    }

    ## Now find median (middle) value
    [array]$sorted = $events | select BootTime | sort BootTime

    ## Now find mode (commonest) value
    [int]$mode = 0
    [int]$lastHighestCount = 0
    [int]$highestCount = 0

    $modes.GetEnumerator() | ForEach-Object `
    {
        if( $_.Value -gt $highestCount )
        {
            $lastHighestCount = $highestCount
            $highestCount = $_.Value
            $mode = $_.Key
        }
    }

    if( $highestCount -eq $lastHighestCount -or ( $highestCount -eq 1 -and $modes.Count -gt 1 ) )
    {
        $mode = 0 ## no single most common boot time
    }

    [int]$median = $sorted[$sorted.Count / 2].BootTime
    [int]$mean = [math]::Round( $totalTime / $events.Count )

    [string]$summary = "Got $($events.Count) events from $($computers.Count) machines : fastest $fastest s slowest $slowest s mean $mean s median $median s mode $mode s ($highestCount instances)"
    
    Write-Output $summary

    if( ! [string]::IsNullOrEmpty( $output ) )
    {
        $events | Export-Csv -Path $output -NoTypeInformation -NoClobber
    }

    [bool]$alert = $false
    [string]$cause = $null
    [bool]$alerting = $meanAbove -ge 0 -or $medianAbove -ge 0 -or $modeAbove -ge 0 -or $slowestAbove -ge 0
    [int]$threshold = 0

    if( $meanAbove -ge 0 -and $mean -gt $meanAbove )
    {
        $alert = $true
        $cause = 'Mean of ' + $mean
        $threshold = $meanAbove
    }
    elseif( $medianAbove -ge 0 -and $median -gt $medianAbove )
    {
        $alert = $true
        $cause = 'Median of ' + $median
        $threshold = $medianAbove
    }
    elseif( $modeAbove -ge 0 -and $mode -gt $modeAbove )
    {
        $alert = $true
        $cause = 'Mode of ' + $mode
        $threshold = $modeAbove
    }
    elseif( $slowestAbove -ge 0 -and $slowest -gt $slowestAbove )
    {
        $alert = $true
        $cause = 'Slowest of ' + $slowest
        $threshold = $slowestAbove
    }
    
    if( ! [string]::IsNullOrEmpty( $mailserver ) -And $recipients -And ( ! $alerting -or ( $alerting -and $alert ) ) )
    {
        ## workaround for scheduled task not passing array through properly
        if( $recipients.Count -eq 1 -And $recipients[0].IndexOf(",") -ge 0 )
        {
            $recipients = $recipients[0] -split ","
        }

        if( $recipients.Count -gt 0 )
        {
            [hashtable]$params = @{}
            if( $alert )
            {
                $params.Add( 'Body' , "$cause seconds exceeds threshold of $threshold for Citrix PVS target boot times since $(Get-Date -Date $startDate -Format F) on $computers" + "`n`n" + $summary )
                $params.Add( 'Subject' , $alertSubject )
            }
            else
            {
                $params.Add( 'Body' , "Citrix PVS target boot times since $(Get-Date -Date $startDate -Format F) on $computers" + "`n`n" + $summary )
                $params.Add( 'Subject' , $subject )
            }
            if( ! [string]::IsNullOrEmpty( $output ) )
            {
                $params.Add( 'Attachments' , $output )
            }

            Send-MailMessage -SmtpServer $mailserver -To $recipients -From $from -UseSsl:$useSSL @params
        }
    }

    if( $gridView )
    {
        $events | Out-GridView -Title $subject
    }
    if( $chartView )
    {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Windows.Forms.DataVisualization

        if( $chartType -lt 0 )
        {
            $chartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Range
        }
        $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
        $chart.width = 900
        $chart.Height = 600
        [void]$chart.Titles.Add( ( $subject + " since $(Get-Date $startDate -Format 'G')" ) )
        $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
        $Chart.ChartAreas.Add($ChartArea)
        $ChartArea.AxisY.Title = "Boot time (seconds)"
        ForEach( $computer in $computers )
        {
            [void]$Chart.Series.Add($computer)
            $Chart.Series[$computer].ChartType = $chartType
            $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
            $legend.name = $computer
            $Chart.Legends.Add($legend)
            $Chart.Series[$computer].ToolTip = $computer

            $events | Where-Object { $_.Server -eq $computer } | ForEach-Object `
            {
                $null = $Chart.Series[$computer].Points.AddXY( $_.TimeCreated , $_.BootTime )
            }
        }

        $AnchorAll = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor
            [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
        $Form = New-Object Windows.Forms.Form
        $Form.Width = $chart.Width
        $Form.Height = $chart.Height + 50
        $Form.controls.add($Chart)
        $Chart.Anchor = $AnchorAll

        $Form.Add_Shown({$Form.Activate()})
        [void]$Form.ShowDialog()
    }
}
else
{
    Write-Output "Found no events"
}
