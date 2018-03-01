#Requires -version 3.0

<#
    Grab all PVS devices and their collections, machine catalogues, delivery groups and more and output to csv or a grid view

    Guy Leech, 2017

    Modification history:

    19/01/18    GL  Fixed issue where not showing not booted off latest if different vdisk assigned from what is booted off
                    Fixed issue with PoSH v5
                    Added -name filter and -tags
                    Added warning for pre PoSH v5 and more than 30 columns in Out-GridView
                    Added error trapping when fails to retrive boot time

    22/01/18    GL  Added action GUI
    
    26/01/18    GL  Added ability to match AD group regex and output group memberships

    01/02/18    GL  Added disk version descriptions and dns lookup option

    22/02/18    GL  Added capability to include machines retrieved from DDCs which are not present in PVS
                    Changed main device collection to .NET array so we can add potential orphans to it
                    Added MaxRecordCount parameter for DDC calls

    25/02/18    GL  Added GUI options to remove from PVS and DDC

    27/02/18    GL  Added saving DDC and PVS servers to registry
#>

<#
.SYNOPSIS

Produce grid view or csv report of Citrix Provisioning Services devices, cross referenced with information from Delivery Controllers and Active Directory

.DESCRIPTION

Alows easy identification of potential problems such as devices booted off the wrong vdisk or vdisk version, devices with no AD account or devices overdue a reboot.
Devices can be selected in the grid view and then shutdown, restarted, booted or maintenance mode turned on or off from a simple GUI

.PARAMETER pvsServers

Comma separated list of PVS servers to contact. Do not specify multiple servers if they use the same SQL database

.PARAMETER ddcs

Comma separated list of Delivery Controllers to contact. Do not specify multiple servers if they use the same SQL database

.PARAMETER registry

Pick up the PVS and DDC servers from the registry rather than the command line. Must previously have saved them with -save and an optional server set name via -serverSet

.PARAMETER serverSet

Pick up the PVS and DDC servers from the registry sub key specified in this paramter rather than the command line. Must previously have saved them with -save and this server set name

.PARAMETER save

Save the PVS and DDC servers to the registry for later use with -registry. Use -serverSet to specify named sets of servers, e.g. preproduction and production

.PARAMETER csv

Path to a csv file that will have the results written to it. If none specified then output will be on screen to a grid view

.PARAMETER noBootTime

Will not try and contact active devices to find their last boot time. If WinRM not setup correctly or other issues mean remote calls will fail then this option can speed the script up

.PARAMETER noOrphans

Do not display machines which are present on DDCs/Studio but not in PVS. May be physical, MCS or VMs with local disks or could be orphans that did exist in PVS but do not any longer.

.PARAMETER maxRecordCount

Citrix DDC cmdlets by default only return 250 items so if there are more machines than this on a DDC then this parameter may need to be increased

.PARAMETER name

Only show devices whose name matches the regular expression specified

.PARAMETER tags

Adds a column containing Citrix tags, if present, for each device

.PARAMETER ADGroups

A regular expression of AD groups to match and will output ones that match for which the device is a member

.PARAMETER provisioningType

The type of catalogue provisioning to check for in machines that are suspected of being orphaned from PVS as they were found on a DDC but not PVS

.PARAMETER dns

Perform DNS ipV4 lookup. This can cause the script to run slower.

.PARAMETER noMenu

Do not display "OK" and "Cancel" buttons on grid view so no action menu is presented after grid view is closed

.PARAMETER messageText

Text to send to users in the action GUI. If none specified and the message option is clicked then you will be prompted for the text

.PARAMETER messageCaption

Caption of the message to send to users in the action GUI. If none specified and the message option is clicked then you will be prompted for the text

.PARAMETER timeout

Timeout in seconds for PVS power commands

.EXAMPLE

& '.\Get PVS device info.ps1' -pvsservers pvsprod01,pvstest01 -ddcs ddctest02,ddcprod03

Retrieve devices from the listed PVS servers, cross reference to the listed DDCs (list order does not matter) and display on screen in a sortable and filterable grid view

& '.\Get PVS device info.ps1' -pvsservers pvsprod01 -ddcs ddcprod03 -name CTXUAT -tags -dns -csv h:\pvs.ctxuat.csv

Retrieve devices matching regular expression CTXUAT from the listed PVS server, cross reference to the listed DDC and output to the given csv file, including Citrix tag information and IPv4 address from DNS query

.NOTES

Uses local PowerShell cmdlets for PVS and DDCs, as well as Active Directory, so run from a machine where both PVS and Studio consoles are installed.

#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='Manual',mandatory=$true,HelpMessage='Comma separated list of PVS servers')]
    [string[]]$pvsServers ,
    [Parameter(ParameterSetName='Manual',mandatory=$true,HelpMessage='Comma separated list of Delivery controllers')]
    [string[]]$ddcs ,
    [Parameter(ParameterSetName='Manual',mandatory=$false)]
    [switch]$save ,
    [Parameter(ParameterSetName='Registry',mandatory=$true,HelpMessage='Use default server set name from registry')]
    [switch]$registry ,
    [string]$serverSet = 'Default' ,
    [string]$csv ,
    [switch]$noBootTime ,
    [switch]$dns ,
    [string]$name ,
    [switch]$tags ,
    [ValidateSet('PVS','MCS','Manual')]
    [string]$provisioningType = 'PVS' ,
    [switch]$noMenu ,
    [switch]$noOrphans ,
    [string]$configRegKey = 'HKCU:\software\Guy Leech\PVS Fetcher' ,
    [string]$messageText ,
    [string]$messageCaption ,
    [int]$maxRecordCount = 2000 ,
    [int]$timeout = 60 ,
    [string]$ADgroups ,
    [string[]]$snapins = @( 'Citrix.Broker.Admin.*'  ) ,
    [string[]]$modules = @( 'ActiveDirectory', "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll" ) 
)

$columns = [System.Collections.ArrayList]( @( 'Name','DomainName','Description','PVS Server','DDC','SiteName','CollectionName','Machine Catalogue','Delivery Group','Registration State','Maintenance Mode','User Sessions','Boot Time','devicemac','active','enabled','Store Name','Disk Version Access','Disk Version Created','AD Account Exists','Disk Name','Booted off vdisk','Booted Disk Version','Vdisk Production Version','Vdisk Latest Version','Latest Version Description','Override Version','Booted off latest','Disk Description','Cache Type','Disk Size (GB)','Write Cache Size (MB)' )  )

if( $dns )
{
   $null = $columns.Add( 'IPv4 address' )
}

if( $tags )
{
   $null = $columns.Add( 'Tags')
}

if( ! [string]::IsNullOrEmpty( $ADgroups ) )
{
   $null = $columns.Add( 'AD Groups')
}


if( $PSVersionTable.PSVersion.Major -lt 5 -and $columns.Count -gt 30 -and [string]::IsNullOrEmpty( $csv ) )
{
    Write-Warning "This version of PowerShell limits the number of columns in a grid view to 30 and we have $($columns.Count) so those from `"$($columns[30])`" will be lost in grid view"
}

if( $snapins -and $snapins.Count -gt 0 )
{
    ForEach( $snapin in $snapins )
    {
        Add-PSSnapin $snapin -ErrorAction Continue
    }
}

if( $modules -and $modules.Count -gt 0 )
{
    ForEach( $module in $modules )
    {
        Import-Module $module -ErrorAction Continue
    }
}

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

$messageWindowXAML = @"
<Window x:Class="Direct2Events.MessageWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="Send Message" Height="414.667" Width="309.333">
    <Grid>
        <TextBox x:Name="txtMessageCaption" HorizontalAlignment="Left" Height="53" Margin="85,20,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="180"/>
        <Label Content="Caption" HorizontalAlignment="Left" Height="24" Margin="10,20,0,0" VerticalAlignment="Top" Width="63"/>
        <TextBox x:Name="txtMessageBody" HorizontalAlignment="Left" Height="121" Margin="85,171,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="180"/>
        <Label Content="Message" HorizontalAlignment="Left" Height="25" Margin="10,167,0,0" VerticalAlignment="Top" Width="56" RenderTransformOrigin="0.47,1.333"/>
        <StackPanel Orientation="Horizontal" Height="43" Margin="10,332,0,0"  Width="283">
            <Button x:Name="btnMessageOk" Content="OK"  Height="20"  Width="89"/>
            <Button x:Name="btnMessageCancel" Content="Cancel"  Height="20" Margin="50,0,0,0"  Width="89"/>
        </StackPanel>
        <ComboBox x:Name="comboMessageStyle" HorizontalAlignment="Left" Height="27" Margin="85,98,0,0" VerticalAlignment="Top" Width="180">
            <ComboBoxItem Content="Information" IsSelected="True"/>
            <ComboBoxItem Content="Question"/>
            <ComboBoxItem Content="Exclamation"/>
            <ComboBoxItem Content="Critical"/>
        </ComboBox>
        <Label Content="Level" HorizontalAlignment="Left" Height="27" Margin="15,98,0,0" VerticalAlignment="Top" Width="58"/>

    </Grid>
</Window>
"@

$pvsDeviceActionerXAML = @"
<Window x:Class="PVSDeviceViewerActions.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PVSDeviceViewerActions"
        mc:Ignorable="d"
        Title="PVS Device Actioner" Height="502.565" Width="401.595">
    <Grid Margin="0,0,-20.333,-103.667">
        <ListView x:Name="lstMachines" HorizontalAlignment="Left" Height="380" Margin="23,22,0,0" VerticalAlignment="Top" Width="140">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Device" DisplayMemberBinding ="{Binding 'Name'}" />
                </GridView>
            </ListView.View>
            <ListBoxItem Content="Device"/>
        </ListView>
        <StackPanel x:Name="stkButtons" Margin="198,32,36,18" Orientation="Vertical">
            <Button x:Name="btnShutdown" Content="_Shutdown" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnPowerOff" Content="_Power Off" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnRestart" Content="_Restart" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnBoot" Content="_Boot" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnMessageUsers" Content="_Message Users" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnMaintModeOn" Content="Maintenance Mode O_n" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnMaintModeOff" Content="Maintenance Mode O_ff" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnRemoveFromDDC" Content="Remove from _DDC" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnRemoveFromPVS" Content="Remove from P_VS" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
        </StackPanel>
    </Grid>
</Window>

"@

Function Save-ConfigToRegistry( [string]$serverSet = 'Default' , [string[]]$DDCs , [string[]]$PVSServers )
{
    [string]$key = Join-Path $configRegKey $serverSet
    if( ! ( Test-Path $key -ErrorAction SilentlyContinue ) )
    {
        $null = New-Item -Path $key -Force
    }
    Set-ItemProperty -Path $key -Name 'DDC' -Value $DDCs
    Set-ItemProperty -Path $key -Name 'PVS' -Value $PVSServers
}   

Function Get-ConfigFromRegistry
{
    Param
    (
        [string]$serverSet = 'Default' , 
        [ref]$DDCs , 
        [ref]$PVSServers
    )
    [string]$key = Join-Path $configRegKey $serverSet
    if( ! ( Test-Path $key -ErrorAction SilentlyContinue ) )
    {
        Write-Warning "Config registry key `"$key`" does not exist"
    }
    $DDCs.value = Get-ItemProperty -Path $key -Name 'DDC' -ErrorAction SilentlyContinue | select -ExpandProperty 'DDC' 
    $PVSServers.value = Get-ItemProperty -Path $key -Name 'PVS' -ErrorAction SilentlyContinue | select -ExpandProperty 'PVS'
}

Function Load-GUI( $inputXaml )
{
    $form = $NULL
    $inputXaml = $inputXaml -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
    [xml]$XAML = $inputXaml
 
    $reader = New-Object Xml.XmlNodeReader $xaml

    try
    {
        $Form = [Windows.Markup.XamlReader]::Load( $reader )
    }
    catch
    {
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        return $null
    }
 
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object `
    {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
    }

    return $form
}

Function Perform-Action
{
    Param
    (
        [ValidateSet('Boot','Reboot','Shutdown','Message','Maintenance Mode On','Maintenance Mode Off','Power Off','Remove From DDC','Remove From PVS')]
        [string]$action ,
        $form
    )
    
    $answer = [Windows.MessageBox]::Show( "Are you sure you want to $action these $($WPFlstMachines.SelectedItems.Count) devices ?" , "Confirm" , 'YesNo' ,'Question' )

    if( $answer -ne 'Yes' )
    {
        return
    }

    if( $action -eq 'Message' )
    {
        $messageForm = Load-GUI $messageWindowXAML
        if( $messageForm )
        {
            ## Load up caption and text and set callbacks
            $WPFtxtMessageCaption.Text = if( [string]::IsNullOrEmpty( $messageCaption ) ) { "Message from $env:USERNAME at $(Get-Date -Format F)" } else { $messageCaption }
            $WPFtxtMessageBody.Text = $messageText
            $WPFbtnMessageOk.Add_Click({
                $messageForm.DialogResult = $true
                $messageForm.Close()
            }) 
            $WPFbtnMessageOk.IsDefault = $true
            $WPFbtnMessageCancel.Add_Click({ $messageForm.Close() })
            $WPFbtnMessageCancel.IsCancel = $true
            $result = $messageForm.ShowDialog()
            if( ! $result )
            {
                return
            }
        }
    }
    
    if( $form )
    {
        $oldCursor = $form.Cursor
        $form.Cursor = [Windows.Input.Cursors]::Wait
    }

    ForEach( $device in $WPFlstMachines.SelectedItems )
    {
        Write-Verbose "Action $action on $($device.Name)"

        switch -regex ( $action )
        {
            'Remove From DDC' { Remove-BrokerMachine -Force -AdminAddress $device.ddc -MachineName $( if( [string]::IsNullOrEmpty( $device.DomainName ) ) { $device.Name } else {  $device.DomainName + '\' +  $device.Name } ) ; break  }
            'Remove From PVS' { Set-PvsConnection -Server $device.'PVS Server'; Remove-PvsDevice -DeviceName $device.Name ; break }
            'Maintenance Mode On'  { Set-BrokerMachine -AdminAddress $device.ddc -InMaintenanceMode $true -MachineName ( $device.DomainName + '\' +  $device.Name ) ; break }
            'Maintenance Mode Off' { Set-BrokerMachine -AdminAddress $device.ddc -InMaintenanceMode $false -MachineName ( $device.DomainName + '\' +  $device.Name ) ; break }
            'Reboot' { Restart-Computer -ComputerName $device.Name ; break }
            'Shutdown' { Stop-Computer -ComputerName $device.Name ; break }
            'Boot|Power Off' ` ##  Can't use New-BrokerHostingPowerAction as may not be known to DDC
            { 
                Set-PvsConnection -Server $device.'PVS Server'
                if( $_ -eq 'Boot' )
                {
                    if( [string]::IsNullOrEmpty( $device.'Disk Name' ) )
                    {
                        $answer = [Windows.MessageBox]::Show( "$($device.Name) has no vdisk assigned so may not boot - continue ?" , "Confirm" , 'YesNo' ,'Question' )

                        if( $answer -ne 'Yes' )
                        {
                            continue
                        }
                    }
                    $thePvsTask = Start-PvsDeviceBoot -DeviceName $device.Name
                }
                else
                {
                    $thePvsTask = Start-PvsDeviceShutdown -DeviceName $device.Name
                }
                $timer = [Diagnostics.Stopwatch]::StartNew()
                [bool]$timedOut = $false
                while ( $thePvsTask -and $thePvsTask.State -eq 0 ) 
                {
                    $percentFinished = Get-PvsTaskStatus -Object $thePvsTask 
                    if( ! $percentFinished -or $percentFinished.ToString() -ne 100 )
                    {
                        Start-Sleep -Milliseconds 500
                        if( $timer.Elapsed.TotalSeconds -gt $timeout )
                        {
                            $timeOut = $true
                            break
                        }
                    }
                    $thePvsTask = Get-PvsTask -Object $thePvsTask
                }
                $timer.Stop()
                
                if ( $timedOut )
                {
                    [Windows.MessageBox]::Show( $device.Name , "Failed to perform $action action - timed out after $timeout seconds" , 'OK' ,'Error' )
                } 
                elseif ( ! $thePvsTask -or $thePvsTask.State -ne 2)
                {
                    [Windows.MessageBox]::Show( $device.Name , "Failed to perform $action action" , 'OK' ,'Error' )
                } 
             } 
            'Message' { Get-BrokerSession -AdminAddress $device.ddc -MachineName ( $env:USERDOMAIN + '\' +  $device.Name ) | Send-BrokerSessionMessage -AdminAddress $device.ddc -Title $WPFtxtMessageCaption.Text -Text $WPFtxtMessageBody.Text -MessageStyle ($WPFcomboMessageStyle.SelectedItem.Content)  }
        }
    }
    
    if( $form )
    {
        $form.Cursor = $oldCursor
    }
}

if( $registry )
{
    Get-ConfigFromRegistry -serverSet $serverSet -DDCs ( [ref] $ddcs ) -PVSServers ( [ref] $pvsServers )
    if( ! $ddcs -or ! $ddcs.Count -or ! $pvsServers -or ! $pvsServers.Count )
    {
        Write-Warning "Failed to get PVS and/or DDC servers from registry key `"$configRegKey`" for server set `"$serverSet`""
        return
    }
}
elseif( $save )
{
    Save-ConfigToRegistry -serverSet $serverSet  -DDCs $ddcs -PVSServers $pvsServers
}

## Get all information from DDCs so we can lookup locally
[hashtable]$machines = @{}

ForEach( $ddc in $ddcs )
{
    $machines.Add( $ddc , [System.Collections.ArrayList] ( Get-BrokerMachine -AdminAddress $ddc -MaxRecordCount $maxRecordCount -ErrorAction SilentlyContinue ) )
}

$devices = New-Object -TypeName System.Collections.ArrayList

ForEach( $pvsServer in $pvsServers )
{
    Set-PvsConnection -Server $pvsServer 

    if( ! $? )
    {
        Write-Output "Cannot connect to PVS server $pvsServer - aborting"
        continue
    }

    ## Cache latest production version for vdisks so don't look up for every device
    [hashtable]$diskVersions = @{}

    # Get all sites that we can see on this server and find all devices and cross ref to Citrix for catalogues and delivery groups
    Get-PvsDevice | Where-Object { $_.Name -match $name } | ForEach-Object `
    {
        $device = $_
        [int]$bootVersion = -1
        $vDisk = Get-PvsDiskInfo -DeviceId $_.DeviceId

        if( $device.Active -and ! $noBootTime )
        {
            $bootTime = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $device.Name | Select -ExpandProperty LastBootUpTime
            if( $bootTime )
            {
                $device | Add-Member -MemberType NoteProperty -Name 'Boot Time' -value $bootTime
            }
            else
            {
                Write-Warning "Failed to get boot time for $($device.Name)"
            }
        }
        $device | Add-Member -MemberType NoteProperty -Name 'PVS Server' -Value $pvsServer
        if( $vdisk )
        {
            $device | Add-Member -MemberType NoteProperty -Name 'Disk Name' -value $vdisk.Name
            $device | Add-Member -MemberType NoteProperty -Name 'Store Name' -value $vdisk.StoreName
            $device | Add-Member -MemberType NoteProperty -Name 'Disk Description' -value $vdisk.Description
            $device | Add-Member -MemberType NoteProperty -Name 'Cache Type' -value $cacheTypes[$vdisk.WriteCacheType]
            $device | Add-Member -MemberType NoteProperty -Name 'Disk Size (GB)' -value ([math]::Round( $vdisk.DiskSize / 1GB , 2 ))
            $device | Add-Member -MemberType NoteProperty -Name 'Write Cache Size (MB)' -value $vdisk.WriteCacheSize
            ## Cache vdisk version info to reduce PVS server hits
            $versions = $diskVersions[ $vdisk.DiskLocatorId ]
            if( ! $versions )
            { 
                try
                {
                    $versions = Get-PvsDiskVersion -DiskLocatorId $vdisk.DiskLocatorId 
                    $diskVersions.Add( $vdisk.DiskLocatorId , $versions )
                }
                catch
                {
                }
            }
            if( $versions )
            {
                ## Now get latest production version of this vdisk
                $override = $versions | Where-Object { $_.Access -eq 3 } 
                $lastestProductionVersion = $versions | Where-Object { $_.Access -eq 0 } | Sort Version -Descending | Select -First 1 | select -ExpandProperty Version
                if( $override )
                {
                    $bootVersion = $override.Version
                }
                else
                {
                    ## Access: Read-only access of the Disk Version. Values are: 0 (Production), 1 (Maintenance), 2 (MaintenanceHighestVersion), 3 (Override), 4 (Merge), 5 (MergeMaintenance), 6 (MergeTest), and 7 (Test) Min=0, Max=7, Default=0
                    $bootVersion = $lastestProductionVersion
                }
                $device | Add-Member -MemberType NoteProperty -Name 'Override Version' -value $( if( $override ) { $bootVersion } else { $null } )    
                $device | Add-Member -MemberType NoteProperty -Name 'Vdisk Latest Version' -value $lastestProductionVersion     
                $device | Add-Member -MemberType NoteProperty -Name 'Latest Version Description' -value ( $versions | Where-Object { $_.Version -eq $lastestProductionVersion } | Select -ExpandProperty Description )         
            }
            $device | Add-Member -MemberType NoteProperty -Name 'Vdisk Production Version' -value $bootVersion
        }
        $deviceInfo = Get-PvsDeviceInfo -DeviceId $device.DeviceId
        if( $deviceInfo )
        {
            $device | Add-Member -MemberType NoteProperty -Name 'Disk Version Access' -value $accessTypes[ $deviceInfo.DiskVersionAccess ]
            if( $device.Active )
            {
                ## Check if booting off the disk we should be as previous info is what is assigned, not what is necessarily being used (e.g. vdisk changed for device whilst it is booted)
                $bootedDiskName = (( Get-PvsDiskVersion -DiskLocatorId $deviceInfo.DiskLocatorId | Select -First 1 | Select -ExpandProperty Name ) -split '\.')[0]
                $device | Add-Member -MemberType NoteProperty -Name 'Booted Disk Version' -value $deviceInfo.DiskVersion
                if( $bootVersion -ge 0 )
                {
                    Write-Verbose "$($device.Name) booted off $bootedDiskName, disk configured $($vDisk.Name)"
                    $device | Add-Member -MemberType NoteProperty -Name 'Booted off latest' -value ( $bootVersion -eq $deviceInfo.DiskVersion -and $bootedDiskName -eq $vdisk.Name )
                    $device | Add-Member -MemberType NoteProperty -Name 'Booted off vdisk' -value $bootedDiskName
                }
            }
            if( $versions )
            {
                try
                {
                    $device | Add-Member -MemberType NoteProperty -Name 'Disk Version Created' -value ( $versions | Where-Object { $_.Version -eq $deviceInfo.DiskVersion } | select -ExpandProperty CreateDate )
                }
                catch
                {
                    $_
                }
            }
        }
        if( Get-Module ActiveDirectory -ErrorAction SilentlyContinue )
        {
            [hashtable]$adparams = @{}
            if( ! [string]::IsNullOrEmpty( $ADgroups ) )
            {
                $adparams.Add( 'Properties' , 'MemberOf' )
            }
            $adAccount = $null
            try
            {
                $adaccount = Get-ADComputer $device.Name -ErrorAction SilentlyContinue @adparams
            }
            catch
            {
            }
            $device | Add-Member -MemberType NoteProperty -Name 'AD Account Exists' -value ( $adAccount -ne $null )

            if( ! [string]::IsNullOrEmpty( $ADgroups ) )
            {
                $device | Add-Member -MemberType NoteProperty -Name 'AD Groups' -value ( ( $adAccount | select -ExpandProperty MemberOf | ForEach-Object { (( $_ -split '^CN=')[1] -split '\,')[0] } | Where-Object { $_ -match $ADgroups } ) -join ' ' )
            }
        }

        if( $device.Active -and $dns )
        {
            [array]$ipv4Address = @( Resolve-DnsName -Name $device.Name -Type A )
            $device | Add-Member -MemberType NoteProperty -Name 'IPv4 address' -Value ( ( $ipv4Address | Select -ExpandProperty IPAddress ) -join ' ' )
        }
            
        if( ( Get-Command -Name Get-BrokerMachine -ErrorAction SilentlyContinue ) )
        {
            ## Need to find a ddc that will return us information on this device
            ForEach( $ddc in $ddcs )
            {
                ## can't use HostedMachineName as only populated if registered
                $machine = $machines[ $ddc ] | Where-Object { $_.MachineName -eq  ( ($device.DomainName -split '\.')[0] + '\' + $device.Name ) } ##Get-BrokerMachine -MachineName ( ($device.DomainName -split '\.')[0] + '\' + $device.Name ) -AdminAddress $ddc -ErrorAction SilentlyContinue
                if( $machine )
                {
                    $device | Add-Member -MemberType NoteProperty -Name 'Machine Catalogue' -value $machine.CatalogName
                    $device | Add-Member -MemberType NoteProperty -Name 'Delivery Group' -value $machine.DesktopGroupName
                    $device | Add-Member -MemberType NoteProperty -Name 'Registration State' -value $machine.RegistrationState
                    $device | Add-Member -MemberType NoteProperty -Name 'User Sessions' -value $machine.SessionCount
                    $device | Add-Member -MemberType NoteProperty -Name 'Maintenance Mode' -value $( if( $machine.InMaintenanceMode ) { 'On' } else { 'Off' } )
                    $device | Add-Member -MemberType NoteProperty -Name 'DDC' -Value $ddc
                    if( $tags )
                    {
                        $device | Add-Member -MemberType NoteProperty -Name 'Tags' -Value ( $machine.Tags -join ',' )
                    }
                    break
                }
            }
        }
        $null = $devices.Add( $device )
    } 
}

## See if we have any devices from DDC machine list which are marked as being in PVS catalogues but not in our devices list so are orphans
if( ! $noOrphans )
{
    $machines.GetEnumerator() | ForEach-Object `
    {
        $ddc = $_.Key
        ## Cache machine catalogues so we can check provisioning type
        [hashtable]$catalogues = @{}
        Get-BrokerCatalog -AdminAddress $ddc | ForEach-Object { $catalogues.Add( $_.Name , $_ ) }

        ## Add to devices so we can display as much detail as possible if PVS provisioned
        ForEach( $machine in $_.Value )
        {
            $domainName,$machineName = $machine.MachineName -split '\\'
            if( [string]::IsNullOrEmpty( $machineName ) )
            {
                $machineName = $domainName
                $domainName = $null
            }
            if( ! [string]::IsNullOrEmpty( $name ) -and $machineName -notmatch $name )
            {
                continue
            }
            ## Now see if have this in devices in which case we ignore it - domain name in device record may be FQDN but domain from catalogue will not be (may also be missing in device)
            $device = $devices | Where-Object { $_.Name -eq $machineName -and ( ! $domainName -or ! $_.DomainName -or ( $domainName -eq ( $_.DomainName -split '\.' )[0] ) ) }
            if( ! $device )
            {
                ## Now check machine catalogues so if ProvisioningType = PVS then we will look to see if it an orphan
                $catalogue = $catalogues[ $machine.CatalogName  ]
                if( ! $catalogue -or $catalogue.ProvisioningType -match $provisioningType )
                {
                    $newItem = New-Object -TypeName PSCustomObject -Property `
                        (@{ 'Name' = ( $machine.MachineName -split '\\' )[-1] 
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
                        'Maintenance Mode' = $( if( $machine.InMaintenanceMode ) { 'On' } else { 'Off' } )
                        'User Sessions' = $machine.SessionCount ; } ) 

                    if( Get-Module ActiveDirectory -ErrorAction SilentlyContinue )
                    {
                        [hashtable]$adparams = @{}
                        if( ! [string]::IsNullOrEmpty( $ADgroups ) )
                        {
                            $adparams.Add( 'Properties' , 'MemberOf' )
                        }
                        $adAccount = $null
                        try
                        {
                            $adaccount = Get-ADComputer $newItem.Name -ErrorAction SilentlyContinue @adparams
                        }
                        catch
                        {
                        }
                        $newItem | Add-Member -MemberType NoteProperty -Name 'AD Account Exists' -value ( $adAccount -ne $null )

                        if( ! [string]::IsNullOrEmpty( $ADgroups ) )
                        {
                            $newItem | Add-Member -MemberType NoteProperty -Name 'AD Groups' -value ( ( $adAccount | select -ExpandProperty MemberOf | ForEach-Object { (( $_ -split '^CN=')[1] -split '\,')[0] } | Where-Object { $_ -match $ADgroups } ) -join ' ' )
                        }
                    }
                    if( ! $noBootTime )
                    {
                        $bootTime = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $newItem.Name | Select -ExpandProperty LastBootUpTime
                        if( $bootTime )
                        {
                            $newItem | Add-Member -MemberType NoteProperty -Name 'Boot Time' -value $bootTime
                        }
                        else
                        {
                            Write-Warning "Failed to get boot time for orphan $($newItem.Name)"
                        }
                    }
                    if( $tags )
                    {
                        $newItem | Add-Member -MemberType NoteProperty -Name 'Tags' -Value ( $machine.Tags -join ',' )
                    }
                    if( $dns )
                    {
                        [array]$ipv4Address = @( Resolve-DnsName -Name $newItem.Name -Type A )
                        $newItem | Add-Member -MemberType NoteProperty -Name 'IPv4 address' -Value ( ( $ipv4Address | Select -ExpandProperty IPAddress ) -join ' ' )
                    }

                    $null = $devices.Add( $newItem )
                }
            }
        }
    }
}

if( $devices -and $devices.Count )
{
    if( ! [string]::IsNullOrEmpty( $csv ) )
    {
        $devices | Select $columns | Sort Name | Export-Csv -Path $csv -NoTypeInformation -NoClobber
    }
    else
    {
        [hashtable]$params = @{}
        if( $noMenu )
        {
            $params.Add( 'Wait' , $true )
        }
        else
        {
            $params.Add( 'PassThru' , $true )
        }
        [string]$title = "$($devices.count) PVS devices via $($pvsServers -join ' ') & ddc $($ddcs -join ' ')"
        if( ! [string]::IsNullOrEmpty( $name ) )
        {
            $title += " matching `"$name`""
        }
        [array]$selected = @( $devices| Select $columns | Sort Name | Out-GridView -Title $title @Params )
        if( $selected -and $selected.Count )
        {
            $mainForm = Load-GUI $pvsDeviceActionerXAML

            if( ! $mainForm )
            {
                return
            }

            $WPFbtnRemoveFromDDC.Add_Click({ Perform-Action -action 'Remove From DDC' -form $mainForm })
            $WPFbtnRemoveFromPVS.Add_Click({ Perform-Action -action 'Remove From PVS' -form $mainForm })
            $WPFbtnMaintModeOff.Add_Click({ Perform-Action -action 'Maintenance Mode Off' -form $mainForm })
            $WPFbtnMaintModeOn.Add_Click({ Perform-Action -action 'Maintenance Mode Off'  -form $mainForm })
            $WPFbtnShutdown.Add_Click({ Perform-Action -action Shutdown -form $mainForm })
            $WPFbtnRestart.Add_Click({ Perform-Action -action Reboot -form $mainForm  })
            $WPFbtnBoot.Add_Click({ Perform-Action -action Boot -form $mainForm })
            $WPFbtnPowerOff.Add_Click({ Perform-Action -action 'Power Off' -form $mainForm  })
            $WPFbtnMessageUsers.Add_Click({ Perform-Action -action Message -form $mainForm })
            $WPFlstMachines.Items.Clear()
            $WPFlstMachines.ItemsSource = $selected
            ## Select all items since already selected them in grid view
            $WPFlstMachines.SelectAll()
            $null = $mainForm.ShowDialog()
        }
    }
}
else
{
    Write-Warning "No PVS devices found via $($pvsServers -join ' ')"
}
