#Requires -version 3.0

<#
    Grab all PVS devices and their collections, machine catalogues, delivery groups and more and output to csv or a grid view

    Use of this script is entirely at the risk of the user. The author accepts no liability or responsibility for any undesired issues it may cause
    
    Guy Leech, 2018

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
                    Added MaxRecordCount parameter for DDC

    25/02/18    GL  Added GUI options to remove from PVS and DDC

    27/02/18    GL  Added saving DDC and PVS servers to registry

    14/03/18    GL  Changed AD Account Exists to Created date

    16/03/18    GL  "Remove from AD" and "Remove from Hypervisor" (VMware) actions added
                    Added ability to identify orphaned VMs in VMware
                    Include VM information if have hypervisor connection (VMware)
                    Added progress bar

    17/03/18    GL  Completion message after actions completed with error count

    18/03/18    GL  Disable Remove from Hypervisor and AD buttons if not available
                    Added profiling and optimised
                    Made MessageBox calls app modal as well as system

    20/03/18    GL  Added remote memory, CPU and hard disk status via WMI/CIM
                    Added -help option

    21/03/18    GL  Changed action UI to use context menus rather than buttons
                    Added vdisk file size

    06/04/18    GL  Retry count, device IP and PVS server booted from fields added
                    -noGridview option added and will produce both csv and gridview if -csv specified unless -nogridview used
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

Path to a csv file that will have the results written to it.

.PARAMETER noGridView

Do not produce an on screen grid view containing the results. Default behaviour will display one.

.PARAMETER cpuSamples

The number of CPU usage samples to gather from a remote system if -noRemoting is not specified. The higher the number, the more accurate the figure but the longer it will take.
Set to zero to not gather CPU but still gather other remote data like memory usage and free disk space

.PARAMETER noRemoting

Will not try and contact active devices to gather information like last boot time. If WinRM not setup correctly or other issues mean remote calls will fail then this option can speed the script up

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

.PARAMETER noProgress

Do not show a progress bar

.PARAMETER hypervisors

Comma separated list of VMware vSphere servers to connect to in order to gather VM information such as memory and CPU configuration.
Will prompt for authentication if no pass thru or saved credentials

.PARAMETER vmPattern

When hypervisors are specified, or retrieved from the registry, a regular expression must be specified to match the VM names in vCenter to prevent all VMs
being included rather than just XenApp/XenDesktop ones

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

.PARAMETER profileCode

Output timings at various points to aid in finding slow parts

.PARAMETER help

Show full help via Get-Help cmdlet

.EXAMPLE

& '.\Get PVS device info.ps1' -pvsservers pvsprod01,pvstest01 -ddcs ddctest02,ddcprod03

Retrieve devices from the listed PVS servers, cross reference to the listed DDCs (list order does not matter) and display on screen in a sortable and filterable grid view

.EXAMPLE

& '.\Get PVS device info.ps1' -pvsservers pvsprod01 -ddcs ddcprod03 -name CTXUAT -tags -dns -csv h:\pvs.ctxuat.csv

Retrieve devices matching regular expression CTXUAT from the listed PVS server, cross reference to the listed DDC and output to the given csv file, including Citrix tag information and IPv4 address from DNS query

.EXAMPLE

& '.\Get PVS device info.ps1' -pvsservers pvsprod01 -ddcs ddcprod03 -ADGroups '^GRP-' -hypervisors vmware01 -vmPattern '^CTX[15]\d{3}$'

Retrieve all devices from the listed PVS server, cross reference to the listed DDC and VMware, including AD groups which start withg "GRP", and output to an on-screen grid view

.NOTES

Uses local PowerShell cmdlets for PVS, DDCs and VMware, as well as Active Directory, so run from a machine where both PVS and Studio consoles and the VMware PowerCLI are installed.

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
    [string[]]$hypervisors ,
    [string]$csv ,
    [switch]$dns ,
    [string]$name ,
    [switch]$tags ,
    [string]$ADgroups ,
    [switch]$noProgress ,
    [switch]$noRemoting ,
    [switch]$noMenu ,
    [switch]$noOrphans ,
    [switch]$noGridView ,
    [ValidateSet('PVS','MCS','Manual')]
    [string]$provisioningType = 'PVS' ,
    [string]$configRegKey = 'HKCU:\software\Guy Leech\PVS Fetcher' ,
    [string]$messageText ,
    [string]$messageCaption ,
    [int]$maxRecordCount = 2000 ,
    [int]$timeout = 60 ,
    [switch]$profileCode ,
    [int]$cpuSamples = 2 ,
    [string]$pvsShare ,
    [switch]$help ,
    [string[]]$snapins = @( 'Citrix.Broker.Admin.*'  ) ,
    [string[]]$modules = @( 'ActiveDirectory', "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll" ,'VMware.VimAutomation.Core'  ) 
)

if( $help )
{
    Get-Help -Name ( & { $myInvocation.ScriptName } ) -Full
    return
}

$columns = [System.Collections.ArrayList]( @( 'Name','DomainName','Description','PVS Server','DDC','SiteName','CollectionName','Machine Catalogue','Delivery Group','Registration State','Maintenance_Mode','User_Sessions','devicemac','active','enabled',
    'Store Name','Disk Version Access','Disk Version Created','AD Account Created','AD Last Logon','AD Description','Disk Name','Booted off vdisk','Booted Disk Version','Vdisk Production Version','Vdisk Latest Version','Latest Version Description','Override Version',
    'Retries','Booted Off','Device IP','Booted off latest','Disk Description','Cache Type','Disk Size (GB)','vDisk Size (GB)','Write Cache Size (MB)' )  )

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

if( ! $noRemoting )
{
    $null = $columns.Add( 'Boot_Time' )
    $null = $columns.Add( 'Available Memory (GB)' )
    $null = $columns.Add( 'Committed Memory %' )
    $null = $columns.Add( 'Free disk space %' )
    if( $cpuSamples -gt 0 )
    {
        $null = $columns.Add( 'CPU Usage %' )
    }
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
        <TextBox x:Name="txtMessageCaption" HorizontalAlignment="Left" Height="53" Margin="85,20,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="180"/>
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

$pvsDeviceActionerXAML = @'
<Window x:Name="formDeviceActioner" x:Class="PVSDeviceViewerActions.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PVSDeviceViewerActions"
        mc:Ignorable="d"
        Title="PVS Device Actioner" Height="522.565" Width="401.595">
    <Grid Margin="0,0,-20.333,-111.667">
        <ListView x:Name="lstMachines" HorizontalAlignment="Left" Height="452" Margin="23,22,0,0" VerticalAlignment="Top" Width="346">
            <ListView.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="VMware" Name="VMwareContextMenu" >
                        <MenuItem Header="Power On" Name="VMwarePowerOnContextMenu" />
                        <MenuItem Header="Power Off" Name="VMwarePowerOffContextMenu" />
                        <MenuItem Header="Restart" Name="VMwareRestartContextMenu" />
                        <MenuItem Header="Delete" Name="VMwareDeleteContextMenu" />
                    </MenuItem>
                    <MenuItem Header="PVS" Name="PVSContextMenu">
                        <MenuItem Header="Boot" Name="PVSBootContextMenu" />
                        <MenuItem Header="Shutdown" Name="PVSShutdownContextMenu" />
                        <MenuItem Header="Restart" Name="PVSRestartContextMenu" />
                        <MenuItem Header="Delete" Name="PVSDeleteContextMenu" />
                    </MenuItem>
                    <MenuItem Header="DDC" Name="DDCContextMenu" >
                        <MenuItem Header="Maintenance Mode On" Name="DDCMaintModeOnContextMenu" />
                        <MenuItem Header="Maintenance Mode Off" Name="DDCMaintModeOffContextMenu" />
                        <MenuItem Header="Message Users" Name="DDCMessageUsersContextMenu" />
                        <MenuItem Header="Delete" Name="DDCDeleteContextMenu" />
                    </MenuItem>
                    <MenuItem Header="AD" Name="ADContextMenu">
                        <MenuItem Header="Delete" Name="ADDeleteContextMenu" />
                    </MenuItem>
                    <MenuItem Header="Windows" Name="WindowsContextMenu">
                        <MenuItem Header="Shutdown" Name="WinShutdownModeOnContextMenu" />
                        <MenuItem Header="Restart" Name="WinRestartContextMenu" />
                    </MenuItem>
                </ContextMenu>
            </ListView.ContextMenu>
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Device" DisplayMemberBinding ="{Binding 'Name'}" />
                    <GridViewColumn Header="Boot Time" DisplayMemberBinding ="{Binding 'Boot_Time'}" />
                    <GridViewColumn Header="Users" DisplayMemberBinding ="{Binding 'User_Sessions'}" />
                    <GridViewColumn Header="Maintenance Mode" DisplayMemberBinding ="{Binding 'Maintenance_Mode'}" />
                </GridView>
            </ListView.View>
            <ListBoxItem Content="Device"/>
        </ListView>
    </Grid>
</Window>

'@

<#
 @"
<Window x:Name="formDeviceActioner" x:Class="PVSDeviceViewerActions.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PVSDeviceViewerActions"
        mc:Ignorable="d"
        Title="PVS Device Actioner" Height="522.565" Width="401.595">
    <Grid Margin="0,0,-20.333,-111.667">
        <ListView x:Name="lstMachines" HorizontalAlignment="Left" Height="452" Margin="23,22,0,0" VerticalAlignment="Top" Width="140">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Device" DisplayMemberBinding ="{Binding 'Name'}" />
                </GridView>
            </ListView.View>
            <ListBoxItem Content="Device"/>
        </ListView>
        <StackPanel x:Name="stkButtons" Margin="198,22,36,28" Orientation="Vertical">
            <Button x:Name="btnShutdown" Content="_Shutdown" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnPowerOff" Content="_Power Off" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnRestart" Content="_Restart" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnBoot" Content="_Boot" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnMessageUsers" Content="_Message Users" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnMaintModeOn" Content="Maintenance Mode O_n" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnMaintModeOff" Content="Maintenance Mode O_ff" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnRemoveFromDDC" Content="Remove from _DDC" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnRemoveFromPVS" Content="Remove from P_VS" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnRemoveFromAD" Content="Remove from _AD" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
            <Button x:Name="btnRemoveFromHypervisor" Content="Remove from _Hypervisor" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="149" Margin="0 0 0 15" />
        </StackPanel>
    </Grid>
</Window>

"@
#>

## Adding so we can make it app modal as well as system
Add-Type @'
using System;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
    public static class Windows
    {
        [DllImport("user32.dll")]
        public static extern int MessageBox(int hWnd, String text, String caption, uint type);
    }
}
'@


Add-Type -TypeDefinition @'
   public enum MessageBoxReturns
   {
      IDYES = 6,
      IDNO = 7 ,
      IDOK = 1 ,
      IDABORT = 3,
      IDCANCEL = 2 ,
      IDCONTINUE = 11 ,
      IDIGNORE = 5 ,
      IDRETRY = 4 ,
      IDTRYAGAIN = 10
   }
'@

Function Show-Profiling( [string]$info , [int]$lineNumber , $timer , [bool]$profileCode )
{
    if( $profileCode )
    {
        "{0}:{1}:{2}" -f $timer.ElapsedMilliSeconds , $lineNumber , $info
    }
}

Function Save-ConfigToRegistry( [string]$serverSet = 'Default' , [string[]]$DDCs , [string[]]$PVSServers , [string[]]$hypervisors )
{
    [string]$key = Join-Path $configRegKey $serverSet
    if( ! ( Test-Path $key -ErrorAction SilentlyContinue ) )
    {
        $null = New-Item -Path $key -Force
    }
    Set-ItemProperty -Path $key -Name 'DDC' -Value $DDCs
    Set-ItemProperty -Path $key -Name 'PVS' -Value $PVSServers
    Set-ItemProperty -Path $Key -Name 'Hypervisor' -Value $hypervisors
}   

Function Get-ConfigFromRegistry
{
    Param
    (
        [string]$serverSet = 'Default' , 
        [ref]$DDCs , 
        [ref]$PVSServers ,
        [ref]$hypervisors
    )
    [string]$key = Join-Path $configRegKey $serverSet
    if( ! ( Test-Path $key -ErrorAction SilentlyContinue ) )
    {
        Write-Warning "Config registry key `"$key`" does not exist"
    }
    $hypervisors.value = Get-ItemProperty -Path $key -Name 'Hypervisor' -ErrorAction SilentlyContinue | select -ExpandProperty 'Hypervisor'
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
        Write-Error "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        return $null
    }
 
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object `
    {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Script
    }

    return $form
}

Function Display-MessageBox( $window , $text , $caption , [System.Windows.MessageBoxButton]$buttons , [System.Windows.MessageBoxImage]$icon )
{
    if( $window -and $window.Handle )
    {
        [int]$modified = switch( $buttons )
            {
                'OK' { [System.Windows.MessageBoxButton]::OK }
                'OKCancel' { [System.Windows.MessageBoxButton]::OKCancel }
                'YesNo' { [System.Windows.MessageBoxButton]::YesNo }
                'YesNoCancel' { [System.Windows.MessageBoxButton]::YesNo }
            }
        [int]$choice = [PInvoke.Win32.Windows]::MessageBox( $Window.handle , $text , $caption , ( ( $icon -as [int] ) -bor $modified ) )  ## makes it app modal so UI blocks
        switch( $choice )
        {
            ([MessageBoxReturns]::IDYES -as [int]) { 'Yes' }
            ([MessageBoxReturns]::IDNO -as [int]) { 'No' }
            ([MessageBoxReturns]::IDOK -as [int]) { 'Ok' } 
            ([MessageBoxReturns]::IDABORT -as [int]) { 'Abort' } 
            ([MessageBoxReturns]::IDCANCEL -as [int]) { 'Cancel' } 
            ([MessageBoxReturns]::IDCONTINUE -as [int]) { 'Continue' } 
            ([MessageBoxReturns]::IDIGNORE -as [int]) { 'Ignore' } 
            ([MessageBoxReturns]::IDRETRY -as [int]) { 'Retry' } 
            ([MessageBoxReturns]::IDTRYAGAIN -as [int]) { 'TryAgain' } 
        }       
    }
    else
    {
        [Windows.MessageBox]::Show( $text , $caption , $buttons , $icon )
    }
}

Function Perform-Action( [string]$action , $form )
{
    $_.Handled = $true

    ## Get HWND so we can make app modal dialogues
    $thisWindow = [System.Windows.Interop.WindowInteropHelper]::new($form)

    $answer = Display-MessageBox -text "Are you sure you want to $action these $($WPFlstMachines.SelectedItems.Count) devices ?" -caption 'Confirm' -buttons YesNo -icon Question -window $thisWindow

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

    [int]$errors = 0
    ForEach( $device in $WPFlstMachines.SelectedItems )
    {
        Write-Verbose "Action $action on $($device.Name)"

        switch -regex ( $action )
        {
            'Message' { Get-BrokerSession -AdminAddress $device.ddc -MachineName ( $env:USERDOMAIN + '\' +  $device.Name ) | Send-BrokerSessionMessage -AdminAddress $device.ddc -Title $WPFtxtMessageCaption.Text -Text $WPFtxtMessageBody.Text -MessageStyle ($WPFcomboMessageStyle.SelectedItem.Content)  ;break }
            'Remove From AD' { Remove-ADComputer -Identity $device.Name -Confirm:$False ;break }
            'Remove From DDC' { Remove-BrokerMachine -Force -AdminAddress $device.ddc -MachineName $( if( [string]::IsNullOrEmpty( $device.DomainName ) ) { $device.Name } else {  $device.DomainName + '\' +  $device.Name } ) ; break  }
            'Remove From PVS' { Set-PvsConnection -Server $device.'PVS Server'; Remove-PvsDevice -DeviceName $device.Name ; break }
            'Maintenance Mode On'  { Set-BrokerMachine -AdminAddress $device.ddc -InMaintenanceMode $true  -MachineName ( $device.DomainName + '\' +  $device.Name ) ; break }
            'Maintenance Mode Off' { Set-BrokerMachine -AdminAddress $device.ddc -InMaintenanceMode $false -MachineName ( $device.DomainName + '\' +  $device.Name ) ; break }
            'Reboot' { Restart-Computer -ComputerName $device.Name ; break }
            'Shutdown' { Stop-Computer -ComputerName $device.Name ; break }
            'Remove from Hypervisor' { Get-VM -Name $device.Name | Remove-VM -DeletePermanently -Confirm:$false ;break }
            'VMware Boot' { Get-VM -Name $device.name | Start-VM -Confirm:$false  ;break }
            'VMware Power Off' { Get-VM -Name $device.name | Stop-VM -Confirm:$false ;break  }
            'VMware Restart' {  Get-VM -Name $device.name | Restart-VM -Confirm:$false ;break }
            'PVS Boot|PVS Power Off|PVS Restart' ` ##  Can't use New-BrokerHostingPowerAction as may not be known to DDC
            { 
                Set-PvsConnection -Server $device.'PVS Server'
                if( $_ -match 'Boot' )
                {
                    if( [string]::IsNullOrEmpty( $device.'Disk Name' ) )
                    {
                        $answer = Display-MessageBox -window $thisWindow -text "$($device.Name) has no vdisk assigned so may not boot - continue ?" -caption 'Confirm' -buttons YesNo -icon Question

                        if( $answer -ne 'Yes' )
                        {
                            continue
                        }
                    }
                    $thePvsTask = Start-PvsDeviceBoot -DeviceName $device.Name
                }
                elseif( $_ -match 'Off' )
                {
                    $thePvsTask = Start-PvsDeviceShutdown -DeviceName $device.Name
                }
                elseif( $_ -match 'Restart' )
                {
                    $thePvsTask = Start-PvsDeviceReboot -DeviceName $device.Name
                }
                $timer = [Diagnostics.Stopwatch]::StartNew()
                [bool]$timedOut = $false
                while ( $thePvsTask -and $thePvsTask.State -eq 0 ) 
                {
                    $percentFinished = Get-PvsTaskStatus -Object $thePvsTask 
                    if( ! $percentFinished -or $percentFinished.ToString() -ne 100 )
                    {
                        Start-Sleep -timer 500
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
                    Display-MessageBox -window $thisWindow -text "Failed to perform action on $($device.Name) - timed out after $timeout seconds" -caption $action -buttons OK -icon Error
                    $errors++
                } 
                elseif ( ! $thePvsTask -or $thePvsTask.State -ne 2)
                {
                    Display-MessageBox -window $thisWindow -text "Failed to perform action on $($device.Name)" -caption $action -buttons OK -icon Error
                    $errors++
                }
             }
            default { Write-Warning "Unknown command `"$action`"" }
        }
        if( ! $? )
        {
            $errors++
        }
    }
    
    if( $form )
    {
        $form.Cursor = $oldCursor
    }

    [string]$status =  [System.Windows.MessageBoxImage]::Information

    if( $errors )
    {
        $status = [System.Windows.MessageBoxImage]::Error
    }
    
    Display-MessageBox -window $thisWindow -text "$errors / $($WPFlstMachines.SelectedItems.Count) errors" -caption "Finished $action" -buttons OK -icon $status
}

Function Get-CurrentLineNumber
{ 
    $MyInvocation.ScriptLineNumber 
}

Function Get-ADMachineInfo
{
    Param
    (
        [string]$name ,
        [hashtable]$adparams ,
        [string]$adGroups ,
        $item
    )
    
    if( $item -and ( Get-Module ActiveDirectory -ErrorAction SilentlyContinue ) )
    {
        try
        {
            Show-Profiling -Info "Getting AD group info" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
            $adaccount = Get-ADComputer $item.Name -ErrorAction SilentlyContinue @adparams
            [string]$groups = $null
            if( ! [string]::IsNullOrEmpty( $ADgroups ) )
            {
               $groups = ( ( $adAccount | select -ExpandProperty MemberOf | ForEach-Object { (( $_ -split '^CN=')[1] -split '\,')[0] } | Where-Object { $_ -match $ADgroups } ) -join ' ' )
            }
            Add-Member -InputObject $item -NotePropertyMembers `
            @{
                'AD Account Created' = $adAccount.Created
                'AD Last Logon' = $adAccount.LastLogonDate
                'AD Description' = $adAccount.Description
                'AD Groups' = $groups
            }
        }
        catch
        {
        }
    }
}

Function Get-RemoteInfo( [string]$computer , [int]$cpuSamples )
{
    [scriptblock]$remoteWork = `
    {
        $osinfo = Get-CimInstance Win32_OperatingSystem
        $logicalDisks = Get-CimInstance -ClassName Win32_logicaldisk -Filter 'DriveType = 3'
        $cpu = $(if( $using:cpuSamples -gt 0 ) { [math]::Round( ( 1..$usage:cpuSamples | ForEach-Object { Get-Counter -Counter '\Processor(*)\% Processor Time'|select -ExpandProperty CounterSamples| Where-Object { $_.InstanceName -eq '_total' }|select -ExpandProperty CookedValue } | Measure-Object -Average ).Average , 1 ) })
        $osinfo,$logicalDisks,$cpu
    }
    try
    {
        $osinfo,$logicalDisks,$cpu = Invoke-Command -ComputerName $computer -ScriptBlock $remoteWork
        @{
            'Boot_Time' = $osinfo.LastBootUpTime
            'Available Memory (GB)' = [Math]::Round( $osinfo.FreePhysicalMemory / 1MB , 1 )
            'Committed Memory %' = 100 - [Math]::Round( ( $osinfo.FreeVirtualMemory / $osinfo.TotalVirtualMemorySize ) * 100 , 1 )
            'CPU Usage %' = $cpu
            'Free disk space %' = ( $logicalDisks | Sort DeviceID | ForEach-Object { [Math]::Round( ( $_.FreeSpace / $_.Size ) * 100 , 1 ) }) -join ' '
        }
    }
    catch
    {
        Write-Error "Failed to get remote info from $computer : $($_.ToString())"
        $null
    }
}

if( $noProgress )
{
    $ProgressPreference = 'SilentlyContinue'
}

if( $registry )
{
    Get-ConfigFromRegistry -serverSet $serverSet -DDCs ( [ref] $ddcs ) -PVSServers ( [ref] $pvsServers ) -hypervisors ( [ref] $hypervisors )
    if( ! $ddcs -or ! $ddcs.Count -or ! $pvsServers -or ! $pvsServers.Count )
    {
        Write-Warning "Failed to get PVS and/or DDC servers from registry key `"$configRegKey`" for server set `"$serverSet`""
        return
    }
}
elseif( $save )
{
    Save-ConfigToRegistry -serverSet $serverSet  -DDCs $ddcs -PVSServers $pvsServers -hypervisors $hypervisors
}

Write-Progress -Activity "Caching information" -PercentComplete 0

## Get all information from DDCs so we can lookup locally
[hashtable]$machines = @{}

ForEach( $ddc in $ddcs )
{
    $machines.Add( $ddc , [System.Collections.ArrayList] ( Get-BrokerMachine -AdminAddress $ddc -MaxRecordCount $maxRecordCount -ErrorAction SilentlyContinue ) )
}

## Make a hashtable so we can index quicker when cross referencing to DDC & VMware
[hashtable]$devices = @{}
#$devices = New-Object -TypeName System.Collections.ArrayList

[hashtable]$vms = @{}

if( $hypervisors -and $hypervisors.Count )
{
    if( [string]::IsNullOrEmpty( $name ) )
    {
        Write-Error "Must specify a VM name pattern via -name when cross referencing to VMware"
        return
    }

    Write-Progress -Activity "Connecting to hypervisors $($hypervisors -split ' ')" -PercentComplete 1

    $null = Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
    if( Connect-VIserver -Server $hypervisors )
    {
        $null = $columns.Add( 'CPUs')
        $null = $columns.Add( 'Memory (GB)')
        $null = $columns.Add( 'Hard Drives (GB)')
        $null = $columns.Add( 'NICs')
        $null = $columns.Add( 'Hypervisor')

        ## Cache all VMs for efficiency
        Get-VM | Where-Object { $_.Name -match $name } | ForEach-Object `
        {
            $vms.Add( $_.Name , $_ )
        }
        Write-Verbose "Got $($vms.Count) vms matching `"$name`" from $($hypervisors -split ' ')"
    }
}

if( $PSVersionTable.PSVersion.Major -lt 5 -and $columns.Count -gt 30 -and [string]::IsNullOrEmpty( $csv ) )
{
    Write-Warning "This version of PowerShell limits the number of columns in a grid view to 30 and we have $($columns.Count) so those from `"$($columns[30])`" will be lost in grid view"
}

[int]$pvsServerCount = 0

[hashtable]$adparams = @{ 'Properties' = @( 'Created' , 'LastLogonDate' , 'Description' )  }
if( ! [string]::IsNullOrEmpty( $ADgroups ) )
{
    $adparams[ 'Properties' ] +=  'MemberOf' 
}

if( $profileCode )
{
    $profiler = [Diagnostics.Stopwatch]::new()
}
else
{
    $profiler = $null
}

ForEach( $pvsServer in $pvsServers )
{
    $pvsServerCount++
    Set-PvsConnection -Server $pvsServer 

    if( ! $? )
    {
        Write-Output "Cannot connect to PVS server $pvsServer - aborting"
        continue
    }

    ## Cache latest production version for vdisks so don't look up for every device
    [hashtable]$diskVersions = @{}

    ## Get Device info in one go as quite slow
    [hashtable]$deviceInfos = @{}
    Get-PvsDeviceInfo | ForEach-Object `
    {
        $deviceInfos.Add( $_.DeviceId , $_ )
    }

    ## Cache store locations so we can look up vdisk sizes
    [hashtable]$stores = @{}
    Get-PvsStore | ForEach-Object `
    {
        $stores.Add( $_.StoreName , $_.Path )
    }

    ## Get all devices so we can do progress
    $pvsDevices = @( Get-PvsDevice | Where-Object { $_.Name -match $name })
    [decimal]$eachDevicePercent = 100 / [Math]::Max( $pvsDevices.Count , 1 ) ## avoid divide by zero if no devices found
    [int]$counter = 0

    # Get all sites that we can see on this server and find all devices and cross ref to Citrix for catalogues and delivery groups
    $pvsDevices | ForEach-Object `
    {
        if( $profileCode )
        {
            $profiler.Restart()
        }
        $counter++
        $device = $_
        [decimal]$percentComplete = $counter * $eachDevicePercent
        Write-Progress -Activity "Processing $($pvsDevices.Count) devices from PVS server $pvsServer" -Status "$($device.name)" -PercentComplete $percentComplete

        [int]$bootVersion = -1
        $vDisk = Get-PvsDiskInfo -DeviceId $_.DeviceId
        [hashtable]$fields = @{}

        Show-Profiling -Info "Got disk Info" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
        if( $vms -and $vms.count )
        {
            Show-Profiling -Info "Getting VMware info" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
            $vm = $vms[ $device.Name ]
            if( $vm )
            {
                $fields += @{
                    'CPUs' = $vm.NumCpu 
                    'Memory (GB)' = $vm.MemoryGB
                    'Hard Drives (GB)' = $( ( Get-HardDisk -VM $vm -ErrorAction SilentlyContinue | sort CapacityGB | select -ExpandProperty CapacityGB ) -join ' ' )
                    'NICS' = $( ( Get-NetworkAdapter -VM $vm -ErrorAction SilentlyContinue | Sort Type | Select -ExpandProperty Type ) -join ' ' )
                    'Hypervisor' = $vm.VMHost
                }
            }
        }

        if( $device.Active -and ! $noRemoting )
        {
            Show-Profiling -Info "Getting remote info" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
            $remoteInfo = Get-RemoteInfo -computer $device.Name -cpuSamples $cpuSamples
            if( $remoteInfo )
            {
                $fields += $remoteInfo
            }
            Show-Profiling -Info "Got remote info" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
        }
        $fields.Add( 'PVS Server' , $pvsServer )
        $versions = $null
        if( $vdisk )
        {
            $fields += @{
                'Disk Name' = $vdisk.Name
                'Store Name' = $vdisk.StoreName
                'Disk Description' = $vdisk.Description
                'Cache Type' = $cacheTypes[$vdisk.WriteCacheType]
                'Disk Size (GB)' = ([math]::Round( $vdisk.DiskSize / 1GB , 2 ))
                'Write Cache Size (MB)' = $vdisk.WriteCacheSize }
            ## Cache vdisk version info to reduce PVS server hits
            $versions = $diskVersions[ $vdisk.DiskLocatorId ]
            if( ! $versions )
            { 
                try
                {
                    Show-Profiling -Info "Getting disk version info" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
                    ## Can't pre-cache since can only retrieve per disk
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
                $vdiskFile = $null
                $latestProduction = $versions | Where-Object { $_.Access -eq 0 } | Sort Version -Descending | Select -First 1 
                if( $latestProduction )
                {
                    $vdiskFile = $latestProduction.DiskFileName
                    $latestProductionVersion = $latestProduction.Version
                }
                else
                {
                    $latestProductionVersion = $null
                }
                if( $override )
                {
                    $bootVersion = $override.Version
                    $vdiskFile = $override.DiskFileName
                }
                else
                {
                    ## Access: Read-only access of the Disk Version. Values are: 0 (Production), 1 (Maintenance), 2 (MaintenanceHighestVersion), 3 (Override), 4 (Merge), 5 (MergeMaintenance), 6 (MergeTest), and 7 (Test) Min=0, Max=7, Default=0
                    $bootVersion = $latestProductionVersion
                }
                if( $vdiskFile)
                {
                    ## Need to see if Store path is local to the PVS server and if so convert to a share so we can get vdisk file info
                    if( $stores[ $vdisk.StoreName ] -match '^([A-z]):(.*$)' )
                    {
                        if( [string]::IsNullOrEmpty( $pvsShare ) )
                        {
                            $vdiskfile = Join-Path ( Join-Path ( '\\' + $pvsServer + '\' + "$($Matches[1])`$"  ) $Matches[2] ) $vdiskFile ## assume regular admin share
                        }
                        else
                        {
                            $vdiskfile = Join-Path ( Join-Path ( '\\' + $pvsServer + '\' + $pvsShare ) ) $vdiskFile
                        }
                    }
                    else
                    {
                        $vdiskFile = Join-Path $stores[ $vdisk.StoreName ] $vdiskFile
                    }
                    if( ( Test-Path $vdiskFile -ErrorAction SilentlyContinue ) )
                    {
                        $fields += @{ 'vDisk Size (GB)' = [math]::Round( (Get-ItemProperty -Path $vdiskFile).Length / 1GB ) }
                    }
                    else
                    {
                        Write-Warning "Could not find disk `"$vdiskFile`" for $($device.name)"
                    }
                }
                if( $latestProductionVersion -eq $null -and $override )
                {
                    ## No production version, only an override so this must be the latest production version
                    $latestProductionVersion = $override.Version
                }
                $fields += @{
                    'Override Version' = $( if( $override ) { $bootVersion } else { $null } ) 
                    'Vdisk Latest Version' = $latestProductionVersion 
                    'Latest Version Description' = $( $versions | Where-Object { $_.Version -eq $latestProductionVersion } | Select -ExpandProperty Description )  
                }      
            }
            $fields.Add( 'Vdisk Production Version' ,$bootVersion )
        }
        
        Show-Profiling -Info "Getting device info" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
        $deviceInfo = $deviceInfos[ $device.DeviceId ]
        if( $deviceInfo )
        {
            $fields.Add( 'Disk Version Access' , $accessTypes[ $deviceInfo.DiskVersionAccess ] )
            $fields.Add( 'Booted Off' , $deviceInfo.ServerName )
            $fields.Add( 'Device IP' , $deviceInfo.IP )
            if( ! [string]::IsNullOrEmpty( $deviceInfo.Status ) )
            {
                $fields.Add( 'Retries' , ($deviceInfo.Status -split ',')[0] -as [int] ) ## scond value is supposedly RAM cache used percent but I've not seen it set
            }
            if( $device.Active )
            {
                ## Check if booting off the disk we should be as previous info is what is assigned, not what is necessarily being used (e.g. vdisk changed for device whilst it is booted)
                $bootedDiskName = (( $diskVersions[ $deviceInfo.DiskLocatorId ] | Select -First 1 | Select -ExpandProperty Name ) -split '\.')[0]
                $fields.Add( 'Booted Disk Version' , $deviceInfo.DiskVersion )
                if( $bootVersion -ge 0 )
                {
                    Write-Verbose "$($device.Name) booted off $bootedDiskName, disk configured $($vDisk.Name)"
                    $fields.Add( 'Booted off latest' , ( $bootVersion -eq $deviceInfo.DiskVersion -and $bootedDiskName -eq $vdisk.Name ) )
                    $fields.Add( 'Booted off vdisk' , $bootedDiskName )
                }
            }
            if( $versions )
            {
                try
                {
                    $fields.Add( 'Disk Version Created' ,( $versions | Where-Object { $_.Version -eq $deviceInfo.DiskVersion } | select -ExpandProperty CreateDate ) )
                }
                catch
                {
                    $_
                }
            }
        }
        else
        {
            Write-Warning "Failed to get PVS device info for id $($device.DeviceId) $($device.Name)"
        }

        Get-ADMachineInfo -name $device.Name -adparams $adparams -adGroups $ADgroups -item $device

        if( $device.Active -and $dns )
        {
            Show-Profiling -Info "Resolving DNS name" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
            [array]$ipv4Address = @( Resolve-DnsName -Name $device.Name -Type A )
            $fields.Add( 'IPv4 address' , ( ( $ipv4Address | Select -ExpandProperty IPAddress ) -join ' ' ) )
        }
        
        Show-Profiling -Info "Getting DDC info" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
        if( ( Get-Command -Name Get-BrokerMachine -ErrorAction SilentlyContinue ) )
        {
            ## Need to find a ddc that will return us information on this device
            ForEach( $ddc in $ddcs )
            {
                ## can't use HostedMachineName as only populated if registered
                $machine = $machines[ $ddc ] | Where-Object { $_.MachineName -eq  ( ($device.DomainName -split '\.')[0] + '\' + $device.Name ) }
                if( $machine )
                {
                    $fields += @{
                        'Machine Catalogue' = $machine.CatalogName
                        'Delivery Group' = $machine.DesktopGroupName
                        'Registration State' = $machine.RegistrationState
                        'User_Sessions' = $machine.SessionCount
                        'Maintenance_Mode' = $( if( $machine.InMaintenanceMode ) { 'On' } else { 'Off' } )
                        'DDC' = $ddc
                    }
                    if( $tags )
                    {
                        $fields.Add( 'Tags' , ( $machine.Tags -join ',' ) )
                    }
                    break
                }
            }
        }
        Add-Member -InputObject $device -NotePropertyMembers $fields
        try
        {
            $devices.Add( $device.Name , $device )
        }
        catch
        {
            Write-Warning "Duplicate device name $($device.Name) found"
        }
        if( $profileCode )
        {
            Show-Profiling -Info "End of loop" -lineNumber (Get-CurrentLineNumber) -timer $profiler -profileCode $profileCode
            $profiler.Stop()
        }
    } 
}

## See if we have any devices from DDC machine list which are marked as being in PVS catalogues but not in our devices list so are orphans
if( ! $noOrphans )
{
    $machines.GetEnumerator() | ForEach-Object `
    {
        $ddc = $_.Key
        Write-Progress -Activity "Checking for orphans on DDC $ddc" -PercentComplete 98

        ## Cache machine catalogues so we can check provisioning type
        [hashtable]$catalogues = @{}
        Get-BrokerCatalog -AdminAddress $ddc | ForEach-Object { $catalogues.Add( $_.Name , $_ ) }

        ## Add to devices so we can display as much detail as possible if PVS provisioned
        $_.Value | ForEach-Object `
        {
            $machine = $_
            $domainName,$machineName = $machine.MachineName -split '\\'
            if( [string]::IsNullOrEmpty( $machineName ) )
            {
                $machineName = $domainName
                $domainName = $null
            }
            if( [string]::IsNullOrEmpty( $name ) -or $machineName -match $name )
            {
                ## Now see if have this in devices in which case we ignore it - domain name in device record may be FQDN but domain from catalogue will not be (may also be missing in device)
                #$device = $devices | Where-Object { $_.Name -eq $machineName -and ( ! $domainName -or ! $_.DomainName -or ( $domainName -eq ( $_.DomainName -split '\.' )[0] ) ) }
                $device = $devices[ $machineName ]
                if( $device ) ## check same domain
                {
                    if( $domainName -and $device.DomainName -and $domainName -ne ( $device.DomainName -split '\.' )[0] )
                    {
                        $device = $null ## doesn't quite match
                    }
                }
                if( ! $device )
                {
                    ## Now check machine catalogues so if ProvisioningType = PVS then we will look to see if it an orphan
                    $catalogue = $catalogues[ $machine.CatalogName  ]
                    if( ! $catalogue -or $catalogue.ProvisioningType -match $provisioningType )
                    {
                        $newItem = [pscustomobject]@{ 
                            'Name' = ( $machine.MachineName -split '\\' )[-1] 
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
                            'Maintenance_Mode' = $( if( $machine.InMaintenanceMode ) { 'On' } else { 'Off' } )
                            'User_Sessions' = $machine.SessionCount ; }

                        Get-ADMachineInfo -name $newItem.Name -adparams $adparams -adGroups $ADgroups -item $newItem

                        if( ! $noRemoting )
                        {
                            Add-Member -InputObject $newItem -NotePropertyMembers ( Get-RemoteInfo -computer $newItem.Name -cpuSamples $cpuSamples )
                        }
                        if( $tags )
                        {
                            Add-Member -InputObject $newItem  -MemberType NoteProperty -Name 'Tags' -Value ( $machine.Tags -join ',' )
                        }
                        if( $dns )
                        {
                            [array]$ipv4Address = @( Resolve-DnsName -Name $newItem.Name -Type A )
                            Add-Member -InputObject $newItem  -MemberType NoteProperty -Name 'IPv4 address' -Value ( ( $ipv4Address | Select -ExpandProperty IPAddress ) -join ' ' )
                        }

                        $devices.Add( $newItem.Name , $newItem )
                    }
                }
            }
        }
    }
    ## if we have VMware details then get those VMs and add if not present here
    if( $hypervisors -and $hypervisors.Count )
    {
        ## will already be connected as have already grabbed VMs
        Write-Progress -Activity "Checking for orphans on hypervisor $($hypervisors -split ' ')" -PercentComplete 99

        [int]$vmCount = 0
        $vms.GetEnumerator() | ForEach-Object `
        {
            $vmwareVM = $_.Value
            $vmCount++
            $existingDevice = $devices[ $vmwareVM.Name ]
            ## Now have to see if we have restricted the PVS device retrieval via -name making $devices a subset of all PVS devices
            if( ! $existingDevice -and ! [string]::IsNullOrEmpty( $name ) )
            {
                $existingDevice = $vmwareVM.Name -notmatch $name
            }
            if( ! $existingDevice )
            {
                $newItem = [pscustomobject]@{ 
                    'Name' = $vmwareVM.Name
                    'Description' = $vmwareVM.Notes
                    'CPUs' = $vmwareVM.NumCpu 
                    'Memory (GB)' = $vmwareVM.MemoryGB
                    'Hard Drives (GB)' = $( ( Get-HardDisk -VM $vmwareVM -ErrorAction SilentlyContinue | sort CapacityGB | select -ExpandProperty CapacityGB ) -join ' ' )
                    'NICS' = $( ( Get-NetworkAdapter -VM $vmwareVM -ErrorAction SilentlyContinue | Sort Type | Select -ExpandProperty Type ) -join ' ' )
                    'Hypervisor' = $vmwareVM.VMHost
                    'Active' = $( if($vmwareVM.PowerState -eq 'PoweredOn') { $true } else { $false } )
                }
                Get-ADMachineInfo -name $newItem.Name -adparams $adparams -adGroups $ADgroups -item $newItem

                if( $vmwareVM.PowerState -eq 'PoweredOn' )
                {
                    if( ! $noRemoting )
                    {
                        Add-Member -InputObject $newItem -NotePropertyMembers ( Get-RemoteInfo -computer $newItem.Name -cpuSamples $cpuSamples )
                    }
                    if( $dns )
                    {
                        [array]$ipv4Address = @( Resolve-DnsName -Name $newItem.Name -Type A )
                        Add-Member -InputObject $newItem  -MemberType NoteProperty -Name 'IPv4 address' -Value ( ( $ipv4Address | Select -ExpandProperty IPAddress ) -join ' ' )
                    }
                }

                $devices.Add( $newItem.Name , $newItem )
            }
            if( ! $vmCount )
            {
                Write-Warning "Found no VMs on $($hypervisors -split ',') matching regex `"$name`""
            }
        }
    }
}

Write-Progress -Activity 'Finished' -Completed -PercentComplete 100

if( $devices -and $devices.Count )
{
    if( ! [string]::IsNullOrEmpty( $csv ) )
    {
        $devices.GetEnumerator() | ForEach-Object { $_.Value } | Select $columns | Sort Name | Export-Csv -Path $csv -NoTypeInformation -NoClobber
    }
    if( ! $noGridView )
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
        [array]$selected = @( $devices.GetEnumerator() | ForEach-Object { $_.Value } | Select $columns | Sort Name | Out-GridView -Title $title @Params )
        if( $selected -and $selected.Count )
        {
            $mainForm = Load-GUI $pvsDeviceActionerXAML

            if( ! $mainForm )
            {
                return
            }
            
            $mainForm.Title += " - $($selected.Count) devices"

            if( $hypervisors -and $hypervisors.Count )
            {
                $WPFVMwarePowerOnContextMenu.Add_Click({ Perform-Action -action 'VMware Boot' -form $mainForm })
                $WPFVMwarePowerOffContextMenu.Add_Click({ Perform-Action -action 'VMware Power Off' -form $mainForm })
                $WPFVMwareRestartContextMenu.Add_Click({ Perform-Action -action 'VMware Restart' -form $mainForm })
                $WPFVMwareDeleteContextMenu.Add_Click({ Perform-Action -action 'Remove From Hypervisor' -form $mainForm })
            }
            else
            {
                $WPFVMwareContextMenu.IsEnabled = $false
            }
            if( Get-Module ActiveDirectory -ErrorAction SilentlyContinue ) 
            {
                $WPFADDeleteContextMenu.Add_Click({ Perform-Action -action 'Remove From AD' -form $mainForm })
            }
            else
            {
                $WPFADContextMenu.IsEnabled = $false
            }
            
            $WPFDDCMessageUsersContextMenu.Add_Click({ Perform-Action -action Message -form $mainForm })
            $WPFDDCDeleteContextMenu.Add_Click({ Perform-Action -action 'Remove From DDC' -form $mainForm })
            $WPFDDCMaintModeOffContextMenu.Add_Click({ Perform-Action -action 'Maintenance Mode Off' -form $mainForm })
            $WPFDDCMaintModeOnContextMenu.Add_Click({ Perform-Action -action 'Maintenance Mode On'  -form $mainForm })

            $WPFWinShutdownModeOnContextMenu.Add_Click({ Perform-Action -action Shutdown -form $mainForm })
            $WPFWinRestartContextMenu.Add_Click({ Perform-Action -action Reboot -form $mainForm  })

            $WPFPVSDeleteContextMenu.Add_Click({ Perform-Action -action 'Remove From PVS' -form $mainForm })
            $WPFPVSBootContextMenu.Add_Click({ Perform-Action -action 'PVS Boot' -form $mainForm })
            $WPFPVSShutdownContextMenu.Add_Click({ Perform-Action -action 'PVS Power Off' -form $mainForm  })
            $WPFPVSRestartContextMenu.Add_Click({ Perform-Action -action 'PVS Restart' -form $mainForm  })

            $WPFlstMachines.Items.Clear()
            $WPFlstMachines.ItemsSource = $selected
            ## Select all items since already selected them in grid view
            $WPFlstMachines.SelectAll()
            $null = $mainForm.ShowDialog()

            ## Put in clipboard so we can paste into something if we need to
            if( $selected )
            {
                $selected | Clip.exe
            }
        }
    }
}
else
{
    Write-Warning "No PVS devices found via $($pvsServers -join ' ')"
}

if( ( Get-Variable global:DefaultVIServers -ErrorAction SilentlyContinue ) -and $global:DefaultVIServers.Count )
{
    Disconnect-VIServer -Server $hypervisors -Force -Confirm:$false
}
