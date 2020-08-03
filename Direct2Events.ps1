#Requires -version 3.0
<#
    Use multiple vendor's PowerShell modules to centralise information for XenApp/XenDesktop 7.x and allow troubleshooting and investigation all from one UI

    Guy Leech, 2018

    Modification History:

    22/05/18  GL  Added "Find User" functionality
#>

[CmdletBinding()]

Param
(
    [string]$configRegKey = 'HKCU:\software\Guy Leech\Citrix Consolidated Config' ,
    [string[]]$snapins = @( 'Citrix.Broker.Admin.*'  ) ,
    [string[]]$modules = @( 'ActiveDirectory' , 'VMware.VimAutomation.Core' , "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll" , 'Guys.Common.Functions.psm1' ) ,
    [int]$maxRecordCount = 1000 ,
    [string]$configFile ,
    [switch]$alwaysOnTop ,
    [int]$secondsWindow = 120 , ## default but overrriden by config/registry
    [string]$messageText = 'Please logoff' ,
    [string]$actionPattern = '##Input' ,
    ## Read from registry or configuration dialog
    [string]$ddc = $null ,
    [string]$PVS = $null ,
    [string]$AMC = $null ,
    [string[]]$hypervisor = @() ,
    [switch]$test
)

$global:hypervisorConnected = $false

[hashtable]$processPriorities = 
@{
  24 = 'Realtime'
  13 = 'High'
  10 = 'Above normal'
  8  = 'Normal' 
  6  = 'Below normal'
  4  = 'Low'
}

$setworkingsetPinvoke = @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
  
    public static class Memory
    {
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool SetProcessWorkingSetSizeEx( IntPtr proc, int min, int max , int flags );
    }
}
'@

#region XAML&Modules

$machineSelectionWindowXML = @"
<Window x:Class="Direct2Events.MachineFilter"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="Machine Filter" Height="757" Width="414">
    <Grid Height="307" VerticalAlignment="Top">
        <ComboBox x:Name="comboDeliveryGroup" HorizontalAlignment="Left" Height="26" Margin="154,33,0,0" VerticalAlignment="Top" Width="210"/>
        <CheckBox x:Name="chkDeliveryGroup" Content="Delivery Group" HorizontalAlignment="Left" Height="26" Margin="10,37,0,0" VerticalAlignment="Top" Width="104"/>
        <ComboBox x:Name="comboMachineCatalogue" HorizontalAlignment="Left" Height="27" Margin="154,76,0,0" VerticalAlignment="Top" Width="210"/>
        <CheckBox x:Name="chkMachineCatalogue" Content="Machine Catalogue" HorizontalAlignment="Left" Height="20" Margin="11,83,0,0" VerticalAlignment="Top" Width="127"/>
        <CheckBox x:Name="chkMultiSession" Content="Multi Session&#xD;&#xA;" HorizontalAlignment="Left" Height="18" Margin="12,133,0,0" VerticalAlignment="Top" Width="102" IsChecked="True"/>
        <CheckBox x:Name="chkSingleSession" Content="Single Session" HorizontalAlignment="Left" Height="27" Margin="10,183,0,0" VerticalAlignment="Top" Width="117"/>
        <Button x:Name="btnConfigurationOk" Content="OK" HorizontalAlignment="Left" Height="22" Margin="19,691,0,-406" VerticalAlignment="Top" Width="77" IsDefault="True"/>
        <Button x:Name="btnConfigurationCancel" Content="Cancel" HorizontalAlignment="Left" Height="22" Margin="114,691,0,-406" VerticalAlignment="Top" Width="77" IsCancel="True"/>
        <CheckBox x:Name="chkPoweredOn" Content="Powered On" HorizontalAlignment="Left" Height="24" Margin="10,222,0,0" VerticalAlignment="Top" Width="113"/>
        <CheckBox x:Name="chkMachineName" Content="Machine Name" HorizontalAlignment="Left" Height="22" Margin="11,265,0,0" VerticalAlignment="Top" Width="118"/>
        <TextBox x:Name="txtMachineName" HorizontalAlignment="Left" Height="28" Margin="164,259,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200"/>
        <CheckBox x:Name="chkUsersConnected" Content="With users connected" HorizontalAlignment="Left" Height="25" Margin="10,413,0,-131" VerticalAlignment="Top" Width="232"/>
        <StackPanel Margin="12,453,170.333,-218" Orientation="Vertical" Name="stkMaintenanceMode">
            <RadioButton x:Name="radInMaintenanceMode" Content="In Maintenance Mode" HorizontalAlignment="Left" Height="24" VerticalAlignment="Top" Width="174" GroupName="MaintenanceMode"/>
            <RadioButton x:Name="radNotInMaintenanceMode" Content="Not in Maintenance Mode" HorizontalAlignment="Left" Height="21" VerticalAlignment="Top" Width="201" GroupName="MaintenanceMode"/>
            <RadioButton x:Name="radEitherMaintenanceMode" Content="Either" HorizontalAlignment="Left" Height="17" VerticalAlignment="Top" Width="151" GroupName="MaintenanceMode" IsChecked="True"/>

        </StackPanel>
        <StackPanel Margin="164,307,40.333,-78" Orientation="Vertical" Name="stkMachinesFrom">
            <RadioButton x:Name="radMachinesFromCitrix" Content="Citrix" HorizontalAlignment="Left" Height="30" VerticalAlignment="Top" Width="179" IsChecked="True" GroupName="MachinesFrom"/>
            <RadioButton x:Name="radMachinesFromHypervisor" Content="Hypervisor" HorizontalAlignment="Left" Height="25" VerticalAlignment="Top" Width="179" GroupName="MachinesFrom"/>
            <RadioButton x:Name="radMachinesFromAD" Content="Active Directory" HorizontalAlignment="Left" Height="22" VerticalAlignment="Top" Width="155" GroupName="MachinesFrom"/>

        </StackPanel>
        <Label Content="Retrieve from:" HorizontalAlignment="Left" Height="32" Margin="45,335,0,-60" VerticalAlignment="Top" Width="93"/>
        <StackPanel x:Name="stkRegistrationState" Margin="12,541,164,-315" Orientation="Vertical">
            <RadioButton x:Name="radRegistered" Content="Registered" HorizontalAlignment="Left" Height="25" VerticalAlignment="Top" Width="230" GroupName="Registered"/>
            <RadioButton x:Name="radUnregistered" Content="Unregistered" HorizontalAlignment="Left" Height="24" VerticalAlignment="Top" Width="230" GroupName="Registered"/>
            <RadioButton x:Name="radEitherRegisteredUnregistered" Content="Either" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="224" IsChecked="True" GroupName="Registered"/>

        </StackPanel>

    </Grid>
</Window>

"@

$progressWindowXML = @"
<Window x:Name="Progress" x:Class="Direct2Events.Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="Busy ..." Height="142.389" Width="303.854">
    <Grid>
        <TextBox x:Name="txtProgress" HorizontalAlignment="Left" Height="67" Margin="10,21,0,0" TextWrapping="Wrap" Text="Doing stuff ..." VerticalAlignment="Top" Width="270"/>

    </Grid>
</Window>
"@

$configurationWindowXML = @"
<Window x:Class="Direct2Events.Configuration"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="Configuration" Height="523.333" Width="343.511">
    <Grid Margin="0,0,0,0          ">
        <TextBox x:Name="txtDDC" HorizontalAlignment="Left" Height="26" Margin="160,30,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="160"/>
        <Label Content="Delivery Controllers" HorizontalAlignment="Left" Height="26" Margin="14,30,0,0" VerticalAlignment="Top" Width="124" AutomationProperties.Name="Delivery Controllers"/>
        <TextBox x:Name="txtPVS" HorizontalAlignment="Left" Height="26" Margin="160,108,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="160"/>
        <Label Content="PVS Servers" HorizontalAlignment="Left" Height="26" Margin="14,108,0,0" VerticalAlignment="Top" Width="124" AutomationProperties.Name="Delivery Controllers"/>
        <TextBox x:Name="txtHypervisor" HorizontalAlignment="Left" Height="26" Margin="160,161,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="160"/>
        <Label Content="Hypervisor" HorizontalAlignment="Left" Height="26" Margin="14,161,0,0" VerticalAlignment="Top" Width="124" AutomationProperties.Name="Delivery Controllers"/>
        <TextBox x:Name="txtAppSense" HorizontalAlignment="Left" Height="26" Margin="160,205,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="160" />
        <Label Content="AppSense" HorizontalAlignment="Left" Height="26" Margin="14,205,0,0" VerticalAlignment="Top" Width="124" AutomationProperties.Name="Delivery Controllers"/>
        <Button x:Name="btnConfigurationOk" Content="OK" HorizontalAlignment="Left" Height="22" Margin="19,452,0,0" VerticalAlignment="Top" Width="77"/>
        <Button x:Name="btnConfigurationCancel" Content="Cancel" HorizontalAlignment="Left" Height="22" Margin="114,452,0,0" VerticalAlignment="Top" Width="77"/>
        <Grid Margin="19,205,198.667,241">
            <TextBox x:Name="txtSeconds" Height="24" Margin="-3,174,74.333,-152" TextWrapping="Wrap" Text="120" VerticalAlignment="Top" UndoLimit="99"/>
            <Label x:Name="labelSeconds" Content="Seconds before/after" HorizontalAlignment="Left" Height="28" Margin="50,170,-167,-150" VerticalAlignment="Top" Width="236"/>

        </Grid>
        <CheckBox x:Name="chkHttps" Content="https" HorizontalAlignment="Left" Height="21" Margin="160,74,0,0" VerticalAlignment="Top" Width="59"/>
        <TextBox x:Name="txtMessageText" HorizontalAlignment="Left" Height="104" Margin="160,261,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="160"/>
        <Label Content="MessageText" HorizontalAlignment="Left" Height="27" Margin="14,261,0,0" VerticalAlignment="Top" Width="128"/>

    </Grid>
</Window>

"@

$healthWindowXML = @"
<Window x:Name="HealthReportWindow" x:Class="Direct2Events.HealthReport"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="Information" Height="540.178" Width="794.03">
    <Grid Margin="0,0,-209.667,-5.333" HorizontalAlignment="Stretch">
        <StackPanel HorizontalAlignment="Left" Height="458" Margin="32,20,0,0" VerticalAlignment="Top" Width="730">
            <ListView x:Name="Health" HorizontalAlignment="Left" Height="454" VerticalAlignment="Top" Width="730" >
                <ListView.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="Remediate" Name="HealthInfoContextMenu" />
                </ContextMenu>
                </ListView.ContextMenu>
                <ListView.ItemContainerStyle>
                    <Style TargetType="{x:Type ListViewItem}">
                        <Style.Triggers>
                            <Trigger Property="IsMouseOver" Value="true">
                                <Setter Property="Foreground" Value="Blue" />
                            </Trigger>
                              <DataTrigger Binding="{Binding Path=Warning}" Value="True">
                                <Setter Property="Foreground" Value="Orange" />
                                <Setter Property="Background" Value="DarkBlue" />
                              </DataTrigger>
                              <DataTrigger Binding="{Binding Path=Critical}" Value="True">
                                <Setter Property="Foreground" Value="Red" />
                                <Setter Property="Background" Value="DarkBlue" />
                              </DataTrigger>	
                        </Style.Triggers>
                    </Style>
                </ListView.ItemContainerStyle>

                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="Property" DisplayMemberBinding ="{Binding Name}"/>
                        <GridViewColumn Header="Value" DisplayMemberBinding ="{Binding Value}"/>
                    </GridView>
                </ListView.View>
            </ListView>
        </StackPanel>

    </Grid>
</Window>

"@

$processesWindowXML = @'
<Window x:Class="Direct2Events.ProcessList"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="Processes" Height="300" Width="600">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition />
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <DataGrid Name="ProcessList">
            <DataGrid.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="Trim" Name="TrimProcessContextMenu" />
                    <MenuItem Header="Set Maximum Working Set" Name="SetMaxWorkingSetProcessContextMenu" />
                    <MenuItem Header="Kill" Name="KillProcessContextMenu" />
                    <MenuItem Header="Set Priority" Name="SetPriorityProcessContextMenu">
                        <MenuItem Header="High" Name="SetHighPriorityProcessContextMenu" />
                        <MenuItem Header="Above Normal" Name="SetAboveNormalPriorityProcessContextMenu" />
                        <MenuItem Header="Normal" Name="SetNormalPriorityProcessContextMenu" />
                        <MenuItem Header="Below Normal" Name="SetBelowNormalPriorityProcessContextMenu" />
                        <MenuItem Header="Low" Name="SetLowPriorityProcessContextMenu" />
                    </MenuItem>
                </ContextMenu>
            </DataGrid.ContextMenu>
            <DataGrid.Columns>
                <DataGridTextColumn Binding="{Binding Name}" Header="Name"/>
                <DataGridTextColumn Binding="{Binding ProcessId}" Header="PID"/>
                <DataGridTextColumn Binding="{Binding ParentProcessId}" Header="Parent Process Id"/>
                <DataGridTextColumn Binding="{Binding SessionId}" Header="Session Id"/>
                <DataGridTextColumn Binding="{Binding Owner}" Header="Owner"/>
                <DataGridTextColumn Binding="{Binding WorkingSetKB}" Header="Working Set (KB) "/>
                <DataGridTextColumn Binding="{Binding PeakWorkingSetKB}" Header="Peak Working Set (KB)"/>
                <DataGridTextColumn Binding="{Binding PageFileUsageKB}" Header="Page File Usage (KB)"/>
                <DataGridTextColumn Binding="{Binding PeakPageFileUsageKB}" Header="Peak Page File Usage (KB)"/>
                <DataGridTextColumn Binding="{Binding IOReadsMB}" Header="IO Reads (MB)"/>
                <DataGridTextColumn Binding="{Binding IOWritesMB}" Header="IO Writes (MB)"/>
                <DataGridTextColumn Binding="{Binding BasePriority}" Header="Base Priority"/>
                <DataGridTextColumn Binding="{Binding ProcessorTime}" Header="Processor Time (s)"/>
                <DataGridTextColumn Binding="{Binding HandleCount}" Header="Handles"/>
                <DataGridTextColumn Binding="{Binding ThreadCount}" Header="Threads"/>
                <DataGridTextColumn Binding="{Binding StartTime}" Header="Start Time"/>
                <DataGridTextColumn Binding="{Binding CommandLine}" Header="Command Line"/>
            </DataGrid.Columns>
        </DataGrid>
        <TextBlock Grid.Row="1">
            <TextBlock.Text>
                <MultiBinding StringFormat="{}{0}, {1}">
                    <Binding ElementName="grid" Path="ActualWidth"/>
                    <Binding ElementName="grid" Path="ActualHeight"/>
                </MultiBinding>
            </TextBlock.Text>
        </TextBlock>
    </Grid>
</Window>
'@

$mainWindowXML = @"
<Window x:Class="Direct2Events.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="Citrix Centralised Console" Height="662.256" Width="1250">
    <Grid Margin="0,0,-21.667,6.667">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="53*"/>
            <ColumnDefinition Width="7*"/>
            <ColumnDefinition Width="458*"/>
        </Grid.ColumnDefinitions>
        <DatePicker x:Name="startDatePicker" HorizontalAlignment="Left" Margin="71,30,0,0" VerticalAlignment="Top" Grid.ColumnSpan="3"/>
        <Button x:Name="btnGetUsers" Content="Get _Users" HorizontalAlignment="Left" Height="25" Margin="24,126,0,0" VerticalAlignment="Top" Width="101" Grid.ColumnSpan="3" IsDefault="True"/>
        <ListView x:Name="Users" HorizontalAlignment="Left" Height="495" Margin="103.333,63,0,0" VerticalAlignment="Top" Width="284" Grid.Column="2" SelectionMode="Single">
            <ListView.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="Info" Name="UserUserInfoContextMenu" />
                    <MenuItem Header="Events" >
                        <MenuItem Header="Events in Time Range" Name="EventsInRangeContextMenu" />
                        <MenuItem Header="Boot Events" Name="MachineBootEventsContextMenu" />
                    </MenuItem>
                    <MenuItem Header="Actions" >
                        <MenuItem Header="Logoff" Name="MachineLogoffContextMenu" />
                        <MenuItem Header="Message" Name="MachineMessageContextMenu" />
                        <MenuItem Header="Disconnect" Name="MachineDisconnectContextMenu" />
                        <MenuItem Header="Restart Service" Name="MachineRestartServiceContextMenu" />
                        <MenuItem Header="Maintenance Mode" Name="MachineMaintenanceModeContextMenu" />
                        <MenuItem Header="Reboot" Name="MachineRebootContextMenu" />
                        <MenuItem Header="Logon" Name="MachineLogonContextMenu" />
                    </MenuItem>
                </ContextMenu>
            </ListView.ContextMenu>
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding 'Value'}" />
                    <GridViewColumn Header="Active Users" DisplayMemberBinding ="{Binding 'Value2'}" />
                    <GridViewColumn Header="Boot Time" DisplayMemberBinding ="{Binding 'Value3'}" />
                </GridView>
            </ListView.View>
        </ListView>
        <Label x:Name="labelUsers" Content="Users" HorizontalAlignment="Left" Height="27" Margin="103.333,27,0,0" VerticalAlignment="Top" Width="284" Grid.Column="2"/>
        <ListView x:Name="Sessions" HorizontalAlignment="Left" Height="496" Margin="508.333,61,0,0" VerticalAlignment="Top" Width="421" Grid.Column="2" SelectionMode="Single" >
            <ListView.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="Information" >
                        <MenuItem Header="User Info" Name="UserInfoContextMenu" />
                        <MenuItem Header="Machine Info" Name="MachineInfoContextMenu" />
                        <MenuItem Header="Session Info" Name="SessionInfoContextMenu" />
                        <MenuItem Header="Processes">
                            <MenuItem Header="Session" Name="SessionProcessesContextMenu" />
                            <MenuItem Header="All" Name="AllProcessesContextMenu" />
                        </MenuItem>
                    </MenuItem>
                    <MenuItem Header="Events" >
                        <MenuItem Header="Logon Events" Name="LogonEventsContextMenu" />
                        <MenuItem Header="Logoff Events" Name="LogoffEventsContextMenu" />
                        <MenuItem Header="Entire Session Events" Name="AllSessionEventsContextMenu" />
                        <MenuItem Header="Boot Events" Name="BootEventsContextMenu" />
                        <MenuItem Header="Events in Time Range" Name="EventsInRangeUserContextMenu" />
                    </MenuItem>
                    <MenuItem Header="Actions" >
                        <MenuItem Header="Logoff" Name="LogoffContextMenu" />
                        <MenuItem Header="Message" Name="MessageContextMenu" />
                        <MenuItem Header="Shadow" Name="ShadowContextMenu" />
                        <MenuItem Header="Disconnect" Name="DisconnectContextMenu" />
                        <MenuItem Header="Restart Service" Name="RestartServiceContextMenu" />
                        <MenuItem Header="Maintenance Mode" Name="MaintenanceModeContextMenu" />
                        <MenuItem Header="Reboot" Name="RebootContextMenu" />
                        <MenuItem Header="Logon" Name="LogonContextMenu" />
                    </MenuItem>
                </ContextMenu>
            </ListView.ContextMenu>
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Start" DisplayMemberBinding ="{Binding 'Column1'}" Width="150" />
                    <GridViewColumn Header="End" DisplayMemberBinding ="{Binding 'Column2'}" Width="150" />
                    <GridViewColumn Header="Server" DisplayMemberBinding ="{Binding 'Column3'}" Width="100"/>
                </GridView>
            </ListView.View>
            <ListView x:Name="listView" Height="100" Width="100">
                <ListView.View>
                    <GridView>
                        <GridViewColumn/>
                    </GridView>
                </ListView.View>
            </ListView>
        </ListView>
        <Label x:Name="labelSessions" Content="Sessions" HorizontalAlignment="Left" Height="27" Margin="508.333,29,0,0" VerticalAlignment="Top" Width="116" Grid.Column="2"/>
        <Button x:Name="btnSettings" Content="Settings" HorizontalAlignment="Left" Height="25" Margin="24,321,0,0" VerticalAlignment="Top" Width="101" AutomationProperties.Name="Settings"/>
        <Label Content="Start" HorizontalAlignment="Left" Margin="17,29,0,0" VerticalAlignment="Top"/>
        <DatePicker x:Name="endDatePicker" HorizontalAlignment="Left" Margin="71,80,0,0" VerticalAlignment="Top" Grid.ColumnSpan="3" Width="101"/>
        <Label Content="End" HorizontalAlignment="Left" Margin="24,78,0,0" VerticalAlignment="Top"/>
        <Button x:Name="btnFindUser" Grid.ColumnSpan="3" Content="Find User" HorizontalAlignment="Left" Height="25" Margin="24,172,0,0" VerticalAlignment="Top" Width="101"/>
        <Button x:Name="btnGetMachines" Grid.ColumnSpan="3" Content="Get _Machines" HorizontalAlignment="Left" Height="25" Margin="24,220,0,0" VerticalAlignment="Top" Width="101"/>
        <Button x:Name="btnTest" Content="Test" HorizontalAlignment="Left" Height="25" Margin="24,383,0,0" VerticalAlignment="Top" Width="101"/>
        <Button x:Name="btnPVSDevices" Content="PVS Devices" HorizontalAlignment="Left" Height="25" Margin="24,268,0,0" VerticalAlignment="Top" Width="101"/>
    </Grid>
</Window>

"@

$generalTextEntry = @'
<Window x:Name="EnterText" x:Class="Direct2Events.EnterText"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="Enter Text" Height="212.986" Width="292.038">
    <Grid FocusManager.FocusedElement="{Binding ElementName=textBoxEnterText}">
        <TextBox x:Name="textBoxEnterText" HorizontalAlignment="Left" Height="29" Margin="30,46,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="227" MaxLines="1"/>
        <Button x:Name="btnTextEntryOk" Content="OK" HorizontalAlignment="Left" Height="24" Margin="30,97,0,0" VerticalAlignment="Top" Width="76" IsDefault="True"/>
        <Button x:Name="btnTextEntryCancel" Content="Cancel" HorizontalAlignment="Left" Height="24" Margin="126,97,0,0" VerticalAlignment="Top" Width="71" IsCancel="True"/>
    </Grid>
</Window>

'@

$messageWindowXML = @"
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

$datePickerXAML = @"
<Window x:Name="DatePicker" x:Class="Direct2Events.DatePicker"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="Pick a date, any date" Height="196" Width="286">
    <Grid Margin="0,0,36,11">
        <DatePicker x:Name="datePicked" HorizontalAlignment="Left" Height="50" Margin="13,21,0,0" VerticalAlignment="Top" Width="221"/>
        <Button x:Name="btnDatePickerOK" Content="OK" HorizontalAlignment="Left" Height="18" Margin="13,116,0,-10" VerticalAlignment="Top" Width="71" IsDefault="True"/>
        <Button x:Name="btnDatePickerCancel" Content="Cancel" HorizontalAlignment="Left" Height="18" Margin="103,116,0,-10" VerticalAlignment="Top" Width="66" IsCancel="True"/>
        <StackPanel Orientation="Horizontal" Height="20" Margin="13,82,17,0" VerticalAlignment="Top">
            <TextBlock HorizontalAlignment="Left" Height="19" Margin="3,0,0,0" TextWrapping="Wrap" Width="31" Text="Time"/>
            <TextBox x:Name="txtTime" Height="19" TextWrapping="Wrap" Text="00:00:00" Width="129" Margin="39,0"/>
        </StackPanel>

    </Grid>
</Window>

"@

$passwordPickerXAML = @"
<Window x:Name="PasswordPicker" x:Class="Direct2Events.PasswordPicker"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Direct2Events"
        mc:Ignorable="d"
        Title="PasswordPicker" Height="287.634" Width="297.71">
    <Grid Margin="0,0,0,143">
        <PasswordBox x:Name="passwordBox" HorizontalAlignment="Left" Height="33" Margin="26,27,0,0" VerticalAlignment="Top" Width="219"/>
        <Button x:Name="btnPasswordOk" Content="OK" HorizontalAlignment="Left" Height="31" Margin="26,80,0,0" VerticalAlignment="Top" Width="95" IsDefault="True"/>
        <Button x:Name="btnPasswordCancel" Content="Cancel" HorizontalAlignment="Left" Height="31" Margin="150,80,0,0" VerticalAlignment="Top" Width="79" IsCancel="True"/>

    </Grid>
</Window>

"@

function Global:Invoke-ODataTransform ($records )
{
    if( $records )
    {
        $propertyNames = ($records | Select -First 1).content.properties |
            Get-Member -MemberType Properties |
            Select -ExpandProperty name

        [int]$timeOffset = Get-DayLightSavingsOffet

        $records | ForEach-Object `
        {
            $record = $_
            $h = @{}
            $h.ID = $record.ID
            $properties = $record.content.properties

            $propertyNames | ForEach-Object `
            {
                $propertyName = $_
                $targetProperty = $properties.$propertyName
                if($targetProperty -is [Xml.XmlElement])
                {
                    try
                    {
                        $h.$propertyName = $targetProperty.'#text'
                        ## see if we need to adjust for daylight savings
                        if( $timeOffset -and ! [string]::IsNullOrEmpty( $h.$propertyName ) -and $targetProperty.type -match 'DateTime' )
                        {
                            $h.$propertyName = (Get-Date -Date $h.$propertyName).AddHours( $timeOffset )
                        }
                    }
                    catch
                    {
                        ##$_
                    }
                }
                else
                {
                    $h.$propertyName = $targetProperty
                }
            }

            [PSCustomObject]$h
        }
    }
}

Function Get-DayLightSavingsOffet
{
    [OutputType([Int])]
    Param()

    if( (Get-Date).IsDaylightSavingTime() )
    {
        1
    }
    else
    {
        0
    }
}

Function Progress-Window( [string]$text , $parent )
{
    if( [string]::IsNullOrEmpty( $text ) )
    {
        ## Called to close progress window
        if( $global:job )
        {
            ##$global:progressWindow.Close() 
            $global:form.Close()
            $global:newThread.Stop()
            $global:newThread.Dispose()
            $global:powershell.Dispose()
        }
        $global:runspace.Close()
        $global:powershell.Dispose()
    }
    else
    {
        Add-Type -AssemblyName System.Windows.Forms

        ## Create async message box to inform user what we're doing
        $global:form = New-Object System.Windows.Forms.Form
        if( $parent )
        {
            $parent.AddChild( $form )
        }
        $form.Size = New-Object System.Drawing.Size(300,150)
        $form.StartPosition = 'CenterParent'
        $form.Location = New-Object Drawing.Point 50 , ( ([Windows.Forms.Screen]::PrimaryScreen).WorkingArea.Height - 400 ) ## want it near the start menu

        $form.AutoSize = $true

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object Drawing.Point 10,20
        $label.Size = New-Object System.Drawing.Point 250,60
        $label.AutoSize = $true
        $label.Font = New-Object System.Drawing.Font( $label.Font.Name, 12 , [Drawing.FontStyle]::Bold ) ## 12 is the point size, default is usualy too small

        $form.Controls.Add( $label )
        $form.TopMost = $true
        $label.text  = $text
        $form.text = "Busy ..."

        $progress = { [void]$form.ShowDialog() }

        ##$global:progressWindow = Load-GUI $progressWindowXML

        if( $true ) ##$global:progressWindow )
        {
            ##$WPFtxtProgress.Text = $text
    
            $global:powershell = [powershell]::Create()
            $global:runspace = [runspacefactory]::CreateRunspace()
            $global:runspace.Open()
            $global:newThread = $powershell.AddScript($progress)
            $global:runspace.SessionStateProxy.SetVariable( 'form' , $form ) # $global:progressWindow )
            $global:powershell.Runspace = $global:runspace
            $global:job = $global:newThread.BeginInvoke() ## This will cause the form to be displayed
        }
    }
}

Function Load-GUI( $inputXml )
{
    $form = $NULL
    $inputXML = $inputXML -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
 
    [xml]$XAML = $inputXML
 
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
 
    $xaml.SelectNodes('//*[@Name]') | ForEach-Object `
    {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
    }

    return $form
}

[void][Reflection.Assembly]::LoadWithPartialName('Presentationframework')

$mainForm = Load-GUI $mainWindowXML

if( ! $mainForm )
{
    return
}

if( $DebugPreference -eq 'Inquire' )
{
    Get-Variable -Name WPF*
}

if( $snapins -and $snapins.Count -gt 0 )
{
    ForEach( $snapin in $snapins )
    {
        Add-PSSnapin $snapin
    }
}

if( $modules -and $modules.Count -gt 0 )
{
   ForEach( $module in $modules )
    {   
        Import-Module $module -ErrorAction SilentlyContinue
        [bool]$loaded = $?
        if( ! $loaded -and $module -notmatch '^[a-z]:\\' -and  $module -notmatch '^\\\\' ) ## only check script folder if not an absolute or UNC path
        {
            ## try same folder as the script if there is no path in the module name
            Import-Module (Join-Path ( & { Split-Path -Path $myInvocation.ScriptName -Parent } ) $module ) -ErrorAction Continue
            $loaded = $?
        }
        if( ! $loaded )
        {
            Write-Warning "Unable to load module `"$module`" so functionality may be limited"
        }
    }
}
#endregion XAML&Modules

Function Action-Item( $item )
{
    ## item is what was pulled from the csv: RightClickName,RightClickAction,RightClickInputs
    ## e.g. "Extend Account,Set-ADAccountExpiration $value -DateTime ##Input1,$                                                                            XAML:Choose New Account Expiry Date" 
    ##    Where DatePickerXAML is XAML to produce UI to return data to replace placeholder ##Input
    ##       
    if( $item )
    {
        $value = $item.value
        if( $item.RightClickAction -match $actionPattern )
        {
            if( $item.RightClickInputs )
            {
                $XAML,$controlWithResult,$header = $item.RightClickInputs -split ':'
                $form = Load-GUI (Get-Variable $XAML).Value
                if( $form )
                {
                    $eventBlock = { 
                        $_.Handled = $true
                        if( $_.psobject.Properties.name -match '^OriginalSource$' -and $_.OriginalSource.psobject.Properties.name -match '^content$' -and ($_.OriginalSource.Content -eq 'OK' -or $_.Source.OriginalSource -eq 'Cancel' ))
                        {
                            if( $_.OriginalSource.Content -eq 'OK' )
                            {
                                $this.DialogResult = $true
                            }
                            $this.Close()
                        }
                    }
                    $eventHandler = [Windows.RoutedEventHandler]$eventBlock
                    $form.AddHandler([Windows.Controls.Button]::ClickEvent, $eventHandler)
                    if( ! [string]::IsNullOrEmpty( $header ) )
                    {
                        $form.Title = $header
                    }
                    if( $form.ShowDialog() )
                    {
                        ## Could look at name of form to figure what $WPFvariable to get or stuff it in the form somewhere
                        $result = Invoke-Expression ( '$WPF' + $controlWithResult )
                        $item.RightClickAction = $item.RightClickAction -replace "$($actionPattern)\d" , $result
                    }
                    else ## cancelled
                    {
                        return
                    }
                }
                else ## UI failed to load
                {
                    ## TODO error
                    return
                }
            }
            else
            {
                ## TODO Syntax error as with have an action pattern to replace but no info as to what to replce it with
                return
            }
        }
        
        ## The command will probably need the user/machine as an argument
        $user = $item.Owner
        $computer = $item.Owner
             
        $Error.Clear()
        Invoke-Expression -Command $item.RightClickAction 
        if( $Error.Count -gt 0 )
        {
            [Windows.MessageBox]::Show( $Error[0].Exception.Message , "Failed to perform action `"$($item.RightClickAction)`"" , 'OK' ,'Error' )
        }
            
    }
}

Function Add-FromConfigFile( [string]$configFile , [string]$scope , $form , [string]$owner )
{
    [hashtable]$actions = @{}
    if( Test-Path $configFile )
    {
        [int]$line = 0
        Import-Csv $configFile | Where-Object { $_.Scope -eq $scope } | ForEach-Object `
        {
            $line++
            $configLine = $_
            Write-Verbose "Config line $line : $_"
            try
            {
                $value = Invoke-Expression $_.Value
                ## add null value if there is a warning or error associated
                if( $_.Warning -or $_.Critical )
                {
                    [bool]$warning = if( ! [string]::IsNullOrEmpty( $_.Warning ) ){ Invoke-Expression $_.Warning } else { $false }
                    [bool]$critical = if( ! [string]::IsNullOrEmpty( $_.Critical ) ){ Invoke-Expression $_.Critical } else { $false }
                }
                ## if we have any remedation items then build them into a table so we can return them to caller, keyed on the property name
                if( ![string]::IsNullOrEmpty( $_.RightClickName ) )
                {
                    ## Store value as doesn't seem to be a way to get it by the time the handler is called
                    $actions.Add( $_.Name , ( New-Object -TypeName PSCustomObject -Property (@{ 'Owner' = $owner ; 'Value' = $value ; 'RightClickName'=$_.RightClickName ; 'RightClickAction' = $_.RightClickAction ; 'RightClickInputs' = $_.RightClickInputs }) ) )
                }
                $form.AddChild( $( New-Object -TypeName PSCustomObject -Property (@{ 'Name' = $_.Name ; 'Value' = $value ; 'Warning' = $warning ; 'Critical' = $critical } )))
            }
            catch
            {
                Write-Warning "$configFile : $line : $configLine`n$_"
            }
        }
    }
    else
    {
        [Windows.MessageBox]::Show( "Failed to find config file `"$configFile`"" , "Machine Information" , 'OK' ,'Error' )
    }
    return $actions
}

Function Process-Processes( $GUIobject , [string]$Operation , [string]$machineName )
{
    $_.Handled = $true
    ## get selected items from control
    [int[]]$pids = @( $GUIobject.selectedItems | Select -ExpandProperty ProcessId )
    if( ! $pids -or ! $pids.Count )
    {
        return
    }
    ## prompt to confirm
    $answer = [Windows.MessageBox]::Show( "Are you sure you want to $operation these $($pids.Count) processes?" , 'Confirm Action' , 'YesNo' ,'Question' )
    if( $answer -ne 'yes' )
    {
        return
    }

    if( $operation -eq 'Kill' -or $operation -match 'Priority' )
    {
        [string]$method = if( $operation -eq 'Kill' ) { 'Terminate' } else { 'SetPriority' }
        [int]$argument = switch( $operation ) ## See https://msdn.microsoft.com/en-us/library/aa393587(v=vs.85).aspx
        {
            'SetHighPriority' { 128 }
            'SetAboveNormalPriority' { 32768 }
            'SetNormalPriority' { 32 }
            'SetBelowNormalPriority' { 16384 }
            'SetLowPriority' { 64 }
            'Kill' { 1 } ## exit code of process killed
        }
        $pids | ForEach-Object `
        {
            $thispid = $_
            Get-WmiObject -Class win32_process -ComputerName $machineName -Filter "ProcessId ='$thispid'" | Invoke-WmiMethod -Name $method -ArgumentList $argument
        }
    }
    else
    {
        [int]$maxWorkingSet = -1
        [int]$minWorkingSet = -1
        [int]$flags = 0

        if( $operation -eq 'SetMaxWorkingSet' )
        {
            ## Need to prompt for working set size to set
            $limitForm = Load-GUI $generalTextEntry
            if( $limitForm )
            {
                $wpfbtnTextEntryOk.Add_Click( {
                    $limitForm.DialogResult = $true 
                    $limitForm.Close() 
                } )
                $limitForm.Title = "Enter Maximum Working Set"
                if( $limitForm.ShowDialog() )
                {
                    $maxWorkingSet = Invoke-Expression $wpftextBoxEnterText.Text.Trim() ## in case MB or GB in there
                    $flags = $flags -bor 0x4 ## QUOTA_LIMITS_HARDWS_MAX_ENABLE
                    $minWorkingSet = 1 ## if a maximum is specified then we must specify a minimum too - this will default to the minimum
                }
                else
                {
                    return
                }
            }
            else
            {
                return
            }
        }
        [scriptblock]$remoteCode = `
        {
            Add-Type $using:setworkingsetPinvoke
            Get-Process -id $using:pids | ForEach-Object { [PInvoke.Win32.Memory]::SetProcessWorkingSetSizeEx( $_.Handle,$using:minWorkingSet,$using:maxWorkingSet,$using:flags) ; [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error() }
        }
        $result,$lastError = Invoke-Command -ComputerName $machineName -ScriptBlock $remoteCode
        if( ! $result )
        {
        }
    }
}

Function Show-UserInformation( [switch]$showProcesses = $false , [switch]$allProcesses )
{
    ## Need to see what is in the $WPFusers control -users or machines
    if( $WPFlabelUsers.Content -and $WPFlabelUsers.Content -match 'users' )
    {
        $selectedUser = $WPFUsers.SelectedItem
    }
    else
    {
        $selectedUser = $WPFSessions.SelectedItem
    }
    if( $selectedUser )
    {
        [string]$userName = $selectedUser.Value
        [string]$fullUserName = $null
        if( $selectedUser.psobject.properties.name -match '^Name$' )
        {
            $fullUserName = $selectedUser.Name
        }
        else
        {
            $fullUserName = $selectedUser.'Full Name'
        }

        if( $showProcesses )
        {
            if( ! (Get-Member -InputObject $WPFSessions.SelectedItem -Name 'EndDate' -ErrorAction SilentlyContinue ) -or [string]::IsNullOrEmpty( $WPFSessions.SelectedItem.EndDate ) )
            {
                ## get processes for this session and display with right click kill action
                ## TODO this won't work with multiple sessions for same user on same server
                [string]$machine = $WPFSessions.SelectedItem.Machine
                [string[]]$quser = (quser.exe $userName /server:$machine | select -Skip 1 -First 1).Trim() -split '\s+'
                [int]$sessionId = -1
                if( $quser -and $quser.Count )
                {
                    if( $quser[0] -eq $userName )
                    {
                        [int]$index = 1
                        if( $quser[1] -match '^[^\d]' ) ## if starts with non-numeric then a winstation name so session not disconnected so session id is in next field
                        {
                            $index++
                        }
                        if( ! [int]::TryParse( $quser[$index] , [ref] $sessionId ) )
                        {
                            $sessionId = -1
                            Write-Warning "Unexpected text `"$($quser[$index]) found when looking for session id for user $userName on $machine"
                        }
                    }
                    else
                    {
                        Write-Warning "Unexpected user $($quser[0]) returned when looking for user $userName on $machine"
                    }
                }
                else
                {
                    Write-Warning "Session not found when looking for user $userName on $machine"
                }
                if( $sessionId -ge 0 -or $allProcesses )
                {
                    [string]$filter = if( $allProcesses ) { "sessionid >= 0" } else { "sessionid = '$sessionid'" }
                    ## Process priorities: Realtime = 24, high = 13 , above normal = 10, normal = , below normal = 6, low = 4
                    ## Some memory sizes are in KB and others not!
                   [array]$processes = @( Get-WmiObject -Class Win32_process -Filter $filter -ComputerName $machine | select Name,ProcessId,@{n='Owner';e={Invoke-WmiMethod -InputObject $_ -Name GetOwner | Select -ExpandProperty user}},ParentProcessId,SessionId,
                        @{n='WorkingSetKB';e={[math]::Round( $_.WorkingSetSize / 1KB)}},@{n='PeakWorkingSetKB';e={$_.PeakWorkingSetSize}},
                        @{n='PageFileUsageKB';e={$_.PageFileUsage}},@{n='PeakPageFileUsageKB';e={$_.PeakPageFileUsage}},
                        @{n='IOReadsMB';e={[math]::Round($_.ReadTransferCount/1MB)}},@{n='IOWritesMB';e={[math]::Round($_.WriteTransferCount/1MB)}},
                        @{n='BasePriority';e={$processPriorities[ [int]$_.Priority ] }},@{n='ProcessorTime';e={[math]::round( ($_.KernelModeTime + $_.UserModeTime) / 10e6 )}},HandleCount,ThreadCount,
                        @{n='StartTime';e={Get-Date ([Management.ManagementDateTimeConverter]::ToDateTime($_.CreationDate)) -Format G}},CommandLine )
                   if( $processes -and $processes.Count )
                   {
                        $processesForm = Load-GUI $processesWindowXML
                        if( $processesForm )
                        {
                            [string]$title = "$($processes.Count) processes $(if( ! $allProcesses ) { "for user $userName in session $sessionId " } )on $machine"
                            $processesForm.Title =  $title
                            $WPFProcessList.ItemsSource = $processes
                            $WPFProcessList.IsReadOnly = $true
                            $WPFProcessList.CanUserSortColumns = $true

                            $WPFTrimProcessContextMenu.Add_Click({ Process-Processes -GUIobject $WPFProcessList -Operation 'Trim' -machineName $machine })
                            $WPFSetMaxWorkingSetProcessContextMenu.Add_Click({ Process-Processes -GUIobject $WPFProcessList -Operation 'SetMaxWorkingSet' -machineName $machine })
                            $WPFKillProcessContextMenu.Add_Click({ Process-Processes -GUIobject $WPFProcessList -Operation 'Kill' -machineName $machine })
                            $WPFSetHighPriorityProcessContextMenu.Add_Click({ Process-Processes -GUIobject $WPFProcessList -Operation 'SetHighPriority' -machineName $machine })
                            $WPFSetAboveNormalPriorityProcessContextMenu.Add_Click({ Process-Processes -GUIobject $WPFProcessList -Operation 'SetAboveNormalPriority' -machineName $machine })
                            $WPFSetNormalPriorityProcessContextMenu.Add_Click({ Process-Processes -GUIobject $WPFProcessList -Operation 'SetNormalPriority' -machineName $machine })
                            $WPFSetBelowNormalPriorityProcessContextMenu.Add_Click({ Process-Processes -GUIobject $WPFProcessList -Operation 'SetBelowNormalPriority' -machineName $machine })
                            $WPFSetLowPriorityProcessContextMenu.Add_Click({ Process-Processes -GUIobject $WPFProcessList -Operation 'SetLowPriority' -machineName $machine })

                            $processesForm.ShowDialog()
                        }
                    }
                    else
                    {
                        [Windows.MessageBox]::Show( "Got no processes in session $sessionId for user $userName on $machine" , 'Processes Information' , 'OK' ,'Warning' )
                    }
                }
            }
            else
            {
                [Windows.MessageBox]::Show( "Session logged off at $($WPFSessions.SelectedItem.EndDate) so cannot get processes" , 'Processes Information' , 'OK' ,'Error' )
            }
        }
        else ## user info requested rather than processes
        {
            $healthForm = Load-GUI $healthWindowXML
            if( $healthForm )
            {          
                $script:userActionRightClicked = $null
                $WPFHealthInfoContextMenu.Add_Click({
                    Action-Item $script:userActionRightClicked
                    $_.Handled = $true
                })

                $healthForm.add_PreviewMouseRightButtonDown({
                    ## $_.OriginalSource gives us the item so we can search config file we read in
                    ## DataContext                 : @{Warning=False; Name=Email; Value=Adrian.mole@life.com; Critical=False}                    
                    ## Text                        : Adrian.mole@life.com
                    [string]$menuLabel = 'No action defined'
                    $script:userActionRightClicked = $null
                    if( $_ -and $_.OriginalSource -and $_.OriginalSource.DataContext )
                    {
                        $action = $script:userActions[ $_.OriginalSource.DataContext.Name ]
                        if( $action )
                        {
                            $menuLabel = $action.RightClickName
                            $script:userActionRightClicked = $action
                        }
                    }
                    $_.Source.ContextMenu.Items[0].Header = $menuLabel
                    $_.Handled = $true
                })

                $healthForm.Title = "$username Information"
                $userDetails = Get-ADUser $userName -Properties *
                if( $userDetails )
                {
                    $script:userActions = Add-FromConfigFile -configFile $configFile -scope 'User' -form $WPFHealth -owner $userName
                    $healthForm.Show()
                }
                else
                {
                    [Windows.MessageBox]::Show( "Failed to retrieve AD info for $username`n$($error[0])" , "Connection Information" , 'OK' ,'Error' )
                }
            }
        }
    }
}

Function Show-SessionInformation
{
    ## TODO Sub menu for processes?
    $selected = $WPFSessions.SelectedItem
    if( $selected )
    {
        $connection = Invoke-ODataTransform( Invoke-RestMethod -Uri "$global:baseURI/Connections()?`$filter=Id eq $($selected.CurrentConnectionId)" -UseDefaultCredentials )
        if( $connection )
        {
            $healthForm = Load-GUI $healthWindowXML
            ##$WPFHealth.Items.Clear()
            $healthForm.Title = 'Session Information' 
            $wpfhealth.ItemsSource =  $connection.psobject.Properties | ForEach-Object `
            { 
                $WPFHealth.AddChild( ( New-Object -TypeName PSCustomObject -Property (@{ 'Name' = $_.Name ; 'Value' = $_.Value })  ) )
                New-Object -TypeName PSCustomObject -Property (@{ 'Name' = $_.Name ; 'Value' = $_.Value })
            }
            $healthForm.Show()
        }
        else
        {
            [Windows.MessageBox]::Show( "Failed to retrieve connection information for from $global:ddc" , "Connection Information" , 'OK' ,'Error' )
        }
    }
}

Function Show-MachineInformation( [string]$machineName , [string]$unqualifiedMachineName )
{
    if( $unqualifiedMachineName -and $machineName )
    {
        Progress-Window "Fetching information from $machineName"
        $machineInfo = @( Invoke-ODataTransform( Invoke-RestMethod -Uri  "$global:baseURI/Machines()?`$filter=Name eq `'$machineName`'" -UseDefaultCredentials ) ) | sort createddate -Descending|select -First 1
        if( ! $machineInfo )
        {
            Write-Warning "Failed to retrieve information for $machineName from $global:ddc"
        }

        $healthForm = Load-GUI $healthWindowXML
        ## hypervisor level, ping
        $healthForm.Title = "$unqualifiedMachineName Information"
        [string]$deliveryGroup = $null
        [string]$machineCatalogue = $null
        if( $machineInfo )
        {
            $deliveryGroup = @( Get-BrokerDesktopGroup -UUid $machineInfo.DesktopGroupId -AdminAddress $global:ddc -ErrorAction SilentlyContinue | select -ExpandProperty Name ) -join "`n"
            $machineCatalogue = @( Get-BrokerCatalog -UUid $machineInfo.CatalogId -AdminAddress $global:ddc -ErrorAction SilentlyContinue | select -ExpandProperty Name ) -join "`n"
        }
        [int]$disconnectedUsers = ( Get-BrokerSession -MachineName $machineName -SessionState Disconnected -AdminAddress $global:ddc) | Measure-Object | select -ExpandProperty Count
        [int]$activeUsers = ( Get-BrokerSession -MachineName $machineName -SessionState Active -AdminAddress $global:ddc) | Measure-Object | select -ExpandProperty Count         
        $brokerMachine = Get-BrokerMachine -MachineName $machineName -AdminAddress $global:ddc
        $machineDetails = Get-ADComputer $unqualifiedMachineName -Properties *
        ## batch up remote commands and so in one go which should be quicker
        $remoteWork = `
        {
            $osinfo = Get-CimInstance Win32_OperatingSystem
            $computerinfo = Get-CimInstance Win32_ComputerSystem
            $vdiskInfo = Get-Content "c:\personality.ini"
            $logicalDisks = Get-CimInstance -ClassName Win32_logicaldisk
            $kms = Get-CimInstance -ClassName SoftwareLicensingProduct 
            $osinfo,$computerInfo,$vdiskInfo,$logicalDisks,$kms
        }
        $osinfo,$computerInfo,$vdiskInfo,$logicalDisks,$kms = Invoke-Command -ComputerName $unqualifiedMachineName -ScriptBlock $remoteWork
        $pvsinfo = $null
        if( ! [string]::IsNullOrEmpty( $global:PVS ) )
        {
            Set-PvsConnection -Server $global:PVS
            $pvsDeviceInfo = Get-PvsDeviceInfo -DeviceName $unqualifiedMachineName
            ## Get currently in use vdisk
            $vdisk = Get-PvsDiskInfo -DiskLocatorId $pvsDeviceInfo.DiskLocatorId
            ## Get vdisks assigned to this device so we can check booted off it
            $vdisks = @( Get-PvsDiskInfo -DeviceId $pvsDeviceInfo.DeviceId )
            ## See if there is an override version as in not booting off latest
            [int]$bootVersion = -1
            $override = Get-PvsDiskVersion -DiskLocatorId $vdisks[0].DiskLocatorId | Where-Object { $_.Access -eq 3 } 
            if( $override )
            {
                $bootVersion = $override.Version
            }
            else
            {
                ## Access: Read-only access of the Disk Version. Values are: 0 (Production), 1 (Maintenance), 2 (MaintenanceHighestVersion), 3 (Override), 4 (Merge), 5 (MergeMaintenance), 6 (MergeTest), and 7 (Test) Min=0, Max=7, Default=0
                $bootVersion = Get-PvsDiskVersion -DiskLocatorId $vdisks[0].DiskLocatorId | Where-Object { $_.Access -eq 0 } | Sort Version -Descending | Select -First 1 | select -ExpandProperty Version
            }
            ##$thisDisk = (( $vdiskInfo | Select-String '^\$DiskName=')-split '=')[1]
        }
        if( ! [string]::IsNullOrEmpty( $global:AMC ) )
        {
            $thisAMC,$thisPort = $global:AMC -split ':'
            if( [string]::IsNullOrEmpty( $thisPort ) )
            {
                $thisPort = 80
            }
            if( ! ( Connect-AppSenseManagementServer -ManagementServer $thisAMC -UseCurrentUser -ErrorAction Stop -Port $thisPort))
            {
                Write-Warning "Can't connect to source AppSense AMC $thisAMC on port $thisPort"
            }
            else
            {
                $appsenseMachine = Get-AppSenseManagementServerMachines -MachineName $unqualifiedMachineName
                if( $appsenseMachine )
                {
                    $appsenseGroup = Get-AppSenseManagementServerDeploymentGroups -DeploymentGroupKey $appsenseMachine.GroupKey -IncludeSummary -IncludeAssignedPackages -IncludeInstallationSchedule -IncludeEnterpiseAuditing
                }
                else
                {
                    Write-Warning "Failed to get machine $unqualifiedMachineName from AppSense AMC $global:AMC"
                }
            }
        }
        
        $vmInfo = $null
        if( $global:hypervisor -and $global:hypervisor.Count )
        {
            if( ! $global:hypervisorConnected )
            {
                $global:hypervisorConnected = Connect-Hypervisor $global:hypervisor
            }
            $vmInfo = Get-VM -Name $unqualifiedMachineName
        }

        Add-FromConfigFile -configFile $configFile -scope 'Machine' -form $WPFHealth

        ##Get-CimInstance Win32_logicaldisk -ComputerName $unqualifiedMachineName | ForEach-Object `
        $logicalDisks  | ForEach-Object `
        {
            if( ! [string]::IsNullOrEmpty( $_.DeviceID) -and $_.Size -gt 0 )
            {
                $WPFHealth.AddChild( $( New-Object -TypeName PSCustomObject -Property (@{ 'Name' = "$($_.DeviceID) Free Space %" ; 'Value' = [math]::Round($_.FreeSpace / $_.Size,2)*100 })))
            }
        }
        ## KMS licence status
        ##$kms = Get-CimInstance -ClassName SoftwareLicensingProduct -ComputerName $unqualifiedMachineName
        [string]$licenceState = 'Unknown'
        if( $kms )
        {
            $licenceState = switch( $kms | Where-Object { $_.LicenseStatus -ne 0 } | Select -ExpandProperty LicenseStatus ) 
            {
                0 { 'Unlicenced' }
                1 { 'Licenced' }
                2 { 'Out of Box Grace' }
                3 { 'Out of Tolerance Grace' }
                4 { 'Non Genuine Grace' }
                5 { 'Notification' }
                6 { 'Extended Grace' }
            }
            $WPFHealth.AddChild( $( New-Object -TypeName PSCustomObject -Property (@{ 'Name' = 'KMS Server' ; 'Value' = ($kms | Where-Object { $_.KeyManagementServiceMachine } | select -ExpandProperty KeyManagementServiceMachine) })  ) )
        }
        $WPFHealth.AddChild( $( New-Object -TypeName PSCustomObject -Property (@{ 'Name' = 'KMS licence status' ; 'Value' = $licenceState })  ) )
        Progress-Window $null
        $healthForm.Show()
    }
}

Function Get-DateTime (
        [string]$title ,
        [ref]$dateTimeChosen 
    )
{
    [bool]$allGood = $false
    $datePicker = Load-GUI $datePickerXAML
    if( $datePicker )
    {
        $eventBlock = { 
                    $_.Handled = $true
                    if( $_.psobject.Properties.name -match '^OriginalSource$' -and $_.OriginalSource.psobject.Properties.name -match '^content$' -and ($_.OriginalSource.Content -eq 'OK' -or $_.Source.OriginalSource -eq 'Cancel' ))
                    {
                        if( $_.OriginalSource.Content -eq 'OK' )
                        {
                            $this.DialogResult = $true
                        }
                        $this.Close()
                    }
                }
        $eventHandler = [Windows.RoutedEventHandler]$eventBlock
        $datePicker.AddHandler([Windows.Controls.Button]::ClickEvent, $eventHandler)
        $datePicker.Title = $title
        if( $datePicker.ShowDialog() )
        {
            [datetime]$datePicked = $WPFDatePicked.SelectedDate
            [datetime]$timePicked = Get-Date

            if( $datePicked -and [datetime]::TryParse( $WPFtxtTime.Text , [ref] $timePicked ) )
            {
                $dateTimeChosen.Value = Get-Date -Date $datePicked -Hour $timePicked.Hour -Minute $timePicked.Minute -Second $timePicked.Second
                $allGood = $true
            }
        }
    }
    $allGood
}

Function Get-FileName( [string]$initialDirectory , [bool]$isOpen = $false )
{   
    [string]$result = $null
    if( $isOpen )
    {
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    }
    else
    {
        $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    }
    if( $OpenFileDialog )
    {
        if( ! [string]::IsNullOrEmpty( $initialDirectory ) )
        {
            $OpenFileDialog.initialDirectory = $initialDirectory
        }
        $OpenFileDialog.filter = "CSV files (*.csv)| *.csv"
        $result = $OpenFileDialog.ShowDialog()
        if( $result -eq 'OK' )
        {
            $result = $OpenFileDialog.filename
        }
    }
    $result
}


Function Process-Events
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateSet('Logon','Logoff','AllSession','Boot','Range')]
        [string]$eventsFrom ,
        $control
    )
    
    ## if none selected then we get the first one
    $selected = $control.SelectedItem
    if( ! $selected )
    {
        return
    }
    
    ## gives context as we can be called from different list views
    if( ! $control )
    {
        return
    }

    [datetime]$start = Get-Date
    [datetime]$end = Get-Date
    [string]$computer = $null
    
    if( $selected.psobject.Properties.name -match 'Machine' )
    {
        ## Get selected session and user so can go to server and get event logs for that period
        $computer = ( $selected.Machine -split '\\')[-1]
    }
    elseif( $selected.psobject.Properties.name -match 'Full Name' )
    {
        $computer = ( $selected.'Full Name' -split '\\')[-1]
    }

    if( ! $computer )
    {
        return
    }

    if( $eventsFrom -eq 'Range' )
    {      
        if( Get-DateTime -title 'Enter start date/time' -dateTimeChosen ( [ref]$start ) )
        {
            if( ! ( Get-DateTime -title 'Enter end date/time' -dateTimeChosen ( [ref]$end ) ) )
            {
                [Windows.MessageBox]::Show( "Failed to get valid start time" , "Information" , 'OK' ,'Error' )
                return
            } 
        } 
        else
        {
            [Windows.MessageBox]::Show( "Failed to get valid start time" , "Information" , 'OK' ,'Error' )
            return
        }
    }
    elseif( $eventsFrom -eq 'boot' )
    {
        ## we already have the boot time from the list view so don't look up again
        if( $selected.psobject.Properties.name -match 'Value3' -and $selected.Value3 )
        {
            $start = Get-Date $selected.Value3
        }
        else
        {
            $start = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computer | Select-Object -ExpandProperty LastBootUpTime
        }
        if( ! $start )
        {
            [Windows.MessageBox]::Show( "Failed to get boot time of $computer" , "Information" , 'OK' ,'Error' )
            return
        }
        else
        {
            $end = $start.AddSeconds( $global:secondsWindow ) 
        }
    }
    elseif( $eventsFrom -eq 'logon' )
    {
        $start = $selected.StartDate
        $end = $start.AddSeconds( $global:secondsWindow ) 
        
        ##  make less if logoff time exists and is less than this
        if( $selected.PSobject.Properties.name -match '^EndDate' -and $selected.EndDate -and $selected.EndDate -lt $end )
        {
            $end = $selected.EndDate
        }
    }
    elseif( $eventsFrom -eq 'AllSession' )
    {
        $start = $selected.StartDate
        if( $selected.PSobject.Properties.name -match '^EndDate' -and $selected.EndDate )
        {
            $end = $selected.EndDate
        }
        else ## Session still active
        {
            $end = Get-Date
        }
    }
    elseif( $eventsFrom -eq 'logoff' )
    {
        if( $selected.PSobject.Properties.name -match '^EndDate' -and ! $selected.EndDate )
        {
            [Windows.MessageBox]::Show( "Session not logged off" , "Information" , 'OK' ,'Warning' )
            return
        }

        $end = $selected.EndDate
        $start = $end.AddSeconds( -$global:secondsWindow )  
    }

    [array]$results = @()
    $oldCursor = $mainForm.Cursor
    $mainForm.Cursor = [Windows.Input.Cursors]::Wait
    
    [string]$startInText = Get-Date $start -Format G
    [string]$endInText = Get-Date $end -Format G

    Progress-Window "Fetching events from $computer`nfrom $startInText`nto   $endInText"
    [int]$eventlogs = 0
    $results = @( Get-WinEvent -ListLog * -EA silentlycontinue -ComputerName $computer | Where-Object { $_.RecordCount -gt 0 } |  ForEach-Object `
    { 
        Write-Verbose "$computer : $($_.LogName) $($_.recordcount) $($_.lastwritetime)"
        ## TODO make the Level configurable since AppSense failure events are reported as information
        ##$result = @( Get-WinEvent -ComputerName $computer -FilterHashtable @{'Logname'=$_.LogName;StartTime=$start;EndTime=$end;Level=1,2,3 } -EA SilentlyContinue  )
        $result = @( Get-WinEvent -ComputerName $computer -FilterHashtable @{'Logname'=$_.LogName;StartTime=$start;EndTime=$end } -EA SilentlyContinue  )
        if( $result -and $result.Count )
        {
            $eventlogs++
            $result
        }
        $result = $null
    } )
    Progress-Window $null
    if( $results -and $results.Count )
    {
        $events = @( $results | Select LogName,ProviderName,TimeCreated,LevelDisplayName,Id,@{n='User';e={([Security.Principal.SecurityIdentifier]($_.UserId)).Translate([Security.Principal.NTAccount]).Value}},Message | Sort-Object -Property TimeCreated | Out-GridView -Title "$($results.count) events on $computer from $eventsFrom at $startInText to $endInText" -PassThru )
        if( $events -and $events.Count )
        {
            ## Offer to save them
            if( [Windows.MessageBox]::Show( "Save $($events.Count) selected events to csv file?" , 'Events' , 'YesNo' ,'Question' ) -eq 'yes' )
            {
                [string]$fileName = Get-FileName
                if( ! [string]::IsNullOrEmpty( $fileName ) )
                {
                    [bool]$writeIt = $true 
                    if( Test-Path -Path $fileName -ErrorAction SilentlyContinue )
                    {
                        $answer = [Windows.MessageBox]::Show( "`"$fileName`" exists - replace?" , 'Events' , 'YesNo' ,'Warning' ) 
                        $writeIt = ( $answer -eq 'yes' )
                    }
                    if( $writeIt )
                    {
                        $events | Export-Csv -Path $fileName -NoTypeInformation -Encoding ASCII 
                        if( ! $? -or ! ( Test-Path -Path $fileName -ErrorAction SilentlyContinue ) )
                        {
                            [Windows.MessageBox]::Show( "Failed to save evetns to `"$fileName`"" , 'Event Save Error' , 'OK' ,'Error' )
                        }
                    }
                }
            }
            $events | clip.exe
        }
    }
    else
    {
        $booted = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computer | Select-Object -ExpandProperty LastBootUpTime
        [Windows.MessageBox]::Show( "No events found between $startInText and $endInText`nBooted at $(Get-Date $booted -Format G)" , "Information" , 'OK' ,'Information' )
    }
    $mainForm.Cursor = $oldCursor
}

Function Save-Config( $form )
{
    if( ! ( Test-Path $configRegKey -ErrorAction SilentlyContinue ) )
    {
        New-Item -Path $configRegKey -Force
    }
    ## What if multiple?
    [string]$oldddc = $global:ddc
    $global:ddc = $WPFtxtDDC.Text
    Set-ItemProperty -Path $configRegKey -Name 'DDC' -Value $global:ddc
    
    $global:PVS = $WPFtxtPVS.Text
    Set-ItemProperty -Path $configRegKey -Name 'PVS' -Value $global:PVS

    $global:Hypervisor = $WPFtxtHypervisor.Text
    Set-ItemProperty -Path $configRegKey -Name 'Hypervisor' -Value $global:Hypervisor

    $global:AMC = $WPFtxtAppSense.Text
    Set-ItemProperty -Path $configRegKey -Name 'AMC' -Value $global:AMC
    
    $global:MessageText = $WPFtxtMessageText.Text
    Set-ItemProperty -Path $configRegKey -Name 'MessageText' -Value $global:MessageText

    $global:secondsWindow = $WPFtxtSeconds.Text
    Set-ItemProperty -Path $configRegKey -Name 'seconds' -Value $global:secondsWindow

    if( $oldddc -ne $global:ddc )
    {       
        [string]$thisddc = $global:ddc
        $global:baseURI = Invoke-Command -ComputerName $global:ddc -ScriptBlock { Add-PSSnapin 'Citrix.Configuration.Admin.*' ; Get-ConfigRegisteredServiceInstance -ServiceType monitor | ?{ $_.Address -match "/$($using:thisddc)[^a-z0-9_\-]" } | Sort Version -Descending | select -first 1 -ExpandProperty address }
        if( [string]::IsNullOrEmpty( $baseURI ) )
        {
            Write-Warning "Failed to get new base URI from $global:ddc"
        }
    }
}

Function Show-Config
{
    $settingsForm = Load-GUI $configurationWindowXML
    if( $settingsForm )
    {
        if( $DebugPreference -eq 'Inquire' )
        {
            Get-Variable -Name WPF*
        }
        $WPFbtnConfigurationOk.Add_Click({ Save-Config( $settingsForm );$settingsForm.Close() })
        $WPFbtnConfigurationOk.IsDefault = $true
        $WPFbtnConfigurationCancel.Add_Click({ $settingsForm.Close() })
        $WPFbtnConfigurationCancel.IsCancel = $true

        $WPFtxtDDC.Text = $global:ddc
        $WPFtxtAppSense.Text = $global:AMC
        $WPFtxtHypervisor.Text = $global:hypervisor -join ','
        $WPFtxtPVS.Text = $global:PVS
        $WPFtxtSeconds.Text = $global:secondsWindow
        $WPFtxtMessageText.Text = $global:messageText
        $WPFchkHttps.IsChecked = ( $global:protocol -eq 'https' )
        $null = $settingsForm.ShowDialog()
    }
}

<#
$WPFSessions.add_MouseDoubleClick({
    Process-Events
})
#>

Function Sort-Columns( $control )
{
    $view =  [Windows.Data.CollectionViewSource]::GetDefaultView($control.ItemsSource)
    [string]$direction = 'Ascending'
    if(  $view -and $view.SortDescriptions -and $view.SortDescriptions.Count -gt 0 )
    {
	    $sort = $view.SortDescriptions[0].Direction
	    $direction = if( $sort -and 'Descending' -eq $sort){'Ascending'} else {'Descending'}
	    $view.SortDescriptions.Clear()
    }

    [string]$column = $_.OriginalSource.Column.DisplayMemberBinding.Path.Path ## has to be name of the binding, not the header unless no binding
    if( [string]::IsNullOrEmpty( $column ) )
    {
        $column = $_.OriginalSource.Column.Header
    }
	$sortDescription = New-Object ComponentModel.SortDescription($column, $direction) 
	$view.SortDescriptions.Add($sortDescription)
}

Function Set-RadioButtonState( $control , [switch]$enable )
{
    $control.Children | Where { $_ -is [windows.controls.radiobutton]  } | ForEach-Object { $_.IsEnabled = $enable }
}

Function Get-LiveSessions( [string]$machineName , [string]$unqualifiedMachineName )
{
    $WPFSessions.ItemsSource = $null
    $sessions = @( Get-BrokerSession -MachineName $machineName )
    if( $sessions -and $sessions.Count )
    {
        ## Change the columns
        $wpfsessions.View.Columns[0].Header = 'Start Time'
        $wpfsessions.View.Columns[1].Header = 'User Name'
        $wpfsessions.View.Columns[1].Width = 100
        $wpfsessions.View.Columns[2].Header = 'Protocol'
        $wpfsessions.View.Columns[2].Width = 50
        $WPFlabelSessions.Content = "$($sessions.Count) sessions"
        $sessions|ForEach-Object `
        { 
            ## we're hijacking the EndTime column for when this control is used for users
            $_| Add-Member -MemberType NoteProperty -Name Column1 -Value ( Get-Date $_.StartTime -Format G)
            $_| Add-Member -MemberType NoteProperty -Name Column2 -Value ( $_.UserName -split '\\' )[-1]
            $_| Add-Member -MemberType NoteProperty -Name Column3 -Value $_.Protocol
            ## Event grabbing uses this field
            $_| Add-Member -MemberType NoteProperty -Name StartDate -Value $_.StartTime
            $_| Add-Member -MemberType NoteProperty -Name Machine -Value $unqualifiedMachineName
            $_| Add-Member -MemberType NoteProperty -Name FullMachine -Value $machineName
            ## These are what the list view for sessions uses
            $_| Add-Member -MemberType NoteProperty -Name Value -Value ( $_.username -split '\\' )[-1]
            $_| Add-Member -MemberType NoteProperty -Name Name -Value $_.username
        }
        if( $wpfsessions.ItemsSource )
        {
            $WPFSessions.ItemsSource = $null
        }
        else
        {
            $wpfsessions.Items.Clear()
        } 
        $wpfsessions.ItemsSource = $sessions
    }
    else
    {
        [Windows.MessageBox]::Show( "No live sessions found for $machineName" , "Session Error" , 'OK' ,'Error' )
        return
    }
}

<#
Function
Handle-ContextMenu ( $sender , $eventArgs , $action )
{
    $action
    $eventArgs.Handled = $true
}
#>

Function Show-PVSDdevices
{
    Progress-Window "Fetching PVS devices from $global:pvs"

    [hashtable]$devices = Get-PVSDevices -pvsservers $global:pvs -ddcs $global:ddc -dns -tags -adgroups '.*' -noProgress -maxRecordCount $maxRecordCount -noOrphans -modules $modules ## no VMware yet

    Progress-Window $null

    if( $devices -and $devices.Count )
    {
        [string[]]$columns = @( 'Name','DomainName','Description','PVS Server','DDC','SiteName','CollectionName','Machine Catalogue','Delivery Group','Load Index','Load Indexes','Registration State','Maintenance_Mode','User_Sessions','devicemac','active','enabled',
            'Store Name','Disk Version Access','Disk Version Created','AD Account Created','AD Account Modified','Domain Membership','AD Last Logon','AD Description','Disk Name','Booted off vdisk','Booted Disk Version','Vdisk Production Version','Vdisk Latest Version',
            'Latest Version Description','Override Version', 'Retries','Booted Off','Device IP','Booted off latest','Disk Description','Cache Type','Disk Size (GB)','vDisk Size (GB)','Write Cache Size (MB)' , 'IPv4 address' , 
            'Tags' , 'AD Groups' , 'Boot_Time' ,'Available Memory (GB)','Committed Memory %','Free disk space %','CPU Usage %' )

        $choices = $devices.GetEnumerator()| ForEach-Object { $_.Value } | Select $columns | Out-GridView -Title "$($devices.count) PVS devices via $global:PVS & ddc $global:ddc" -PassThru
        ## Now what shall we do with the device(s) returned?
    }
    else
    {
        [Windows.MessageBox]::Show( "No PVS devices found via $global:PVS" , "PVS Devices" , 'OK' ,'Error' )
    }
}

Function Connect-Hypervisor( [string[]] $hypervisors )
{    
    $null = Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false

    Connect-VIserver -Server $hypervisors
}

$WPFbtnPVSDevices.Add_Click({
    Show-PVSDdevices
})

Function Get-UserSessions( [string]$machineName )
{
    ## Now get session count - do it this was so we don't get an error if there are no users
    [int]$users = 0
    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = "quser.exe"
    $pinfo.Arguments = "/server:$machineName"
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.CreateNoWindow = $true
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $pinfo
    $null = $process.Start()
    [string[]]$output = $process.StandardOutput.ReadToEnd() -split "`r`n" | Where-Object { ! [string]::IsNullOrEmpty( $_.Trim() ) }
    $process.WaitForExit()
    if( $output -and $output.Count )
    {
        $users = $output.Count - 1 ## account for heading
    }
    $users
}

$WPFbtnGetMachines.Add_Click({   
    [string]$deliveryGroup = $null
    [string]$machineCatalogue = $null
    [string]$machineName = $null
    [bool]$poweredOn = $false
    [bool]$multiSession = $false
    [bool]$singleSession = $false
    [bool]$script:cancelled = $false
    
    $_.Handled = $true

    ## Prompt filtering on delivery group/machine catalogue for multisession/singlesession (XenApp vs XenDesktop)     
    $machineSelectionForm = Load-GUI $machineSelectionWindowXML

    if( $machineSelectionForm )
    {
        if( $DebugPreference -eq 'Inquire' )
        {
            Get-Variable -Name WPF*
        }
        ## Set event handlers for button clicks so we can populate the dropdown lists
        $WPFchkDeliveryGroup.add_Checked({
            $wpfcomboDeliveryGroup.Clear()
            $wpfcomboDeliveryGroup.IsEnabled = $true
            ## this is so we can show machines not in a delivery group
            $wpfcomboDeliveryGroup.items.add( '<None>' )
            Get-BrokerDesktopGroup -AdminAddress $global:ddc | Select -ExpandProperty Name | Sort | ForEach-Object { $wpfcomboDeliveryGroup.items.add( $_ ) }
        })
        $WPFchkDeliveryGroup.add_Unchecked({
            $wpfcomboDeliveryGroup.IsEnabled = $false
        })
        $WPFchkMachineCatalogue.add_Checked({
            $WPFcomboMachineCatalogue.Clear()
            $WPFcomboMachineCatalogue.IsEnabled = $true
            $WPFcomboMachineCatalogue.items.add( '<None>' )
            Get-BrokerCatalog -AdminAddress $global:ddc | Select -ExpandProperty Name | Sort | ForEach-Object { $WPFcomboMachineCatalogue.items.add( $_ ) }
        })
        $WPFchkMachineCatalogue.add_Checked({
            $WPFcomboMachineCatalogue.IsEnabled = $true
        })
        $WPFbtnConfigurationOk.Add_Click({ $cancelled = $false ;$machineSelectionForm.Close()
        })
        $WPFbtnConfigurationOk.IsDefault = $true
        $WPFbtnConfigurationCancel.Add_Click({ 
            $script:cancelled = $true ## needs to be scoped otherwise local to handler so doesn't propagate
            $machineSelectionForm.Close() 
        })
        $WPFchkMachineName.add_Checked({ ## enable machine source radio buttons 
            Set-RadioButtonState $wpfStkMachinesFrom -enable
        })
        $WPFchkMachineName.add_unChecked({ ## disable mchine source radio buttons  
            Set-RadioButtonState $wpfStkMachinesFrom
        })
        Set-RadioButtonState $wpfStkMachinesFrom -enable:$($WPFchkMachineName.IsChecked)

        ## if AD source selected then delivery group not valid
        $WPFradMachinesFromAD.add_Checked({
            $WPFchkDeliveryGroup.IsEnabled = $false
            $WPFcomboDeliveryGroup.IsEnabled = $false
        })
        $WPFradMachinesFromAD.add_Unchecked({
            $WPFchkDeliveryGroup.IsEnabled = $true
            $WPFcomboDeliveryGroup.IsEnabled = $true
        })
        $WPFbtnConfigurationCancel.IsCancel = $true
        $null = $machineSelectionForm.ShowDialog()

        if( $script:cancelled )
        {
            return
        }
    }
   
    [string]$extraInfo = ''
    $deliveryGroup = $( if( $WPFchkDeliveryGroup.IsChecked ) { $extraInfo += " delivery group `"$($WPFcomboDeliveryGroup.SelectedItem)`""; $WPFcomboDeliveryGroup.SelectedItem } else { $null } )
    $machineCatalogue =  $( if( $WPFchkMachineCatalogue.IsChecked ) { $extraInfo += " catalog `"$($WPFcomboMachineCatalogue.SelectedItem)`""; $WPFcomboMachineCatalogue.SelectedItem } else { $null } )
    $poweredOn = $WPFchkPoweredOn.IsChecked
    $multiSession = $WPFchkMultiSession.IsChecked
    $singleSession = $WPFchkSingleSession.IsChecked
    $machineName = $( if ( $WPFchkMachineName.IsChecked ) { $WPFtxtMachineName.Text } else { $null } )
    [string]$source = $global:ddc
    [array]$items = @()

    ## if radio button for AD pressed
    [string]$machineSourceButton = $wpfStkMachinesFrom.Children | Where { $_ -is [windows.controls.radiobutton] -and $_.IsChecked } | Select -ExpandProperty Name
    if( ! [string]::IsNullOrEmpty( $machineSourceButton ) )
    {
        ## Should we check if Citrix servers? Probably not so tool can be used for other things
        if( $machineSourceButton -eq 'radMachinesFromAD' )
        {
            if( ! $WPFchkMachineName.IsChecked -or [string]::IsNullOrEmpty( $machineName ) )
            {
                [Windows.MessageBox]::Show( "Must specify machine name in AD mode" , "Filter Problem" , 'OK' ,'Error' )
                return
            }
            else
            {
                ## TODO set filter based on multisession/singlesession check boxes
                $computers = Get-ADComputer -Filter *| Where-Object { $_.Name -match $machineName }|Get-ADComputer -Properties *
                $source = 'Active Directory'
            }
        }
        elseif( $machineSourceButton -eq 'radMachinesFromHypervisor' )
        {
            if( ! $WPFchkMachineName.IsChecked -or [string]::IsNullOrEmpty( $machineName ) )
            {
                [Windows.MessageBox]::Show( "Must specify machine name in hyerpvisor mode" , "Filter Problem" , 'OK' ,'Error' )
                return
            }
            if( ! $global:hypervisorConnected )
            {
                $global:hypervisorConnected = Connect-Hypervisor $global:hypervisor
                if( ! $global:hypervisorConnected )
                {
                    [System.Windows.MessageBox]::Show( "Unable to connect to hypervisor $($global:hypervisor -join ' ' )" , "Information" , 'OK' ,'Error' )
                    return
                }
            }
            $computers = Get-VM | Where-Object { $_.Name -match $machineName }
            $source = 'hypervisor'
        }
    }
    ## else will default to from Citrix

    Progress-Window "Fetching machines from $source"

    if( $WPFUsers.ItemsSource )
    {
        $WPFUsers.ItemsSource = $null
    }
    else
    {
        $WPFUsers.Items.Clear()
    }

    [hashtable]$params = @{}
    
    if( ! [string]::IsNullOrEmpty( $deliveryGroup ) -and $deliveryGroup -ne '<None>' )
    {
        $params.Add( 'DesktopGroupName' , $deliveryGroup )
    }
    if( ! [string]::IsNullOrEmpty( $machineCatalogue )  -and $machineCatalogue -ne '<None>' )
    {
        $params.Add( 'CatalogName' , $machineCatalogue )
    }
    if( $multiSession -and ! $singleSession )
    {
        $params.Add( 'SessionSupport' , 'MultiSession' )
    }
    elseif( $singleSession -and ! $multiSession )
    {
        $params.Add( 'SessionSupport' , 'SingleSession' )
    }

    switch( $wpfStkMaintenanceMode.Children | Where {
            $_ -is [windows.controls.radiobutton] -and $_.IsChecked
        } | Select -ExpandProperty Name )
    {
         'radInMaintenanceMode'    { $params.Add( 'InMaintenanceMode' , $true ) ; $extraInfo += ' maintenance mode' }    
         'radNotInMaintenanceMode' { $params.Add( 'InMaintenanceMode' , $false ) ; $extraInfo += ' not maintenance mode'  }
         ## 'radInMaintenanceMode' ignore this as default is to show machines regardless of maintenance mode
    }

    if( $machineSourceButton -eq 'radMachinesFromCitrix' )
    {
        switch( $wpfstkRegistrationState.Children | Where {
                $_ -is [windows.controls.radiobutton] -and $_.IsChecked
            } | Select -ExpandProperty Name )
        {
             'radRegistered'    { $params.Add( 'RegistrationState' , 'Registered' ) ; $extraInfo += " registered" }    
             'radUnregistered' { $params.Add( 'Filter' ,"{ RegistrationState -ne 'Registered' }" ) ; $extraInfo += " unregistered" }
             ## 'radEitherRegisteredUnregistered' ignore this as default is to show machines regardless of registration state
        }

        $machines = @( Get-BrokerMachine -AdminAddress $global:ddc @params -MaxRecordCount $maxRecordCount )
        if( ! [string]::IsNullOrEmpty( $machineName ) )
        {
            $machines = @( $machines | Where-Object { $_.MachineName -match $machineName } )
        }
        if( $deliveryGroup -and $deliveryGroup -eq '<None>' )
        {
            $machines = @( $machines | Where-Object { ! $_.DesktopGroupName -or ! $_.DesktopGroupName.Length } )
        }
        
        Write-Verbose "Got $($machines.Count) machines"
        [int]$counter = 1
        $items = @( $machines | ForEach-Object `
            {
                [string]$shortName = ($_.MachineName -split '\\')[-1]
                [int]$sessionCount = $_.SessionCount
                if( ! $wpfchkUsersConnected.IsChecked -or $sessionCount )
                {
                    $object = New-Object -TypeName PSCustomObject -Property (@{ `
                        'Value' = $shortName ; 
                        'Full Name' = $_.MachineName  ; ## not displayed but we need it later for Citrix cmdlets
                        'Value2' = $SessionCount })
                    if( $_.PowerState -ne 'Off' )
                    {
                        try
                        {
                            Write-Verbose "$counter / $($machines.Count) : getting boot time from $shortName"
                            [datetime]$booted = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $shortName | Select-Object -ExpandProperty LastBootUpTime )
                            $object | Add-Member -MemberType NoteProperty -Name 'Value3' -Value (Get-Date $booted -Format G)
                        }
                        catch{}
                    }
                    $object
                }
                $counter++
            })
    }
    else ##if( $machineSourceButton -eq 'radMachinesFromAD' )
    {
        ## Now build our items array that is put into list view
        ## For each computer for AD we check if it is in a machine catalogue because if not it may be an orphaned citrix server, depending on name
        [string[]]$machinesInCatalogues = if(  $machineCatalogue -eq '<None>' ) { Get-BrokerMachine -AdminAddress $global:ddc -MaxRecordCount $maxRecordCount | Select -ExpandProperty MachineName } else { $null }
        Write-Verbose "Got $($computers.Count) computers from AD to check against $($machinesInCatalogues.Count) from Citrix"
        $items = @( $computers | ForEach-Object `
        {
            ## may need to specify a splitter character like _ if VM names aren't just NetBIOS names
            [string]$fullName = If( $machineSourceButton -eq 'radMachinesFromAD' ) { ( $_.DNSHostName -split '\.' )[1] + '\' +  ( $_.DNSHostName -split '\.' )[0] } else { $env:USERDOMAIN + '\' + $_.Name }
            if( ! $machinesInCatalogues -or $machinesInCatalogues -notcontains $fullName )
            {
                Write-Verbose "Processing $fullName as not in catalogue on $global:ddc"
                ## ToDo should we also check against PVS - tick box?
                $users = Get-UserSessions $_.Name

                if( ! $wpfchkUsersConnected.IsChecked -or $users )
                {
                    ## not in a machine catalogue so not in Citrix
                    $object = New-Object -TypeName PSCustomObject -Property (@{ `
                        'Value' = $_.Name ; 
                        'Full Name' = $fullName  ; ## not displayed but we need it later for Citrix cmdlets
                        'Value2' = $users })
                    ## TODO do we go to hypervisor to get power state?
                    if( $true )
                    {
                        $object | Add-Member -MemberType NoteProperty -Name 'Value3' -Value (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $_.Name | Select-Object -ExpandProperty LastBootUpTime )
                    }
                    $object
                }
            }
        })
    }
        
    $sortMachineColumnsHandler = { Sort-Columns $_.Source }
    $eventHandler = [Windows.RoutedEventHandler]$sortMachineColumnsHandler
    $wpfusers.AddHandler([Windows.Controls.GridViewColumnHeader]::ClickEvent, $eventHandler)

    $wpfusers.View.Columns[0].Header = 'Machine Name'
    $wpfusers.View.Columns[1].Header = 'Sessions'
    $wpfusers.View.Columns[1].Width = 50
    $wpfusers.View.Columns[2].Header = 'Last Boot'
    $wpfusers.View.Columns[2].Width = 200
    ## Set up for sorting on columns 

    $wpfusers.ItemsSource = $items
    $WPFlabelUsers.Content = "$($WPFUsers.ItemsSource.Count) machines $extraInfo"
    <#
    ## This gets machines from DDC
    Get-MachinesFromDDC
    if( ! $machines -or ! $machines.Count )
    {
        Progress-Window $null
        [System.Windows.MessageBox]::Show( "No machines found via $global:ddc" , "Information" , 'OK' ,'Information' )
        return
    }
    ## Populate the user's list with machines

    $WPFlabelUsers.Content = "$($machines.Count) machines"
    $machines.GetEnumerator() | Select -ExpandProperty Value | Sort | ForEach-Object `
    {
        $WPFUsers.AddChild( (  New-Object -TypeName PSCustomObject -Property (@{ 'Value' = ($_ -split '\\')[-1] ; 'Full Name' = $_ })  ) )
    }
    #>
    ## Set context menu for machine content not user
    <#
    $contextmenu = New-Object -TypeName System.Windows.Controls.ContextMenu
    $menuItem = New-Object -TypeName System.Windows.Controls.MenuItem
    $menuItem.Header = 'Test 1'
    $menuItem.Add_Click({ 
        param($sender, $eventArgs)
        Handle-ContextMenu $sender $eventArgs 'Action 1'                     
    })
    $contextmenu.AddChild( $menuItem )
    $menuItem = New-Object -TypeName System.Windows.Controls.MenuItem
    $menuItem.Header = 'Test 2'
    $menuItem.Add_Click({ 
        param($sender, $eventArgs)
        Handle-ContextMenu $sender $eventArgs 'Action 2'                      
    })
    $contextmenu.AddChild( $menuItem )

    $WPFUsers.ContextMenu = $contextmenu
    #>
    ## Need to enable the context menus we need for machines which all of them
    $wpfusers.ContextMenu.Items.GetEnumerator() | ForEach-Object `
    {
        $_.IsEnabled = $true
    }
    Progress-Window $null
})

$WPFbtnTest.Add_Click({
    $WPFUsers.Items.Clear()
    ForEach( $dummy in @( 'one','two','three','simon.mulrain','guy.leech' ) )
    {
        $WPFUsers.AddChild((  New-Object -TypeName PSCustomObject -Property (@{ 'Value' = "$dummy" })  ) )
    }
    $_.Handled = $true
})

Function Process-UserDoubleClicked
{
    if( $WPFlabelUsers.Content -and $WPFlabelUsers.Content -match 'users' )
    {
        Get-Sessions
    }
    else
    {
        $selected = $this.SelectedItem
        if( $selected )
        {
            Get-LiveSessions $selected.'Full Name'  $selected.value
        }
    }
    $_.Handled = $true
}

$WPFUsers.add_MouseDoubleClick({
    Process-UserDoubleClicked
})

$WPFSessionInfoContextMenu.Add_Click({
    Show-SessionInformation
    $_.Handled = $true
})

$WPFSessionProcessesContextMenu.Add_Click({
    Show-UserInformation -showProcesses 
    $_.Handled = $true
})

$WPFAllProcessesContextMenu.Add_Click({
    Show-UserInformation -showProcesses -allProcesses
    $_.Handled = $true
})

$WPFMachineInfoContextMenu.Add_Click({
    $selected = $WPFSessions.SelectedItem
    if( $selected )
    {
        Show-MachineInformation $selected.FullMachine  $selected.Machine
    }
    $_.Handled = $true
})

$WPFUserInfoContextMenu.Add_Click({
    Show-UserInformation
    $_.Handled = $true
})

$WPFUserUserInfoContextMenu.Add_Click({
    if( $WPFlabelUsers.Content -and $WPFlabelUsers.Content -match 'users' )
    {
        Show-UserInformation
    }
    elseif( $WPFlabelUsers.Content -and $WPFlabelUsers.Content -match 'machines' )
    {
        $selected = $WPFUsers.SelectedItem
        Show-MachineInformation $selected.'Full Name' $selected.Value
    }
    $_.Handled = $true
})

$WPFbtnSettings.Add_Click({
    Show-Config
    $_.Handled = $true
})

$WPFLogonEventsContextMenu.Add_Click({
    Process-Events Logon $WPFSessions
    $_.Handled = $true
})

$WPFLogoffEventsContextMenu.Add_Click({
    Process-Events Logoff $WPFSessions
    $_.Handled = $true
})

$WPFAllSessionEventsContextMenu.Add_Click({
    Process-Events AllSession $WPFSessions
    $_.Handled = $true
})

$WPFBootEventsContextMenu.Add_Click({
    Process-Events Boot $WPFSessions
    $_.Handled = $true
})

$WPFMachineBootEventsContextMenu.Add_Click({
    Process-Events Boot $wpfusers
    $_.Handled = $true
})

$wpfEventsInRangeContextMenu.Add_Click({
    Process-Events Range $wpfusers
    $_.Handled = $true
})

$wpfEventsInRangeUserContextMenu.Add_Click({
    Process-Events Range $WPFSessions
    $_.Handled = $true
})

$WPFMessageContextMenu.Add_Click({
    Perform-Action -Context User -Action Message
    $_.Handled = $true
})

$WPFLogoffContextMenu.Add_Click({
    Perform-Action -Context User -Action Logoff
    $_.Handled = $true
})

$WPFShadowContextMenu.Add_Click({
    Perform-Action -Context User -Action Shadow
    $_.Handled = $true
})

$WPFDisconnectContextMenu.Add_Click({
    Perform-Action -Context User -Action Disconnect
    $_.Handled = $true
})

$WPFRestartServiceContextMenu.Add_Click({
    Perform-Action -Context Machine -Action RestartService
    $_.Handled = $true
})

$WPFMaintenanceModeContextMenu.Add_Click({
    Perform-Action -Context Machine -Action MaintenanceMode
    $_.Handled = $true
})

$WPFLogonContextMenu.Add_Click({
    Perform-Action -Context Machine -Action Logon
    $_.Handled = $true
})

$WPFRebootContextMenu.Add_Click({
    Perform-Action -Context Machine -Action Reboot
    $_.Handled = $true
})

Function Perform-Action
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateSet('User','Machine')]
        [string]$Context ,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Message','Logoff','Shadow','Disconnect','RestartService','MaintenanceMode','Logon','Reboot')]
        [string]$Action 
    )

    if( $Context -eq 'user' )
    {    
        $selected = $WPFSessions.SelectedItem
        if( ! $selected )
        {
            [Windows.MessageBox]::Show( "Please select a session" , "Information" , 'OK' ,'Error' )
            return
        }
        elseif( $selected.PSobject.Properties.name -match '^EndDate' -and $selected.EndDate )       
        {
            [Windows.MessageBox]::Show( "Session has ended" , "Information" , 'OK' ,'Error' )
            return
        }
        [string]$computer = ( $selected.Machine -split '\\')[-1]
        [string]$fullUsername = $users[ $selected.UserId ]
        $sessions = @( Get-BrokerSession -UserName $fullUsername -AdminAddress $global:ddc | Where-Object { $_.MachineName -eq $Selected.FullMachine } )
        if( ! $sessions -or ! $sessions.Count )    
        {
            [Windows.MessageBox]::Show( "No sessions found for user $fullUsername on $computer" , "Information" , 'OK' ,'Error' )
            return
        }
        if( $Action -eq 'Logoff' )
        {
            $answer = [Windows.MessageBox]::Show( "Logoff $fullUsername on $computer" , "Confirm Action" , 'YesNo' ,'Question' )
            if( $answer -eq 'Yes' )
            {
                $sessions | Stop-BrokerSession -AdminAddress $global:ddc
                if( ! $? )
                {
                    [Windows.MessageBox]::Show( "Failed to logoff $fullUsername from $computer`n$($Error[0])" , "Result" , 'OK' , 'Error' )
                }
            }
        } 
        elseif( $Action -eq 'Disconnect' )
        {
            $answer = [Windows.MessageBox]::Show( "Disconnect $fullUsername on $computer" , "Confirm Action" , 'YesNo' ,'Question' )
            if( $answer -eq 'Yes' )
            {
                $sessions | Disconnect-BrokerSession -AdminAddress $global:ddc
                if( ! $? )
                {
                    [Windows.MessageBox]::Show( "Failed to disconnect $fullUsername from $computer`n$($Error[0])" , "Result" , 'OK' , 'Error' )
                }
            }
        }
        elseif( $Action -eq 'Message' )
        {
            $messageForm = Load-GUI $messageWindowXML
            if( $messageForm )
            {
                ## Load up caption and text and set callbacks
                $WPFtxtMessageCaption.Text = "Message from $env:USERNAME at $(Get-Date -Format F)"
                $WPFtxtMessageBody.Text = $global:messageText
                $WPFbtnMessageOk.Add_Click({
                    $sessions | Send-BrokerSessionMessage -AdminAddress $global:ddc -Title $WPFtxtMessageCaption.Text -Text $WPFtxtMessageBody.Text -MessageStyle ($WPFcomboMessageStyle.SelectedItem.Content) ;$messageForm.Close() 
                }) 
                $WPFbtnMessageOk.IsDefault = $true
                $WPFbtnMessageCancel.Add_Click({ $messageForm.Close() })
                $WPFbtnMessageCancel.IsCancel = $true
                $messageForm.ShowDialog()
            }
        }

    }
    elseif( $Context -eq 'machine' )
    {
    }
}

Function Get-Sessions
{
    if( $WPFlabelUsers.Content -and $WPFlabelUsers.Content -match 'users' )
    {
    }
    elseif( $WPFlabelUsers.Content -and $WPFlabelUsers.Content -match 'machines' )
    {
    }
    ## get selected user from user list
    [int]$userid = -1
    if( $WPFUsers.SelectedItem )
    {
        ForEach( $user in $users.GetEnumerator() )
        {
            if( $user.Value -eq $WPFUsers.SelectedItem.Name )
            {
                $userid = $user.Name
                break
            }
        }
    }
    else ## no selected item so we can't retrieve sessions
    {
        return
    }
    
    if( $userid -lt 0 )
    {
        [Windows.MessageBox]::Show( "Unable to find user $($WPFUsers.SelectedItem.Name)" , "Information" , 'OK' ,'Error' )
        return
    }

    [datetime]$start = [datetime]::Parse( $WPFstartDatePicker.ToString() )
    [datetime]$end = Get-Date -Date ([datetime]::Parse( $WPFendDatePicker.ToString() )) -hour 23 -Minute 59 -Second 59
    [int]$dstOffset = Get-DayLightSavingsOffet
    if( $dstOffset )
    {
        $start = $start.AddMinutes( $dstOffset )
        $end = $end.AddMinutes( $dstOffset )
    }
    ## we don't care about the end time, only that the session started in the date range selected
    [array]$sessions = @( Invoke-ODataTransform( Invoke-RestMethod -Uri "$global:baseURI/Sessions()?`$filter=UserId eq $userid and StartDate ge datetime'$(Get-Date -date $start -format s)' and StartDate le datetime'$(Get-Date -date $end -format s)'" -UseDefaultCredentials ) )

    Write-Verbose "Got $($sessions.Count) for user id $userid"

    $WPFlabelSessions.Content = "$($sessions.Count) sessions"
    ## Create new bindings so we can ensure date shows correct format for region since cannot find a way to change UI culture away from en-US
    <#
    $newBinding  = [System.Windows.Data.BindingBase]( $WPFSessions.View.Columns[0].DisplayMemberBinding )
    $newBinding.StringFormat = 'dd/MM/yy hh:MM:ss'
    $WPFSessions.View.Columns[0].DisplayMemberBinding = $newBinding
#>
    if( ! $wpfsessions.ItemsSource )
    {
        $wpfsessions.Items.Clear()
    }
    $wpfsessions.View.Columns[0].Header = 'Start'
    $wpfsessions.View.Columns[1].Header = 'End'
    $wpfsessions.View.Columns[2].Header = 'Server'
    $wpfsessions.View.Columns[0].Width = 150
    $wpfsessions.View.Columns[1].Width = 150
    $wpfsessions.View.Columns[2].Width = 100
    ##ForEach( $session in $sessions )
    $sessions | ForEach-Object `
    {
        $session = $_
        $session | Add-Member -MemberType NoteProperty -Name Column1 -Value ( Get-Date $session.StartDate -Format G )
        try
        {
            $session | Add-Member -MemberType NoteProperty -Name Column2 -Value ( Get-Date $session.EndDate -Format G )
        }
        catch
        {
            ## Will get an exception if no end date because session still active
        }
        $session | Add-Member -MemberType NoteProperty -Name Machine -Value ( $machines[ $session.MachineId ].Name -split '\\' )[-1] 
        $session | Add-Member -MemberType NoteProperty -Name Column3 -Value $session.Machine
        $session | Add-Member -MemberType NoteProperty -Name FullMachine -Value $machines[ $session.MachineId ].Name
        ##$WPFSessions.AddChild( $session )
    }
    $WPFSessions.ItemsSource = $sessions
}

## Keyboard jumping doesn't work OOB so do our own
$WPFUsers.Add_PreviewKeyDown({
    $_.Handled = $true
    if( $this.Items.Count )
    {
        if( $_.Key -eq 'Return' )
        {
            Process-UserDoubleClicked
        }
        elseif( $_.Key.ToString().Length -eq 1 ) ## single character, not special key
        {
            [int]$index = $this.SelectedIndex + 1 ## start search after current selection
            if( ($this.Items.Item($this.SelectedIndex).Value.ToUpper())[0] -gt $_.Key.ToString().ToUpper() )
            {
                $index = 0
            }
            While( $index -lt $this.Items.Count )
            {
                if( $this.Items.Item($index).Value[0] -eq $_.Key.ToString() )
                {
                    $this.SelectedIndex = $index
                    $this.ScrollIntoView( $this.SelectedItem )
                    break
                }
                $index++
            }
        }
    }
})

<#
$WPFUsers.Add_KeyDown({
    $_
    $_.Handled = $true
})
#>

Function Get-Users( $GUIobject , [bool]$findUser )
{
    [string]$username = $null

    if( $findUser )
    {
        ## Prompt for user so we only get data in the time window for them
        $userForm = Load-GUI $generalTextEntry
        if( $userForm )
        {
            $wpfbtnTextEntryOk.Add_Click( {
                $userForm.DialogResult = $true 
                $userForm.Close() 
            } )
            $userForm.Title = "Enter user name"
            if( $userForm.ShowDialog() )
            {
                $username = $wpftextBoxEnterText.Text.Trim()
            }
            else
            {
                if( $GUIobject )
                {
                    $GUIobject.Handled = $true
                }
                return
            }
        }
    }

    if( $WPFUsers.ItemsSource )
    {
        $WPFUsers.ItemsSource = $null
    }
    else
    {
        $WPFUsers.Items.Clear()
    }
    ## We may also have items in the sessions list so clear those here too   
    if( $wpfsessions.ItemsSource )
    {
        $WPFSessions.ItemsSource = $null
    }
    else
    {
        $wpfsessions.Items.Clear()
    } 
    $WPFlabelSessions.Content = '0 sessions'

    [datetime]$start = [datetime]::Parse( $WPFstartDatePicker.ToString() )
    [datetime]$end = [datetime]::Parse( $WPFEndDatePicker.ToString() )
    $end = $end.AddHours(24)
    
    [int]$dstOffset = Get-DayLightSavingsOffet
    if( $dstOffset )
    {
        $start = $start.AddMinutes( $dstOffset )
        $end = $end.AddMinutes( $dstOffset )
    }

    Progress-Window "Fetching users from $global:ddc"

    # Always get users in case has changed
    $global:users = Get-UsersFromDDC -username $username
    if( ! $global:users -or ! $global:users.Count )
    {
        Progress-Window $null
        [Windows.MessageBox]::Show( "No users found via $global:ddc" , "Information" , 'OK' ,'Information' )
        return
    }

    $global:machines = Get-MachinesFromDDC
    if( ! $global:machines -or ! $global:machines.Count )
    {
        Progress-Window $null
        [Windows.MessageBox]::Show( "No machines found via $global:ddc" , "Information" , 'OK' ,'Information' )
        return
    }

    ## Get all sessions within that period
    [string]$query = "$global:baseURI/Sessions()?`$filter=StartDate ge datetime'$(Get-Date -date $start -format s)' and StartDate le datetime'$(Get-Date -date $end -format s)'"
    if( ! [string]::IsNullOrEmpty( $username  ) )
    {
        if( $global:users.Count -eq 1 )
        {
            $query += " and UserId eq $($global:users.GetEnumerator()|select -ExpandProperty Name|Select -First 1)" ## Name is the id
        }
        ## else we can't filter on more than one user in the query
    }
    [array]$sessions = @( Invoke-ODataTransform( Invoke-RestMethod -Uri $query -UseDefaultCredentials ) )

    if( ! $sessions -or ! $sessions.Count )
    {
        Progress-Window $null
        [Windows.MessageBox]::Show( "No sessions found for period via $global:ddc" , "Information" , 'OK' ,'Information' )
        return
    }

    Write-Verbose "Got $($sessions.Count) sessions from $start to $end"

    ## Now get list of unique users who have logged on on the given date so we can add to picker after translating user id to name from users list we fetched earlier
    [hashtable]$usersInWindow = @{}
    ##ForEach( $session in $sessions )
    $sessions | ForEach-Object `
    {
        $session = $_
        try
        {
            [string]$user = $users[ $session.UserId ]
            if( ! [string]::IsNullOrEmpty( $user.Trim() ) )
            {
                $usersInWindow.Add( $users[ $session.UserId ] , ( $users[ $session.UserId ] -split '\\')[-1] ) ## add name and fqdn so we can use this object in the listview
            }
        }
        catch
        {
            ## probably duplicate so ignore
            ##$error[0]
        }
    }
    $WPFlabelUsers.Content = "$($usersInWindow.Count) users"
    
    Progress-Window $null
    ##$WPFUsers.Items.Clear()
    $wpfusers.View.Columns[0].Header = 'User Name'
    ## columns are used by machine view not user view
    $wpfusers.View.Columns[1].Header = $wpfusers.View.Columns[2].Header = ''
    $wpfusers.View.Columns[1].Width =  $wpfusers.View.Columns[2].Width = 0

    $usersInWindow.GetEnumerator() | sort  Value | ForEach-Object { $WPFUsers.AddChild( $_ ) } ## "value", which is username without domain, is bound in listview to "name" column
    
    ## disable all context menus except the Info one since the other are for one we have machines in list, not users
    $wpfusers.ContextMenu.Items.GetEnumerator() | ForEach-Object `
    {
        if( $_.Header -and $_.Header -ne 'Info' )
        {
            $_.IsEnabled = $false
        }
    }
    if( $GUIobject )
    {
        $GUIobject.Handled = $true
    }
}

$WPFbtnFindUser.Add_Click({
    Get-Users -GUIobject $_ -findUser $true
})

$WPFbtnGetUsers.Add_Click({
    Get-Users -GUIobject $_ -findUser $false
})

Function Get-UsersFromDDC( [string]$username )
{
    ## Get all users so we can map user id in session data to a name
    [hashtable]$users = @{}
    [string]$query = "$global:baseURI/Users"
    if( ! [string]::IsNullOrEmpty( $username ) )
    {
        $query += "?`$filter=substringof(tolower('$username'),tolower(UserName)) eq true"
    }
    Invoke-ODataTransform(Invoke-RestMethod -Uri $query -UseDefaultCredentials) | ForEach-Object `
    {
        try
        {
            [string]$domainQualifiedName = $( 
                if( [string]::IsNullOrEmpty( $_.Domain ) )
                    {
                        $_.username
                    }
                    else
                    {
                        $_.Domain + '\' + $_.username 
                    }).Trim()

            if( $domainQualifiedName.Length  )
            {
                $users.Add( $_.Id , $domainQualifiedName )
            }
        }
        catch
        {
            $error[0]
        }
    }
    $users
}

Function Get-MachinesFromDDC
{
    ## Get all machines so we can map id to name
    $uri = "$global:baseURI/Machines"
    [hashtable]$machines = @{}
    Invoke-ODataTransform(Invoke-RestMethod -Uri $uri -UseDefaultCredentials) | ForEach-Object `
    {
        $machine = $_
        try
        {
            $machines.Add(  $machine.Id , $machine )
        }
        catch
        {
            $error[0]
        }
    }
    $machines
}

Function Get-ConfigFromRegistry
{
    $global:ddc = Get-ItemProperty -Path $configRegKey -Name 'DDC' -ErrorAction SilentlyContinue | select -ExpandProperty 'DDC' 
    $global:pvs = Get-ItemProperty -Path $configRegKey -Name 'PVS' -ErrorAction SilentlyContinue | select -ExpandProperty 'PVS'
    $global:hypervisor = ( Get-ItemProperty -Path $configRegKey -Name 'Hypervisor' -ErrorAction SilentlyContinue | select -ExpandProperty 'Hypervisor' ) -split ','
    $global:AMC = Get-ItemProperty -Path $configRegKey -Name 'AMC' -ErrorAction SilentlyContinue | select -ExpandProperty 'AMC'
    $global:messagteText = Get-ItemProperty -Path $configRegKey -Name 'MessageText' -ErrorAction SilentlyContinue | select -ExpandProperty 'MessageText'
    [int]$storedValue = Get-ItemProperty -Path $configRegKey -Name 'seconds' -ErrorAction SilentlyContinue| select -ExpandProperty 'seconds'
    if( $? )
    {
        $global:secondsWindow = $storedValue
    }
    ## else no value so leave as is
    ## TODO Where multiple items listed, figure out which is reachable and use that except hypervisors since VMware can work on an array
    return $global:ddc ## This is the only mandatory item
}

$WPFstartDatePicker.SelectedDate = $WPFendDatePicker.SelectedDate = Get-Date

## Read config from registry
if( ! ( Get-ConfigFromRegistry ) )
{
    Show-Config
}

[string]$thisddc = $global:ddc
[string]$global:baseURI = Invoke-Command -ComputerName $global:ddc -ScriptBlock { Add-PSSnapin 'Citrix.Configuration.Admin.*' ; Get-ConfigRegisteredServiceInstance -ServiceType monitor | ?{ $_.Address -match "/$($using:thisddc)[^a-z0-9_\-]" } | Sort Version -Descending | select -first 1 -ExpandProperty address }
if( [string]::IsNullOrEmpty( $baseURI ) )
{
    Write-Error "Failed to get base URI from $global:ddc"
    return
}

$global:baseURI += '/Data'

if( [string]::IsNullOrEmpty( $configFile ) )
{
    $configFile =  ( & { $myInvocation.ScriptName } ) -replace '\.ps1$' , '.csv' 
}
if( $alwaysOnTop )
{
    $mainForm.Topmost = $true
}

$WPFbtnTest.IsEnabled = $test 
$WPFbtnTest.Visibility = if( $test ) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Hidden }

$result = $mainForm.ShowDialog()
