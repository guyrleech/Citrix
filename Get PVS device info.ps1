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
    
    17/04/18    GL  Implemented timeout when getting remote information as Invoke-Command can take up to 25 minutes to timeout.
                    Added remote domain health check via Test-ComputerSecureChannel and AD account modification time

    30/05/18    GL  Added Load indexes

    01/06/18    GL  Changed to run device queries in runspaces to speed up

    19/06/18    GL  Split main code into module so can re-use in other scripts

    20/06/18    GL  Added option to split VMware VM names in case have description, etc after an _ character or similar

    21/01/21    GL  Added remove from delivery group option
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

.PARAMETER jobTimeout

Timeout in seconds for the command to retrieve information remotely from a device.

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

.PARAMETER splitVM

Split VMware name on a specific character such as an underscore where the second and subsequent parts of the name are descriptions, etc

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
    [ValidateSet('PVS','MCS','Manual','Any')]
    [string]$provisioningType = 'PVS' ,
    [string]$configRegKey = 'HKCU:\software\Guy Leech\PVS Fetcher' ,
    [string]$messageText ,
    [string]$messageCaption ,
    [int]$maxRecordCount = 2000 ,
    [int]$jobTimeout = 120 ,
    [int]$timeout = 60 ,
    [int]$cpuSamples = 2 ,
    [int]$maxThreads = 10 ,
    [string]$pvsShare ,
    [string]$splitVM ,
    [switch]$help ,
    [string[]]$snapins = @( 'Citrix.Broker.Admin.*'  ) ,
    [string[]]$modules = @( 'ActiveDirectory', "$env:ProgramFiles\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll" , 'Guys.Common.Functions.psm1' ) ,
    [string]$vmWareModule = 'VMware.VimAutomation.Core'
)

if( $help )
{
    Get-Help -Name ( & { $myInvocation.ScriptName } ) -Full
    return
}

$columns = [System.Collections.ArrayList]( @( 'Name','DomainName','Description','PVS Server','DDC','SiteName','CollectionName','Machine Catalogue','Delivery Group','Load Index','Load Indexes','Registration State','Maintenance_Mode','User_Sessions','devicemac','active','enabled',
    'Store Name','Disk Version Access','Disk Version Created','AD Account Created','AD Account Modified','Domain Membership','AD Last Logon','AD Description','Disk Name','Booted off vdisk','Booted Disk Version','Vdisk Production Version','Vdisk Latest Version','Latest Version Description','Override Version',
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
                        <MenuItem Header="Reconfigure" >
                            <MenuItem Header="CPUs" Name="VMwareReconfigCPUsContextMenu" />
                            <MenuItem Header="Memory" Name="VMwareReconfigMemoryContextMenu" />
                        </MenuItem>
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
                        <MenuItem Header="Remove from Delivery Group" Name="DDCRemoveDeliveryGroupContextMenu" />
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
            'Remove From Delivery Group' { Remove-BrokerMachine -Force -DesktopGroup $device.'Delivery Group' -AdminAddress $device.ddc -MachineName $( if( [string]::IsNullOrEmpty( $device.DomainName ) ) { $device.Name } else {  $device.DomainName + '\' +  $device.Name } ) ; break  }
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
  
if( $hypervisors -and $hypervisors.Count )
{
    if( [string]::IsNullOrEmpty( $name ) )
    {
        Write-Error "Must specify a VM name pattern via -name when cross referencing to VMware"
        exit
    }

    Import-Module $vmWareModule -ErrorAction Stop

    Write-Progress -Activity "Connecting to hypervisors $($hypervisors -split ' ')" -PercentComplete 1

    $null = Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
    if( Connect-VIserver -Server $hypervisors )
    {
        $null = $columns.Add( 'CPUs')
        $null = $columns.Add( 'Memory (GB)')
        $null = $columns.Add( 'Hard Drives (GB)')
        $null = $columns.Add( 'NICs')
        $null = $columns.Add( 'Hypervisor')
        ## we pass the modules list to our function so that they can be loaded into the runspaces
        $modules += $vmWareModule
    }
    else
    {
        Write-Error "Failed to connect to vmware $($hypervisors -join ' ')"
        exit
    }
}

if( $PSVersionTable.PSVersion.Major -lt 5 -and $columns.Count -gt 30 -and [string]::IsNullOrEmpty( $csv ) )
{
    Write-Warning "This version of PowerShell limits the number of columns in a grid view to 30 and we have $($columns.Count) so those from `"$($columns[30])`" will be lost in grid view"
}

[datetime]$startTime = Get-Date

[hashtable]$devices = Get-PVSDevices -pvsservers $pvsServers -ddcs $ddcs -hypervisors $hypervisors -dns:$dns -name $name -tags:$tags -adgroups $ADgroups -noRemoting:$noRemoting -cpusamples:$cpuSamples -noProgress:$noprogress `
    -maxThreads $maxThreads -timeout $timeout -jobTimeout $jobTimeout -maxRecordCount $maxRecordCount -provisioningType $provisioningType -noOrphans:$noOrphans -pvsShare $pvsShare -modules $modules -splitVM $splitVM

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
        [string]$title = "$(Get-Date -Format G) : $($devices.count) PVS devices via $($pvsServers -join ' ') & ddc $($ddcs -join ' ')"
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
            
            $WPFDDCRemoveDeliveryGroupContextMenu.Add_Click({ Perform-Action -action 'Remove From Delivery Group' -form $mainForm })
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

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUUJIFakZfLcNPKbCkn1xDh6UU
# La+gggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
# AQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# +NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ
# 1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0
# sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6s
# cKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4Tz
# rGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg
# 0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIB
# ADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUH
# AQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYI
# KwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaG
# NGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYE
# FFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06
# GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5j
# DhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgC
# PC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIy
# sjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4Gb
# T8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIFTzCC
# BDegAwIBAgIQBP3jqtvdtaueQfTZ1SF1TjANBgkqhkiG9w0BAQsFADByMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBMB4XDTIwMDcyMDAwMDAwMFoXDTIzMDcyNTEyMDAwMFowgYsx
# CzAJBgNVBAYTAkdCMRIwEAYDVQQHEwlXYWtlZmllbGQxJjAkBgNVBAoTHVNlY3Vy
# ZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMRgwFgYDVQQLEw9TY3JpcHRpbmdIZWF2
# ZW4xJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAr20nXdaAALva07XZykpRlijxfIPk
# TUQFAxQgXTW2G5Jc1YQfIYjIePC6oaD+3Zc2WN2Jrsc7bj5Qe5Nj4QHHHf3jopLy
# g8jXl7Emt1mlyzUrtygoQ1XpBBXnv70dvZibro6dXmK8/M37w5pEAj/69+AYM7IO
# Fz2CrTIrQjvwjELSOkZ2o+z+iqfax9Z1Tv82+yg9iDHnUxZWhaiEXk9BFRv9WYsz
# qTXQTEhv8fmUI2aZX48so4mJhNGu7Vp1TGeCik1G959Qk7sFh3yvRugjY0IIXBXu
# A+LRT00yjkgMe8XoDdaBoIn5y3ZrQ7bCVDjoTrcn/SqfHvhEEMj1a1f0zQIDAQAB
# o4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0O
# BBYEFE16ovlqIk5uX2JQy6og0OCPrsnJMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUE
# DDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUw
# QzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNp
# Z25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAU9zO
# 9UpTkPL8DNrcbIaf1w736CgWB5KRQsmp1mhXbGECUCCpOCzlYFCSeiwH9MT0je3W
# aYxWqIpUMvAI8ndFPVDp5RF+IJNifs+YuLBcSv1tilNY+kfa2OS20nFrbFfl9QbR
# 4oacz8sBhhOXrYeUOU4sTHSPQjd3lpyhhZGNd3COvc2csk55JG/h2hR2fK+m4p7z
# sszK+vfqEX9Ab/7gYMgSo65hhFMSWcvtNO325mAxHJYJ1k9XEUTmq828ZmfEeyMq
# K9FlN5ykYJMWp/vK8w4c6WXbYCBXWL43jnPyKT4tpiOjWOI6g18JMdUxCG41Hawp
# hH44QHzE1NPeC+1UjTGCAigwggIkAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EAT946rb3bWrnkH02dUhdU4wCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFMvCyoaarpXf1m3I9rXz
# +mGEcDt9MA0GCSqGSIb3DQEBAQUABIIBABAZbhYs8iZ9RLxXClyggrVfqjWAPRza
# ZNWivDAQSnUZzTa8CPVCfDJe/uuIHiXzwS0SNvm1dpC7frDkz+bOKilWkDVbWWiC
# fmwtbV9rV38srPs2smJMznwmrHr4H42VJqCQaBGDnTvaMgSkMosJsmc2ZCEWjtPc
# zKsvtP8wvBmp+bOXJybNWx51DjDVRD+0dHR+dIpZEk5L39LWh5eNmpvl1ONeu/Lt
# ffAvesA06kS0PQZ7YsO45B0b/y2qEkNh6fTT3tGzvTVv4xgJSW2NwIrmYPw1TPfy
# AEtcPa2D+ejBnC+Fsdn2O380+EBHE/Xkk7VnYXGsdXfxMrhVLVlDGNA=
# SIG # End signature block
