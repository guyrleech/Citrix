#requires -version 3.0

<#
    Retrieve all user accounts which have Citrix Studio admin access

    Guy Leech
#>

<#
.SYNOPSIS

Show all individual Citrix XenApp 7.x admins as defined in Studio by using the Active Directory PowerShell module to recursively expand groups.

.DESCRIPTION

Will write to a csv file or display on screen in a grid view, including various AD attributes for each user/

.PARAMETER ddc

The Desktop Delivery Controller to connect to to retrieve the list of administrators from.

.PARAMETER csv

The name of a csv file to write the results to. Must not exist already. If not specified then results will be dispplayed in an on-screen grid view

.PARAMETER ADProperties

A commas separated list of the AD properties to include as returned by Get-ADUser

.PARAMETER name

The name of a single account (user or group) that you wish to only report on. This is as displayed in Studio.

.EXAMPLE

& '.\scripts\Show Studio access.ps1' -ddc ctxddc01 

Retrieve the list of Citrix admins from Delivery Controller cxtddc01 and display every user in an on-screen grid view

.EXAMPLE

& '.\scripts\Show Studio access.ps1' -ddc ctxddc01 -csv h:\citrix.admins.csv

Retrieve the list of Citrix admins from Delivery Controller cxtddc01 and save to file h:\citrix.admins.csv

.NOTES

The Citrix XenApp 7.x PowerShell cmdlets must be available so run where Studio is installed.
Also requires the ActiveDirectory PowerShell module.

#>

[CmdletBinding()]

Param
(
    [string]$ddc = 'localhost' ,
    [string]$csv ,
    [string[]]$ADProperties = @( 'Description','Office','Info','LockedOut','Created','Enabled','LastLogonDate' ) ,
    [string]$name 
)

Function Get-ADProperties( [string[]]$ADProperties , [string]$SamAccountName , $ADObject )
{
    [hashtable]$properties = @{}

    if( ! $ADObject )
    {
        $ADObject = Get-ADUser -Identity $SamAccountName -Properties $ADProperties
    }
    if( $ADObject )
    {
        ForEach( $ADProperty in $ADProperties )
        {
            $properties.Add( $ADProperty , ( $ADObject | select -ExpandProperty $ADProperty ) )
        }
    }
    else
    {
        Write-Warning "Failed to get $SamAccountName from AD"
    }
    $properties
}

[string[]]$snapins = @( 'Citrix.DelegatedAdmin.Admin.*'  ) 
[string[]]$modules = @( 'ActiveDirectory' ) 

if( $snapins -and $snapins.Count -gt 0 )
{
    ForEach( $snapin in $snapins )
    {
        Add-PSSnapin $snapin -ErrorAction Stop
    }
}

if( $modules -and $modules.Count -gt 0 )
{
    ForEach( $module in $modules )
    {
        Import-Module $module -ErrorAction Stop
    }
}

[hashtable]$params = @{}
if( ! [string]::IsNullOrEmpty( $name ) )
{
    $params.Add( 'Name' , $name )
}

$admins = @( Get-AdminAdministrator -AdminAddress $ddc @params )

Write-Verbose "Got $($admins.Count) admin entries from $ddc"

[int]$counter = 0

$results = @( ForEach( $admin in $admins )
{
    $counter++
    ## Now get all user accounts of this entity
    [string]$account = ($admin.Name -split '\\')[-1]
    Write-Verbose "$counter / $($admins.Count) : $account"
    $user = $null
    $group = $null
    [string]$role,[string]$scope = $admin.Rights -split ':'
    [hashtable]$commonProperties = @{ 'Role' = $role ; 'Scope' = $scope }

    try
    {
        $user = Get-ADUser -Identity $account -Properties $ADProperties
    }
    catch
    {
        $group = Get-ADGroupMember -Identity $account -Recursive
    }
       
    if( $group )
    {
        $group | ForEach-Object `
        {
            $thisGroup = $_
            $result = [pscustomobject][ordered]@{ 'Name'=$thisGroup.SamAccountName ; 'Via Group'= $account }
            Add-Member -InputObject $result -NotePropertyMembers $commonProperties
            $extras = Get-ADProperties -ADProperties $ADProperties -SamAccountName $thisGroup.SamAccountName -ADObject $null
            if( $extras -and $extras.Count )
            {
                Add-Member -InputObject $result -NotePropertyMembers $extras
            }
            $result
        }
    }
    elseif( $user )
    {
        $result = [pscustomobject][ordered]@{ 'Name'=$user.SamAccountName ; 'Via Group'= $null  }
        Add-Member -InputObject $result -NotePropertyMembers $commonProperties
        $extras = Get-ADProperties -ADProperties $ADProperties -SamAccountName $user.SamAccountName -ADObject $user
        if( $extras -and $extras.Count )
        {
            Add-Member -InputObject $result -NotePropertyMembers $extras
        }
        $result
    }
    else
    {
        Write-Warning "Unable to find AD entity `"$account`""
    }
})

[string]$message = "Got $($results.count) individual admins via $($admins.Count) entries via $ddc"

Write-Verbose $message

if( $results -and $results.Count )
{
    if( [string]::IsNullOrEmpty( $csv ) )
    {
        $selected = $results | Out-GridView -Title $message -PassThru
        if( $selected -and $selected.Count )
        {
            $selected | clip.exe
        }
    }
    else
    {
        $results | Export-Csv -Path $csv -NoTypeInformation -NoClobber
    }
}
