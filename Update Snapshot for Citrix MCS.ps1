<#
.SYNOPSIS

    Change the snapshot being used for a specified Citrix MCS machine catalog and optionally reboot the machines in the updated machine catalog

.PARAMETER catalog

    The name, or pattern, of the single machine catalog to update

.PARAMETER snapshotName

    The name of the snapshot to use (regular expressions supported)

.PARAMETER vmName

    The name of the virtual machine to use the snapshot from. If not specified will use the existing VM set for the catalog.

.PARAMETER ddc

    The delivery controller to use

.PARAMETER reportOnly

    Do not change anything, just report on the current status and snapshot

.PARAMETER rebootdurationMinutes

    The period, in minutes, over which to reboot all machines in the catalog once updated

.PARAMETER warningDurationMinutes

    Time in minutes prior to a machine reboot at which a warning message is displayed in all user sessions on that machine

.PARAMETER warningTitle

    The title of the warning message

.PARAMETER warningMessage

    The contents of the warning message. Environment variables will be expanded

.PARAMETER warningRepeatIntervalMinutes

    Repeat the warning message at this interval in minutes

.PARAMETER async

    Do not wait for the catalog update to complete. Cannot be used with reboot parameters.

.EXAMPLE

    & '.\Update Snapshot for Citrix MCS.ps1' -catalog "MCS Server 2019" -ddc grl-xaddc02 -reportonly

    Show the current snapshot, application date and other information about the machine catalog via the delivery controller grl-xaddc02

.EXAMPLE

    & '.\Update Snapshot for Citrix MCS.ps1' -catalog "MCS Server 2019" -ddc grl-xaddc02 -snapshotName "FSlogix 2.9.7838.44263, WU" -rebootdurationMinutes 60

    Update the named machine catalog with the named snapshot and when complete initiate a reboot cycle of the machines in the catlog which should last no longer than 60 minutes.

.EXAMPLE

    & '.\Update Snapshot for Citrix MCS.ps1' -catalog "MCS Server 2019" -ddc grl-xaddc02 -async

    Update the named machine catalog with the latest snapshot but do not initiate reboots of the machines in the catalog.
    Note that rebooting within the OS or hypervisor will not cause the new snapshot to be used - the reboot must be initiated via Studio or Start-BrokerRebootCycle

.NOTES

    Requires CVAD PowerShell cmdlets (installed with Studio or available as separate msi files on the product ISO)

    Use -confirm:$false to suppress prompting to take actions (use at your own risk)

    https://support.citrix.com/article/CTX129205

    Modification History

    2021/12/05 @guyrleech  Fixed multiple -verbose to Publish-ProvMasterVMImage
    2022/01/08 @guyrleech  Added extra fields to reportonly
#>

<#
Copyright © 2021 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='Name of machine catalog to list/update')]
    [string]$catalog ,
    [Parameter(Mandatory=$false,ParameterSetName='Async')]
    [Parameter(Mandatory=$false,ParameterSetName='Reboot')]
    [string]$snapshotName ,
    [string]$vmName ,
    [string]$hostingUnitName ,
    [string]$ddc ,
    [string]$profileName ,
    [Parameter(Mandatory=$true,ParameterSetName='ReportOnly')]
    [switch]$reportOnly ,
    [Parameter(Mandatory=$true,ParameterSetName='Reboot')]
    [int]$rebootdurationMinutes ,
    [Parameter(Mandatory=$false,ParameterSetName='Reboot')]
    [int]$warningDurationMinutes ,
    [Parameter(Mandatory=$false,ParameterSetName='Reboot')]
    [string]$warningTitle ,
    [Parameter(Mandatory=$false,ParameterSetName='Reboot')]
    [string]$warningMessage ,
    [Parameter(Mandatory=$false,ParameterSetName='Reboot')]
    [int]$warningRepeatIntervalMinutes ,
    [Parameter(Mandatory=$true,ParameterSetName='Async')]
    [switch]$async
)

Add-PSSnapin -Name Citrix.MachineCreation.* , Citrix.Host.*

[hashtable]$citrixParams = @{ 'Verbose' = $false }

if( $PSBoundParameters[ 'profileName' ] )
{
    Get-XDAuthentication -ProfileName $profileName -ErrorAction Stop
    $ddc = 'Cloud'
}
elseif( $PSBoundParameters[ 'ddc' ] )
{
    $citrixParams.Add( 'AdminAddress' , $ddc )
}

[array]$machineCatalogs = @( Get-BrokerCatalog -Name $catalog -ProvisioningType MCS @citrixParams )

if( ! $machineCatalogs -or ! $machineCatalogs.Count )
{
    Throw "Found no MCS machine catalogues for `"$catalog`""
}

if( $machineCatalogs.Count -ne 1 )
{
    Throw "Found $($machineCatalogs.Count) MCS machine catalogues for `"$catalog`" - `"$(($machineCatalogs|Select-Object -ExpandProperty Name) -join '" , "')`""
}

$machineCatalog = $machineCatalogs[0]

if( $provScheme = Get-ProvScheme @citrixParams -ProvisioningSchemeUid $machineCatalog.ProvisioningSchemeId )
{
    ## get requested or deepest snapshot - Filter is not supported and -Include doesn't match everything we need

    ## The Get-ChildItem is very verbose so we'll lose that
    [array]$snapshots = @( Get-ChildItem -Path "$(($provScheme.MasterImageVM -split '\.vm\\')[0]).vm" -Recurse -Verbose:$false | Where-Object { $_.Name -match $snapshotName } -Verbose:$false ) ## if snapshot is null , will match everything

    if( ! $PSBoundParameters[ 'snapshotName' ] -and $snapshots.Count )
    {
        ## if not matching on name then get deepest snapshot only
        $snapshots = @( $snapshots[ -1 ] )
    }

    [string]$currentSnapshot = ($provScheme.MasterImageVM -split '\\')[-1] -replace '\.snapshot$'

    if( $PSBoundParameters[ 'vmName' ] )
    {
        if( ! $PSBoundParameters[ 'hostingUnitName' ] )
        {
            $hostingUnitName = $provScheme.HostingUnitName
        }

        ## Find this VM and get its snapshots instead
        if( ! ( $baseVM = Get-ChildItem -Path "XDHyp:\HostingUnits\$hostingUnitName\" | Where-Object { $_.Objecttype -eq 'vm' -and $_.PSChildName -eq "$vmName.vm" } ) )
        {
            Throw "Unable to find vm $vmName in hosting unit $hostingUnitName"
        }
        if( $baseVM -is [array] -and $baseVM.Count -gt 1 )
        {
            Throw "Found $($baseVM.Count) vms for $vmName in hosting unit $hostingUnitName"
        }
        $snapshots = @( Get-ChildItem -Path $baseVM.PSPath -Recurse -Verbose:$false | Where-Object { $_.Name -match $snapshotName } ) ## if snapshotis null , will match everything
        if( ! $snapshots -or ! $snapshots.Count )
        {
            [string]$exception = "No snapshots found for $vmName"
            if( ! [string]::IsNullOrEmpty( $snapshotName ) )
            {
                $exception += " for snapshot `"$snapshotName`""
            }
            Throw $exception
        }
        if( ! $PSBoundParameters[ 'snapshotName' ] -and $snapshots.Count )
        {
            ## if not matching on name then get deepest snapshot only
            $snapshots = @( $snapshots[ -1 ] )
        }
    }

    Write-Verbose -Message "Current snapshot is '$currentSnapshot'"

    if( $reportOnly )
    {
        New-Object -TypeName pscustomobject -ArgumentList @{
            ## XDHyp:\HostingUnits\Internal Network\GLCTXMCSMAST19.vm\CVAD VDA 2012, VMtools 11.2.1.snapshot\Updated UWM 2020.3 & FSlogix, CVAD 2012.snapshot
            ##                                      ^^^^^^^^^^^^^^
            'VM' = $provScheme.MasterImageVM -replace '^XDhyp:\\[^\\]+\\[^\\]+\\([^\\]+)\.vm\\.*$' , '$1'
            'Current Snapshot' = $currentSnapshot
            'Applied Date' = $provScheme.MasterImageVMDate
            'vCPU' = $provScheme.CpuCount
            'MemoryGB' = [math]::Round( $provScheme.MemoryMB / 1024 , 2 )
            'DiskGB' = $provScheme.DiskSize
            'HostingUnitName' = $provScheme.HostingUnitName
            'WriteBackCacheDiskSize' = $provScheme.WriteBackCacheDiskSize
            'WriteBackCacheMemorySize' = $provScheme.WriteBackCacheMemorySize
        }
    }
    elseif( ! $snapshots -or ! $snapshots.Count )
    {
        [array]$allSnapshots = ForEach( $snapshot in ((Get-ChildItem -Path "$(($provScheme.MasterImageVM -split '\.vm\\')[0]).vm" -Recurse -Verbose:$false)|Select-Object -ExpandProperty FullPath  )) { ($snapshot -split '\\')[-1] -replace '\.snapshot$' }
        [string]$extraMessage = if( $PSBoundParameters[ 'snapshotName' ] ) { " out of the $($allSnapshots.Count) found matching `"$snapshotName`", snapshots are '$( $allSnapshots -join "' , '" )'" }
        Throw "Found no snapshots$extraMessage"
    }
    elseif( $snapshots.Count -gt 1 )
    {
        [string]$snapshotDetails =  "Found $($snapshots.Count) matching snapshots for `"$snapshotName`" - '$(((Split-Path -Path $snapshots.FullPath -Leaf -Verbose:$false) -replace '\.snapshot$') -join "' , '" )'"
        Throw $snapshotDetails
    }
    else
    {
        $snapshot = $snapshots[ 0 ]
        [string]$newSnapshotName = $snapshot.FullPath -replace '\.snapshot(\\|$)' , '\\' -replace '\\+$'

        if( $snapshot.FullPath -eq $provScheme.MasterImageVM )
        {
            Throw "Already using snapshot `"$($snapshot.FullPath)`""
        }
        elseif( $PSCmdlet.ShouldProcess( "Catalog '$catalog'" , "Publish snapshot '$newSnapshotName'" ) )
        {
            Write-Verbose -Message "$(Get-Date -Format G): starting provisioning snapshot '$(Split-Path -Path $newSnapshotName -Leaf -Verbose:$false)' ..."
            if( ! ( $publishedResult = Publish-ProvMasterVMImage -ProvisioningSchemeName $provScheme.IdentityPoolName -MasterImageVM $snapshot.FullPath -RunAsynchronously:$async @citrixParams ) )
            {
                Throw "Null returned from Publish-ProvMasterVMImage - provisoning most likely failed"
            }
            elseif( ! [string]::IsNullOrEmpty( $publishedResult.TerminatingError )  )
            {
                Throw "Provisioning task failed with `"$($publishedResult.TerminatingError)`" at $(Get-Date -Date $publishedResult.DateFinished -Format G) after $($publishedResult.ActiveElapsedTime) seconds"
            }
            elseif( ! $async )
            {
                if( $PSBoundParameters[ 'rebootdurationMinutes' ] )
                {
                    if( $PSCmdlet.ShouldProcess( "Catalogue `"$($machineCatalog.Name)`" with $($machineCatalog.UsedCount) machines" , 'Start Reboots' ) )
                    {
                        [hashtable]$rebootArguments = @{ InputObject = $machineCatalog ; RebootDuration = $rebootdurationMinutes }
                        $rebootArguments += $citrixParams
                        if( $PSBoundParameters[ 'warningDurationMinutes' ] )
                        {
                            $rebootArguments.Add( 'WarningDuration' , $warningDurationMinutes )
                        }
                        if( $PSBoundParameters[ 'warningRepeatIntervalMinutes' ] )
                        {
                            $rebootArguments.Add( 'WarningRepeatInterval' , $warningRepeatIntervalMinutes )
                        }
                        if( ! [string]::IsNullOrEmpty( $warningTitle ) )
                        {
                            $rebootArguments.Add( 'WarningTitle' , $warningTitle )
                        }
                        if( ! [string]::IsNullOrEmpty( $warningMessage ) )
                        {
                            $rebootArguments.Add( 'WarningMessage' , [System.Environment]::ExpandEnvironmentVariables( $warningMessage ) )
                        }
                        if( ! ( $rebootCycle = Start-BrokerRebootCycle @rebootArguments ) )
                        {
                            Write-Warning -Message "Failed to initiate reboot cycle for catalogue `"$($machineCatalog.Name)`""
                        }
                        else
                        {
                            $rebootCycle
                        }
                    }
                    else
                    {
                        Write-Warning -Message "Updated to snapshot ok but unable to find catalogue with provisioning scheme uid $($publishedResult.ProvisioningSchemeUid)"
                    }
                }
                Write-Verbose -Message "Finished ok at $(Get-Date -Date $publishedResult.DateFinished -Format G), taking $($publishedResult.ActiveElapsedTime) seconds"
            }
            elseif( ! ($taskStatus = Get-ProvTask @citrixParams -TaskId $publishedResult ) )
            {
                Throw "Failed to get task status"
            }
            elseif( ! $taskStatus.Active -or $taskStatus.Status -ne 'Running' -or $taskStatus.TerminatingError )
            {
                Throw "Async task not running - status is $($taskStatus.Status) error is '$($taskStatus.TerminatingError)'"
            }
            else
            {
                $taskStatus
            }
        }
    }
}
else
{
    Throw "Failed to find provisioning scheme for delivery group `"$catalog`""
}

# SIG # Begin signature block
# MIIZsAYJKoZIhvcNAQcCoIIZoTCCGZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU1j4Nj/wEFWI5JCy6qIO5frbn
# Jv+gghS+MIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTAwggQY
# oAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsx
# SRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawO
# eSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJ
# RdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEc
# z+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whk
# PlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8l
# k9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQD
# AgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARI
# MEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdp
# Y2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg
# +S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG
# 9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/E
# r4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3
# nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpo
# aK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW
# 6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ
# 92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTEwggQZoAMCAQICEAqhJdbW
# Mht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIG
# A1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTE2MDEwNzEyMDAw
# MFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnFOVQoV7YjSsQOB0Uz
# URB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQAOPcuHjvuzKb2Mln+
# X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhisEeTwmQNtO4V8CdPu
# XciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQjMF287DxgaqwvB8z9
# 8OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+fMRTWrdXyZMt7HgXQ
# hBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW/5MCAwEAAaOCAc4w
# ggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEFBQcBAQRtMGswJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBQBgNV
# HSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cu
# ZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggEB
# AHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafDDiBCLK938ysfDCFa
# KrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6HHssIeLWWywUNUME
# aLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4H9YLFKWA1xJHcLN1
# 1ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHKeZR+WfyMD+NvtQEm
# tmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIoxhhWz0E0tmZdtnR7
# 9VYzIi8iNrJLokqV2PWmjlIwggVPMIIEN6ADAgECAhAE/eOq2921q55B9NnVIXVO
# MA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwNzIwMDAw
# MDAwWhcNMjMwNzI1MTIwMDAwWjCBizELMAkGA1UEBhMCR0IxEjAQBgNVBAcTCVdh
# a2VmaWVsZDEmMCQGA1UEChMdU2VjdXJlIFBsYXRmb3JtIFNvbHV0aW9ucyBMdGQx
# GDAWBgNVBAsTD1NjcmlwdGluZ0hlYXZlbjEmMCQGA1UEAxMdU2VjdXJlIFBsYXRm
# b3JtIFNvbHV0aW9ucyBMdGQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCvbSdd1oAAu9rTtdnKSlGWKPF8g+RNRAUDFCBdNbYbklzVhB8hiMh48LqhoP7d
# lzZY3YmuxztuPlB7k2PhAccd/eOikvKDyNeXsSa3WaXLNSu3KChDVekEFee/vR29
# mJuujp1eYrz8zfvDmkQCP/r34Bgzsg4XPYKtMitCO/CMQtI6Rnaj7P6Kp9rH1nVO
# /zb7KD2IMedTFlaFqIReT0EVG/1ZizOpNdBMSG/x+ZQjZplfjyyjiYmE0a7tWnVM
# Z4KKTUb3n1CTuwWHfK9G6CNjQghcFe4D4tFPTTKOSAx7xegN1oGgifnLdmtDtsJU
# OOhOtyf9Kp8e+EQQyPVrV/TNAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7
# KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUTXqi+WoiTm5fYlDLqiDQ4I+uyckw
# DgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4w
# NaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3Mt
# ZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1
# cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUF
# BwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYI
# KwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAA
# MA0GCSqGSIb3DQEBCwUAA4IBAQBT3M71SlOQ8vwM2txshp/XDvfoKBYHkpFCyanW
# aFdsYQJQIKk4LOVgUJJ6LAf0xPSN7dZpjFaoilQy8Ajyd0U9UOnlEX4gk2J+z5i4
# sFxK/W2KU1j6R9rY5LbScWtsV+X1BtHihpzPywGGE5eth5Q5TixMdI9CN3eWnKGF
# kY13cI69zZyyTnkkb+HaFHZ8r6binvOyzMr69+oRf0Bv/uBgyBKjrmGEUxJZy+00
# 7fbmYDEclgnWT1cRROarzbxmZ8R7Iyor0WU3nKRgkxan+8rzDhzpZdtgIFdYvjeO
# c/IpPi2mI6NY4jqDXwkx1TEIbjUdrCmEfjhAfMTU094L7VSNMYIEXDCCBFgCAQEw
# gYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1
# cmVkIElEIENvZGUgU2lnbmluZyBDQQIQBP3jqtvdtaueQfTZ1SF1TjAJBgUrDgMC
# GgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG
# 9w0BCQQxFgQU5ZXw7eYZG/RYZvBUoLM1CrAMHKMwDQYJKoZIhvcNAQEBBQAEggEA
# o7tCC8LSJXak2eVke0Q9KbF3nVaqazeVlZzyKxGf06Gq3yGAw6ABMV3+Y7MhAy4V
# JHyf6P4MaRYNjNcStYrnlayBGNeXk8lwIY3R/+KS9jvbVW4pxUn2FZUw0l3PsMxQ
# qicTkfPL3Bch2m+spJxmNMO1D6Rc56Ocs6UhFolFaGZmrSCxWCBIt+ZcyPo29X+4
# DKpS6iUCMFi5qpeoLo0fRoP1ZBrv/k2YvXpJ0lejF2w8qxRpwe5GOVzlzl7+3J01
# 1nZMhJCe41G0Hy8XtnMSu0SIy2OKFYPThz3iJK9wkL/J9ijHj9JETN43yy3eqfBY
# iiR9HwyMHqNNrD44UUJmDqGCAjAwggIsBgkqhkiG9w0BCQYxggIdMIICGQIBATCB
# hjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgVGltZXN0YW1waW5nIENBAhANQkrgvjqI/2BAIc4UAPDdMA0GCWCGSAFl
# AwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjIwMTA4MTYxMzM4WjAvBgkqhkiG9w0BCQQxIgQgJNZT9BwdwzlGc5hzeQj8
# 8F7qg4/HSBDQrtrgCxFKdIgwDQYJKoZIhvcNAQEBBQAEggEAunloKUHWmaGM1TVI
# SBQg5Mj0BeiPojGCjgm93CRlIRo/Ak4XAzYEvqK1U2xVVc6pW2xkGWrzHM3xptuc
# CC0rStfBJhXohfLYwVTzhWR17FQVcvjBqeWtwaFAZL1cImniF///RSIXmwNgk7wV
# kegT8HwBEVfwFaEWX+uFGOTN+7mEWROgNhISXLxTv2pA7qP3Tcbjzu6JmIYdSCI1
# bcqYeOcioqAq7jMk8Sn883BhpMxuM8dVPRvluS68X0h7M3Xqb6SgHyCSPOZZGOGq
# JapbWrUd1/o/2GRdqX8Xy6DMaAze6o7cWUXaN9ZgNI12vfS+q52OTH4TP2d2nToI
# JA6b8A==
# SIG # End signature block
