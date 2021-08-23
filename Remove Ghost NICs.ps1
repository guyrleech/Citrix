
<#
.SYNOPSIS

    Remove any ghost NICs as in those which show in device manager when view->show hidden devices is enabled

.DESCRIPTION

    Ghost NICs can cause Citrix Provisioning Services (PVS) target devices to fail to boot

.PARAMETER nicRegex

    A regular expression to match the name of ghost NICs to remove. If not specified, all ghost NICs will be removed. It is recommended that this parameter is used to specify the expected ghost NICs

.EXAMPLE

    . '.\Remove Ghost NICs.ps1' -nicRegex Intel.*Gigabit

    Remove all ghost NICS containing the given regular expression in their name after prompting for confirmation to perform the action

.EXAMPLE

    . '.\Remove Ghost NICs.ps1' -nicRegex Intel.*Gigabit -Confirm:$false

    Remove all ghost NICS containing the given regular expression in their name without prompting for confirmation to perform the action

.NOTES

    Removal code adapted from http://www.dwarfsoft.com/blog/2012/12/09/network-interface-removal-and-renaming/

    NICs may be ghosts because the driver is installed but currently there is no device for it which may change if a NIC of that type is added so do not remove ghost NICs unless you are certain
    that a NIC of that type will never be used on this machine.

    Another cause of PVS target device boot failure can be that the PCI slot number for the NIC has changed - https://github.com/guyrleech/VMware/blob/master/Change%20NIC%20PCI%20slot%20number.ps1

    Modification History:

        2021/08/23  @guyrleech  Initial release
#>

<#
Copyright © 2021 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [String]$nicRegex
)

function RemoveDevice([string]$DeviceID)
{
$RemoveDeviceSource = @'

    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    namespace Microsoft.Windows.Diagnosis
    {
        public sealed class DeviceManagement_Remove
        {
            public const UInt32 ERROR_CLASS_MISMATCH = 0xE0000203;
 
            [DllImport("setupapi.dll", SetLastError = true, EntryPoint = "SetupDiOpenDeviceInfo", CharSet = CharSet.Auto)]
            static extern UInt32 SetupDiOpenDeviceInfo(IntPtr DeviceInfoSet, [MarshalAs(UnmanagedType.LPWStr)]string DeviceID, IntPtr Parent, UInt32 Flags, ref SP_DEVINFO_DATA DeviceInfoData);
 
            [DllImport("setupapi.dll", SetLastError = true, EntryPoint = "SetupDiCreateDeviceInfoList", CharSet = CharSet.Unicode)]
            static extern IntPtr SetupDiCreateDeviceInfoList(IntPtr ClassGuid, IntPtr Parent);
 
            [DllImport("setupapi.dll", SetLastError = true, EntryPoint = "SetupDiDestroyDeviceInfoList", CharSet = CharSet.Unicode)]
            static extern UInt32 SetupDiDestroyDeviceInfoList(IntPtr DevInfo);
 
            [DllImport("setupapi.dll", SetLastError = true, EntryPoint = "SetupDiRemoveDevice", CharSet = CharSet.Auto)]
            public static extern int SetupDiRemoveDevice(IntPtr DeviceInfoSet, ref SP_DEVINFO_DATA DeviceInfoData);
 
            [StructLayout(LayoutKind.Sequential)]
            public struct SP_DEVINFO_DATA
            {
                public UInt32 Size;
                public Guid ClassGuid;
                public UInt32 DevInst;
                public IntPtr Reserved;
            }
 
            private DeviceManagement_Remove()
            {
            }
 
            public static UInt32 GetDeviceInformation(string DeviceID, ref IntPtr DevInfoSet, ref SP_DEVINFO_DATA DevInfo)
            {
                DevInfoSet = SetupDiCreateDeviceInfoList(IntPtr.Zero, IntPtr.Zero);
                if (DevInfoSet == IntPtr.Zero)
                {
                    return (UInt32)Marshal.GetLastWin32Error();
                }
 
                DevInfo.Size = (UInt32)Marshal.SizeOf(DevInfo);
 
                if(0 == SetupDiOpenDeviceInfo(DevInfoSet, DeviceID, IntPtr.Zero, 0, ref DevInfo))
                {
                    SetupDiDestroyDeviceInfoList(DevInfoSet);
                    return ERROR_CLASS_MISMATCH;
                }
                return 0;
            }
 
            public static void ReleaseDeviceInfoSet(IntPtr DevInfoSet)
            {
                SetupDiDestroyDeviceInfoList(DevInfoSet);
            }
 
            public static UInt32 RemoveDevice(string DeviceID)
            {
                UInt32 ResultCode = 0;
                IntPtr DevInfoSet = IntPtr.Zero;
                SP_DEVINFO_DATA DevInfo = new SP_DEVINFO_DATA();
 
                ResultCode = GetDeviceInformation(DeviceID, ref DevInfoSet, ref DevInfo);
 
                if (0 == ResultCode)
                {
                    if (1 != SetupDiRemoveDevice(DevInfoSet, ref DevInfo))
                    {
                        ResultCode = (UInt32)Marshal.GetLastWin32Error();
                    }
                    ReleaseDeviceInfoSet(DevInfoSet);
                }
 
                return ResultCode;
            }
        }
    }
'@
    Add-Type -TypeDefinition $RemoveDeviceSource
 
    $DeviceManager = [Microsoft.Windows.Diagnosis.DeviceManagement_Remove]
    $ErrorCode = $DeviceManager::RemoveDevice($DeviceID)
    return $ErrorCode
}

[array]$ghostNICs = @( Get-WmiObject -Class win32_networkadapter -Filter "ServiceName IS NULL" -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $nicRegex } )

if( ! $ghostNICs -or ! $ghostNICs.Count )
{
    [string]$warning = "No ghost NICs found"
    if( $PSBoundParameters[ 'nicRegex' ] )
    {
        $warning += " matching $nicRegex"
    }
    Write-Warning -Message $warning
    exit 0
}

Write-Verbose -Message "Found $($ghostNICs.Count) ghost NICs"

if( ! ( $netclass = Get-ChildItem -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Class\*" | Get-ItemProperty -Name Class | Where-Object Class -ceq 'Net' ) )
{
    Throw "Unable to find Net class in registry"
}

ForEach( $ghostNIC in $ghostNICs )
{
    [string]$index = "$($ghostNIC.Index)".PadLeft( 4 , '0' )

    if( ! ( $nic = Get-ItemProperty -Path (Join-Path -Path $netclass.PSPath -ChildPath $index) ) )
    {
        Write-Warning -Message "Unable to find `"$($ghostNIC.Name)`" index $index in $($class.PSPath)"
    }
    elseif( $PSCmdlet.ShouldProcess( $ghostNIC.Name , 'Remove' ) )
    {
        $guid = $nic.NetCfgInstanceId
        $devid = $nic.DeviceInstanceId
        Write-Verbose -Message "Removing $($ghostNIC.Name) : GUID $guid device id $devid"

        [int]$result = RemoveDevice( $devid )
        if( $result -ne 0 )
        {
            Write-Error -Message "Error $result returned from removing device"
        }
        else
        {
            if( ( $netcfg = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Network\$($netclass.PSChildName)\$guid\Connection" -ErrorAction SilentlyContinue ) -and $netcfg.PnpInstanceID -eq $devid)
            {
                Remove-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Network\$($netclass.PSChildName)\$guid" -Recurse
            }
            else
            {
                Write-Warning -Message "No corresponding Connection key found in registry so network registry key not removed for guid $guid"
            }
        }
    }
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUX5r/s/zrgGS8460HGVBdhCIk
# 5omgggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFGW/X1vxLa/aHVA/28yq
# 3wVMN0CKMA0GCSqGSIb3DQEBAQUABIIBAIu2dH+Aw/VogyUHbn3fZpWXgzQup4u/
# +dqYvnjAfkRYYHtpKzeCPxaAH3K6IK3Uo45JwbaWjDq0GHXW7WMEp0JfD2ETnA+f
# 3tm3mNncYMgR8UJu0YARcMO0mmqR70O+lq7uQBB7/LNcZXakRXFuU7M62/+V6fov
# PLJTQWwwx1WhB5vtOxEXXdAsJBMgUBg4swp7gcBXdd+OWd1bxckelzJxhwH+rPw/
# J9mHpGuyZbPfp4wOvL4/u4VmxTImn+sYKsxKPTc6o6PRg3pRbWKpVjqPAXo0Vm5/
# 2KmfGSvZu6b74Iuu79wJ5xMCef5YQb729LPsAZOA4h0J8FD2WSl4gvY=
# SIG # End signature block
