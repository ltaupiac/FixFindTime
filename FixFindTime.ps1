<#
.SYNOPSIS
Add findtime.microsoft.com to trusted zone to avoid pb with proxy

.DESCRIPTION
Add findtime.microsoft.com to trusted zone to avoid pb with proxy

.FUNCTIONALITY
On-demand

.NOTES
Context:            LocalSystem
Version:            1.0.0.0 - Initial release
Last modified: 		2019/06/24 16:19:08
#>


#
# Constants definition
#
New-Variable -Name 'REMOTE_ACTION_DLL_PATH' `
    -Value "$env:NEXTHINK\RemoteActions\nxtremoteactions.dll" `
    -Option Constant -Scope Script

New-Variable -Name 'REG_FINDTIME' `
    -Value 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\microsoft.com\findtime\' `
    -Option Constant -Scope Script

New-Variable -Name 'REG_KEY' `
    -Value 'https' `
    -Option Constant -Scope Script

#
# Environment checks
#
function Add-NexthinkDLL {
    if (-not (Test-Path -Path $REMOTE_ACTION_DLL_PATH)) { throw 'Nexthink Remote Action DLL not found. ' }
    Add-Type -Path $REMOTE_ACTION_DLL_PATH
}

function Test-RunningAsInteractiveUser {
    $currentIdentity = Get-CurrentIdentity
    if ($currentIdentity -eq $LOCAL_SYSTEM_IDENTITY) {
        throw 'This script must be run as InteractiveUser. '
    }
}

function Get-CurrentIdentity {
    return [Security.Principal.WindowsIdentity]::GetCurrent().User.ToString()
}

function Test-SupportedOSVersion {
    $OSVersion = (Get-OSVersion) -as [version]
    if (-not ($OSVersion)) { throw 'This script could not return OS version. ' }
    if (($OSVersion.Major -ne 6 -or $OSVersion.Minor -ne 1) ) {
        throw 'This script is compatible with Windows 7 only. '
    }
}

function Get-OSVersion {
    return Get-WmiObject -Class Win32_OperatingSystem -Filter 'ProductType = 1' -ErrorAction Stop `
        | Select-Object -ExpandProperty Version
}


#
# Registry management
#
function Test-RegistryKey ([string]$Key, [string]$Property) {
    return $null -ne (Get-ItemProperty -Path $Key `
                                       -Name $Property `
                                       -ErrorAction SilentlyContinue)
}

function Get-RegistryKey ([string]$Key, [string]$Property) {
    return (Get-ItemProperty -Path $Key `
                             -Name $Property `
                             -ErrorAction SilentlyContinue).$Property
}

function Set-RegistryKey ([string]$Key, [string]$Property, [string]$Type, [string]$Value) {
    if (-not (Test-Path -Path $Key)) { [void](New-Item -Path $Key -Force) }
    [void](New-ItemProperty -Path $Key `
                            -Name $Property `
                            -PropertyType $Type `
                            -Value $Value -Force)
}


#
# O365 Information
#
function Set-FindTime () {
    if( (get-RegistryKey -key $REG_FINDTIME -property $REG_KEY) -ne 2) {
        Set-RegistryKey -Key $REG_FINDTIME -Property $REG_KEY -Value 2 -Type 'DWord'
    }
}

#
# Main script flow
#
$ExitCode = 0
try {
    Add-NexthinkDLL
    Test-RunningAsInteractiveUser
    Test-SupportedOSVersion
    Set-FindTime
} catch {
    $host.ui.WriteErrorLine($_.ToString())
    $ExitCode = 1
} finally {
    [Environment]::Exit($ExitCode)
} 


# SIG # Begin signature block
# MIIMXQYJKoZIhvcNAQcCoIIMTjCCDEoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUY3dZTNfOdsPA+hlaMM8X0B6L
# CVuggglTMIIEdDCCA1ygAwIBAgIBAzANBgkqhkiG9w0BAQUFADCBvjELMAkGA1UE
# BhMCRlIxDzANBgNVBAgTBkZyYW5jZTEdMBsGA1UEBxMUQm91bG9nbmUtQmlsbGFu
# Y291cnQxHjAcBgNVBAoTFUxhIEZyYW5jYWlzZSBEZXMgSmV1eDESMBAGA1UECxMJ
# TEZESiBJU0VDMSowKAYDVQQDEyFMYSBGcmFuY2Fpc2UgRGVzIEpldXggLSBBQyBS
# YWNpbmUxHzAdBgkqhkiG9w0BCQEWEHNlY29wZXJAbGZkai5jb20wHhcNMDUwMTE0
# MTc0NDE5WhcNMjUwMTA5MTc0NDE5WjCByzELMAkGA1UEBhMCRlIxDzANBgNVBAgT
# BkZyYW5jZTEdMBsGA1UEBxMUQm91bG9nbmUtQmlsbGFuY291cnQxHjAcBgNVBAoT
# FUxhIEZyYW5jYWlzZSBEZXMgSmV1eDESMBAGA1UECxMJTEZESiBJU0VDMTcwNQYD
# VQQDEy5MYSBGcmFuY2Fpc2UgRGVzIEpldXggLSBBQyBTaWduYXR1cmUgUHJvZ3Jh
# bW1lMR8wHQYJKoZIhvcNAQkBFhBzZWNvcGVyQGxmZGouY29tMIGfMA0GCSqGSIb3
# DQEBAQUAA4GNADCBiQKBgQC3OJ6csswr2soUXhkj38CM4dNN8LFWZV2KdyUrHjha
# BGdx53OYQxRtlv27QtAnwPR7NOH3wFZGcWIjnxxitkLTGagrLc2DxaU5ePUy+SpU
# tezxbFfDYP9sO+viv1qCuVD/1Oot2v3U+P6ISli3Zhn6XD+gm3ssfZLL4k/g93VN
# bQIDAQABo4HxMIHuMAwGA1UdEwQFMAMBAf8wCwYDVR0PBAQDAgEGMGoGA1UdHwRj
# MGEwLqAsoCqGKGh0dHBzOi8vY2VydHMuZmRqZXV4LmNvbS9jYS9yb290L2NybC5j
# cmwwL6AtoCuGKWh0dHBzOi8vY2VydHMucHJvZC5mZGouZnIvY2Evcm9vdC9jcmwu
# Y3JsMEIGCWCGSAGG+EIBDQQ1FjNUaGlzIGNlcnRpZmljYXRlIGlzIHVzZWQgZm9y
# IGlzc3VlaW5nIHN1Yi1DQSBjZXJ0cy4wIQYJYIZIAYb4QgECBBQWEkBjcnRfcm9v
# dF9iYXNlX3VybDANBgkqhkiG9w0BAQUFAAOCAQEATwm7OA/XusW4p76dXvHhq3yT
# /WOkK+hoK1IEtGUHdAezqcP8FYQhQ0ivXMHytXtOP0DoagDy14JBUeoUu9jZLBUV
# /T5NAW1XD4DqsS5bsRxB8/tlkahIROq5pNPftel1wGjH4N6/bCOw4SpICuOdvgqZ
# 36KjtZL5PhBzyRlcwgoUltukIF2uAXvKHS9c15R1vWpRxxw27LiP9kOOq/yYNWj3
# Pr2za3erkUYPuswc+Qgjni5f0ocosvFCZmggEwpBfGDocywLhTjIS0mFBF/OOIDh
# EGdoDgkHPBVFckzpfmkiaVFDe+fGyBr9wxt0q+8huqwQjSVTV2z/CFOE1u3HEjCC
# BNcwggRAoAMCAQICAgDmMA0GCSqGSIb3DQEBCwUAMIHLMQswCQYDVQQGEwJGUjEP
# MA0GA1UECBMGRnJhbmNlMR0wGwYDVQQHExRCb3Vsb2duZS1CaWxsYW5jb3VydDEe
# MBwGA1UEChMVTGEgRnJhbmNhaXNlIERlcyBKZXV4MRIwEAYDVQQLEwlMRkRKIElT
# RUMxNzA1BgNVBAMTLkxhIEZyYW5jYWlzZSBEZXMgSmV1eCAtIEFDIFNpZ25hdHVy
# ZSBQcm9ncmFtbWUxHzAdBgkqhkiG9w0BCQEWEHNlY29wZXJAbGZkai5jb20wHhcN
# MTgwMTIyMTYxMjEzWhcNMjMwMTIxMTYxMjEzWjCBqDELMAkGA1UEBhMCRlIxDzAN
# BgNVBAgMBkZSQU5DRTEYMBYGA1UEBwwPTW91c3N5IGxlIHZpZXV4MR4wHAYDVQQK
# DBVMYSBGcmFuY2Fpc2UgRGVzIEpldXgxDDAKBgNVBAsMA0RORTEeMBwGA1UEAwwV
# bnh0YXNzaWduLnByb2QuZmRqLmZyMSAwHgYJKoZIhvcNAQkBFhFsdGF1cGlhY0Bs
# ZmRqLmNvbTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAL5TNiXlpodP
# jg3wv2B0F9kMcVWEE3gBGoJbW06B6SKdYMVBQ6ba/XxaLkOGcjp8DUOepbLTsh4I
# ffmW7W+yZuKWAe0VTwyGa16sU6mqVEw7y+tVK5WS1otguK8tSsBOK3RWYraqD9So
# xsc3XUPnIcPlmvkscdgOQhmCBBfvHLWT9So5XT7UkSl9Xs214Akw4JjmTTEphlmb
# S+rqyE5PxvkGdyz7ZOVSBe4KLoYafv3a+bXSwx5fxpCSbx+5KzbVIVBEQU8HFWSb
# TGW5ixDK9vYCyEK8ZBbjJ+4K9D7hz8TnR4fiWhYjGTLHge5gsjSi2YAhw9MJOXA0
# qO36uNukYuUCAwEAAaOCAWUwggFhMAsGA1UdDwQEAwIGwDATBgNVHSUEDDAKBggr
# BgEFBQcDAzBYBgNVHRIEUTBPhiRodHRwOi8vY2VydHMubGZkai5jb20vY2Evc2ln
# bi9jYS5jcnSGJ2h0dHA6Ly9jZXJ0cy5wcm9kLmZkai5mci9jYS9zaWduL2NhLmNy
# dDBmBgNVHR8EXzBdMCugKaAnhiVodHRwOi8vY2VydHMubGZkai5jb20vY2Evc2ln
# bi9jcmwuY3JsMC6gLKAqhihodHRwOi8vY2VydHMucHJvZC5mZGouZnIvY2Evc2ln
# bi9jcmwuY3JsMEUGCWCGSAGG+EIBDQQ4FjZUaGlzIGNlcnRpZmljYXRlIGlzIHVz
# ZWQgZm9yIENvZGVTaWduaW5nQ2VydHMgc2lnbmluZy4wIQYJYIZIAYb4QgECBBQW
# EkBjcnRfc2lnbl9iYXNlX3VybDARBglghkgBhvhCAQEEBAMCBBAwDQYJKoZIhvcN
# AQELBQADgYEAZRxUbMgwzp9GJyXcPOffsfB38YcJzOsXx3L9vTVR1HDnmsHY4aAx
# 86HXOA7Cd7pYNgeMQvHul0Rz6KlRnP0/nWWfgCy2JuH6MJM5buE7FO4t8rdKPMTD
# W9sC2SxcwCkF9tM+0DoJVuO/Evpcm2DX3j/Xgo/N0pZMP5MJFTltUJ8xggJ0MIIC
# cAIBATCB0jCByzELMAkGA1UEBhMCRlIxDzANBgNVBAgTBkZyYW5jZTEdMBsGA1UE
# BxMUQm91bG9nbmUtQmlsbGFuY291cnQxHjAcBgNVBAoTFUxhIEZyYW5jYWlzZSBE
# ZXMgSmV1eDESMBAGA1UECxMJTEZESiBJU0VDMTcwNQYDVQQDEy5MYSBGcmFuY2Fp
# c2UgRGVzIEpldXggLSBBQyBTaWduYXR1cmUgUHJvZ3JhbW1lMR8wHQYJKoZIhvcN
# AQkBFhBzZWNvcGVyQGxmZGouY29tAgIA5jAJBgUrDgMCGgUAoHgwGAYKKwYBBAGC
# NwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUihnUwFtb
# b4qUnv/G/5hR7qLVqZAwDQYJKoZIhvcNAQEBBQAEggEAGnEwDEOnN2EPHOxyNRsl
# drWVqXnG5q5+myazAbpc9deET2aXiLzxE1vweg1SpJWlLEdBYQ5iUXgjf796hDWA
# s9igSxpWNTPVSJqGE6RZUALytLPRgfDAOnl97I+JTiP/oBhS0zw/MTJcaTe7FXe9
# q44W/yAtSeh3D+cP1iEJX7UqfochKQugL3GOSx/ihhIYj1TzzccR+mPWWIdGoS1r
# sA3uwHGfm0FUnJUnCc0Xvli+zdS6McBvl4rCMIbtnuPt4/+hQHWf8NswzAFMqmqe
# rW67kvC6KHhoVUsdUEry/5qAr5JGMoSTf3o4ZWe7LumjVN2lKbf+SElo6Wey81+O
# TQ==
# SIG # End signature block
