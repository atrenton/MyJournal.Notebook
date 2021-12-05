#Get-OneNote-Desktop-Info.ps1

#requires -version 5.1

# REF: https://github.com/PowerShell/PowerShell/issues/8076
Function exist {
    param($path)
    return ( [string]::Empty -ne $path -and ( Test-Path $path ))
}

Function Get-OneNote-AppPath
{
    $ErrorActionPreference = 'Stop'
    $appPaths = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths'
    $program = 'OneNote.exe'
    $key = $null
    try
    {
        $key = Get-ItemProperty -Path "HKLM:\$appPaths\$program"
    }
    catch [System.Management.Automation.ItemNotFoundException]
    {
        Write-Warning "App Paths registry key is missing for $program"
        Press-Any-Key
        Exit
    }
    return $key.'(default)'
}

Function Get-OneNote-Bitness
{
    param ($ExeFilePath = $(Get-OneNote-AppPath))

    if (exist $ExeFilePath) {
        # The following code was inspired by:
        # https://github.com/guyrleech/Microsoft/blob/master/Get%20file%20bitness.ps1
        [int]$MACHINE_OFFSET = 4
        [int]$PE_POINTER_OFFSET = 60

        [hashtable]$bitness = @{
            0x014c = '32-bit'
            0x8664 = '64-bit'
        }
        $data = New-Object System.Byte[] 4096
        $stream = New-Object System.IO.FileStream -ArgumentList $ExeFilePath,Open,Read
        $stream.Read($data, 0, $data.Count) | Out-Null
        $stream.Close()

        [int]$PE_HEADER_ADDR = [System.BitConverter]::ToInt32($data, $PE_POINTER_OFFSET)
        [int]$typeOffset = $PE_HEADER_ADDR + $MACHINE_OFFSET
        [uint16]$machineType = [System.BitConverter]::ToUInt16($data, $typeOffset)

        return $bitness[[int]$machineType]
    } else {
        throw 'ONENOTE.EXE not found'
    }
}

Function Get-OneNote-ProductName
{
    param(
        [ValidateScript({
            if( $_ -match 'Office\d\d' ) {
                $true
            } else {
                throw "Invalid value: $_"
            }
        })][string] $OfficeVersion
    )

    [hashtable]$productName = @{
        12 = '2007'
        14 = '2010'
        15 = '2013'
        16 = '2016'
    }

    $len = $OfficeVersion.Length
    [int]$i = $OfficeVersion.Substring($len - 2)
    return $productName[$i]
}

#REF: https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/reading-installed-software-from-registry
Function Get-Software
{
    param([string]$DisplayName='*')

    $keys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 
            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    
    Get-ItemProperty -Path $keys | 
        Where-Object { $_.DisplayName } |
        Select-Object -Property DisplayName, DisplayVersion |
        Where-Object { $_.DisplayName -match $DisplayName }
}

Function Get-Office-Properties
{
    param([string]$PropertiesToMatch='*')

	$Esc = $([char]27)
    $key = 'HKLM:\SOFTWARE\Microsoft\Office\'
    $useANSI = $Host.Name -notmatch 'ISE'

#   REF: https://docs.microsoft.com/en-us/windows/console/console-virtual-terminal-sequences#cursor-visibility
	if ($useANSI) {
        Write-Host "Searching . . .  $Esc[?25l" -NoNewline
    }
    $registry = Get-ChildItem -Path $key -Recurse

    $activity = "Searching $key"
    $count = 0
    $obj = New-Object PSObject
    $progress = [Math]::Round($registry.Count * 0.1)

    foreach ($item in $registry) {
        $item.Property | Where-Object { $_ -match $PropertiesToMatch } |
        ForEach-Object {
            $property = @{
                MemberType = 'NoteProperty'
                Name = "$_"
                Value = "$($item.GetValue($_))"
            }
            $obj | Add-Member @property -Force
        }
        $pct = ((($count++) / $registry.Count) * 100)
        if ($count % $progress -eq 0) {
            $status = "$([Math]::Round($pct))% Complete"
            Write-Progress -Activity $activity -Status $status -PercentComplete $pct
        }
    }

    Write-Progress -Completed -Activity [string]::Empty
	if ($useANSI) {
    	Write-Host "`r$Esc[?25h" -NoNewline
    }
    return $obj
}

Function Press-Any-Key
{
    if (($Host.Name -eq 'ConsoleHost') -and (-not $env:WT_SESSION)) {
        Write-Host "`nPress any key to continue. . ." -NoNewline
        $Host.UI.RawUI.ReadKey('NoEcho, IncludeKeyDown') | Out-Null
    }
}

try
{
    Clear-Host
    echo 'OneNote Desktop Software Information'
    echo '------------------------------------'

    $appPath = Get-OneNote-AppPath
    $bitness = Get-OneNote-Bitness $appPath
    $productName = Get-OneNote-ProductName ($appPath -split '\\')[-2]

    echo "App Path : $appPath"
    echo "Bitness  : $bitness"
    echo "`nOneNote $productName is installed."

    echo "`nOffice Software Information"
    echo '---------------------------'
    (Get-Software -DisplayName '^Microsoft (365|Office|OneNote)+\s' |
        Format-List | Out-String).Trim()

    echo "`nOffice Deployment Properties"
    echo '----------------------------'
    (Get-Office-Properties -PropertiesToMatch '^(ProductReleaseIds|Platform)$' |
        Format-List | Out-String).Trim()
}
catch
{
    Write-Host -ForegroundColor Red $_.Exception.Message
}
Press-Any-Key

# SIG # Begin signature block
# MIIlGwYJKoZIhvcNAQcCoIIlDDCCJQgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCg8MeRXW8Mjq4U
# fMu5KQazh7OjCezi0wkqYuULzpSMrKCCHt0wggVgMIIESKADAgECAhEA6JzdWUZA
# uzxpjz0C2ZP+JDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJHQjEbMBkGA1UE
# CBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRgwFgYDVQQK
# Ew9TZWN0aWdvIExpbWl0ZWQxJDAiBgNVBAMTG1NlY3RpZ28gUlNBIENvZGUgU2ln
# bmluZyBDQTAeFw0yMDA4MjYwMDAwMDBaFw0yMzExMjUyMzU5NTlaMIGPMQswCQYD
# VQQGEwJVUzEOMAwGA1UEEQwFNDQwNjAxDTALBgNVBAgMBE9oaW8xDzANBgNVBAcM
# Bk1lbnRvcjEeMBwGA1UECQwVNzc5NiBIaWRkZW4gSG9sbG93IERyMRcwFQYDVQQK
# DA5BcnRodXIgVHJlbnRvbjEXMBUGA1UEAwwOQXJ0aHVyIFRyZW50b24wggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC3oQNrBAhdwheKdemNPV7JDkfD27Mv
# e3MhGCuelZrYmG/CFVKC3ikRRwOQdYZT9pETEfCkn5gHCcpYE8bk09O5bNmaew/2
# gdMtdKNSGihp3jLy/xOLpHqbsDirCFVapEdmNmq4HJEsnsWorVco+dVs1wFOROO5
# f/pDc/KGJPupb/gl/Aj0ck0NVsZsf2E5WkUZ3RmxOgQmlIGksqqtyKzNkoBh7ntA
# v/Du9g9ZMqMdGpKd6+wFzEF3QKXbHxIXJ/qXQoK/ZBv20Dh2IDG0A49lMlvBOP4d
# x9pjLYVXu+zcRPWMVJ4KYeeO3jI34fH9Ili437ReTepat4y8lLQdfxllAgMBAAGj
# ggHHMIIBwzAfBgNVHSMEGDAWgBQO4TqoUzox1Yq+wbutZxoDha00DjAdBgNVHQ4E
# FgQUskFrIBdn9M2Ez2ZsWp2+YVbmVgIwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB
# /wQCMAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwEQYJYIZIAYb4QgEBBAQDAgQQMEoG
# A1UdIARDMEEwNQYMKwYBBAGyMQECAQMCMCUwIwYIKwYBBQUHAgEWF2h0dHBzOi8v
# c2VjdGlnby5jb20vQ1BTMAgGBmeBDAEEATBDBgNVHR8EPDA6MDigNqA0hjJodHRw
# Oi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29SU0FDb2RlU2lnbmluZ0NBLmNybDBz
# BggrBgEFBQcBAQRnMGUwPgYIKwYBBQUHMAKGMmh0dHA6Ly9jcnQuc2VjdGlnby5j
# b20vU2VjdGlnb1JTQUNvZGVTaWduaW5nQ0EuY3J0MCMGCCsGAQUFBzABhhdodHRw
# Oi8vb2NzcC5zZWN0aWdvLmNvbTA1BgNVHREELjAsgSo0ODA2NDE4MSthdHJlbnRv
# bkB1c2Vycy5ub3JlcGx5LmdpdGh1Yi5jb20wDQYJKoZIhvcNAQELBQADggEBAAHy
# cT7H0dyT69FCbmyEMD+WisJwFriQ23QmrHukwyIEgN1figdnGbkL9dTTVoRI0ORE
# mbFIP5yw2SwufZmMxdRvQ6s2rIW1GMppUYUttMhgIgHaT8CHerglYdomuKOLhrDH
# aKdeWDGaFQjVfYNT+Q9t3esvQ/76VJDcB9SoMiOgrrnvpRUzeX2ZavCt4Ve5Ei2s
# 7cO/3lzdJlAHUem04xrTidoI1t5/M21tpY6Y/JmLaQm3QAyFqHIDQNCm+ZcT33xp
# Xr+j2dUT20l4JqmIYD4rh2+YWzc1JKrQiO3mzUGxYRGNTJvnz5GuZ+0baIzeG7Kq
# cwFf97QETZeC0CBkFjgwggWBMIIEaaADAgECAhA5ckQ6+SK3UdfTbBDdMTWVMA0G
# CSqGSIb3DQEBDAUAMHsxCzAJBgNVBAYTAkdCMRswGQYDVQQIDBJHcmVhdGVyIE1h
# bmNoZXN0ZXIxEDAOBgNVBAcMB1NhbGZvcmQxGjAYBgNVBAoMEUNvbW9kbyBDQSBM
# aW1pdGVkMSEwHwYDVQQDDBhBQUEgQ2VydGlmaWNhdGUgU2VydmljZXMwHhcNMTkw
# MzEyMDAwMDAwWhcNMjgxMjMxMjM1OTU5WjCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCk5ldyBKZXJzZXkxFDASBgNVBAcTC0plcnNleSBDaXR5MR4wHAYDVQQKExVU
# aGUgVVNFUlRSVVNUIE5ldHdvcmsxLjAsBgNVBAMTJVVTRVJUcnVzdCBSU0EgQ2Vy
# dGlmaWNhdGlvbiBBdXRob3JpdHkwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQCAEmUXNg7D2wiz0KxXDXbtzSfTTK1Qg2HiqiBNCS1kCdzOiZ/MPans9s/B
# 3PHTsdZ7NygRK0faOca8Ohm0X6a9fZ2jY0K2dvKpOyuR+OJv0OwWIJAJPuLodMkY
# tJHUYmTbf6MG8YgYapAiPLz+E/CHFHv25B+O1ORRxhFnRghRy4YUVD+8M/5+bJz/
# Fp0YvVGONaanZshyZ9shZrHUm3gDwFA66Mzw3LyeTP6vBZY1H1dat//O+T23LLb2
# VN3I5xI6Ta5MirdcmrS3ID3KfyI0rn47aGYBROcBTkZTmzNg95S+UzeQc0PzMsNT
# 79uq/nROacdrjGCT3sTHDN/hMq7MkztReJVni+49Vv4M0GkPGw/zJSZrM233bkf6
# c0Plfg6lZrEpfDKEY1WJxA3Bk1QwGROs0303p+tdOmw1XNtB1xLaqUkL39iAigmT
# Yo61Zs8liM2EuLE/pDkP2QKe6xJMlXzzawWpXhaDzLhn4ugTncxbgtNMs+1b/97l
# c6wjOy0AvzVVdAlJ2ElYGn+SNuZRkg7zJn0cTRe8yexDJtC/QV9AqURE9JnnV4ee
# UB9XVKg+/XRjL7FQZQnmWEIuQxpMtPAlR1n6BB6T1CZGSlCBst6+eLf8ZxXhyVeE
# Hg9j1uliutZfVS7qXMYoCAQlObgOK6nyTJccBz8NUvXt7y+CDwIDAQABo4HyMIHv
# MB8GA1UdIwQYMBaAFKARCiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBRTeb9a
# qitKz1SA4dibwJ3ysgNmyzAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB
# /zARBgNVHSAECjAIMAYGBFUdIAAwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2Ny
# bC5jb21vZG9jYS5jb20vQUFBQ2VydGlmaWNhdGVTZXJ2aWNlcy5jcmwwNAYIKwYB
# BQUHAQEEKDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20w
# DQYJKoZIhvcNAQEMBQADggEBABiHUdx0IT2ciuAntzPQLszs8ObLXhHeIm+bdY6e
# cv7k1v6qH5yWLe8DSn6u9I1vcjxDO8A/67jfXKqpxq7y/Njuo3tD9oY2fBTgzfT3
# P/7euLSK8JGW/v1DZH79zNIBoX19+BkZyUIrE79Yi7qkomYEdoiRTgyJFM6iTcky
# s7roFBq8cfFb8EELmAAKIgMQ5Qyx+c2SNxntO/HkOrb5RRMmda+7qu8/e3c70sQC
# kT0ZANMXXDnbP3sYDUXNk4WWL13fWRZPP1G91UUYP+1KjugGYXQjFrUNUHMnREd/
# EF2JKmuFMRTE6KlqTIC8anjPuH+OdnKZDJ3+15EIFqGjX5UwggX1MIID3aADAgEC
# AhAdokgwb5smGNCC4JZ9M9NqMA0GCSqGSIb3DQEBDAUAMIGIMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKTmV3IEplcnNleTEUMBIGA1UEBxMLSmVyc2V5IENpdHkxHjAc
# BgNVBAoTFVRoZSBVU0VSVFJVU1QgTmV0d29yazEuMCwGA1UEAxMlVVNFUlRydXN0
# IFJTQSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTAeFw0xODExMDIwMDAwMDBaFw0z
# MDEyMzEyMzU5NTlaMHwxCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1h
# bmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEkMCIGA1UEAxMbU2VjdGlnbyBSU0EgQ29kZSBTaWduaW5nIENBMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAhiKNMoV6GJ9J8JYvYwgeLdx8nxTP
# 4ya2JWYpQIZURnQxYsUQ7bKHJ6aZy5UwwFb1pHXGqQ5QYqVRkRBq4Etirv3w+Bis
# p//uLjMg+gwZiahse60Aw2Gh3GllbR9uJ5bXl1GGpvQn5Xxqi5UeW2DVftcWkpwA
# L2j3l+1qcr44O2Pej79uTEFdEiAIWeg5zY/S1s8GtFcFtk6hPldrH5i8xGLWGwuN
# x2YbSp+dgcRyQLXiX+8LRf+jzhemLVWwt7C8VGqdvI1WU8bwunlQSSz3A7n+L2U1
# 8iLqLAevRtn5RhzcjHxxKPP+p8YU3VWRbooRDd8GJJV9D6ehfDrahjVh0wIDAQAB
# o4IBZDCCAWAwHwYDVR0jBBgwFoAUU3m/WqorSs9UgOHYm8Cd8rIDZsswHQYDVR0O
# BBYEFA7hOqhTOjHVir7Bu61nGgOFrTQOMA4GA1UdDwEB/wQEAwIBhjASBgNVHRMB
# Af8ECDAGAQH/AgEAMB0GA1UdJQQWMBQGCCsGAQUFBwMDBggrBgEFBQcDCDARBgNV
# HSAECjAIMAYGBFUdIAAwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC51c2Vy
# dHJ1c3QuY29tL1VTRVJUcnVzdFJTQUNlcnRpZmljYXRpb25BdXRob3JpdHkuY3Js
# MHYGCCsGAQUFBwEBBGowaDA/BggrBgEFBQcwAoYzaHR0cDovL2NydC51c2VydHJ1
# c3QuY29tL1VTRVJUcnVzdFJTQUFkZFRydXN0Q0EuY3J0MCUGCCsGAQUFBzABhhlo
# dHRwOi8vb2NzcC51c2VydHJ1c3QuY29tMA0GCSqGSIb3DQEBDAUAA4ICAQBNY1Dt
# RzRKYaTb3moqjJvxAAAeHWJ7Otcywvaz4GOz+2EAiJobbRAHBE++uOqJeCLrD0bs
# 80ZeQEaJEvQLd1qcKkE6/Nb06+f3FZUzw6GDKLfeL+SU94Uzgy1KQEi/msJPSrGP
# JPSzgTfTt2SwpiNqWWhSQl//BOvhdGV5CPWpk95rcUCZlrp48bnI4sMIFrGrY1rI
# FYBtdF5KdX6luMNstc/fSnmHXMdATWM19jDTz7UKDgsEf6BLrrujpdCEAJM+U100
# pQA1aWy+nyAlEA0Z+1CQYb45j3qOTfafDh7+B1ESZoMmGUiVzkrJwX/zOgWb+W/f
# iH/AI57SHkN6RTHBnE2p8FmyWRnoao0pBAJ3fEtLzXC+OrJVWng+vLtvAxAldxU0
# ivk2zEOS5LpP8WKTKCVXKftRGcehJUBqhFfGsp2xvBwK2nxnfn0u6ShMGH7EezFB
# cZpLKewLPVdQ0srd/Z4FUeVEeN0B3rF1mA1UJP3wTuPi+IO9crrLPTru8F4Xkmht
# yGH5pvEqCgulufSe7pgyBYWe6/mDKdPGLH29OncuizdCoGqC7TtKqpQQpOEN+BfF
# tlp5MxiS47V1+KHpjgolHuQe8Z9ahyP/n6RRnvs5gBHN27XEp6iAb+VT1ODjosLS
# Wxr6MiYtaldwHDykWC6j81tLB9wyWfOHpxptWDCCBuwwggTUoAMCAQICEDAPb6zd
# Zph0fKlGNqd4LbkwDQYJKoZIhvcNAQEMBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkgQ2l0eTEeMBwGA1UEChMV
# VGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVVU0VSVHJ1c3QgUlNBIENl
# cnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTE5MDUwMjAwMDAwMFoXDTM4MDExODIz
# NTk1OVowfTELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3Rl
# cjEQMA4GA1UEBxMHU2FsZm9yZDEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSUw
# IwYDVQQDExxTZWN0aWdvIFJTQSBUaW1lIFN0YW1waW5nIENBMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAyBsBr9ksfoiZfQGYPyCQvZyAIVSTuc+gPlPv
# s1rAdtYaBKXOR4O168TMSTTL80VlufmnZBYmCfvVMlJ5LsljwhObtoY/AQWSZm8h
# q9VxEHmH9EYqzcRaydvXXUlNclYP3MnjU5g6Kh78zlhJ07/zObu5pCNCrNAVw3+e
# olzXOPEWsnDTo8Tfs8VyrC4Kd/wNlFK3/B+VcyQ9ASi8Dw1Ps5EBjm6dJ3VV0Rc7
# NCF7lwGUr3+Az9ERCleEyX9W4L1GnIK+lJ2/tCCwYH64TfUNP9vQ6oWMilZx0S2U
# TMiMPNMUopy9Jv/TUyDHYGmbWApU9AXn/TGs+ciFF8e4KRmkKS9G493bkV+fPzY+
# DjBnK0a3Na+WvtpMYMyou58NFNQYxDCYdIIhz2JWtSFzEh79qsoIWId3pBXrGVX/
# 0DlULSbuRRo6b83XhPDX8CjFT2SDAtT74t7xvAIo9G3aJ4oG0paH3uhrDvBbfel2
# aZMgHEqXLHcZK5OVmJyXnuuOwXhWxkQl3wYSmgYtnwNe/YOiU2fKsfqNoWTJiJJZ
# y6hGwMnypv99V9sSdvqKQSTUG/xypRSi1K1DHKRJi0E5FAMeKfobpSKupcNNgtCN
# 2mu32/cYQFdz8HGj+0p9RTbB942C+rnJDVOAffq2OVgy728YUInXT50zvRq1naHe
# lUF6p4MCAwEAAaOCAVowggFWMB8GA1UdIwQYMBaAFFN5v1qqK0rPVIDh2JvAnfKy
# A2bLMB0GA1UdDgQWBBQaofhhGSAPw0F3RSiO0TVfBhIEVTAOBgNVHQ8BAf8EBAMC
# AYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggrBgEFBQcDCDARBgNV
# HSAECjAIMAYGBFUdIAAwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC51c2Vy
# dHJ1c3QuY29tL1VTRVJUcnVzdFJTQUNlcnRpZmljYXRpb25BdXRob3JpdHkuY3Js
# MHYGCCsGAQUFBwEBBGowaDA/BggrBgEFBQcwAoYzaHR0cDovL2NydC51c2VydHJ1
# c3QuY29tL1VTRVJUcnVzdFJTQUFkZFRydXN0Q0EuY3J0MCUGCCsGAQUFBzABhhlo
# dHRwOi8vb2NzcC51c2VydHJ1c3QuY29tMA0GCSqGSIb3DQEBDAUAA4ICAQBtVIGl
# M10W4bVTgZF13wN6MgstJYQRsrDbKn0qBfW8Oyf0WqC5SVmQKWxhy7VQ2+J9+Z8A
# 70DDrdPi5Fb5WEHP8ULlEH3/sHQfj8ZcCfkzXuqgHCZYXPO0EQ/V1cPivNVYeL9I
# duFEZ22PsEMQD43k+ThivxMBxYWjTMXMslMwlaTW9JZWCLjNXH8Blr5yUmo7Qjd8
# Fng5k5OUm7Hcsm1BbWfNyW+QPX9FcsEbI9bCVYRm5LPFZgb289ZLXq2jK0KKIZL+
# qG9aJXBigXNjXqC72NzXStM9r4MGOBIdJIct5PwC1j53BLwENrXnd8ucLo0jGLmj
# wkcd8F3WoXNXBWiap8k3ZR2+6rzYQoNDBaWLpgn/0aGUpk6qPQn1BWy30mRa2Coi
# wkud8TleTN5IPZs0lpoJX47997FSkc4/ifYcobWpdR9xv1tDXWU9UIFuq/DQ0/yy
# sx+2mZYm9Dx5i1xkzM3uJ5rloMAMcofBbk1a0x7q8ETmMm8c6xdOlMN4ZSA7D0Gq
# H+mhQZ3+sbigZSo04N6o+TzmwTC7wKBjLPxcFgCo0MR/6hGdHgbGpm0yXbQ4CStJ
# B6r97DDa8acvz7f9+tCjhNknnvsBZne5VhDhIG7GrrH5trrINV0zdo7xfCAMKneu
# taIChrop7rRaALGMq+P5CslUXdS5anSevUiumDCCBwcwggTvoAMCAQICEQCMd6AA
# j/TRsMY9nzpIg41rMA0GCSqGSIb3DQEBDAUAMH0xCzAJBgNVBAYTAkdCMRswGQYD
# VQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDElMCMGA1UEAxMcU2VjdGlnbyBSU0EgVGltZSBT
# dGFtcGluZyBDQTAeFw0yMDEwMjMwMDAwMDBaFw0zMjAxMjIyMzU5NTlaMIGEMQsw
# CQYDVQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQH
# EwdTYWxmb3JkMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxLDAqBgNVBAMMI1Nl
# Y3RpZ28gUlNBIFRpbWUgU3RhbXBpbmcgU2lnbmVyICMyMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEAkYdLLIvB8R6gntMHxgHKUrC+eXldCWYGLS81fbvA
# +yfaQmpZGyVM6u9A1pp+MshqgX20XD5WEIE1OiI2jPv4ICmHrHTQG2K8P2SHAl/v
# xYDvBhzcXk6Th7ia3kwHToXMcMUNe+zD2eOX6csZ21ZFbO5LIGzJPmz98JvxKPiR
# mar8WsGagiA6t+/n1rglScI5G4eBOcvDtzrNn1AEHxqZpIACTR0FqFXTbVKAg+Zu
# SKVfwYlYYIrv8azNh2MYjnTLhIdBaWOBvPYfqnzXwUHOrat2iyCA1C2VB43H9QsX
# Hprl1plpUcdOpp0pb+d5kw0yY1OuzMYpiiDBYMbyAizE+cgi3/kngqGDUcK8yYIa
# IYSyl7zUr0QcloIilSqFVK7x/T5JdHT8jq4/pXL0w1oBqlCli3aVG2br79rflC7Z
# GutMJ31MBff4I13EV8gmBXr8gSNfVAk4KmLVqsrf7c9Tqx/2RJzVmVnFVmRb945S
# D2b8mD9EBhNkbunhFWBQpbHsz7joyQu+xYT33Qqd2rwpbD1W7b94Z7ZbyF4UHLmv
# hC13ovc5lTdvTn8cxjwE1jHFfu896FF+ca0kdBss3Pl8qu/CdkloYtWL9QPfvn2O
# DzZ1RluTdsSD7oK+LK43EvG8VsPkrUPDt2aWXpQy+qD2q4lQ+s6g8wiBGtFEp8z3
# uDECAwEAAaOCAXgwggF0MB8GA1UdIwQYMBaAFBqh+GEZIA/DQXdFKI7RNV8GEgRV
# MB0GA1UdDgQWBBRpdTd7u501Qk6/V9Oa258B0a7e0DAOBgNVHQ8BAf8EBAMCBsAw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDBABgNVHSAEOTA3
# MDUGDCsGAQQBsjEBAgEDCDAlMCMGCCsGAQUFBwIBFhdodHRwczovL3NlY3RpZ28u
# Y29tL0NQUzBEBgNVHR8EPTA7MDmgN6A1hjNodHRwOi8vY3JsLnNlY3RpZ28uY29t
# L1NlY3RpZ29SU0FUaW1lU3RhbXBpbmdDQS5jcmwwdAYIKwYBBQUHAQEEaDBmMD8G
# CCsGAQUFBzAChjNodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29SU0FUaW1l
# U3RhbXBpbmdDQS5jcnQwIwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28u
# Y29tMA0GCSqGSIb3DQEBDAUAA4ICAQBKA3iQQjPsexqDCTYzmFW7nUAGMGtFavGU
# DhlQ/1slXjvhOcRbuumVkDc3vd/7ZOzlgreVzFdVcEtO9KiH3SKFple7uCEn1KAq
# MZSKByGeir2nGvUCFctEUJmM7D66A3emggKQwi6Tqb4hNHVjueAtD88BN8uNovq4
# WpquoXqeE5MZVY8JkC7f6ogXFutp1uElvUUIl4DXVCAoT8p7s7Ol0gCwYDRlxOPF
# w6XkuoWqemnbdaQ+eWiaNotDrjbUYXI8DoViDaBecNtkLwHHwaHHJJSjsjxusl6i
# 0Pqo0bglHBbmwNV/aBrEZSk1Ki2IvOqudNaC58CIuOFPePBcysBAXMKf1TIcLNo8
# rDb3BlKao0AwF7ApFpnJqreISffoCyUztT9tr59fClbfErHD7s6Rd+ggE+lcJMfq
# RAtK5hOEHE3rDbW4hqAwp4uhn7QszMAWI8mR5UIDS4DO5E3mKgE+wF6FoCShF0DV
# 29vnmBCk8eoZG4BU+keJ6JiBqXXADt/QaJR5oaCejra3QmbL2dlrL03Y3j4yHiDk
# 7JxNQo2dxzOZgjdE1CYpJkCOeC+57vov8fGP/lC4eN0Ult4cDnCwKoVqsWxo6Srk
# ECtuIf3TfJ035CoG1sPx12jjTwd5gQgT/rJkXumxPObQeCOyCSziJmK/O6mXUczH
# RDKBsq/P3zGCBZQwggWQAgEBMIGRMHwxCzAJBgNVBAYTAkdCMRswGQYDVQQIExJH
# cmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEkMCIGA1UEAxMbU2VjdGlnbyBSU0EgQ29kZSBTaWduaW5n
# IENBAhEA6JzdWUZAuzxpjz0C2ZP+JDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEE
# AYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwG
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCY6Vqv
# qv5JPwwTJAh+7ax91WeyAulPafWE2z3NmKqbDDANBgkqhkiG9w0BAQEFAASCAQBT
# W97E4+IEEVtV79vKm2VaPVJnfnWXZueU1lt00ajrrpG83DaRlqCKi6SAcPLLtRjw
# yZk6/9O9DJbXCM7L8dShY5FETKumdpTVQJH3/CxE856zN12L8OClbvqhbtAeKTPh
# vzIucy7J4/OSrhoAr1O+QVFZaRsJ9ASZL2yvWNpdGdhzZsVpld5cRMnw850QdtVI
# rhiSSJ+gxVsPcswDpJwwPtybJlPb1dSJJd7U37DW5biig1QSpOkKBpkgX40i4dyh
# 9xP4ZQKD4fMFMsLYyASLOQdEE21Dx6DOHZliVokM9PGYKIE/XNgQeccrvSYVNx75
# XF6JXb03kKcQBTYugVDooYIDTDCCA0gGCSqGSIb3DQEJBjGCAzkwggM1AgEBMIGS
# MH0xCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAO
# BgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDElMCMGA1UE
# AxMcU2VjdGlnbyBSU0EgVGltZSBTdGFtcGluZyBDQQIRAIx3oACP9NGwxj2fOkiD
# jWswDQYJYIZIAWUDBAICBQCgeTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwG
# CSqGSIb3DQEJBTEPFw0yMTEwMjkxNzMwMzVaMD8GCSqGSIb3DQEJBDEyBDCC2RIX
# hxcsc+gZmczTR1rC2Gm8a1Qy7PF5QtwJ7TT/K8ydVbBGFpuNY0FjI2HbKLowDQYJ
# KoZIhvcNAQEBBQAEggIAUH6SR9Y3+9ISPOThhDjpza2ODfWqhAWioVNx1KjFZo3y
# 2kjzxh7frhHwr2Q924u84SczY8F67cjJ3K1A0Bsch0dJ/Jc9fx8sgz01OnWIH1Vi
# OoR4CMWigXLnBTjz8cOE8byVta0DadeyATO2cLbOUOCwfJpWjUkIl6PgI6cI5doq
# nlWeWEt376+bwa4lij3jOLinO6zRI48QTo2nLLyuoJW+m2hFfQVL3YwacA6NXzbB
# vdIP2dy+53ENrT5E8IF7dIKypcirI4YQUmLFohktQ0nKflGWC4S3/Dtorx3TKQIK
# xy29WvQYQzRBuvWBpbHAy54d7NLQrFtxhJzyxmD5xpcEmSURncmQbpUHzH460wJL
# /lO5WPfXWGhgXETW9rPWT5BeVOYqTyfYzNJBzA+Js0DvfF1lRhpptxAhYb0mnDlF
# qC48fNKIGj6y7PGA2WqFCmeqBN0YrclQUhHgF5VnKmeA9c15Kpd/UK4rw8nf24zW
# apt12jPs8QCLM2bMu8dnYxW2cozkhVMpwEzJYgSgOShAmiYUsd/08HgcgsaxNAcf
# cHBo9BjDf8KJPpVpV03Znnizr9GHVkl0Iho7C4OCYlPJlpxJQziAJazLDaZ4TytM
# mBjFK+pIroaLIunjHqgYp9bAC5Rr9CmPRmTGlrziOIWUiVdVUqSeQ92VvjP2usY=
# SIG # End signature block
