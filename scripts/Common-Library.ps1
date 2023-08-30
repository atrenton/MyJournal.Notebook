# Common-Library.ps1

#requires -version 5.1

# Common Code
if ($global:CommonCode -eq $null) {
    $code = @"
using System;
using System.Windows.Forms;

namespace CommonCode
{
    // To ensure a dialog window shows on top, you need a window handle from the
    // owner process. This handle must implement System.Windows.Forms.IWin32Window.
    // The code below creates a wrapper class that implements IWin32Window. The
    // IWin32Window interface contains only a single property that must be implemented
    // to expose the underlying handle.
    public class Win32Window : IWin32Window
    {
        public Win32Window(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; private set; }
    }
}
"@
    $type = Add-Type -TypeDefinition $code `
        -ReferencedAssemblies 'System.Windows.Forms.dll' -PassThru

    Set-Variable -Name CommonCode -Value $type -Option Constant -Scope Global
}

# Windows API Functions
if ($global:Win32API -eq $null) {
    $signatures = '
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);
    '
    $type = Add-Type -MemberDefinition $signatures -Name WindowsAPI -PassThru
    Set-Variable -Name Win32API -Value $type -Option Constant -Scope Global
}

# Set Globally Unique Identifier (GUID) for My Journal Notebook COM Add-in
# Code reference: MyJournal.Notebook.ComClassId.Connect constant value
if ($global:AddIn_GUID -eq $null) {
    Set-Variable `
        -Name AddIn_GUID -Value '{B899BB4F-3A1E-4E6E-9040-9B9B65969180}' `
        -Option Constant -Scope Global
}

Function Create-Semantic-Version {
    param(
        [Parameter(Mandatory,Position=0)][UInt16]$Major,
        [Parameter(Mandatory,Position=1)][UInt16]$Minor,
        [Parameter(Mandatory,Position=2)][UInt16]$Patch,
        [Parameter(Mandatory,Position=3)]
        [AllowEmptyString()]
# REF: https://semver.org/#is-there-a-suggested-regular-expression-regex-to-check-a-semver-string
        [ValidatePattern('((0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*)(\.(0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*))*)?')]
        [string]$PreRelease
    )
    if ($PreRelease -eq [string]::Empty) {
        return "{0}.{1}.{2}" -f $Major, $Minor, $Patch
    } else {
        return "{0}.{1}.{2}-{3}" -f $Major, $Minor, $Patch, $PreRelease
    }
}

# Inspired by https://ss64.com/ps/syntax-msgbox.html
Function Display-MsgBox {
    param (
        [string]$message,
        [string]$title = $MyInvocation.PSCommandPath.`
                Substring($MyInvocation.PSScriptRoot.Length + 1)
    )
    Add-Type -AssemblyName System.Windows.Forms | Out-Null

    $buttons = [Windows.Forms.MessageBoxButtons]::OK
    $icon = [Windows.Forms.MessageBoxIcon]::Information
    return [Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)
}

# REF: https://github.com/PowerShell/PowerShell/issues/8076
Function exist {
    param($path)
    return ( [string]::Empty -ne $path -and ( Test-Path $path ))
}
Function not-exist { return ( -not (exist $Args) ) }

Function Find-Dotnet-InstallLocation {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('x86', 'x64')]
        [string]${Specify "x86" or "x64" platform}
    )
    $arch = ${Specify "x86" or "x64" platform}

    #REF: https://github.com/dotnet/designs/blob/main/accepted/2020/install-locations.md#globally-registered-install-location-new
    if ([IntPtr]::Size * 8 -eq 64) {
        $path = "HKLM:\SOFTWARE\WOW6432Node\dotnet\Setup\InstalledVersions\$arch"
    } else {
        $path = "HKLM:\SOFTWARE\dotnet\Setup\InstalledVersions\$arch"
    }
    $key = Get-ItemProperty $path -ErrorAction SilentlyContinue
    if ($key) {
        return $key.InstallLocation
    } else {
        $errorParams = @{
            Message = "Can't find install location for dotnet.exe"
            Category = 'NotInstalled'
            ErrorAction  = 'Stop'
        }
        Write-Error @errorParams
    }
}

Function Find-MSBuild {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('x86', 'x64')]
        [string]${Specify "x86" or "x64" platform}
    )
    $Platform = ${Specify "x86" or "x64" platform}

    #vswhere is included with the installer as of Visual Studio 2017 Update 2
    $vswhere = Join-Path $(Get-ProgramFilesPath-x86) `
            'Microsoft Visual Studio\Installer\vswhere.exe'

    if (exist $vswhere) {
    # REF: https://github.com/microsoft/vswhere/wiki/Find-MSBuild#powershell
        $msbuild = (& $vswhere @('-latest',
            '-requires', 'Microsoft.Component.MSBuild',
            '-find', $(& "Get-MSBuild-$Platform-FindPath"))) |
            Select-Object -First 1

        if (exist $msbuild) {
            return $msbuild
        } else {
            Write-Host -ForegroundColor Red 'MSBuild.exe not found'
            break
        }
    } else {
        Write-Host -ForegroundColor Red 'vswhere.exe not found'
        break
    }
}

Function Get-Build-OutputPath {
    [OutputType([string])]
    Param(
        [Parameter(Mandatory,HelpMessage="The configuration to be selected, either 'Debug' or 'Release'")]
        [ValidateSet('Debug', 'Release')]
        [string]$configuration
    )
    Process {
        if ($OneNote_Bitness -eq '64-bit') {
            $platform = 'x64'
        } else {
            $platform = 'x86'
        }
        $addInPath = Resolve-Path "$($MyInvocation.PSScriptRoot)\..\src\add-in"
        $binPath = Join-Path $addInPath 'bin'
        $csproj = Join-Path $addInPath 'Add-In.csproj'

        $xml = [xml](Get-Content $csproj)
        $tfm = $xml.Project.PropertyGroup.TargetFramework

        return "$binPath\$platform\$configuration\$tfm"
    }
}

Function Get-ComHost {
    $ErrorActionPreference = 'Stop'
    $registryPath = "HKLM:\SOFTWARE\Classes\CLSID\$($AddIn_GUID)\InprocServer32"
    $key = $null
    try
    {
        $key = Get-ItemProperty $registryPath
    }
    catch [System.Management.Automation.ItemNotFoundException]
    {
        Write-Warning 'A COM host for the add-in is not registered!'
        Press-Any-Key
        break
    }
    return $key.'(Default)'
}

Function Get-ComAddIn-CodeBase {
    return $(Get-ComHost) -Replace '.comhost',''
}

Function Get-Dotnet-FileNamePath {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('x86', 'x64')]
        [string]${Specify "x86" or "x64" platform}
    )
    $platform = ${Specify "x86" or "x64" platform}
    $dotnet = Join-Path $(Find-Dotnet-InstallLocation $platform) 'dotnet.exe'

    if (exist $dotnet) {
        return $dotnet
    } else {
        $errorParams = @{
            Message = 'dotnet.exe'
            Category = 'ObjectNotFound'
            ErrorAction  = 'Stop'
        }
        Write-Error @errorParams
    }
}

Function Get-MSBuild-x64-FindPath {
    return 'MSBuild\**\Bin\amd64\MSBuild.exe'
}

Function Get-MSBuild-x86-FindPath {
    return 'MSBuild\**\Bin\MSBuild.exe'
}

Function Get-OneNote-Bitness
{
    param (
        [string]$ExeFilePath = $(Get-OneNote-AppPath)
    )
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
        break
    }
    return $key.'(Default)'
}

Function Get-ProgramFilesPath-x86 {
    if ($env:PROCESSOR_ARCHITECTURE -eq 'x86') {
        if ([string]::IsNullOrEmpty($env:PROCESSOR_ARCHITEW6432)) {
            return $env:ProgramFiles
        }
    }
    return ${env:ProgramFiles(x86)}
}

Function Git-Latest-Commit {
    $ErrorActionPreference = 'SilentlyContinue'
    $value = git rev-parse --short HEAD
    if ($LASTEXITCODE -ne 0) { Handle-NativeCommandError }
    return $value
}

Function Handle-NativeCommandError {
    Write-Host -ForegroundColor Red $Error[0].Exception.Message
    $Error.Clear()
    break
}

Function Install-NET-6-Desktop-Runtime {
# REF: https://learn.microsoft.com/en-us/answers/questions/944165/net-desktop-runtime-6-download.html
    param(
        [Parameter(Mandatory)]
        [ValidateSet('x86', 'x64')]
        [string]${Specify "x86" or "x64" platform}
    )
    $platform = ${Specify "x86" or "x64" platform}
    $runtimeFile = "windowsdesktop-runtime-win-$platform.exe"
    $downloadUri = "https://aka.ms/dotnet/6.0/$runtimeFile"

    $tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) `
                -ChildPath ([System.IO.Path]::GetRandomFileName())

    $tmpDirectory = @{
        ItemType = 'Directory'
        Path = $tempDir
        Force = $true
        ErrorAction = 'SilentlyContinue'
    }
    $null = New-Item @tmpDirectory

    Write-Host -ForegroundColor 'Green' -Object "Downloading $downloadUri"

    $exeFilePath = Join-Path -Path $tempDir -ChildPath $runtimeFile
    Invoke-WebRequest -Uri $downloadUri -OutFile $exeFilePath

    Start-Process $exeFilePath -Verb 'RunAs'
}

Function Press-Any-Key
{
    if (($Host.Name -eq 'ConsoleHost') -and (-not $env:WT_SESSION)) {
        Write-Host "`nPress any key to continue. . ." -NoNewline
        $Host.UI.RawUI.ReadKey('NoEcho, IncludeKeyDown') | Out-Null
    }
}

Function Show-OpenFileDialog {
    param (
        [string]$WindowTitle='Please select a file',
        [string]$InitialDirectory = "$PWD.Path",
        [string]$Filter = 'All files (*.*)|*.*',
        [switch]$AllowMultiSelect = $false
    )
    Add-Type -AssemblyName System.Windows.Forms
    $objForm = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = $InitialDirectory
        Filter = $Filter
        Multiselect = $AllowMultiSelect
        Title = $WindowTitle
    }
    if ($Host.Name -eq 'ConsoleHost') { $objForm.ShowHelp = $true }

    $handle = $Win32API::GetForegroundWindow()
    $owner = [CommonCode.Win32Window]::new($handle)
    $objForm.ShowDialog($owner) | Out-Null

    if ($AllowMultiSelect) {
        return $objForm.FileNames
    } else {
        return $objForm.FileName
    }
}

Function Sign-Git-Tag {
    param(
        [Parameter(Mandatory=$true)][string]$TagName,
        [Parameter(Mandatory=$true)][string]$SemVer,
        [string]$M = "`"$(&{
            if ($SemVer.Contains('-')) {
                'Pre-release'
            } else {
                'Release'
            }
        }) version $SemVer`""
    )
    $ErrorActionPreference = 'SilentlyContinue'
    $cmdLine = "git tag -s -m $M $TagName"; $cmdLine
    Invoke-Expression $cmdLine
    if ($LASTEXITCODE -ne 0) { Handle-NativeCommandError }
    git tag -v $TagName
    if ($LASTEXITCODE -ne 0) { Handle-NativeCommandError }
}

Function YesNo-Prompt {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Caption,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Message
    )
    $type = 'System.Management.Automation.Host.ChoiceDescription'
    $yes = New-Object -TypeName $type -ArgumentList '&Yes'
    $no = New-Object -TypeName $type -ArgumentList '&No'
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)

    $result = $Host.UI.PromptForChoice($Caption, $Message, $choices, 0)
    switch ($result)
    {
        0 {return $true}
        1 {return $false}
    }
}

if ($global:OneNote_Bitness -eq $null) {
    Set-Variable `
        -Name OneNote_Bitness -Value $(Get-OneNote-Bitness) `
        -Option Constant -Scope Global
}

# SIG # Begin signature block
# MIIlCAYJKoZIhvcNAQcCoIIk+TCCJPUCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC7H1gIRWZ6Kumz
# 4K5mVZihYiGPykDyilUu5TbaBcunx6CCHsswggVgMIIESKADAgECAhEA6JzdWUZA
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
# taIChrop7rRaALGMq+P5CslUXdS5anSevUiumDCCBvUwggTdoAMCAQICEDlMJeF8
# oG0nqGXiO9kdItQwDQYJKoZIhvcNAQEMBQAwfTELMAkGA1UEBhMCR0IxGzAZBgNV
# BAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMHU2FsZm9yZDEYMBYGA1UE
# ChMPU2VjdGlnbyBMaW1pdGVkMSUwIwYDVQQDExxTZWN0aWdvIFJTQSBUaW1lIFN0
# YW1waW5nIENBMB4XDTIzMDUwMzAwMDAwMFoXDTM0MDgwMjIzNTk1OVowajELMAkG
# A1UEBhMCR0IxEzARBgNVBAgTCk1hbmNoZXN0ZXIxGDAWBgNVBAoTD1NlY3RpZ28g
# TGltaXRlZDEsMCoGA1UEAwwjU2VjdGlnbyBSU0EgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgIzQwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCkkyhSS88nh3ak
# KRyZOMDnDtTRHOxoywFk5IrNd7BxZYK8n/yLu7uVmPslEY5aiAlmERRYsroiW+b2
# MvFdLcB6og7g4FZk7aHlgSByIGRBbMfDCPrzfV3vIZrCftcsw7oRmB780yAIQrNf
# v3+IWDKrMLPYjHqWShkTXKz856vpHBYusLA4lUrPhVCrZwMlobs46Q9vqVqakSgT
# Nbkf8z3hJMhrsZnoDe+7TeU9jFQDkdD8Lc9VMzh6CRwH0SLgY4anvv3Sg3MSFJua
# TAlGvTS84UtQe3LgW/0Zux88ahl7brstRCq+PEzMrIoEk8ZXhqBzNiuBl/obm36I
# h9hSeYn+bnc317tQn/oYJU8T8l58qbEgWimro0KHd+D0TAJI3VilU6ajoO0ZlmUV
# KcXtMzAl5paDgZr2YGaQWAeAzUJ1rPu0kdDF3QFAaraoEO72jXq3nnWv06VLGKEM
# n1ewXiVHkXTNdRLRnG/kXg2b7HUm7v7T9ZIvUoXo2kRRKqLMAMqHZkOjGwDvorWW
# nWKtJwvyG0rJw5RCN4gghKiHrsO6I3J7+FTv+GsnsIX1p0OF2Cs5dNtadwLRpPr1
# zZw9zB+uUdB7bNgdLRFCU3F0wuU1qi1SEtklz/DT0JFDEtcyfZhs43dByP8fJFTv
# bq3GPlV78VyHOmTxYEsFT++5L+wJEwIDAQABo4IBgjCCAX4wHwYDVR0jBBgwFoAU
# GqH4YRkgD8NBd0UojtE1XwYSBFUwHQYDVR0OBBYEFAMPMciRKpO9Y/PRXU2kNA/S
# lQEYMA4GA1UdDwEB/wQEAwIGwDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoG
# CCsGAQUFBwMIMEoGA1UdIARDMEEwNQYMKwYBBAGyMQECAQMIMCUwIwYIKwYBBQUH
# AgEWF2h0dHBzOi8vc2VjdGlnby5jb20vQ1BTMAgGBmeBDAEEAjBEBgNVHR8EPTA7
# MDmgN6A1hjNodHRwOi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29SU0FUaW1lU3Rh
# bXBpbmdDQS5jcmwwdAYIKwYBBQUHAQEEaDBmMD8GCCsGAQUFBzAChjNodHRwOi8v
# Y3J0LnNlY3RpZ28uY29tL1NlY3RpZ29SU0FUaW1lU3RhbXBpbmdDQS5jcnQwIwYI
# KwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28uY29tMA0GCSqGSIb3DQEBDAUA
# A4ICAQBMm2VY+uB5z+8VwzJt3jOR63dY4uu9y0o8dd5+lG3DIscEld9laWETDPYM
# nvWJIF7Bh8cDJMrHpfAm3/j4MWUN4OttUVemjIRSCEYcKsLe8tqKRfO+9/YuxH7t
# +O1ov3pWSOlh5Zo5d7y+upFkiHX/XYUWNCfSKcv/7S3a/76TDOxtog3Mw/FuvSGR
# GiMAUq2X1GJ4KoR5qNc9rCGPcMMkeTqX8Q2jo1tT2KsAulj7NYBPXyhxbBlewoNy
# kK7gxtjymfvqtJJlfAd8NUQdrVgYa2L73mzECqls0yFGcNwvjXVMI8JB0HqWO8NL
# 3c2SJnR2XDegmiSeTl9O048P5RNPWURlS0Nkz0j4Z2e5Tb/MDbE6MNChPUitemXk
# 7N/gAfCzKko5rMGk+al9NdAyQKCxGSoYIbLIfQVxGksnNqrgmByDdefHfkuEQ81D
# +5CXdioSrEDBcFuZCkD6gG2UYXvIbrnIZ2ckXFCNASDeB/cB1PguEc2dg+X4yiUc
# RD0n5bCGRyoLG4R2fXtoT4239xO07aAt7nMP2RC6nZksfNd1H48QxJTmfiTllUqI
# jCfWhWYd+a5kdpHoSP7IVQrtKcMf3jimwBT7Mj34qYNiNsjDvgCHHKv6SkIciQPc
# 9Vx8cNldeE7un14g5glqfCsIo0j1FfwET9/NIRx65fWOGtS5QDGCBZMwggWPAgEB
# MIGRMHwxCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIx
# EDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEkMCIG
# A1UEAxMbU2VjdGlnbyBSU0EgQ29kZSBTaWduaW5nIENBAhEA6JzdWUZAuzxpjz0C
# 2ZP+JDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAA
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgor
# BgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCIIR1YTTgfFJz7H1MoqDCVQnGTkZ3E
# QCP7kGdsSXQTKDANBgkqhkiG9w0BAQEFAASCAQAj+Y8ZfF71lBfen9tQXteCM4Zs
# qrDOkniQRycMXXNkN8crmNxX03RQgcrdQso+m8UwFWJCQRZnTYgq0bPJAWdJJkRa
# nvjjZqR0HsAKIHUphqMSqN2BCSCBEVi+aZWKl+KEb22Wq+bF7QDbl+4LwjEqlHj6
# zisHknpkWqVTC1b6JFM/Zic1aq8fKD511RGwSlcCC+hoE4rIyjaHsJfo7b1ZhfHI
# tcxew/Ib81yQddu4vKGZ1Zgj3xx8dIP2b9bbDqvEngsG/+q474j4TIy59/ouy2Ak
# JLBmc0kSJxm8jOUqWTlvmgV5H1zDYT39coBF78FsvxprIYkKa1D4x4MRz0w5oYID
# SzCCA0cGCSqGSIb3DQEJBjGCAzgwggM0AgEBMIGRMH0xCzAJBgNVBAYTAkdCMRsw
# GQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAW
# BgNVBAoTD1NlY3RpZ28gTGltaXRlZDElMCMGA1UEAxMcU2VjdGlnbyBSU0EgVGlt
# ZSBTdGFtcGluZyBDQQIQOUwl4XygbSeoZeI72R0i1DANBglghkgBZQMEAgIFAKB5
# MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIzMDgy
# OTA1NDIzNVowPwYJKoZIhvcNAQkEMTIEMFYxlThjsePl7cLGz2EPq4rk/t9hyTmp
# qBMmtCW+uhTR5TxOhwIk5UzzOb+IRaWk1jANBgkqhkiG9w0BAQEFAASCAgBrf+W1
# mc1ZzxAJTCBuLgQsFRNSqdHLIjHwUxnrVpTNqFdGzPGr7RGdH1QBxEbP7Y4gNYj5
# SGCqd4W3ZULc56lYAt3HBfUY5jdKGtAbQ46bDxOqeU5kVY3PJMg+nuAx15/2DP+S
# fGLruQxlA/M2EjJYvnUwb1UcGFCZwb6Qne3EdY5JLtiFVdOWZ6pJk8MU5mhvcntJ
# sP2aUHgs+YPBNlpxNyul6tD6nUb+auoba+Awln2C25miixa5L/K2wJye8I2yDJmd
# cocWxLuHOT+u77IU4FXI4utXmXOL0I7k8lbx5jqDnm8fgNnYUy2ngnU8GXl+FFUw
# /k1+Z0GjjZpA8yYj+EvEz5ltY8D3UPApArhzmokGMAe/Fq9GlLb4mFYfgxPfrvsW
# sdkjIOyL1ERyxHiA8XY56wZ7bLUtu7Yakb55BvrsMDrg/UQDOUl3YMD0Ji51qWXv
# mK5fiVgNhsKlV9t0DldqE8e1mwvzaX7bGow4Oq0D+tIf4gW6iJjeVEBibPw0wiik
# qQHq47+75dBnfwhMyPy9LOFioFjamujmnNNFirIiX07vssXHjy/wKeql/JqjBMFd
# 8NGTqUcIxSbK2HedA23fc2XPUofr9epSTqqVP5MhR5dOMMkBqZAcjQb/Ql5Z5TcL
# LdtGmiX4IGJWGI66bX6z+wGzOmFi09qD5grk6g==
# SIG # End signature block
