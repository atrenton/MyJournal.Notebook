# Common-Library.ps1
#requires -Version 5.1

# Set Globally Unique Identifier (GUID) for My Journal Notebook COM Add-in
# Code reference: MyJournal.Notebook.Component.ProgId_Guid constant value
if ($global:AddIn_GUID -eq $null) {
    Set-Variable `
        -Name AddIn_GUID -Value '{B899BB4F-3A1E-4E6E-9040-9B9B65969180}' `
        -Option Constant -Scope Global
}

Function Create-Semantic-Version {
    param(
        [Parameter(Mandatory=$true)][UInt16]$Major,
        [Parameter(Mandatory=$true)][UInt16]$Minor,
        [Parameter(Mandatory=$true)][UInt16]$Patch
    )
    return "{0}.{1}.{2}" -f $Major, $Minor, $Patch
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
        $installationPath = & $vswhere @('-latest', '-products', '*',
            '-requires', 'Microsoft.Component.MSBuild',
            '-version', '15.0', '-property', 'installationPath')

        $getChildPath = "Get-MSBuild-$Platform-ChildPath"
        $msbuild = Join-Path $installationPath $(& $getChildPath '15.0')
        if (not-exist $msbuild) {
            $msbuild = Join-Path $installationPath $(& $getChildPath 'Current')
        }

        if (exist $msbuild) {
            return $msbuild
        } else {
            Write-Host -ForegroundColor Red 'MSBuild.exe not found'
            Exit
        }
    } else {
        Write-Host -ForegroundColor Red 'vswhere.exe not found'
        Exit
    }
}

Function Get-ComAddIn-CodeBase {
    $ErrorActionPreference = 'Stop'
    $fileProtocol = 'file:///'
    $registryPath = "HKLM:\SOFTWARE\Classes\CLSID\$($AddIn_GUID)\InprocServer32"
    $codebase = $null
    try
    {
        $codebase = Get-ItemPropertyValue $registryPath -Name CodeBase
    }
    catch [Management.Automation.ItemNotFoundException]
    {
        Write-Warning 'COM Add-in class is not registered!'
        Press-Any-Key
        Exit
    }

    if ($codebase.StartsWith($fileProtocol)) {
        $codebase = $codebase.Substring($fileProtocol.Length)
    }
    return $codebase.Replace('/', '\')
}

Function Get-MSBuild-x64-ChildPath {
    return "MSBuild\$Args\Bin\amd64\MSBuild.exe"
}

Function Get-MSBuild-x86-ChildPath {
    return "MSBuild\$Args\Bin\MSBuild.exe"
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
        Exit
    }
    return $key.'(default)'
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
    Exit
}

Function Press-Any-Key
{
    if ($Host.Name -notmatch 'ISE') {
        Write-Host 'Press any key to continue. . .' -NoNewline
        $Host.UI.RawUI.ReadKey('NoEcho, IncludeKeyDown') | Out-Null
    }
}

Function Sign-Git-Tag {
    param(
        [Parameter(Mandatory=$true)][string]$TagName,
        [Parameter(Mandatory=$true)][string]$SemVer,
        [string]$M = "`"Release version $SemVer`""
    )
    $ErrorActionPreference = 'SilentlyContinue'
    $cmdLine = "git tag -s -m $M $TagName"; $cmdLine
    Invoke-Expression $cmdLine
    if ($LASTEXITCODE -ne 0) { Handle-NativeCommandError }
    git tag -v $TagName
    if ($LASTEXITCODE -ne 0) { Handle-NativeCommandError }
}

if ($global:OneNote_Bitness -eq $null) {
    Set-Variable `
        -Name OneNote_Bitness -Value $(Get-OneNote-Bitness) `
        -Option Constant -Scope Global
}
