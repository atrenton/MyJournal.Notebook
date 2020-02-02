# Common-Library.ps1

# Set Globally Unique Identifier (GUID) for My Journal Notebook COM Add-in
# Code reference: MyJournal.Notebook.Component.ProgId_GUID constant value
if ($global:AddIn_GUID -eq $null) {
    Set-Variable `
        -Name AddIn_GUID -Value '{B899BB4F-3A1E-4E6E-9040-9B9B65969180}' `
        -Option Constant -Scope Global
}

Function not-exist { -not (Test-Path $args) }
Set-Alias !exist not-exist
Set-Alias exist Test-Path

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

Function Find-MSBuild-v15 {
    #vswhere is included with the installer as of Visual Studio 2017 Update 2 and later
    $vswhere = "$(Get-ProgramFilesPath-x86)\Microsoft Visual Studio\Installer\vswhere.exe"

    if (exist $vswhere) {
        $installationPath = & $vswhere @('-latest', '-products', '*',
            '-requires', 'Microsoft.Component.MSBuild',
            '-version', '[15.0,16.0)', '-property', 'installationPath')

        $msbuild = "$installationPath\MSBuild\15.0\Bin\MSBuild.exe"

        if (exist $msbuild) {
            return $msbuild
        } else {
            Write-Host -ForegroundColor red 'MSBuild.exe not found'
            Exit
        }
    } else {
        Write-Host -ForegroundColor red 'vswhere.exe not found'
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
        $codebase = (Get-ItemProperty -Path $registryPath -Name CodeBase).Codebase
    }
    catch [Management.Automation.ItemNotFoundException]
    {
        Write-Warning 'COM Add-in class is not registered!'
        pause
        Exit
    }

    if ($codebase.StartsWith($fileProtocol)) {
        $codebase = $codebase.Substring($fileProtocol.Length)
    }
    return $codebase.Replace('/', '\')
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
    Write-Host -ForegroundColor red $Error[0].Exception.Message
    $Error.Clear()
    Exit
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
