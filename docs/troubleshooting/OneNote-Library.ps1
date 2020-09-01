# OneNote-Library.ps1
#requires -Version 5.1

# REF: https://github.com/PowerShell/PowerShell/issues/8076
Function exist {
    param($path)
    return ( [string]::Empty -ne $path -and ( Test-Path $path ))
}

Function Find-RegAsm {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('x86', 'x64')]
        [string]${Specify "x86" or "x64" platform}
    )
    $Platform = ${Specify "x86" or "x64" platform}

    $dotNetPath = "$env:windir\Microsoft.NET"
    $getChildPath = "Get-RegAsm-$Platform-ChildPath"
    $regAsm = Join-Path $dotNetPath $(& $getChildPath)

    if (exist $regAsm) {
        return $regAsm
    } else {
        Write-Host -ForegroundColor Red 'RegAsm.exe not found'
        Exit
    }
}

Function Get-Assembly-Path
{
    param(
        [Parameter(Mandatory)]
        [string]${Specify path to MyJournal.Notebook.dll assembly (wildcards OK)}
    )
    $path = ${Specify path to MyJournal.Notebook.dll assembly (wildcards OK)}

    $assembly = (Resolve-Path $path).Path

    if (exist $assembly) {
        return $assembly
    } else {
        throw "$assembly not found"
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

Function Get-RegAsm-x64-ChildPath {
    return 'Framework64\v4.0.30319\RegAsm.exe'
}

Function Get-RegAsm-x86-ChildPath {
    return 'Framework\v4.0.30319\RegAsm.exe'
}

Function Press-Any-Key
{
    if ($Host.Name -notmatch 'ISE') {
        Write-Host 'Press any key to continue. . .' -NoNewline
        $Host.UI.RawUI.ReadKey('NoEcho, IncludeKeyDown') | Out-Null
    }
}
