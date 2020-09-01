# Edit-MyJournal-Notebook-Config.ps1

# Load the common script library
. "$PSScriptRoot\Common-Library.ps1"

[int]$bits=[IntPtr]::Size * 8

if (( $(Get-OneNote-Bitness) -eq '32-bit') -and ( $bits -ne 32 ))
{
    $this = $MyInvocation.MyCommand.Path
    Write-Host Loading 32-bit PowerShell. . .
    & "$env:windir\SysWOW64\WindowsPowerShell\v1.0\PowerShell.exe" -File $this
    Exit
}

# Check .config file access to determine if we need to Run As Administrator
$config = "$(Get-ComAddIn-CodeBase).config"
$ErrorActionPreference = 'Stop'
$verb = ‘Open’
try { [System.IO.File]::OpenWrite($config).Close() }
catch { $verb = ‘RunAs’ }

# Edit the .config file using Notepad
$process = new-object System.Diagnostics.Process
$process.StartInfo.FileName = 'notepad.exe'
$process.StartInfo.Arguments = $config
$process.StartInfo.Verb = $verb
[void]$process.Start()
