# List-OneNote-Notebooks.ps1

# Load the OneNote common script library
. "$PSScriptRoot\OneNote-Library.ps1"

[int]$bits=[IntPtr]::Size * 8

if (( $(Get-OneNote-Bitness) -eq '32-bit') -and ( $bits -ne 32 ))
{
    $this = $MyInvocation.MyCommand.Path
    Write-Host Loading 32-bit PowerShell. . .
    & "$env:windir\SysWOW64\WindowsPowerShell\v1.0\PowerShell.exe" -File $this
    Exit
}

$ErrorActionPreference = 'Stop'
$OneNote = $null
try
{
    $OneNote = New-Object -ComObject OneNote.Application
}
catch [System.Runtime.InteropServices.COMException]
{
    Write-Host -ForegroundColor Red 'OneNote COM API is not available.'
    Press-Any-Key
    Exit
}

$SpecialLocation = [string]::Empty
$OneNote.GetSpecialLocation('slDefaultNotebookFolder', [ref]$SpecialLocation)
echo "Default Notebook Folder: $SpecialLocation`r`n"

[xml]$Hierarchy = $null
$OneNote.GetHierarchy([string]::Empty, 'hsNotebooks', [ref]$Hierarchy)

echo 'OneNote Notebooks'
echo '-----------------'
# Filters out sensitive OneDrive path information
$Hierarchy.Notebooks.Notebook | Format-Table name, nickname, `
    @{Name = "path"; Expression = {$_.path `
        -replace 'https://d.docs.live.net/[a-fA-F0-9]+/', 'OneDrive:/'}}
Press-Any-Key
