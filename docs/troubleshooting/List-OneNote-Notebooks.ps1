# List-OneNote-Notebooks.ps1

If ( [IntPtr]::Size * 8 -ne 32 )
{
    $this = $MyInvocation.MyCommand.Path
    & "$env:windir\SysWOW64\WindowsPowerShell\v1.0\PowerShell.exe" -File $this
    Exit
    # NOTE: The OneNote.Application COM object must be loaded in a 32-bit
    # instance of PowerShell.
}

# Load the OneNote common script library
. "$PSScriptRoot\OneNote-Library.ps1"

$ErrorActionPreference = 'Stop'
$OneNote = $null
try
{
    $OneNote = New-Object -ComObject OneNote.Application
}
catch [System.Runtime.InteropServices.COMException]
{
    Write-Host -ForegroundColor red 'OneNote COM API is not available.'
    Press-Any-Key
    Exit
}

[xml]$Hierarchy = $null
$OneNote.GetHierarchy([string]::Empty, 'hsNotebooks', [ref]$Hierarchy)

echo 'OneNote Notebooks'
echo '-----------------'
# Filters out sensitive OneDrive path information
$Hierarchy.Notebooks.Notebook | Format-Table name, nickname, `
    @{Name = "path"; Expression = {$_.path `
        -replace 'https://d.docs.live.net/[a-fA-F0-9]+/', 'OneDrive:/'}}
Press-Any-Key
