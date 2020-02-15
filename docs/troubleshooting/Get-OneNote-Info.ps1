# Get-OneNote-Info.ps1

# Load the OneNote common script library
. "$PSScriptRoot\OneNote-Library.ps1"

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

$file = $key.'(default)'
$path = $key.'Path'

try
{
    echo 'OneNote Desktop Software Information'
    echo '------------------------------------'
    echo "(Default) : $file"
    echo "Path      : $path"
    echo "Bitness   : $(Get-OneNote-Bitness -ExeFilePath $file)"
    Display-OneNote-Version -OfficeProductName ($file -split '\\')[-2]
}
catch
{
    Write-Host -ForegroundColor Red $_.Exception.Message
}
Press-Any-Key
