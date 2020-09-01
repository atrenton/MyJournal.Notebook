# Get-OneNote-Info.ps1

# Load the OneNote common script library
. "$PSScriptRoot\OneNote-Library.ps1"

try
{
    $file = Get-OneNote-AppPath
    $bitness = Get-OneNote-Bitness $file
    $productName = Get-OneNote-ProductName ($file -split '\\')[-2]

    echo 'OneNote Desktop Software Information'
    echo '------------------------------------'
    echo "App Path : $file"
    echo "Bitness  : $bitness"
    echo "`r`nOneNote $productName is installed."
}
catch
{
    Write-Host -ForegroundColor Red $_.Exception.Message
}
Press-Any-Key
