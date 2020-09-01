# About-MyJournal-Notebook.ps1

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

# Load the COM Add-in assembly
$assembly = [Reflection.Assembly]::LoadFile("$(Get-ComAddIn-CodeBase)")

$description = [Reflection.CustomAttributeExtensions]::GetCustomAttribute(`
    $assembly, [Reflection.AssemblyDescriptionAttribute]).Description

$copyright = [Reflection.CustomAttributeExtensions]::GetCustomAttribute(`
    $assembly, [Reflection.AssemblyCopyrightAttribute]).Copyright

$version = [Reflection.CustomAttributeExtensions]::GetCustomAttribute(`
    $assembly, [Reflection.AssemblyInformationalVersionAttribute]).`
    InformationalVersion

$msg = "{0}`r`n{1}`r`nProduct Version: {2}" -f $description, $copyright, $version
Display-MsgBox $msg | Out-Null
