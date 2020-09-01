# Unregister-MyJournal-Notebook.ps1
#requires -RunAsAdministrator

# Load the common script library
. "$PSScriptRoot\OneNote-Library.ps1"

$assembly = Get-Assembly-Path
$typelib = [System.IO.Path]::ChangeExtension($assembly, '.tlb')

if ($global:RegAsm_EXE -eq $null) {
    if ( $(Get-OneNote-Bitness) -eq '64-bit' ) {
        $Platform = 'x64'
    } else {
        $Platform = 'x86'
    }
    Set-Variable `
        -Name RegAsm_EXE  -Value $(Find-RegAsm $Platform) `
        -Option Constant -Scope Global
}

# Unregister .NET assembly and type library
# REF: https://docs.microsoft.com/en-us/previous-versions/dotnet/netframework-4.0/tzat5yw6(v=vs.100)
& $RegAsm_EXE @("`"$assembly`"", "/tlb:`"$typelib`"", '/unregister')
