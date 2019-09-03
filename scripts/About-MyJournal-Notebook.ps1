# About-MyJournal-Notebook.ps1

If ( [IntPtr]::Size * 8 -ne 32 )
{
    $this = $MyInvocation.MyCommand.Path
    C:\Windows\SysWOW64\WindowsPowerShell\v1.0\PowerShell.exe -File $this
    Exit
    # NOTE: The Assembly::LoadFile statement below must run in 32-bit instance
    # of PowerShell in order to successfully load the COM Add-in assembly.
}

# Load the common script library
. "$PSScriptRoot\Common-Library.ps1"

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
