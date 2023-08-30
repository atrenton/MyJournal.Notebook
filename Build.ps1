# Build.ps1 -- Builds MyJournal.Notebook VS2017+ Solution using MSBuild

# Load the common script library
. "$PSScriptRoot\scripts\Common-Library.ps1"

if ($OneNote_Bitness -eq '64-bit') {
    $Platform = 'x64'
} else {
    $Platform = 'x86'
}

if ($global:MSBuild_EXE -eq $null) {
    Set-Variable `
        -Name MSBuild_EXE  -Value $(Find-MSBuild $Platform) `
        -Option Constant -Scope Global
}

Set-Location $PSScriptRoot

#-------------------------------------------------------------------------------
# MSBuild properties
#-------------------------------------------------------------------------------
# NOTE: THE FOLLOWING 2 PROPERTIES ARE MUTUALLY EXCLUSIVE; USE ONE OR THE OTHER
#
# To specify a semantic version, use the /p:Version property:
# EXAMPLE: '/p:Version=16.0.0-rc.1'
#
# To specify Git Commit SHA-1 hash, use the /p:SourceRevisionId property:
# EXAMPLE: "/p:SourceRevisionId=g$(Git-Latest-Commit)"
#-------------------------------------------------------------------------------
$properties = @('/p:Configuration=Release', "/p:Platform=$Platform",
                "/p:SourceRevisionId=g$(Git-Latest-Commit)")

$sln = '"{0}"' -f "$PSScriptRoot\src\MyJournal.Notebook.sln"

& $MSBuild_EXE $sln $properties /t:'Clean;Build' /restore
