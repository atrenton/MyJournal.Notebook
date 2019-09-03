# Build.ps1 -- Builds MyJournal.Notebook VS2017 Solution using MSBuild
#Requires –Version 4
#Requires -RunAsAdministrator

# Load the common script library
. "$PSScriptRoot\scripts\Common-Library.ps1"

if ($global:MSBuild_EXE -eq $null) {
    Set-Variable `
        -Name MSBuild_EXE  -Value $(Find-MSBuild-v15) `
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
$properties = @('/p:Configuration=Release', '/p:Platform=x86',
                "/p:SourceRevisionId=g$(Git-Latest-Commit)")

$sln = '"{0}"' -f "$PSScriptRoot\src\MyJournal.Notebook.sln"

& $MSBuild_EXE $sln $properties /t:Restore

# MSBuild registers the OneNote COM Add-in; requires Run as Administrator option
& $MSBuild_EXE $sln $properties /t:'Clean;Build'
