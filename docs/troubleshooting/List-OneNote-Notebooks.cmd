@echo OFF
rem PowerShell List-OneNote-Notebooks.ps1 Script Wrapper
    SETLOCAL
    color 1F
    set PS_FILE="%~dpn0.ps1"
    TITLE Executing PowerShell %~n0.ps1 script . . .

    powershell.exe -ExecutionPolicy Bypass -NoLogo -File %PS_FILE%

:ExitScript
    ENDLOCAL
