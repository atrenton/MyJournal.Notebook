@echo OFF
rem PowerShell Build.ps1 Script Wrapper
    SETLOCAL
    color 1F
    set PS_FILE="%~dpn0.ps1"
    TITLE Executing PowerShell %PS_FILE% script . . .

    if NOT exist %SystemRoot%\system32\config\system (
       echo ERROR: You must run this script as Administrator
       goto :ExitScript
    )

    powershell.exe -ExecutionPolicy RemoteSigned -NoLogo -File %PS_FILE%

:ExitScript
    ENDLOCAL
    pause
