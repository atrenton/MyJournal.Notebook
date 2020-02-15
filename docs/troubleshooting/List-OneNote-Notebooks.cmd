@echo OFF
rem PowerShell List-OneNote-Notebooks.ps1 Script Wrapper
    SETLOCAL
    color 1F
    set PS_FILE="%~dpn0.ps1"
    TITLE Executing PowerShell %~n0.ps1 script . . .

    call :Set_PS_32-Bit_EXE
    %PS_EXE% -ExecutionPolicy Bypass -NoLogo -File %PS_FILE%

:ExitScript
    ENDLOCAL
    goto :EOF

:Set_PS_32-Bit_EXE
::: ============================================================================
::: SUBROUTINE: Set PowerShell's 32-Bit EXE File Name
::: ============================================================================
    set SYSWOW64_DIR=%windir%\SysWOW64

    if exist "%SYSWOW64_DIR%" (
      set PS_EXE="%SYSWOW64_DIR%\WindowsPowerShell\v1.0\powershell.exe"
    ) else (
      set PS_EXE=powershell.exe
    )
    goto :EOF
