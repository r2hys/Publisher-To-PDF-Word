@echo off
REM ====================================================================
REM  Launcher for Publisher_To_PDF_WORD.ps1
REM
REM  WHY THIS FILE EXISTS
REM
REM  When a .ps1 is started with "Run with PowerShell", the console
REM  belongs to the script. If PowerShell refuses to run the script at
REM  all, for example because of the execution policy or an AppLocker
REM  rule, the error appears and the window closes instantly, so there
REM  is nothing to read.
REM
REM  A .cmd file owns its own window. The "pause" at the bottom runs no
REM  matter what happened, so the error always stays on screen.
REM
REM  HOW TO USE
REM  Keep this file in the SAME FOLDER as Publisher_To_PDF_WORD.ps1,
REM  then double-click this file instead of the .ps1.
REM ====================================================================

setlocal

title Publisher to PDF / Word converter

echo ====================================================
echo  Publisher to PDF / Word converter - launcher
echo ====================================================
echo.

REM %~dp0 is the folder this .cmd lives in, so the script is found
REM regardless of the current working directory.
set "SCRIPT=%~dp0Publisher_To_PDF_WORD.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: Publisher_To_PDF_WORD.ps1 was not found next to this launcher.
    echo.
    echo   Expected here: %SCRIPT%
    echo.
    echo Keep Run-Publisher_To_PDF_WORD.cmd and Publisher_To_PDF_WORD.ps1
    echo in the same folder.
    echo.
    pause
    exit /b 1
)

echo Script: %SCRIPT%
echo.

REM -NoProfile        ignores user profile scripts, which can fail or slow startup
REM -ExecutionPolicy Bypass  applies to THIS process only and does not
REM                   change any machine or user policy. It is ignored if
REM                   the policy is locked by Group Policy.
REM -File             runs the script

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"

set "PSEXIT=%ERRORLEVEL%"

echo.
echo ====================================================
echo  PowerShell finished with exit code %PSEXIT%
echo ====================================================

if not "%PSEXIT%"=="0" (
    echo.
    echo The script did not finish cleanly.
    echo.
    echo If nothing appeared above, PowerShell was most likely blocked
    echo from running the script at all. Common causes on managed
    echo school devices:
    echo.
    echo   * Execution policy locked by Group Policy
    echo   * AppLocker or App Control blocking script files
    echo   * The .ps1 file is "blocked" after being copied from a share
    echo.
    echo Ask your IT team to confirm which applies, or run
    echo Test-PubConvert.ps1 on this account to check.
)

echo.
pause

endlocal
exit /b %PSEXIT%
