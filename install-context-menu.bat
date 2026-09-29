@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
set "REGKEY=HKCU\Software\Classes\SystemFileAssociations\.docx\shell\DocToMdConverter"
set "LAUNCHER=%~dp0start.bat"
if not exist "%LAUNCHER%" (
    echo [ERROR] start.bat not found alongside this installer.
    goto :error
)
rem Per-user only: no UAC/admin rights, no background app or Explorer restart.
reg add "%REGKEY%" /ve /t REG_SZ /d "Конвертировать DOCX в Markdown" /f >nul
if errorlevel 1 goto :error
rem Explorer replaces %%1 with the clicked DOCX path, keeping quotes for spaces.
reg add "%REGKEY%\command" /ve /t REG_SZ /d "\"%LAUNCHER%\" \"%%1\" --context-menu" /f >nul
if errorlevel 1 (
    reg delete "%REGKEY%" /f >nul 2>nul
    goto :error
)
echo.
echo [OK] Context menu installed for .docx files (current Windows user).
echo Right-click a DOCX, choose "Show more options", then "Конвертировать DOCX в Markdown".
echo Keep the project folder at this location, or rerun the installer after moving it.
if not defined CI pause
exit /b 0

:error
echo.
echo [ERROR] Could not register the DOCX context menu.
if not defined CI pause
exit /b 1
