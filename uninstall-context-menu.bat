@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
set "REGKEY=HKCU\Software\Classes\SystemFileAssociations\.docx\shell\DocToMdConverter"
reg query "%REGKEY%" >nul 2>nul
if errorlevel 1 (
    echo [OK] The context menu is already absent.
    if not defined CI pause
    exit /b 0
)
rem Delete only our own verb, never DOCX associations or other apps.
reg delete "%REGKEY%" /f >nul
if errorlevel 1 (
    echo [ERROR] Could not remove the context menu.
    if not defined CI pause
    exit /b 1
)
echo [OK] DOCX to Markdown context menu removed.
if not defined CI pause
exit /b 0
