@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
set "LAUNCHER=%~dp0start.bat"
if not exist "%LAUNCHER%" (
    echo [ERROR] start.bat not found alongside this installer.
    goto :error
)
rem Register context actions only for this Windows user, not system-wide.
for %%E in (.docx .pdf) do (
    reg add "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\DocToMdConverter" /ve /t REG_SZ /d "Конвертировать DOCX/PDF в Markdown" /f >nul
    if errorlevel 1 goto :rollback
    rem The quoted %%1 placeholder is replaced by Explorer with the selected file.
    reg add "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\DocToMdConverter\command" /ve /t REG_SZ /d "\"%LAUNCHER%\" \"%%1\" --context-menu" /f >nul
    if errorlevel 1 goto :rollback
)
echo.
echo [OK] Context menu registered for both .docx and .pdf (current user only).
echo Right-click either file type, choose "Show more options", then the converter.
echo After moving the project folder, rerun this installer to update its path.
if not defined CI pause
exit /b 0

:rollback
for %%E in (.docx .pdf) do reg delete "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\DocToMdConverter" /f >nul 2>nul
:error
echo [ERROR] Could not register the context menu.
if not defined CI pause
exit /b 1
