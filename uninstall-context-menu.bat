@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
rem Remove only this application's user-specific command; do not touch file associations.
for %%E in (.docx .pdf) do (
    reg query "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\DocToMdConverter" >nul 2>nul
    if not errorlevel 1 (
        reg delete "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\DocToMdConverter" /f >nul
        if errorlevel 1 goto :error
    )
)
echo [OK] DOCX/PDF to Markdown context menu removed.
if not defined CI pause
exit /b 0
:error
echo [ERROR] Could not remove the DOCX/PDF context menu.
if not defined CI pause
exit /b 1
