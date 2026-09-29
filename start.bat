@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"
if errorlevel 1 goto :error

echo [DOCX to Markdown] Checking Python...
call :find_python
if defined PYTHON goto :python_ready

where winget >nul 2>nul
if errorlevel 1 goto :missing_python
echo [Setup] Installing Python 3.13 using Windows Package Manager...
winget install --id Python.Python.3.13 --exact --source winget --accept-source-agreements --accept-package-agreements
if errorlevel 1 goto :missing_python
call :find_python
if not defined PYTHON goto :missing_python

:python_ready
echo [Setup] Python found.
set "VENV=.venv"
if exist "%VENV%\Scripts\python.exe" (
    "%VENV%\Scripts\python.exe" -c "import sys; assert sys.version_info >= (3,11)" >nul 2>nul
    if errorlevel 1 set "VENV=.venv-win"
)
if not exist "%VENV%\Scripts\python.exe" (
    echo [Setup] Creating local Python environment...
    %PYTHON% -m venv --system-site-packages "%VENV%"
    if errorlevel 1 goto :error
)
set "VPY=%CD%\%VENV%\Scripts\python.exe"
"%VPY%" -c "import sys; assert sys.version_info >= (3,11)" >nul 2>nul
if errorlevel 1 goto :venv_error

"%VPY%" -m pip --version >nul 2>nul
if errorlevel 1 (
    "%VPY%" -m ensurepip --upgrade
    if errorlevel 1 goto :error
)

rem Install ONLY missing or incompatible libraries. The venv reuses system packages.
"%VPY%" -c "from importlib.metadata import version; import re; v=tuple(map(int,re.match(r'^(\d+)\.(\d+)',version('python-docx')).groups())); assert (1,1) <= v < (2,0)" >nul 2>nul
if errorlevel 1 (
    echo [Setup] Installing python-docx...
    "%VPY%" -m pip install --disable-pip-version-check "python-docx>=1.1,<2"
    if errorlevel 1 goto :error
)
"%VPY%" -c "from importlib.metadata import version; import re; v=tuple(map(int,re.match(r'^(\d+)\.(\d+)',version('lxml')).groups())); assert (4,9) <= v < (7,0)" >nul 2>nul
if errorlevel 1 (
    echo [Setup] Installing lxml...
    "%VPY%" -m pip install --disable-pip-version-check "lxml>=4.9,<7"
    if errorlevel 1 goto :error
)
"%VPY%" -c "from importlib.metadata import version; import re; v=int(re.match(r'^\d+',version('setuptools')).group()); assert v >= 68; version('wheel')" >nul 2>nul
if errorlevel 1 (
    echo [Setup] Installing local build tools...
    "%VPY%" -m pip install --disable-pip-version-check "setuptools>=68" wheel
    if errorlevel 1 goto :error
)

rem Rebuild from local source without dependency downloads or build isolation.
echo [Setup] Building the application...
"%VPY%" -m pip install --quiet --disable-pip-version-check --no-deps --no-build-isolation -e .
if errorlevel 1 goto :error

if not "%~1"=="" goto :launch
"%VPY%" -c "import tkinter" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Tkinter is missing. Install Python from python.org with the Tcl/Tk option enabled.
    goto :error
)

:launch
echo [Setup] Starting DOCX to Markdown...
"%VPY%" -m doc_to_md_converter %*
if errorlevel 1 goto :error
exit /b 0

:find_python
set "PYTHON="
python -c "import sys; assert sys.version_info >= (3,11)" >nul 2>nul
if not errorlevel 1 (
    set "PYTHON=python"
    exit /b 0
)
py -3 -c "import sys; assert sys.version_info >= (3,11)" >nul 2>nul
if not errorlevel 1 (
    set "PYTHON=py -3"
    exit /b 0
)
for %%P in ("%LOCALAPPDATA%\Programs\Python\Python313\python.exe" "%ProgramFiles%\Python313\python.exe" "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%LOCALAPPDATA%\Programs\Python\Python311\python.exe") do (
    if exist "%%~P" (
        "%%~P" -c "import sys; assert sys.version_info >= (3,11)" >nul 2>nul
        if not errorlevel 1 (
            set PYTHON="%%~P"
            exit /b 0
        )
    )
)
exit /b 0

:missing_python
echo [ERROR] Python 3.11+ is required. Install it from https://www.python.org/downloads/windows/
echo [ERROR] If winget was not found, install Windows App Installer or Python manually.
goto :error

:venv_error
echo [ERROR] An incompatible virtual environment was found. Remove .venv-win and retry.
goto :error

:error
echo.
echo [ERROR] Application could not start. See the messages above.
pause
exit /b 1
