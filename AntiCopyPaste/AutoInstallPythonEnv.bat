@echo off
setlocal enabledelayedexpansion

REM Get the latest Python version
echo Fetching the latest Python version...
for /f "tokens=*" %%i in ('curl -s https://www.python.org/ftp/python/ ^| findstr /r /c:"<a href=""[0-9]+\.[0-9]+\.[0-9]+/""') do (
    set "line=%%i"
    for /f "tokens=2 delims=/" %%j in ("!line!") do (
        set "version=%%j"
        if "!version!" gtr "!latest_version!" set "latest_version=!version!"
    )
)

REM Download Python installer
echo Downloading Python %latest_version%...
curl -o python-installer.exe https://www.python.org/ftp/python/%latest_version%/python-%latest_version%-amd64.exe

REM Install Python silently
echo Installing Python...
python-installer.exe /quiet InstallAllUsers=1 PrependPath=1

REM Verify Python installation
echo Verifying Python installation...
python --version
if %ERRORLEVEL% neq 0 (
    echo Python installation failed!
    exit /b 1
)

REM Upgrade pip
echo Upgrading pip...
python -m ensurepip
python -m pip install --upgrade pip

REM Create virtual environment in the current folder
echo Creating virtual environment...
python -m venv venv

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate

REM Install required libraries in virtual environment
echo Installing required libraries...
pip install openpyxl pandas python-pptx PyPDF2 PyMuPDF python-docx pillow

echo Installation complete!
pause
