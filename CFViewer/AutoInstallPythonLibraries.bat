@echo off
REM Download Python installer
echo Downloading Python...
curl -o python-installer.exe https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe

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

REM Install required libraries
echo Installing required libraries...
pip install openpyxl
pip install pandas
pip install python-pptx
pip install PyPDF2
pip install pymupdf
pip install python-docx
pip install pillow
pip install python-pptx
pip install PyQt5
echo Installation complete!
pause