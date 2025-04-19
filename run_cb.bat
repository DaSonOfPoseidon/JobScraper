@echo off
cd /d %~dp0

REM === CHECK IF PYTHON IS INSTALLED LOCALLY ===
IF NOT EXIST "embedded_python\python.exe" (
    echo Python not found. Downloading embedded Python...
    powershell -Command ^
        "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.2/python-3.12.2-embed-amd64.zip' -OutFile 'python_embed.zip'"

    echo Extracting embedded Python...
    powershell -Command ^
        "Expand-Archive -Path 'python_embed.zip' -DestinationPath 'embedded_python'"

    echo Cleaning up...
    del python_embed.zip
)

REM === SETUP LOCAL PYTHON ENVIRONMENT ===
IF NOT EXIST "embedded_python\Scripts\pip.exe" (
    echo Setting up pip in embedded Python...
    embedded_python\python.exe -m ensurepip --default-pip
)

REM === INSTALL MISSING DEPENDENCIES ONLY ===
echo Checking Python packages...
SETLOCAL ENABLEDELAYEDEXPANSION
SET PACKAGES=pandas selenium openpyxl tkinterdnd2 tkcalendar python-dateutil
FOR %%P IN (%PACKAGES%) DO (
    embedded_python\python.exe -c "import %%P" 2>NUL
    IF ERRORLEVEL 1 (
        echo Installing: %%P
        embedded_python\python.exe -m pip install %%P --target=embedded_python\lib
    ) ELSE (
        echo Found: %%P
    )
)
ENDLOCAL

REM === LAUNCH THE TOOL ===
echo Starting tool...
embedded_python\python.exe main.py

pause
