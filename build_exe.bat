@echo off
REM ============================================================
REM   Bulk Folder Creator - Windows EXE builder
REM   Just double-click this file on a Windows PC that has
REM   Python 3 installed. The .exe lands in the "dist" folder.
REM ============================================================

echo.
echo Installing build tools (pyinstaller + openpyxl)...
python -m pip install --upgrade pip
python -m pip install pyinstaller openpyxl
if errorlevel 1 (
    echo.
    echo ERROR: Could not install packages. Is Python on your PATH?
    echo Get Python from https://www.python.org/downloads/ and tick
    echo "Add python.exe to PATH" during install.
    pause
    exit /b 1
)

echo.
echo Building the executable...
pyinstaller --onefile --windowed --name "BulkFolderCreator" folder_creator.py
if errorlevel 1 (
    echo.
    echo ERROR: Build failed. See messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   DONE!  Your program is here:
echo   dist\BulkFolderCreator.exe
echo ============================================================
echo.
pause
