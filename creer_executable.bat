@echo off
echo ========================================
echo    CREATING CSV TO ASC EXECUTABLE
echo ========================================
echo.

if not exist "csv_to_asc.py" (
    echo ERROR: csv_to_asc.py file does not exist in this folder!
    pause
    exit /b 1
)

echo [1/4] Checking PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not installed. Installing...
    pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: Cannot install PyInstaller!
        pause
        exit /b 1
    )
) else (
    echo PyInstaller already installed.
)

echo.
echo [2/4] Cleaning old files...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "csv_to_asc.spec" del "csv_to_asc.spec"

echo.
echo [3/4] Creating executable...
if exist "icon.ico" (
    echo Custom icon detected: icon.ico
    python -m PyInstaller --onefile --windowed --icon=icon.ico --name "CSV_to_ASC" csv_to_asc.py
) else (
    echo No custom icon found, using default
    python -m PyInstaller --onefile --windowed --name "CSV_to_ASC" csv_to_asc.py
)

if errorlevel 1 (
    echo ERROR: Executable creation failed!
    pause
    exit /b 1
)

echo.
echo [4/4] Checking result...
if exist "dist\CSV_to_ASC.exe" (
    echo.
    echo ========================================
    echo        SUCCESS!
    echo ========================================
    echo.
    echo Executable created successfully:
    echo  ^> dist\CSV_to_ASC.exe
    echo.
    echo File size:
    for %%A in ("dist\CSV_to_ASC.exe") do echo  ^> %%~zA bytes
    echo.
    echo You can now distribute this .exe file
    echo without needing Python installed!
    echo.
    
    set /p choice="Do you want to open the folder containing the executable? (y/n): "
    if /i "%choice%"=="y" (
        explorer "dist"
    )
    
) else (
    echo ERROR: Executable was not created!
)

echo.
echo Cleaning temporary files...
if exist "build" rmdir /s /q "build"
if exist "__pycache__" rmdir /s /q "__pycache__"

echo.
echo Press any key to close...
pause >nul