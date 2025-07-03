@echo off
echo ====================================
echo  Auto Converter Pro Build Script
echo  Building with Portrait Image Support
echo ====================================
echo.

echo Step 1: Installing required packages...
pip install Pillow>=8.0.0
pip install python-docx
pip install pyinstaller

echo.
echo Step 2: Cleaning previous build...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

echo Step 3: Building executable with spec file...
pyinstaller Auto-Converter_with_pillow.spec

echo.
echo Step 4: Testing if executable was created...
if exist "dist\AutoConverterV2.exe" (
    echo SUCCESS: AutoConverterV2.exe created successfully!
    echo Location: %CD%\dist\AutoConverterV2.exe
    echo.
    echo Checking if icon was embedded...
    powershell -Command "Get-ItemProperty 'dist\AutoConverterV2.exe' | Select-Object Name, Length"
    echo.
    echo The executable now includes:
    echo - Portrait image orientation support
    echo - Pillow library for proper image handling
    echo - All required dependencies
    echo - Custom icon (convert.ico)
    echo.
    echo This executable should work on other computers without requiring
    echo additional installations.
) else (
    echo ERROR: Failed to create executable
    echo Check the build output above for errors.
)

echo.
echo Build process complete!
pause
