@echo off
chcp 65001 >nul
title Driver Manager — Build EXE

echo ============================================
echo  Driver Manager — Building portable EXE
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Install from https://python.org
    pause & exit /b 1
)

:: Install / upgrade PyInstaller
echo [1/3] Installing PyInstaller...
python -m pip install --upgrade pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pip failed. Make sure pip is available.
    pause & exit /b 1
)

:: Clean previous build
echo [2/3] Cleaning previous build...
if exist dist\DriverManager.exe del /f /q dist\DriverManager.exe >nul 2>&1
if exist build rmdir /s /q build >nul 2>&1

:: Build
echo [3/3] Building DriverManager.exe ...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "DriverManager" ^
    --icon NONE ^
    --uac-admin ^
    driver_manager.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. See output above.
    pause & exit /b 1
)

echo.
echo ============================================
echo  SUCCESS!  dist\DriverManager.exe is ready
echo ============================================
echo.
echo Copy dist\DriverManager.exe to any PC — no install needed.
echo.
pause
