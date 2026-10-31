@echo off
REM Switch to the folder where this bat file lives.
REM (Running "as administrator" starts cmd in C:\Windows\System32,
REM  which would otherwise break relative paths.)
cd /d "%~dp0"

REM ============================================================
REM  HWInfo Tool - one-click build script (PyInstaller, single EXE)
REM  Recommended: run under 32-bit Python 3.8.10
REM  (covers Windows 7 ~ Windows 11, both 32/64-bit targets)
REM  NOTE: keep this file ASCII-only. Do NOT add Chinese text,
REM  or cmd.exe (GBK codepage) will corrupt the commands.
REM ============================================================

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] python not found in PATH.
    echo Please install 32-bit Python 3.8.10 from python.org
    echo and check "Add Python to PATH" during setup.
    pause
    exit /b 1
)

echo [1/3] Python version:
python --version
echo.

echo [2/3] Installing dependencies...
python -m pip install --upgrade pip
python -m pip install pywin32==306 pyinstaller==5.13.2
if errorlevel 1 (
    echo [ERROR] Dependency installation failed. Check your network.
    pause
    exit /b 1
)
echo.

echo [3/3] Building EXE...
REM Close any running instance first, otherwise Windows locks
REM dist\HWInfoTool.exe and the build fails with WinError 5.
taskkill /f /im HWInfoTool.exe >nul 2>nul
python -m PyInstaller --noconfirm --clean --onefile --windowed --name HWInfoTool --hidden-import pythoncom --hidden-import pywintypes hwinfo_tool.py
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build OK: dist\HWInfoTool.exe
echo ============================================================
pause
