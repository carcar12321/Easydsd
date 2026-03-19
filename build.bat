@echo off
chcp 65001 > nul 2>&1
title easydsd Builder

echo.
echo  ====================================
echo   easydsd EXE Builder
echo  ====================================
echo.

where python > nul 2>&1
if %errorlevel% neq 0 (
    echo  [ERROR] Python not found.
    echo.
    echo  Please install Python 3.x:
    echo    www.python.org/downloads
    echo.
    echo  Check "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo  Python found:
python --version
echo.

python "%~dp0build_exe.py"

