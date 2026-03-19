@echo off
chcp 65001 > nul 2>&1
title easydsd - DART 변환 도구

:: Python 설치 확인
where python > nul 2>&1
if %errorlevel% neq 0 goto NO_PYTHON

:: 라이브러리 자동 설치 (최초 1회)
python -c "import flask, openpyxl" > nul 2>&1
if %errorlevel% neq 0 (
    echo  필요한 라이브러리 설치 중... (최초 1회만 실행됩니다)
    python -m pip install flask openpyxl --quiet
)

:: 실행
python "%~dp0dart_gui.py"
goto END

:NO_PYTHON
echo.
echo  Python이 설치되어 있지 않습니다.
echo.
echo  www.python.org/downloads 에서 설치 후
echo  설치 시 "Add Python to PATH" 를 반드시 체크하세요.
echo.
pause

:END
