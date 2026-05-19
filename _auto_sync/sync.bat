@echo off
chcp 65001 > nul
title SMT·건웅 데이터 동기화
cd /d "%~dp0"
echo.
echo ============================================================
echo   SMT서울기연·건웅 양사 데이터 자동 동기화
echo ============================================================
echo.
python sync.py
if errorlevel 1 (
    echo.
    echo [오류] Python 실행 실패. 아래를 확인하세요:
    echo   1) Python이 설치되어 있나요? cmd에서 python --version
    echo   2) openpyxl이 깔려있나요? cmd에서 pip install openpyxl
    pause
)
