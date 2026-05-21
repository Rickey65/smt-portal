@echo off
chcp 65001 > nul
title 위하고 데이터 동기화
cd /d "%~dp0"
echo.
echo ============================================================
echo   SMT서울기연·건웅 위하고 데이터 자동 동기화
echo ============================================================
echo.

REM python launcher(py)와 python.exe 둘 다 시도
where py >nul 2>&1
if %errorlevel%==0 (
    py sync.py
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        python sync.py
    ) else (
        echo [오류] Python을 찾을 수 없습니다.
        echo.
        echo Windows 시작메뉴에서 "cmd" 검색 후 다음 명령을 입력해 보세요:
        echo    python --version
        echo    py --version
        echo.
        echo 둘 다 안 되면 Python을 재설치하고
        echo 반드시 [Add Python to PATH] 체크박스를 선택하세요.
    )
)
echo.
pause
