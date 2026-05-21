@echo off
setlocal
cd /d "%~dp0"

echo.
echo ============================================================
echo    SMT + GW Wehago Daily Sync
echo ============================================================
echo.

set "PYCMD="
where py >nul 2>nul && set "PYCMD=py"
if not defined PYCMD where python >nul 2>nul && set "PYCMD=python"

if not defined PYCMD (
    echo [ERROR] Python not found in PATH.
    echo.
    echo Open cmd and type:  py --version
    echo If that fails, reinstall Python with [Add to PATH] checked.
    pause
    exit /b 1
)

echo Using Python launcher: %PYCMD%
echo.
%PYCMD% sync.py
set RC=%errorlevel%
echo.
echo ============================================================
if %RC%==0 (
    echo  DONE. Site will refresh in 1-2 minutes.
) else (
    echo  ERROR. Python script exited with code %RC%.
)
echo ============================================================
pause
endlocal
