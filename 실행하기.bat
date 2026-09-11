@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM --- 파이썬 명령 찾기 ---
set "PYCMD="
py -3 --version >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD python --version >nul 2>nul && set "PYCMD=python"
if not defined PYCMD goto NOPYTHON

REM --- 실제 프로그램 실행 (드래그한 파일이 있으면 그대로 넘겨줍니다) ---
%PYCMD% "%~dp0update_calendar.py" %*

echo.
echo ------------------------------------------------------------
echo  프로그램이 끝났습니다. 이 창은 아무 키나 누르면 닫힙니다.
echo ------------------------------------------------------------
pause
goto END

:NOPYTHON
echo [오류] 파이썬이 설치되어 있지 않거나, 설치할 때 PATH 등록을 안 했습니다.
echo        먼저 최초설치.bat 을 실행하거나, 사용설명서.md 의 1단계를 따라 하세요.
echo.
pause

:END
