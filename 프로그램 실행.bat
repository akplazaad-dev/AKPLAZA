@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM ============================================================
REM  카드 사은혜택 프로그램 (창 GUI) 실행
REM  - 창(GUI)으로 열립니다. 검은 명령창 없이 뜨도록 pythonw 사용.
REM ============================================================

REM 파이썬 런처(py)가 있으면 창 전용 pyw 로 실행 (검은 창 없음)
py -3 --version >nul 2>nul && ( start "" pyw -3 "%~dp0사은혜택_프로그램.py" & goto :end )

REM py 가 없으면 python/pythonw 로 시도
python --version >nul 2>nul && ( start "" pythonw "%~dp0사은혜택_프로그램.py" & goto :end )

echo [오류] 파이썬이 설치되어 있지 않거나, 설치할 때 PATH 등록을 안 했습니다.
echo        먼저 "최초설치.bat" 을 실행하거나, "사용설명서.md" 의 1단계를 따라 하세요.
echo.
pause

:end
