@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo   카드 사은혜택 프로그램 - 최초 1회 설치
echo ============================================================
echo.
echo  이 창은 프로그램이 쓰는 '부품'을 설치합니다.
echo  컴퓨터에 딱 한 번만 하면 됩니다. 잠시 기다려 주세요...
echo.

REM --- 파이썬(Python)이 설치돼 있는지 확인합니다 ---
set "PYCMD="
py -3 --version >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD python --version >nul 2>nul && set "PYCMD=python"

if not defined PYCMD (
  echo [오류] 파이썬(Python)이 설치되어 있지 않습니다.
  echo.
  echo   먼저 "사용설명서.md" 파일의 [1단계: 파이썬 설치]를 따라 해 주세요.
  echo   설치할 때 "Add Python to PATH" 체크를 꼭 켜야 합니다.
  echo.
  pause
  exit /b 1
)

echo  (사용하는 파이썬 명령: %PYCMD%)
echo.

%PYCMD% -m pip install --upgrade pip
%PYCMD% -m pip install -r "%~dp0requirements.txt"

echo.
if errorlevel 1 (
  echo [오류] 설치 중 문제가 생겼습니다.
  echo        인터넷 연결을 확인한 뒤 이 파일을 다시 실행해 주세요.
) else (
  echo ============================================================
  echo   설치 완료! 이제 "실행하기.bat" 을 사용하시면 됩니다.
  echo ============================================================
)
echo.
pause
