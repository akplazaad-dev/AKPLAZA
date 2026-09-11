@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM --- 파이썬(Python) 명령 찾기 ---
set "PYCMD="
py -3 --version >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD python --version >nul 2>nul && set "PYCMD=python"

if not defined PYCMD (
  echo [오류] 파이썬(Python)이 설치되어 있지 않습니다.
  echo        "사용설명서.md" 의 [1단계: 파이썬 설치] 를 먼저 따라 해 주세요.
  pause
  exit /b 1
)

REM --- 실제 프로그램 실행 (드래그한 파일이 있으면 그대로 넘겨줍니다) ---
%PYCMD% "%~dp0update_calendar.py" %*

REM 프로그램이 오류로 멈췄을 때 창이 바로 닫히지 않게 합니다.
if errorlevel 1 pause
