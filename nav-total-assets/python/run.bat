@echo off
chcp 65001 > nul
rem 사용법: 대상 엑셀과 기준가격대장 폴더를 이 파일 위로 함께 끌어다 놓거나,
rem        명령창에서  run.bat 대상.xlsx 장부폴더 --base-date 2026-09-30
set PYTHONUTF8=1
cd /d "%~dp0"
python -m nav_total_assets %*
echo.
pause
