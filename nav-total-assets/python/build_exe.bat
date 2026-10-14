@echo off
chcp 65001 > nul
rem Windows PC에서 exe를 직접 만들 때: Python 3.9 이상이 설치되어 있어야 합니다.
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
python build_exe.py
pause
