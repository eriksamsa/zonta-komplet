@echo off
cd /d "%~dp0\.."
if not exist .venv (
  py -m venv .venv
)
call .venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python -m pip install pyinstaller
pyinstaller --onefile --windowed --name ZONTA_KOMPLET src\main.py
pause
