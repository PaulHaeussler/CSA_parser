@echo off
call .\venv\Scripts\activate.bat
git pull
python build_table.py
pause