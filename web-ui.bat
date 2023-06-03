@echo off

cd /d %~dp0

set FLASK_APP=main.py
set FLASK_ENV=development
.\env\Scripts\python -m flask run