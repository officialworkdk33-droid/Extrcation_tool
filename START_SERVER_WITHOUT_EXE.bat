@echo off
title PERMIT SERVER
cd /d "%~dp0"
python -m pip install -q -r requirements.txt
python server.py
pause
