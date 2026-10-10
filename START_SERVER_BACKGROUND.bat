@echo off
rem Starts the Permit Server with NO window - you can close this window, the server keeps running.
rem Stop it with STOP_SERVER.bat (or Settings > Stop the server on this PC).
cd /d "%~dp0"
if exist PERMIT_SERVER.exe ( PERMIT_SERVER.exe --background %* ) else ( python -m pip install -q -r requirements.txt & python server.py --background %* )
timeout /t 8 >nul
