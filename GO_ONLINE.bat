@echo off
rem Starts the server in the background AND its internet address (Cloudflare Tunnel). You can close this window.
rem The internet address is on the Settings page (and the GitHub fixed link if you set it up).
cd /d "%~dp0"
if exist PERMIT_SERVER.exe ( PERMIT_SERVER.exe --background --online ) else ( python -m pip install -q -r requirements.txt & python server.py --background --online )
timeout /t 8 >nul
