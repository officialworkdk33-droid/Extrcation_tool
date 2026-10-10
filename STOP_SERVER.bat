@echo off
rem Stops the Permit Server that runs in the background.
cd /d "%~dp0"
if exist PERMIT_SERVER.exe ( PERMIT_SERVER.exe --stop ) else ( python server.py --stop )
pause
