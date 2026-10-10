@echo off
rem Run as administrator (right-click > Run as administrator) - lets other office PCs open the server.
set PORT=8080
if not "%1"=="" set PORT=%1
net session >nul 2>&1
if errorlevel 1 ( echo Please right-click this file and choose "Run as administrator". & pause & exit /b 1 )
netsh advfirewall firewall delete rule name="Permit Server" >nul 2>&1
netsh advfirewall firewall add rule name="Permit Server" dir=in action=allow protocol=TCP localport=%PORT% profile=private,domain
echo Done - port %PORT% is open for the office network (private / domain networks only).
pause
