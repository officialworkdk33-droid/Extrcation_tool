@echo off
rem The server starts in the background (no window) every time you log in to Windows.
cd /d "%~dp0"
if not exist PERMIT_SERVER.exe ( echo Build PERMIT_SERVER.exe first ^(BUILD_SERVER_EXE.bat^). & pause & exit /b 1 )
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Startup')+'\Permit Server.lnk'); $s.TargetPath='%~dp0PERMIT_SERVER.exe'; $s.Arguments='--background --no-browser'; $s.WorkingDirectory='%~dp0'; $s.WindowStyle=7; $s.Save()"
echo Done - the server starts in the background when you log in. (To go online too, edit server_config.json: "online": true)
echo Delete "Permit Server" from the Startup folder to stop that.
pause
