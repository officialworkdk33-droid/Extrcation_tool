@echo off
setlocal
title PERMIT_SERVER - build the server exe (OCR included)
cd /d "%~dp0"
echo ============================================================
echo  PERMIT SERVER - build ONE exe: web server + database + all permit tools
echo  (Danfoss, Festo, Agilent, MSD export, any supplier) + OCR + UPDATE HS CODE list
echo ============================================================
python --version >nul 2>&1
if errorlevel 1 (
  echo Python is not installed on THIS PC ^(needed only for building^).
  echo Install Python 3.10 or newer from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" during install, then run this file again.
  pause
  exit /b 1
)
for %%F in (server.py toolrun.py danfoss_engine.py danfoss_docs.py danfoss_festo_docs.py festo_engine.py festo_docs.py agilent_engine.py permit_engine.py readers.py export_docs.py msd_engine.py msd_server_engine.py jfood_server_engine.py sfa_rules.py doc_analysis.py declaration_rules.py liquor_rules.py permit_core.py embedded_data.py PLUGIN_TEMPLATE.py HOW_TO_ADD_A_SUPPLIER.txt ocr_runner.py xml_export.py frtimport_skeleton.xml dec-upload-template-with-description.xlsx web\index.html web\app.js web\tradenet.js web\app.css web\theme.css web\dashboard.js web\logo.png web\mark.png web\favicon.png) do (
  if not exist "%%F" ( echo %%F is missing - unzip ALL files into this folder. & pause & exit /b 1 )
)

rem ---------------------------------------------------------------- 1. Tesseract OCR -> folder Tesseract-OCR here
if exist "Tesseract-OCR\tesseract.exe" goto tess_check
for %%D in ("%ProgramFiles%\Tesseract-OCR" "%ProgramFiles(x86)%\Tesseract-OCR" "%LOCALAPPDATA%\Programs\Tesseract-OCR" "C:\Tesseract-OCR") do (
  if exist "%%~D\tesseract.exe" (
    echo Copying Tesseract from %%~D ...
    robocopy "%%~D" "Tesseract-OCR" /E /NFL /NDL /NJH /NJS /NP >nul
    goto tess_check
  )
)
echo Tesseract OCR not found on this PC - downloading it ^(one time, about 50 MB^) ...
set "TESS_URL=https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri $env:TESS_URL -OutFile 'tesseract_setup.exe'"
if not exist tesseract_setup.exe goto tess_winget
echo Installing Tesseract into the folder Tesseract-OCR here ^(click YES if Windows asks^) ...
start "" /wait tesseract_setup.exe /S /D=%CD%\Tesseract-OCR
del /q tesseract_setup.exe >nul 2>&1
if exist "Tesseract-OCR\tesseract.exe" goto tess_check
:tess_winget
echo Trying winget ...
winget install -e --id UB-Mannheim.TesseractOCR --accept-package-agreements --accept-source-agreements --silent >nul 2>&1
for %%D in ("%ProgramFiles%\Tesseract-OCR" "%LOCALAPPDATA%\Programs\Tesseract-OCR") do (
  if exist "%%~D\tesseract.exe" (
    robocopy "%%~D" "Tesseract-OCR" /E /NFL /NDL /NJH /NJS /NP >nul
    goto tess_check
  )
)
echo.
echo *** Could not get Tesseract automatically ^(no internet / blocked^).
echo *** Install it once from https://github.com/UB-Mannheim/tesseract/wiki and run this file again,
echo *** or press a key to build WITHOUT OCR ^(text PDFs still work^).
pause
goto tess_done
:tess_check
if not exist "Tesseract-OCR\tessdata\eng.traineddata" (
  echo *** Tesseract-OCR\tessdata\eng.traineddata is missing - OCR will not work. Reinstall Tesseract with English.
  pause
  goto tess_done
)
"Tesseract-OCR\tesseract.exe" --version >nul 2>&1
if errorlevel 1 ( echo *** Tesseract-OCR\tesseract.exe does not start on this PC - check the folder. & pause & goto tess_done )
echo Tesseract OCR OK - it will be packed INSIDE the exe.
set "HAVE_OCR=1"
:tess_done

rem ---------------------------------------------------------------- 2. Python packages
echo Installing required packages...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 ( echo Package install failed. & pause & exit /b 1 )

rem ---------------------------------------------------------------- 3. data packed into the exe
if exist UPDATE_HS_CODE.xlsx echo Converting UPDATE_HS_CODE.xlsx into built-in HS data ...
if exist UPDATE_HS_CODE.xlsx python make_hs_data.py
if errorlevel 1 ( echo HS list conversion failed. & pause & exit /b 1 )
set "EXTRA= --add-data web;web --add-data masters;masters --add-data dec-upload-template-with-description.xlsx;. --add-data frtimport_skeleton.xml;. --add-data hs_data.json;. --add-data logo.png;. --add-data PLUGIN_TEMPLATE.py;. --add-data HOW_TO_ADD_A_SUPPLIER.txt;. --add-data TRAINED_DOCUMENTS.txt;. --add-data plugins;plugins"
if defined HAVE_OCR ( set "EXTRA=%EXTRA% --add-data Tesseract-OCR;Tesseract-OCR" )
if exist logo.ico ( set "EXTRA=%EXTRA% --icon logo.ico" )
if defined HAVE_OCR (
  python -c "import ocr_runner as o, sys; ok=o.setup(); print('OCR self-test:', 'OK - '+o.STATE['version'] if ok else 'FAILED - '+o.STATE['error']); sys.exit(0 if ok else 1)"
  if errorlevel 1 ( echo *** OCR self-test failed - see the message above. & pause )
)

rem ---------------------------------------------------------------- 4. build (console window = the server)
echo Building PERMIT_SERVER.exe ...
python -m PyInstaller --noconfirm --clean --onefile --console --name PERMIT_SERVER %EXTRA% --hidden-import waitress --collect-all waitress --hidden-import flask --hidden-import xlrd --hidden-import pypdf --hidden-import pdfplumber --collect-all pdfplumber --collect-all pdfminer --hidden-import openpyxl --hidden-import numpy --collect-all pypdfium2 --collect-all pypdfium2_raw --hidden-import ocr_runner --hidden-import xml_export --hidden-import toolrun --hidden-import danfoss_engine --hidden-import danfoss_docs --hidden-import danfoss_festo_docs --hidden-import festo_engine --hidden-import festo_docs --hidden-import agilent_engine --hidden-import permit_engine --hidden-import readers --hidden-import export_docs --hidden-import declaration_rules --hidden-import msd_engine --hidden-import msd_server_engine --hidden-import jfood_server_engine --hidden-import sfa_rules --hidden-import doc_analysis --hidden-import liquor_rules --hidden-import permit_core --hidden-import embedded_data --hidden-import csv --hidden-import difflib --hidden-import decimal --hidden-import statistics --hidden-import unicodedata --hidden-import fnmatch --hidden-import email.parser server.py
if errorlevel 1 ( echo Build failed. & pause & exit /b 1 )
copy /y "dist\PERMIT_SERVER.exe" "PERMIT_SERVER.exe" >nul
echo.
echo ============================================================
echo DONE - PERMIT_SERVER.exe is in this folder.
if defined HAVE_OCR ( echo OCR is INSIDE the exe. ) else ( echo WARNING: built WITHOUT OCR - scanned PDFs need Tesseract on this PC. )
echo.
echo Next: 1. Right-click OPEN_FIREWALL_PORT.bat - Run as administrator ^(one time, lets other PCs connect^)
echo       2. Double-click PERMIT_SERVER.exe - the browser opens, create the admin account
echo       3. Optional: START_WITH_WINDOWS.bat so the server starts when this PC starts
echo ============================================================
pause
