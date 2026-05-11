@echo off
setlocal

set "TARGET_DIR=C:\GitHub\TateByWater-DocumentManagement\msaccess\TBCMS"
set "PACKAGE_URL=http://localhost:5173/api/Extractor/DownloadPackage"
set "ZIP_PATH=%TEMP%\msaccess-extractor.zip"

echo ===============================================
echo PCAPPS MS Access Extractor - Local Setup
echo ===============================================
echo Target folder: %TARGET_DIR%
if not "%PACKAGE_URL%"=="" echo Package URL: %PACKAGE_URL%
echo.

if not exist "%TARGET_DIR%" (
  echo Creating target folder...
  mkdir "%TARGET_DIR%"
)

cd /d "%TARGET_DIR%"

if exist "extract_msaccess.py" (
  set "EXTRACTOR_DIR=%TARGET_DIR%"
) else if exist "code\extract_msaccess.py" (
  set "EXTRACTOR_DIR=%TARGET_DIR%\code"
) else (
  echo extract_msaccess.py not found in target folder.
  if "%PACKAGE_URL%"=="" (
    echo ERROR: No package URL was provided.
    echo Download and extract the package into %TARGET_DIR% first.
    echo Expected:
    echo   %TARGET_DIR%\extract_msaccess.py
    echo   or
    echo   %TARGET_DIR%\code\extract_msaccess.py
    pause
    exit /b 1
  )

  echo Downloading extractor package...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri '%PACKAGE_URL%' -OutFile '%ZIP_PATH%'"
  if errorlevel 1 (
    echo ERROR: Failed to download package from %PACKAGE_URL%.
    pause
    exit /b 1
  )

  echo Extracting package...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Path '%ZIP_PATH%' -DestinationPath '%TARGET_DIR%' -Force"
  if errorlevel 1 (
    echo ERROR: Failed to extract package into %TARGET_DIR%.
    pause
    exit /b 1
  )

  if exist "extract_msaccess.py" (
    set "EXTRACTOR_DIR=%TARGET_DIR%"
  ) else if exist "code\extract_msaccess.py" (
    set "EXTRACTOR_DIR=%TARGET_DIR%\code"
  ) else (
    echo ERROR: Package extracted but extract_msaccess.py was still not found.
    echo Check package contents and folder permissions.
    pause
    exit /b 1
  )
)

echo Using extractor directory: %EXTRACTOR_DIR%
cd /d "%EXTRACTOR_DIR%"

set "PY_CMD="
where py >nul 2>&1
if not errorlevel 1 (
  set "PY_CMD=py -3"
)
if "%PY_CMD%"=="" (
  where python >nul 2>&1
  if not errorlevel 1 (
    set "PY_CMD=python"
  )
)
if "%PY_CMD%"=="" (
  echo ERROR: Python launcher was not found (py or python).
  echo Install Python 3.10+ and ensure it is on PATH.
  pause
  exit /b 1
)

if exist ".venv\Scripts\python.exe" (
  echo Reusing existing virtual environment.
) else (
  echo Creating virtual environment...
  %PY_CMD% -m venv .venv
  if errorlevel 1 (
    echo ERROR: Failed to create virtual environment.
    pause
    exit /b 1
  )
)

set "VENV_PY=.venv\Scripts\python.exe"
set "VENV_PIP=.venv\Scripts\pip.exe"
if not exist "%VENV_PY%" (
  echo ERROR: Virtual environment python executable not found.
  pause
  exit /b 1
)

if exist "requirements.txt" (
  echo Installing dependencies...
  "%VENV_PIP%" install --upgrade pip
  "%VENV_PIP%" install -r requirements.txt
  if errorlevel 1 (
    echo ERROR: Dependency installation failed.
    pause
    exit /b 1
  )
) else (
  echo WARNING: requirements.txt was not found. Skipping pip install.
)

echo Running smoke test...
"%VENV_PY%" extract_msaccess.py --help >nul 2>&1
if errorlevel 1 (
  echo WARNING: Smoke test failed. Check Python environment and Access setup.
  pause
  exit /b 1
)

echo.
echo Setup completed successfully.
echo Return to the web app and click "Run Preflight" again.
pause
exit /b 0