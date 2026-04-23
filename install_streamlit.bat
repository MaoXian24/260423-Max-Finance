@echo off
setlocal

pushd "%~dp0"

set "PY_CMD="
set "PY_ARGS="

set "CONDA_PY=%USERPROFILE%\anaconda3\envs\wrds312\python.exe"
if exist "%CONDA_PY%" (
  set "PY_CMD=%CONDA_PY%"
)

if "%PY_CMD%"=="" (
  where py >nul 2>&1
  if not errorlevel 1 (
    set "PY_CMD=py"
    set "PY_ARGS=-3"
  )
)

if "%PY_CMD%"=="" (
  where python >nul 2>&1
  if not errorlevel 1 set "PY_CMD=python"
)

if "%PY_CMD%"=="" (
  echo [ERROR] Python not found in PATH.
  echo Please install Python 3.10+ and re-run.
  pause
  exit /b 1
)

echo Using Python command: %PY_CMD% %PY_ARGS%
%PY_CMD% %PY_ARGS% -m pip --version >nul 2>&1
if errorlevel 1 (
  echo pip not found, trying ensurepip...
  %PY_CMD% %PY_ARGS% -m ensurepip --upgrade >nul 2>&1
)

%PY_CMD% %PY_ARGS% -m pip install --upgrade pip

if exist "requirements.txt" (
  %PY_CMD% %PY_ARGS% -m pip install -r requirements.txt
) else (
  %PY_CMD% %PY_ARGS% -m pip install --upgrade streamlit wrds pandas matplotlib openpyxl numpy
)

echo.
echo Dependency installation completed.
%PY_CMD% %PY_ARGS% -m pip show streamlit

echo.
echo To run web app:
echo %PY_CMD% %PY_ARGS% -m streamlit run streamlit_app.py

pause
popd
endlocal
