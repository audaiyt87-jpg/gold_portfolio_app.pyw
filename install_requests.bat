@echo off
setlocal
pushd "%~dp0"
set "PY=python"
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PY=py -3"
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
    echo Installation failed. The app can still use its built-in HTTP fallback.
) else (
    echo Requests installed. Restart the app to use it.
)
pause
popd
