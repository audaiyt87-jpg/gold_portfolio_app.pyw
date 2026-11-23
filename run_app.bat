@echo off
chcp 65001 >nul
setlocal
pushd "%~dp0"
set "PY="
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PY=py -3"
if not defined PY (
    python -c "import sys" >nul 2>&1
    if not errorlevel 1 set "PY=python"
)
if not defined PY (
    echo Python 3 was not found. Install Python 3 with Tcl/Tk, or use a built EXE.
    pause
    popd
    exit /b 1
)
echo Starting GoldPortfolio PriceFix-5-LiveSources ...
%PY% -u gold_portfolio_app.pyw > startup.log 2>&1
set "RESULT=%errorlevel%"
if not "%RESULT%"=="0" (
    echo The app could not start. Details:
    type startup.log
    echo Please share the error message, not your account data.
    pause
)
popd
exit /b %RESULT%
