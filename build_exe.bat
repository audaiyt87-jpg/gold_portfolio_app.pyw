@echo off
chcp 65001 >nul
setlocal
pushd "%~dp0"
set "PY=python"
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PY=py -3"
%PY% -m pip install --upgrade pyinstaller
if errorlevel 1 goto failed
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto failed
%PY% -m PyInstaller --noconfirm --collect-data certifi --hidden-import requests --hidden-import live_chart --onefile --windowed --name GoldPortfolio-PriceFix-10-CleanStatus --icon app.ico --add-data "app.ico;." --add-data "brand_badge.png;." gold_portfolio_app.pyw
if errorlevel 1 goto failed
echo Done: dist\GoldPortfolio-PriceFix-10-CleanStatus.exe
pause
popd
exit /b 0
:failed
echo Build failed. Read the error above; no successful build is claimed.
pause
popd
exit /b 1
