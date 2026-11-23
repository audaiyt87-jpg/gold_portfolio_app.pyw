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
%PY% -m PyInstaller --noconfirm --collect-data certifi --hidden-import requests --hidden-import live_chart --onefile --windowed --name GoldPortfolio-PriceFix-19-Contrast --icon app.ico --add-data "app.ico;." --add-data "brand_badge.png;." --add-data "nav_dark_main.png;." --add-data "nav_dark_chart.png;." --add-data "nav_dark_history.png;." --add-data "nav_dark_alerts.png;." --add-data "nav_dark_zakat.png;." --add-data "nav_dark_settings.png;." --add-data "nav_dark_account.png;." --add-data "nav_light_main.png;." --add-data "nav_light_chart.png;." --add-data "nav_light_history.png;." --add-data "nav_light_alerts.png;." --add-data "nav_light_zakat.png;." --add-data "nav_light_settings.png;." --add-data "nav_light_account.png;." gold_portfolio_app.pyw
if errorlevel 1 goto failed
echo Done: dist\GoldPortfolio-PriceFix-19-Contrast.exe
pause
popd
exit /b 0
:failed
echo Build failed. Read the error above; no successful build is claimed.
pause
popd
exit /b 1
