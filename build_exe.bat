@echo off
chcp 65001 >nul
echo Building GoldPortfolio.exe ...
python -m pip install --upgrade pyinstaller openpyxl
pyinstaller --noconfirm --onefile --windowed --name GoldPortfolio --icon app.ico --add-data "app.ico;." gold_portfolio_app.pyw
echo.
echo Done. EXE is in the dist\ folder.
pause
