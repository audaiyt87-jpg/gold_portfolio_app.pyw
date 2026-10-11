@echo off
REM بناء ملف تنفيذي .exe مستقل لويندوز
chcp 65001 >nul
echo جارٍ تثبيت PyInstaller ...
python -m pip install --upgrade pyinstaller
echo جارٍ بناء التطبيق ...
python -m PyInstaller --onefile --noconsole --name GoldPortfolio "%~dp0gold_portfolio_app.pyw"
echo.
echo تم البناء. ستجد الملف GoldPortfolio.exe داخل مجلد dist
pause
