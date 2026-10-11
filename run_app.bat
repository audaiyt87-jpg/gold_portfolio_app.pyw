@echo off
REM تشغيل التطبيق مباشرة (يتطلب تثبيت Python 3)
chcp 65001 >nul
python "%~dp0gold_portfolio_app.pyw"
if errorlevel 1 (
  echo.
  echo تعذر التشغيل. تأكد من تثبيت Python 3 من python.org
  pause
)
