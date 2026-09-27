@echo off
:: Starts the State Gas launcher now and every time you log in to Windows.
:: Safe to run again at any time.
set "HERE=%~dp0"
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
(
  echo @echo off
  echo cd /d "%HERE%"
  echo start "" pyw "%HERE%launcher.py"
) > "%STARTUP%\stategas_launcher.bat"
cd /d "%HERE%"
start "" pyw "%HERE%launcher.py" --pair
echo.
echo  Launcher installed and started. It will also start whenever you log in.
echo  Your browser is opening the control app to pair with it.
echo.
pause
