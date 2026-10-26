@echo off
REM Crea dist\SerialNumberChecker\SerialNumberChecker.exe (necesita Python 3 para Windows).
REM Se genera en carpeta (no un solo archivo): Windows Defender lo bloquea mucho menos.
cd /d "%~dp0"

python -m pip install --upgrade pyinstaller openpyxl || goto :error

python -m PyInstaller --noconfirm --onedir --windowed --noupx ^
    --name SerialNumberChecker ^
    --icon icon.ico ^
    --add-data "icon.ico;." ^
    --version-file version_info.txt ^
    serial_number_checker.py || goto :error

echo.
echo Listo: dist\SerialNumberChecker\SerialNumberChecker.exe
pause
exit /b 0

:error
echo.
echo Hubo un error al crear el ejecutable.
pause
exit /b 1
