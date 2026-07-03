@echo off
chcp 65001 > nul
echo ====================================
echo   KPI-ORC v6.0 - Compilation EXE
echo ====================================
echo.
echo [1/2] Installation des dependances...
pip install pyinstaller openpyxl flask pywebview --quiet
if %ERRORLEVEL% neq 0 (
    echo ERREUR: pip install a echoue
    pause & exit /b 1
)
echo.
echo [2/2] Compilation PyInstaller...
pyinstaller --onefile --windowed --name "KPI-ORC" ^
  --hidden-import flask ^
  --hidden-import webview ^
  --collect-all webview ^
  main.py
if %ERRORLEVEL% neq 0 (
    echo ERREUR de compilation!
    pause & exit /b 1
)
echo.
echo ====================================
echo  SUCCESS! Fichier: dist\KPI-ORC.exe
echo ====================================
pause
