@echo off
chcp 65001 > nul
echo ====================================
echo   KPI-ORC - Compilation EXE
echo ====================================
echo.
echo [1/3] Installation des dependances...
pip install pyinstaller openpyxl Pillow --quiet
if %ERRORLEVEL% neq 0 (
    echo ERREUR: pip install a echoue
    pause & exit /b 1
)
echo.
echo [2/3] Conversion logo DODO en icone...
set ICON_ARG=
if exist dodo_logo.png (
    python -c "from PIL import Image; img=Image.open('dodo_logo.png').convert('RGBA'); img.save('dodo_logo.ico', format='ICO', sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)])" 2>nul
    if exist dodo_logo.ico (
        set ICON_ARG=--icon=dodo_logo.ico
        echo    Logo DODO converti avec succes.
    ) else (
        echo    Conversion echouee, icone par defaut utilisee.
    )
) else (
    echo    dodo_logo.png non trouve, icone par defaut utilisee.
    echo    Placez dodo_logo.png dans ce dossier pour utiliser le logo DODO.
)
echo.
echo [3/3] Compilation PyInstaller...
pyinstaller --onefile --windowed --name "KPI-ORC" %ICON_ARG% main.py
if %ERRORLEVEL% neq 0 (
    echo ERREUR de compilation!
    pause & exit /b 1
)
echo.
echo ====================================
echo  SUCCESS! Fichier: dist\KPI-ORC.exe
echo ====================================
pause
