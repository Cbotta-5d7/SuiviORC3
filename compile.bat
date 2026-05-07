@echo off
chcp 65001 > nul
title Compilation Suivi Production ORC

echo.
echo ==========================================================================
echo   COMPILATION DE L'APPLICATION SUIVI PRODUCTION ORC
echo ==========================================================================
echo.

REM Verifier que Python est installe
where python >nul 2>nul
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo.
    echo Telechargez et installez Python depuis :
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANT : Pendant l'installation, COCHEZ la case
    echo "Add Python to PATH" en bas de la fenetre d'installation.
    echo.
    pause
    exit /b 1
)

echo [1/4] Python detecte :
python --version
echo.

echo [2/4] Installation des dependances (peut prendre 1-2 minutes)...
echo.
python -m pip install --upgrade pip
python -m pip install customtkinter openpyxl Pillow pyinstaller
if errorlevel 1 (
    echo.
    echo [ERREUR] L'installation des dependances a echoue.
    echo Verifiez votre connexion internet et reessayez.
    pause
    exit /b 1
)
echo.

echo [3/4] Compilation de l'executable (peut prendre 2-3 minutes)...
echo.
python -m PyInstaller --onefile --windowed --name "SuiviProductionORC" --noconfirm app.py
if errorlevel 1 (
    echo.
    echo [ERREUR] La compilation a echoue.
    pause
    exit /b 1
)
echo.

echo [4/4] Nettoyage et organisation des fichiers...
echo.

REM Deplacer l'exe a la racine
if exist "dist\SuiviProductionORC.exe" (
    move /Y "dist\SuiviProductionORC.exe" "SuiviProductionORC.exe" > nul
    echo - SuiviProductionORC.exe genere avec succes
)

REM Nettoyer les fichiers temporaires
if exist "build" rmdir /S /Q "build"
if exist "dist" rmdir /S /Q "dist"
if exist "SuiviProductionORC.spec" del /F /Q "SuiviProductionORC.spec"

echo.
echo ==========================================================================
echo   COMPILATION TERMINEE
echo ==========================================================================
echo.
echo L'application a ete generee : SuiviProductionORC.exe
echo.
echo Vous pouvez maintenant double-cliquer dessus pour la lancer.
echo Au premier lancement, l'application vous demandera de selectionner ou
echo creer un fichier Excel pour stocker les donnees.
echo.
pause
