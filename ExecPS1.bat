:: Fix pour BOM UTF-8

@echo off
setlocal EnableDelayedExpansion

:: Efface tout affichage précédent
cls

:: Se placer dans le dossier du .bat
cd /d "%~dp0"

echo ============================================
echo      LANCEUR DE SCRIPTS POWERSHELL
echo ============================================
echo.
echo Dossier courant :
echo   %cd%
echo.

set "idx=0"

:: Recherche des scripts PS1 sans afficher la boucle
for %%F in ("*.ps1") do (
    set /a idx+=1
    set "SCRIPT_!idx!=%%~fF"
)

:: Aucun script trouvé
if %idx%==0 (
    echo Aucun fichier .ps1 detecte dans ce dossier.
    echo Place ce fichier BAT dans un dossier contenant des scripts.
    echo.
    pause
    exit /b 1
)

echo Scripts disponibles :
echo ---------------------
for /L %%I in (1,1,%idx%) do (
    call echo   %%I^) !SCRIPT_%%I!
)
echo.

:ASK_CHOICE
set "CHOIX="
set /p "CHOIX=Entrez le numero du script a executer (1-%idx%) : "

:: Vérification choix vide
if "%CHOIX%"=="" (
    echo Veuillez entrer un numero.
    goto ASK_CHOICE
)

:: Vérification numérique
echo %CHOIX%| findstr /R "^[0-9][0-9]*$" >nul || (
    echo Entrez uniquement des chiffres.
    goto ASK_CHOICE
)

:: Vérification plage
if %CHOIX% LSS 1 (
    echo Numero hors plage.
    goto ASK_CHOICE
)
if %CHOIX% GTR %idx% (
    echo Numero hors plage.
    goto ASK_CHOICE
)

:: Récupération du script choisi
call set "SCRIPT_PATH=!SCRIPT_%CHOIX%!"

cls
echo ============================================
echo        EXECUTION DU SCRIPT SELECTIONNE
echo ============================================
echo.
echo Script :
echo   %SCRIPT_PATH%
echo --------------------------------------------
echo.

:: Exécution PowerShell propre
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"

echo.
echo --------------------------------------------
echo Execution terminee.
echo ============================================
echo.
pause

endlocal
exit /b
