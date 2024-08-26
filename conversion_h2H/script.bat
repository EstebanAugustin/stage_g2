@echo off
chcp 65001
setlocal enabledelayedexpansion

REM Configuration des chemins pour les programmes tiers
set "EXIFTOOL_PATH=C:\micmac\binaire-aux\windows\exiftool.exe"
set "CIRCE_PATH=C:\Program Files (x86)\IGN\Circe 5-4-4\circeFR.exe"
set "CIRCE_METADATA_PATH=C:\ProgramData\IGN\Circe\5-4-4\Service-public\Metropole\"

REM Demander le répertoire de travail
set /p "repertoire=Veuillez glisser le dossier contenant les photos : "

REM Vérifier l'existence du répertoire
if not exist "%repertoire%" (
    echo Le répertoire spécifié n'existe pas. Veuillez vérifier le chemin et réessayer.
    pause
    exit /b
)

REM Définir les chemins des fichiers de coordonnées
set "fichier_cor_origine=%repertoire%\coordonnees_lon_lat_h.txt"
set "fichier_cor_converties=%repertoire%\coordonnees_lon_lat_alt.txt"

REM Extraire les coordonnées GPS à partir des images avec ExifTool
echo Extraction des coordonnées GPS...
"%EXIFTOOL_PATH%" -GPSLongitude# -GPSLatitude# -GPSAltitude# -csv "%repertoire%\*.JPG" > "%fichier_cor_origine%"

REM Vérifier si le fichier d'origine a été créé
if not exist "%fichier_cor_origine%" (
    echo L'extraction des coordonnées a échoué. Aucune donnée n'a été trouvée.
    pause
    exit /b
)

REM Supprimer les en-têtes du fichier d'origine
echo Suppression des en-têtes...
more +1 "%fichier_cor_origine%" > "%fichier_cor_origine%.tmp"
move /y "%fichier_cor_origine%.tmp" "%fichier_cor_origine%"

echo Extraction terminée. Les coordonnées ont été enregistrées dans "%fichier_cor_origine%".

REM Convertir les coordonnées avec Circe
echo Conversion des coordonnées avec Circe...
"%CIRCE_PATH%" --metadataFile="%CIRCE_METADATA_PATH%DataFRnew.txt" --boundaryFile="%CIRCE_METADATA_PATH%PB2002_plates.txt" --sourceCRS=RGF93v2bG. --sourceFormat=ILPH.METERS.DEGREES --targetCRS=RGF93v2bG.IGN69 --targetFormat=ILPV.METERS.DMS --displayPrecision=0.001 --separator=Virgule --gridLoading=BINARY --sourcePathname="%fichier_cor_origine%" --targetPathname="%fichier_cor_converties%"

REM Vérifier si la conversion a réussi
if not exist "%fichier_cor_converties%" (
    echo La conversion des coordonnées a échoué. Veuillez vérifier Circe.
    pause
    exit /b
)

REM Réinjecter les altitudes dans les images originales
echo Réinjection des altitudes dans les fichiers JPEG...
for /f "skip=17 tokens=1,2,3,4* delims= " %%a in ('type "%fichier_cor_converties%"') do (
    set "image_file=%%a"
    if exist "!image_file!" (
        echo Mise à jour de "!image_file!" avec l'altitude %%d...
        "%EXIFTOOL_PATH%" -GPSAltitude=%%d -overwrite_original "!image_file!"
    ) else (
        echo Le fichier "!image_file!" n'existe pas. Vérifiez le nom ou le chemin.
    )
)

echo Processus terminé. Les altitudes ont été réinjectées dans les images.
endlocal
pause
