@echo off
SETLOCAL ENABLEEXTENSIONS

:: ====== Configuration ======
:: Direct download link from OneDrive
SET "DOWNLOAD_URL=https://onedrive.live.com/download?resid=XXXXXXX&id=XXXXXXX"

:: ISO filename
SET "ISO_FILENAME=file.iso"

:: Required free space in MB (7 GB = 7168 MB)
SET "REQUIRED_SPACE_MB=7168"

:: ====== Paths ======
SET "DEST_FOLDER=%CD%"
SET "DEST_FILE=%DEST_FOLDER%\%ISO_FILENAME%"

:: ====== Check if file already exists ======
IF EXIST "%DEST_FILE%" (
    echo File "%ISO_FILENAME%" already exists in this directory. Skipping download.
    GOTO END
)

:: ====== Check free space ======
FOR /F "tokens=3" %%A IN ('fsutil volume diskfree %DEST_FOLDER:~0,2%') DO (
    SET "FREE_BYTES=%%A"
    GOTO CHECK_SPACE
)

:CHECK_SPACE
:: Convert bytes to MB (1 MB = 1048576 bytes)
SET /A FREE_MB=%FREE_BYTES:~0,-6%

IF %FREE_MB% LSS %REQUIRED_SPACE_MB% (
    echo Not enough free space on drive %DEST_FOLDER:~0,2%.
    echo Required: %REQUIRED_SPACE_MB% MB, Available: %FREE_MB% MB
    GOTO END
)

:: ====== Download file using PowerShell ======
echo Downloading ISO from OneDrive to "%DEST_FILE%"...
powershell -Command "Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%DEST_FILE%'"

:: ====== Check if download succeeded ======
IF EXIST "%DEST_FILE%" (
    echo Download completed successfully.
) ELSE (
    echo Download failed. Please check the URL or internet connection.
)

:END
pause
ENDLOCAL
