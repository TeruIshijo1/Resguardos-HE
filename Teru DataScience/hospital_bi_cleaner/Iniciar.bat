@echo off
title Hospital BI Cleaner & Recommender

echo ===================================================
echo   Levantando Hospital BI Cleaner...
echo ===================================================

:: Verificar si existe el entorno virtual
if not exist "venv\Scripts\activate.bat" (
    echo [ERROR] No se encontro el entorno virtual "venv".
    echo Por favor, asegurate de haber ejecutado: python -m venv venv
    echo y luego haber instalado los requirements.txt
    pause
    exit /b
)

:: Activar entorno virtual
call venv\Scripts\activate.bat

:: Abrir el navegador automaticamente
echo Abriendo el navegador en http://127.0.0.1:5000...
start http://127.0.0.1:5000

:: Iniciar la aplicacion
python app.py

pause
