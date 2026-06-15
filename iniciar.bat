@echo off
title Resguardos - Hospital Escandon v3
color 0A
cd /d "%~dp0"
echo.
echo  =====================================================
echo   RESGUARDOS HOSPITAL ESCANDON v3 - Area de Sistemas
echo  =====================================================
echo.
netsh advfirewall firewall show rule name="ResguardosHospital" >nul 2>&1
if %errorlevel% neq 0 (
    echo  Configurando firewall...
    netsh advfirewall firewall add rule name="ResguardosHospital" dir=in action=allow protocol=TCP localport=5000 >nul 2>&1
)
echo  Instalando dependencias...
pip install -r requirements.txt -q --disable-pip-version-check
echo.
echo  -------------------------------------------------------
echo   LOCAL:  http://localhost:5000
echo   RED:    http://192.168.254.150:5000
echo   (Comparte la URL de RED con tus companeros)
echo  -------------------------------------------------------
echo.
python app.py
pause
