@echo off
title Reiniciar Monitor SRI (Web)
echo ==================================================
echo   REINICIANDO EL MONITOR SRI (WEB) Y LIMPIANDO PROCESOS
echo ==================================================
echo.
echo 1. Buscando y terminando procesos de python colgados...
powershell -Command "Get-CimInstance Win32_Process -Filter \"name = 'python.exe'\" | ForEach-Object { if ($_.CommandLine -like '*check_sri_change.py*' -and $_.CommandLine -like '*MonitorSri-Web*') { Stop-Process -Id $_.ProcessId -Force; Write-Output \"Proceso colgado $_.ProcessId terminado.\" } }"

echo.
echo 2. Limpiando archivos de bloqueo obsoletos...
if exist "%~dp0sri_monitor.lock" (
    del "%~dp0sri_monitor.lock"
    echo Archivo de bloqueo sri_monitor.lock eliminado.
) else (
    echo No se encontraron archivos de bloqueo.
)

echo.
echo 3. Asegurando que la tarea programada esté habilitada...
powershell -Command "Enable-ScheduledTask -TaskName 'MonitorSRIVehiculosWeb'" 2>nul
if %errorlevel% neq 0 (
    echo [INFO] La tarea programada local 'MonitorSRIVehiculosWeb' no está creada en este equipo.
) else (
    echo.
    echo 4. Iniciando una nueva ejecución de comprobación del SRI...
    powershell -Command "Start-ScheduledTask -TaskName 'MonitorSRIVehiculosWeb'"
)

echo.
echo ==================================================
echo   PROCESOS REINICIADOS. EL MONITOR ESTÁ ACTIVO.
echo ==================================================
echo.
pause
