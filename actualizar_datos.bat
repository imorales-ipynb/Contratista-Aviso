@echo off
:: ─────────────────────────────────────────────────────────────
:: actualizar_datos.bat
:: Programar con el Programador de Tareas de Windows.
:: Ejemplo: cada día a las 07:00 y 13:00.
:: ─────────────────────────────────────────────────────────────

cd /d "C:\Users\ivan.morales\Desktop\Contratista Aviso"

:: Crear carpeta de logs si no existe
if not exist "logs" mkdir logs

echo [%date% %time%] Iniciando actualización... >> logs\update.log

python update_data.py >> logs\update.log 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] ERROR: el script terminó con error %ERRORLEVEL% >> logs\update.log
) else (
    echo [%date% %time%] Actualización exitosa. >> logs\update.log
)
