@echo off
cd /d "C:\Users\fabia\OneDrive\Escritorio\Data Science\Python\00 Proyectos\mails-automaticos"

"venv\Scripts\python.exe" envio_reporte_ventas.py >> logs\reporte.log 2>&1

echo Ejecutado: %date% %time% >> logs\reporte.log