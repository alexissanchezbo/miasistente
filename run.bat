@echo off
cd /d "%~dp0"
echo Iniciando aplicacion PyG...
echo.
pip install -r requirements.txt -q
echo.
echo Abriendo navegador en http://localhost:8501
echo Presiona Ctrl+C para cerrar la aplicacion.
echo.
python -m streamlit run app.py --server.headless false --browser.gatherUsageStats false
pause
