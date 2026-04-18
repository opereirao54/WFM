@echo off
echo.
echo  ============================================
echo   WFM ENGINE - Iniciando servidor...
echo  ============================================
echo.
pip install flask numpy scipy -q
echo.
echo  Acesse no navegador: http://localhost:5000
echo.
python app.py
pause
