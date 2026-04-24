@echo off
chcp 65001 >nul 2>&1
echo.
echo  ============================================
echo   WFM ENGINE - Iniciando servidor...
echo  ============================================
echo.

REM Tenta encontrar o Python
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON_CMD=python
    goto :found
)

where python3 >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON_CMD=python3
    goto :found
)

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON_CMD=py
    goto :found
)

REM Tenta caminhos comuns do Anaconda/Miniconda
if exist "%USERPROFILE%\anaconda3\python.exe" (
    set PYTHON_CMD=%USERPROFILE%\anaconda3\python.exe
    goto :found
)
if exist "%USERPROFILE%\miniconda3\python.exe" (
    set PYTHON_CMD=%USERPROFILE%\miniconda3\python.exe
    goto :found
)
if exist "C:\ProgramData\anaconda3\python.exe" (
    set PYTHON_CMD=C:\ProgramData\anaconda3\python.exe
    goto :found
)
if exist "C:\Python312\python.exe" (
    set PYTHON_CMD=C:\Python312\python.exe
    goto :found
)
if exist "C:\Python311\python.exe" (
    set PYTHON_CMD=C:\Python311\python.exe
    goto :found
)
if exist "C:\Python310\python.exe" (
    set PYTHON_CMD=C:\Python310\python.exe
    goto :found
)

echo.
echo  [ERRO] Python nao foi encontrado!
echo.
echo  Instale o Python 3.10+ em: https://www.python.org/downloads/
echo  IMPORTANTE: Marque "Add Python to PATH" durante a instalacao.
echo.
pause
exit /b 1

:found
echo  Python encontrado: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

REM Instalar dependências
echo  Instalando dependencias...
%PYTHON_CMD% -m pip install flask numpy scipy openpyxl -q
echo.

REM Iniciar o servidor
echo  Iniciando servidor em http://localhost:5000
echo  Pressione Ctrl+C para parar.
echo.

cd /d "%~dp0"
set PYTHONPATH=%~dp0src
%PYTHON_CMD% src\app.py --port 5000

echo.
echo  Servidor encerrado.
pause
