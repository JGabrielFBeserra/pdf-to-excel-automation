@echo off
echo =========================================
echo    Extrator PDF para Excel
echo =========================================
echo.
echo Iniciando o programa...
echo.

REM Vai para a pasta onde este .bat está localizado
cd /d "%~dp0"

REM Executa o script Python
python projeto.py

echo.
echo =========================================
echo Programa finalizado!
echo =========================================
pause
