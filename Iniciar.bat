@echo off
:: Verificar si pywin32 está instalado usando pip
pip show pywin32 >nul 2>nul

:: Comprobar si el paquete no está instalado
if %ERRORLEVEL% neq 0 (
    echo pywin32 no esta instalado. Instalando...
    pip install pywin32
) else (
    echo pywin32 ya esta instalado.
)


python replace_str_in_word.py

pause
