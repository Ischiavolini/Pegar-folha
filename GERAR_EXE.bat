@echo off
title Gerador do BuscadorCNPJ.exe
color 0A

echo ================================================
echo   Gerador do BuscadorCNPJ.exe
echo ================================================
echo.
echo Este script vai instalar as dependencias necessarias
echo e gerar o arquivo BuscadorCNPJ.exe automaticamente.
echo Nao e necessario ser administrador da maquina.
echo.
echo Certifique-se de que o arquivo buscador_cnpj.py esta
echo na mesma pasta que este script antes de continuar.
echo.
pause

echo.
echo [1/4] Verificando instalacao do Python...
python --version
if errorlevel 1 (
    echo.
    echo ERRO: Python nao encontrado!
    echo Instale o Python em https://www.python.org/downloads/
    echo Lembre de marcar "Add python.exe to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

echo.
echo [2/4] Instalando biblioteca openpyxl...
pip install --user openpyxl
if errorlevel 1 (
    echo.
    echo ERRO ao instalar openpyxl.
    echo.
    pause
    exit /b 1
)

echo.
echo [3/4] Instalando PyInstaller...
pip install --user pyinstaller
if errorlevel 1 (
    echo.
    echo ERRO ao instalar PyInstaller.
    echo.
    pause
    exit /b 1
)

echo.
echo [4/4] Gerando BuscadorCNPJ.exe... (pode demorar 1-2 minutos)
python -m PyInstaller --onefile --windowed --name BuscadorCNPJ buscador_cnpj.py
if errorlevel 1 (
    echo.
    echo ERRO ao gerar o .exe. Verifique se o arquivo buscador_cnpj.py
    echo esta na mesma pasta que este script.
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================
echo   PRONTO! .exe gerado com sucesso!
echo ================================================
echo.
echo O arquivo BuscadorCNPJ.exe esta na pasta "dist"
echo que foi criada aqui nesta mesma pasta.
echo.
echo Voce pode copiar o BuscadorCNPJ.exe para qualquer
echo computador Windows e rodar com duplo clique,
echo sem precisar instalar Python nele.
echo.
pause
