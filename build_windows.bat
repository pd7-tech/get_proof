@echo off
REM Script para criar executável do Extrator de Comprovantes para Windows

echo ====================================
echo Criando executavel para Windows...
echo ====================================
echo.

REM Verificar se ambiente virtual existe
if not exist "venv" (
    echo Criando ambiente virtual...
    python -m venv venv
)

REM Ativar ambiente virtual
echo Ativando ambiente virtual...
call venv\Scripts\activate.bat

REM Instalar dependências
echo Instalando dependencias...
pip install --upgrade pip
pip install pandas openpyxl xlrd PyPDF2 pdfplumber Pillow customtkinter darkdetect pyinstaller

REM Criar executável
echo Compilando executavel...
pyinstaller build_windows.spec

REM Desativar ambiente virtual
deactivate

echo.
echo ====================================
echo Executavel criado com sucesso!
echo ====================================
echo.
echo Localizacao: dist\PD7Lab_ExtractorPDF.exe
echo.
echo Para distribuir:
echo   - Copie o arquivo dist\PD7Lab_ExtractorPDF.exe para outras maquinas Windows
echo   - Duplo clique no .exe para executar
echo.
pause