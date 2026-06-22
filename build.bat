@echo off
echo.
echo  ====================================
echo   FoliaExchange -- Build
echo  ====================================
echo.
echo [1/2] Instalando dependencias...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo ERROR: fallo al instalar dependencias.
    pause
    exit /b 1
)
echo.
echo [2/2] Construyendo FoliaExchange.exe...
pyinstaller --onefile --windowed --name FoliaExchange ^
  --collect-all fitz ^
  --collect-all docx ^
  --collect-all mammoth ^
  --collect-all xhtml2pdf ^
  --hidden-import PIL ^
  main.py
if %errorlevel% neq 0 (
    echo ERROR: fallo en PyInstaller.
    pause
    exit /b 1
)
echo.
echo  ====================================
echo   Listo! dist\FoliaExchange.exe
echo  ====================================
echo.
pause
