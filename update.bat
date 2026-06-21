@echo off
echo.
echo  ====================================
echo   FoliaExchange -- Actualizar
echo  ====================================
echo.
echo [1/2] Descargando ultima version...
git pull origin main
if %errorlevel% neq 0 (
    echo ERROR: fallo en git pull.
    pause
    exit /b 1
)
echo.
echo [2/2] Reconstruyendo ejecutable...
call build.bat
