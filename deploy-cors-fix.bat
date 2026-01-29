@echo off
echo ===================================
echo Desplegando fix CRITICO de CORS...
echo ===================================
cd /d "%~dp0"
echo.
echo [1/2] Haciendo commit...
git add api/proxy.js
git commit -m "Fix CRITICAL: Disable logging to fix @libsql/client error"
echo.
echo [2/2] Pushing to GitHub...
git push origin main
echo.
echo ===================================
echo Deployment completado!
echo Espera 1 minuto y prueba el LOGIN.
echo ===================================
pause
