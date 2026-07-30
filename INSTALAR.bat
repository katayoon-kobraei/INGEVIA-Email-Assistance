@echo off
setlocal
cd /d "%~dp0"

echo ================================================
echo   Asistente de Email IA - Instalacion
echo ================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo No se ha encontrado Python en este ordenador.
    echo.
    echo Instala Python desde https://www.python.org/downloads/
    echo IMPORTANTE: durante la instalacion, marca la casilla
    echo "Add Python to PATH" antes de darle a Instalar.
    echo.
    echo Cuando lo tengas instalado, vuelve a hacer doble clic en este archivo.
    echo.
    pause
    exit /b 1
)

if not exist ".env" (
    if exist ".env.example" (
        copy ".env.example" ".env" >nul
    )
    echo Se ha creado el archivo de configuracion .env
    echo Se va a abrir en el Bloc de notas: rellena tu clave de Gemini
    echo ^(y los demas datos si hace falta^) y GUARDA el archivo antes de cerrarlo.
    echo.
    notepad ".env"
    echo.
    echo Cuando hayas guardado el archivo .env, pulsa una tecla para continuar.
    pause >nul
)

echo.
echo [1/4] Instalando el programa ^(puede tardar uno o dos minutos^)...
powershell -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
if errorlevel 1 (
    echo.
    echo Algo ha fallado durante la instalacion. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo [2/4] Comprobando la conexion con Outlook y con la IA...
echo ^(Asegurate de que Outlook esta abierto y con la sesion iniciada^)
".venv\Scripts\python.exe" "tests\test_classify_latest_inbox.py"
if errorlevel 1 (
    echo.
    echo La comprobacion no ha funcionado. Las causas mas habituales son
    echo una clave de Gemini incorrecta en .env, o que Outlook este cerrado.
    echo Puedes seguir con la instalacion y solucionarlo despues -- la
    echo revision automatica volvera a intentarlo sola cada 2 minutos.
    echo.
    pause
)

echo.
echo [3/4] Activando la revision automatica de correo ^(cada 2 minutos^)...
powershell -NoProfile -ExecutionPolicy Bypass -File ".\setup_scheduler.ps1"

echo.
echo [4/4] Activando la revision diaria de calidad...
powershell -NoProfile -ExecutionPolicy Bypass -File ".\setup_eval_scheduler.ps1"

echo.
echo ================================================
echo   Instalacion completada.
echo.
echo   A partir de ahora el correo se clasifica solo
echo   cada 2 minutos, incluso si cierras esta ventana.
echo.
echo   Usa "ABRIR PANEL.bat" cuando quieras ver el estado.
echo ================================================
echo.
pause
