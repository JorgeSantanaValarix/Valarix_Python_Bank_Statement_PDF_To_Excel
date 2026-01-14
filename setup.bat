@echo off
setlocal enabledelayedexpansion

:: Configuración
set PYTHON_MIN_VERSION=3.7
set TESSERACT_URL=https://github.com/UB-Mannheim/tesseract/wiki
set TESSERACT_DEFAULT_PATH=C:\Program Files\Tesseract-OCR\tesseract.exe
set TESSERACT_X86_PATH=C:\Program Files (x86)\Tesseract-OCR\tesseract.exe

:: Variables de estado
set INSTALL_SUCCESS=0
set PYTHON_OK=0
set DEPENDENCIES_OK=0
set TESSERACT_OK=0
set VERIFY_RESULT=1

:: Banner
echo.
echo ========================================
echo   CONFIGURACION PDF TO EXCEL
echo   Instalacion de dependencias
echo ========================================
echo.
echo Este script instalara:
echo   - Dependencias Python (pip)
echo   - Tesseract OCR (si no esta instalado)
echo.
echo Tiempo estimado: 5-10 minutos
echo.
pause

:: PASO 1: Verificar Python
echo.
echo [PASO 1/5] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado
    echo.
    echo Por favor instala Python desde:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANTE: Durante la instalacion, marca:
    echo   - "Add Python to PATH"
    echo.
    set INSTALL_SUCCESS=1
    goto :end
)

:: Verificar versión mínima
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo [OK] Python encontrado: %PYTHON_VERSION%
set PYTHON_OK=1

:: PASO 2: Actualizar pip
echo.
echo [PASO 2/5] Actualizando pip...
python -m pip install --upgrade pip --quiet
if errorlevel 1 (
    echo [ADVERTENCIA] Error actualizando pip, continuando...
) else (
    echo [OK] pip actualizado
)

:: PASO 3: Instalar dependencias Python
echo.
echo [PASO 3/5] Instalando dependencias Python...
if not exist requirements.txt (
    echo [ERROR] requirements.txt no encontrado
    echo Asegurate de ejecutar este script desde la raiz del proyecto
    set INSTALL_SUCCESS=1
    goto :end
)

python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Error instalando dependencias Python
    set INSTALL_SUCCESS=1
    goto :end
)

echo [OK] Dependencias Python instaladas
set DEPENDENCIES_OK=1

:: PASO 4: Verificar/Instalar Tesseract OCR
echo.
echo [PASO 4/5] Verificando Tesseract OCR...

:: Verificar si Tesseract ya está instalado en PATH
where tesseract >nul 2>&1
if not errorlevel 1 (
    echo [OK] Tesseract ya esta instalado
    tesseract --version
    set TESSERACT_OK=1
    goto :verify_tesseract
)

:: Verificar rutas comunes
if exist "%TESSERACT_DEFAULT_PATH%" (
    echo [OK] Tesseract encontrado en: %TESSERACT_DEFAULT_PATH%
    set TESSERACT_OK=1
    goto :verify_tesseract
)

if exist "%TESSERACT_X86_PATH%" (
    echo [OK] Tesseract encontrado en: %TESSERACT_X86_PATH%
    set TESSERACT_OK=1
    goto :verify_tesseract
)

:: Intentar instalar con Chocolatey
echo [INFO] Tesseract no encontrado, intentando instalar...
where choco >nul 2>&1
if not errorlevel 1 (
    echo [INFO] Chocolatey encontrado, instalando Tesseract...
    choco install tesseract -y
    if not errorlevel 1 (
        echo [OK] Tesseract instalado con Chocolatey
        :: Actualizar PATH en esta sesión
        call refreshenv >nul 2>&1
        set TESSERACT_OK=1
        :: Descargar idiomas después de instalar
        call :download_tesseract_languages
        goto :verify_tesseract
    )
)

:: Instalación manual necesaria
echo.
echo ========================================
echo   INSTALACION MANUAL DE TESSERACT OCR
echo ========================================
echo.
echo Tesseract OCR no esta instalado.
echo.
echo Por favor sigue estos pasos:
echo.
echo 1. Abre tu navegador y ve a:
echo    %TESSERACT_URL%
echo.
echo 2. Descarga el instalador para Windows:
echo    - Para 64-bit: tesseract-ocr-w64-setup-5.x.x.exe
echo    - Para 32-bit: tesseract-ocr-w32-setup-5.x.x.exe
echo.
echo 3. Ejecuta el instalador y durante la instalacion:
echo    - Marca "Add Tesseract to PATH"
echo    - Selecciona "Spanish" en idiomas adicionales
echo    - Selecciona "English" tambien
echo.
echo 4. IMPORTANTE: Despues de instalar, CIERRA esta ventana
echo    y abre una nueva para que se actualice el PATH
echo.
echo 5. Ejecuta este script nuevamente para verificar
echo.
pause
set INSTALL_SUCCESS=1
goto :end

:verify_tesseract
:: Verificar instalación de Tesseract
where tesseract >nul 2>&1
if errorlevel 1 (
    echo [ADVERTENCIA] Tesseract no encontrado en PATH
    echo Verifica que lo hayas instalado correctamente
    echo y que hayas cerrado y reabierto esta ventana
    set TESSERACT_OK=0
) else (
    echo [OK] Tesseract encontrado en PATH
    set TESSERACT_OK=1
)

:: Descargar idiomas si Tesseract está instalado
if %TESSERACT_OK%==1 (
    call :download_tesseract_languages
)

:: PASO 5: Verificación final
echo.
echo [PASO 5/5] Verificando instalacion...
echo.

:: Crear script de verificación temporal
(
echo import sys
echo import importlib
echo.
echo def check_package^(package_name, import_name=None^):
echo     if import_name is None:
echo         import_name = package_name
echo     try:
echo         importlib.import_module^(import_name^)
echo         return True
echo     except ImportError:
echo         return False
echo.
echo def main^(^):
echo     print^("=" * 60^)
echo     print^("VERIFICACION DE INSTALACION"^)
echo     print^("=" * 60^)
echo.
echo     packages = [
echo         ^("pdfplumber", "pdfplumber"^),
echo         ^("pandas", "pandas"^),
echo         ^("openpyxl", "openpyxl"^),
echo         ^("pymupdf", "fitz"^),
echo         ^("pytesseract", "pytesseract"^),
echo         ^("Pillow", "PIL"^),
echo     ]
echo.
echo     all_ok = True
echo     for package, import_name in packages:
echo         if check_package^(package, import_name^):
echo             print^(f"OK {package}"^)
echo         else:
echo             print^(f"ERROR {package} - NO INSTALADO"^)
echo             all_ok = False
echo.
echo     print^("\n" + "=" * 60^)
echo     print^("VERIFICACION TESSERACT OCR"^)
echo     print^("=" * 60^)
echo.
echo     try:
echo         import pytesseract
echo         version = pytesseract.get_tesseract_version^(^)
echo         print^(f"OK Tesseract encontrado: version {version}"^)
echo.
echo         try:
echo             langs = pytesseract.get_languages^(^)
echo             print^(f"OK Idiomas disponibles: {len^(langs^)}"^)
echo             if 'spa' in langs:
echo                 print^("OK Espanol ^(spa^) disponible"^)
echo             else:
echo                 print^("ADVERTENCIA Espanol ^(spa^) NO disponible"^)
echo             if 'eng' in langs:
echo                 print^("OK Ingles ^(eng^) disponible"^)
echo             else:
echo                 print^("ADVERTENCIA Ingles ^(eng^) NO disponible"^)
echo         except Exception as e:
echo             print^(f"ADVERTENCIA Error verificando idiomas: {e}"^)
echo     except Exception as e:
echo         print^(f"ERROR Tesseract NO encontrado: {e}"^)
echo         all_ok = False
echo.
echo     print^("\n" + "=" * 60^)
echo     if all_ok:
echo         print^("OK INSTALACION COMPLETA"^)
echo         return 0
echo     else:
echo         print^("ERROR INSTALACION INCOMPLETA"^)
echo         return 1
echo.
echo if __name__ == "__main__":
echo     sys.exit^(main^(^)^)
) > verify_installation.py

:: Ejecutar verificación
python verify_installation.py
set VERIFY_RESULT=%ERRORLEVEL%

:: Limpiar archivo temporal
del verify_installation.py >nul 2>&1

goto :end

:: Función para descargar archivos de idioma de Tesseract
:download_tesseract_languages
echo.
echo [INFO] Verificando idiomas de Tesseract...

:: Determinar ruta de tessdata
set TESSDATA_DIR=
if exist "%TESSERACT_DEFAULT_PATH%" (
    set TESSDATA_DIR=C:\Program Files\Tesseract-OCR\tessdata
)
if exist "%TESSERACT_X86_PATH%" (
    set TESSDATA_DIR=C:\Program Files (x86)\Tesseract-OCR\tessdata
)

:: Si no se encontró, intentar desde PATH
if "%TESSDATA_DIR%"=="" (
    for /f "tokens=*" %%i in ('where tesseract 2^>nul') do (
        set TESSERACT_EXE=%%i
        for %%j in ("%%~dpi..") do set TESSDATA_DIR=%%~fj\tessdata
    )
)

if "%TESSDATA_DIR%"=="" (
    echo [ADVERTENCIA] No se pudo determinar la ruta de tessdata
    goto :end_language_check
)

:: Verificar si los idiomas ya están instalados
set NEED_ENG=1
set NEED_SPA=1

if exist "%TESSDATA_DIR%\eng.traineddata" (
    echo [OK] Ingles ^(eng^) ya esta instalado
    set NEED_ENG=0
)

if exist "%TESSDATA_DIR%\spa.traineddata" (
    echo [OK] Espanol ^(spa^) ya esta instalado
    set NEED_SPA=0
)

if %NEED_ENG%==0 if %NEED_SPA%==0 (
    echo [OK] Todos los idiomas necesarios estan instalados
    goto :end_language_check
)

:: Verificar si PowerShell está disponible para descargar
where powershell >nul 2>&1
if errorlevel 1 (
    echo [ADVERTENCIA] PowerShell no encontrado, no se pueden descargar idiomas automaticamente
    goto :manual_language_instructions
)

:: Crear carpeta tessdata si no existe
if not exist "%TESSDATA_DIR%" (
    echo [INFO] Creando carpeta tessdata...
    mkdir "%TESSDATA_DIR%" >nul 2>&1
    if errorlevel 1 (
        echo [ADVERTENCIA] No se pudo crear la carpeta tessdata, se requieren permisos de administrador
        goto :manual_language_instructions
    )
)

:: Descargar inglés si es necesario
if %NEED_ENG%==1 (
    echo [INFO] Descargando ingles ^(eng^)...
    powershell -Command "try { Invoke-WebRequest -Uri 'https://github.com/tesseract-ocr/tessdata/raw/main/eng.traineddata' -OutFile '%TESSDATA_DIR%\eng.traineddata' -ErrorAction Stop; Write-Host 'OK' } catch { Write-Host 'ERROR' }" >nul 2>&1
    if exist "%TESSDATA_DIR%\eng.traineddata" (
        echo [OK] Ingles ^(eng^) descargado e instalado
    ) else (
        echo [ERROR] Error descargando ingles ^(eng^)
        goto :manual_language_instructions
    )
)

:: Descargar español si es necesario
if %NEED_SPA%==1 (
    echo [INFO] Descargando espanol ^(spa^)...
    powershell -Command "try { Invoke-WebRequest -Uri 'https://github.com/tesseract-ocr/tessdata/raw/main/spa.traineddata' -OutFile '%TESSDATA_DIR%\spa.traineddata' -ErrorAction Stop; Write-Host 'OK' } catch { Write-Host 'ERROR' }" >nul 2>&1
    if exist "%TESSDATA_DIR%\spa.traineddata" (
        echo [OK] Espanol ^(spa^) descargado e instalado
    ) else (
        echo [ERROR] Error descargando espanol ^(spa^)
        goto :manual_language_instructions
    )
)

goto :end_language_check

:manual_language_instructions
echo.
echo [ADVERTENCIA] No se pudieron descargar los idiomas automaticamente
echo.
echo Para instalar los idiomas manualmente:
echo 1. Ve a: https://github.com/tesseract-ocr/tessdata
echo 2. Descarga los archivos:
echo    - eng.traineddata ^(para ingles^)
echo    - spa.traineddata ^(para espanol^)
echo 3. Copia los archivos a: %TESSDATA_DIR%
echo.

:end_language_check
goto :eof

:end
echo.
echo ========================================
echo   RESUMEN DE INSTALACION
echo ========================================
echo.
if %PYTHON_OK%==1 (
    echo [OK] Python instalado
) else (
    echo [ERROR] Python no encontrado
)
if %DEPENDENCIES_OK%==1 (
    echo [OK] Dependencias Python instaladas
) else (
    echo [ERROR] Dependencias Python no instaladas
)
if %TESSERACT_OK%==1 (
    echo [OK] Tesseract OCR instalado
) else (
    echo [ADVERTENCIA] Tesseract OCR no instalado o no en PATH
    echo    El script funcionara pero sin soporte OCR
)
echo.
if %INSTALL_SUCCESS%==0 (
    if %VERIFY_RESULT%==0 (
        echo ========================================
        echo   INSTALACION COMPLETADA EXITOSAMENTE
        echo ========================================
        echo.
        echo Puedes usar pdf_to_excel.py ahora:
        echo   python pdf_to_excel.py "ruta\al\archivo.pdf"
        echo.
    ) else (
        echo ========================================
        echo   INSTALACION COMPLETADA CON ADVERTENCIAS
        echo ========================================
        echo.
        echo Revisa los mensajes anteriores para detalles
        echo.
    )
) else (
    echo ========================================
    echo   INSTALACION INCOMPLETA
    echo ========================================
    echo.
    echo Revisa los errores anteriores y vuelve a intentar
    echo Consulta CONFIGURACION.md para mas ayuda
    echo.
)
pause

