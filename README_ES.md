# Configuración e Instalación - PDF to Excel

Este documento explica cómo configurar e instalar todas las dependencias necesarias para ejecutar `pdf_to_excel.py` en Windows Server.

## ¿Qué hace este script?

`pdf_to_excel.py` convierte estados de cuenta bancarios en formato PDF a archivos Excel (.xlsx) con múltiples pestañas organizadas.

## Requisitos del Sistema

- Windows Server o Windows 10/11
- Python 3.7 o superior
- Conexión a Internet (para descargar dependencias)
- Permisos de administrador (opcional, solo para instalar Tesseract OCR)

---

## Requisitos Previos

### Python

- **Versión mínima:** Python 3.7
- **Descarga:** https://www.python.org/downloads/
- **IMPORTANTE:** Durante la instalación, marca la opción **"Add Python to PATH"**

### Permisos

- Permisos de lectura/escritura en la carpeta del proyecto
- Permisos de administrador (opcional, solo si quieres instalar Tesseract automáticamente con Chocolatey)

---

## Instalación Rápida (Recomendada)

### Paso 1: Clonar el repositorio

```bash
git clone <url-del-repositorio>
cd Valarix_Python_Bank_Statement_PDF_To_Excel
```

### Paso 2: Ejecutar el script de instalación

```bash
setup.bat
```

El script hará todo automáticamente:
1. ✅ Verificará que Python esté instalado
2. ✅ Actualizará pip
3. ✅ Instalará todas las dependencias Python
4. ✅ Intentará instalar Tesseract OCR (si Chocolatey está disponible)
5. ✅ Descargará automáticamente los paquetes de idioma inglés y español para Tesseract
6. ✅ Verificará que todo esté correcto

**Tiempo estimado:** 5-10 minutos

---

## Instalación Manual

Si el script automático falla o prefieres instalar manualmente, sigue estos pasos:

### Paso 1: Instalar Python

1. Descarga Python desde: https://www.python.org/downloads/
2. Ejecuta el instalador
3. **IMPORTANTE:** Marca "Add Python to PATH"
4. Completa la instalación
5. Verifica la instalación:
   ```bash
   python --version
   ```

### Paso 2: Instalar Dependencias Python

```bash
# Actualizar pip
python -m pip install --upgrade pip

# Instalar dependencias
python -m pip install -r requirements.txt
```

**Dependencias que se instalarán:**
- `pdfplumber` - Extracción de texto de PDFs
- `pandas` - Manipulación de datos
- `openpyxl` - Escritura de archivos Excel
- `pymupdf>=1.23.0` - Conversión de PDF a imágenes para OCR
- `pytesseract>=0.3.10` - Wrapper de Python para Tesseract OCR
- `Pillow>=10.0.0` - Procesamiento de imágenes

### Paso 3: Instalar Tesseract OCR

#### Opción A: Con Chocolatey (Recomendado)

Si tienes Chocolatey instalado:

```powershell
choco install tesseract
```

#### Opción B: Instalador de Windows

1. Ve a: https://github.com/UB-Mannheim/tesseract/wiki
2. Descarga el instalador más reciente:
   - **64-bit:** `tesseract-ocr-w64-setup-5.x.x.exe`
   - **32-bit:** `tesseract-ocr-w32-setup-5.x.x.exe`
3. Ejecuta el instalador
4. **IMPORTANTE:** Durante la instalación:
   - ✅ Marca **"Add Tesseract to PATH"**
   - ✅ Selecciona **"Spanish"** en idiomas adicionales
   - ✅ Selecciona **"English"** también
5. Completa la instalación
6. **CIERRA y REABRE** PowerShell/CMD para que se actualice el PATH
7. Verifica la instalación:
   ```bash
   tesseract --version
   ```

Deberías ver algo como:
```
tesseract 5.3.0
 leptonica-1.83.0
```

---

## Verificación de Instalación

### Verificación Automática

El script `setup.bat` incluye verificación automática. Si quieres verificar manualmente:

```bash
python -c "import pdfplumber, pandas, openpyxl, fitz, pytesseract, PIL; print('✅ Todas las dependencias instaladas')"
```

### Verificar Tesseract OCR

```bash
tesseract --version
```

### Verificar Idiomas de Tesseract

```python
import pytesseract
langs = pytesseract.get_languages()
print(f"Idiomas disponibles: {langs}")
print(f"Español disponible: {'spa' in langs}")
print(f"Inglés disponible: {'eng' in langs}")
```

---

## Uso del Script

Una vez completada la instalación, puedes usar el script:

```bash
python pdf_to_excel.py "ruta\al\archivo.pdf"
```

El script generará un archivo Excel con el mismo nombre que el PDF pero con extensión `.xlsx`.

### Ejemplo

```bash
python pdf_to_excel.py "Test\Bank Statement\BBVA.pdf"
```

Esto generará: `Test\Bank Statement\BBVA.xlsx`

### Salida del Script

El script mostrará:
- 🏦 Banco detectado: [nombre del banco]
- 📊 Exporting to Excel...
- ✅ VALIDACIÓN: TODO CORRECTO (o HAY DIFERENCIAS)
- ✅ Excel file created -> [ruta del archivo]

---

## Troubleshooting

### Error: "Python no encontrado"

**Problema:** Python no está instalado o no está en PATH.

**Solución:**
1. Instala Python desde https://www.python.org/downloads/
2. Asegúrate de marcar "Add Python to PATH" durante la instalación
3. Cierra y reabre PowerShell/CMD
4. Verifica con: `python --version`

### Error: "pip no encontrado"

**Problema:** pip no está instalado o no está en PATH.

**Solución:**
```bash
python -m ensurepip --upgrade
```

### Error: "Tesseract no encontrado"

**Problema:** Tesseract OCR no está instalado o no está en PATH.

**Solución:**
1. Instala Tesseract siguiendo las instrucciones en la sección "Instalación Manual"
2. Asegúrate de marcar "Add Tesseract to PATH" durante la instalación
3. **CIERRA y REABRE** PowerShell/CMD
4. Verifica con: `tesseract --version`

### Error: "Español (spa) NO disponible"

**Problema:** El idioma español no está instalado en Tesseract.

**Solución:**
1. Reinstala Tesseract
2. Durante la instalación, selecciona "Spanish" en idiomas adicionales
3. O ejecuta `setup.bat` nuevamente - descargará automáticamente los paquetes de idioma faltantes

### Error al instalar dependencias Python

**Problema:** Error de conexión o permisos.

**Solución:**
1. Verifica tu conexión a Internet
2. Intenta actualizar pip: `python -m pip install --upgrade pip`
3. Si usas proxy, configura pip:
   ```bash
   pip config set global.proxy http://proxy:port
   ```
4. Ejecuta PowerShell/CMD como administrador

### El script funciona pero sin OCR

**Problema:** Tesseract no está instalado o no está en PATH.

**Solución:**
- El script funcionará pero solo procesará PDFs con texto legible
- Para soporte OCR completo, instala Tesseract siguiendo las instrucciones

### Error: "requirements.txt no encontrado"

**Problema:** El script se ejecutó desde un directorio incorrecto.

**Solución:**
1. Asegúrate de ejecutar `setup.bat` desde la raíz del proyecto
2. Verifica que `requirements.txt` esté en el mismo directorio

---

## Preguntas Frecuentes

### ¿Necesito instalar Tesseract OCR?

**Respuesta:** No es estrictamente necesario, pero altamente recomendado. Sin Tesseract:
- ✅ El script funcionará para PDFs con texto legible
- ❌ No podrá procesar PDFs con texto escaneado o ilegible

### ¿Puedo usar el script sin conexión a Internet?

**Respuesta:** Una vez instaladas todas las dependencias, sí. Pero necesitas Internet para:
- Instalar dependencias Python (primera vez)
- Instalar Tesseract OCR (primera vez)

### ¿Funciona en Windows Server?

**Respuesta:** Sí, el script está diseñado para funcionar en Windows Server y Windows normal.

### ¿Necesito permisos de administrador?

**Respuesta:** 
- **Para dependencias Python:** No
- **Para Tesseract OCR con Chocolatey:** Sí
- **Para Tesseract OCR manual:** No (pero necesitas permisos de instalación)

### ¿Qué bancos soporta?

**Respuesta:** El script soporta múltiples bancos mexicanos:
- BBVA
- Banamex
- HSBC
- Santander
- Scotiabank
- Inbursa
- Konfio
- Clara
- Banregio
- Banorte
- Banbajío
- Base

### ¿Puedo ejecutar el script desde Windows Service?

**Respuesta:** Sí, el script está diseñado para ejecutarse desde línea de comandos:
```bash
python pdf_to_excel.py "ruta\al\archivo.pdf"
```

Esto lo hace compatible con Windows Services que ejecuten comandos.

**Códigos de salida:**
- `0` = Éxito (Excel creado correctamente)
- `1` = Error (ver mensajes en consola)

Ejemplo para Windows Service:
```batch
python pdf_to_excel.py "\\ruta\network\archivo.pdf"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Procesamiento falló
)
```

### ¿El Excel generado tiene alguna marca?

**Respuesta:** Sí, el Excel generado tiene "CONTAAYUDA" como autor en las propiedades del archivo. Puedes verlo en:
- Windows: Propiedades del archivo → Detalles → Autor
- Excel: Archivo → Información → Propiedades → Autor

---

## Enlaces Útiles

- **Python:** https://www.python.org/downloads/
- **Tesseract OCR Windows:** https://github.com/UB-Mannheim/tesseract/wiki
- **Chocolatey:** https://chocolatey.org/
- **Documentación openpyxl:** https://openpyxl.readthedocs.io/
- **Documentación pandas:** https://pandas.pydata.org/

---

## Soporte

Si encuentras problemas durante la instalación:

1. Revisa la sección de **Troubleshooting** arriba
2. Verifica que todos los requisitos previos estén cumplidos
3. Asegúrate de haber cerrado y reabierto PowerShell/CMD después de instalar Python o Tesseract
4. Ejecuta `setup.bat` nuevamente para verificar la instalación

---

## Próximos Pasos

Una vez completada la instalación:

1. Prueba el script con un PDF de prueba:
   ```bash
   python pdf_to_excel.py "Test\Bank Statement\BBVA.pdf"
   ```

2. Verifica que el Excel se haya generado correctamente

3. Revisa las pestañas del Excel:
   - Summary
   - Bank Statement Report
   - Data Validation

¡Listo para usar! 🎉

