import sys
import os
import re
import pdfplumber
import pandas as pd

# NUEVOS IMPORTS para Tesseract OCR (mínimos):
try:
    import pytesseract
    import fitz  # PyMuPDF - para convertir PDF a imágenes
    from PIL import Image
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False
    print("[ADVERTENCIA] Tesseract OCR no disponible. Instala: pip install pytesseract pymupdf pillow")


# Bank configurations with column coordinate ranges (X-axis)
# Use find_coordinates.py to get the exact ranges for your PDF
# python find_coordinates.py <pdf_path> <page_number>

def configure_tesseract():
    """
    Configura la ruta a Tesseract OCR si no está en PATH.
    Detecta automáticamente la ubicación común en Windows.
    """
    if not TESSERACT_AVAILABLE:
        return False
    
    default_paths = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    ]
    
    for path in default_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return True
    
    # Si no se encuentra, intentar usar el que está en PATH
    try:
        pytesseract.get_tesseract_version()
        return True
    except:
        return False

BANK_CONFIGS = {
    "BBVA": {
        "name": "BBVA",
        "columns": {
            "fecha": (9, 38),              # Columna Fecha de Operación
            "liq": (52, 81),                # Columna LIQ (Liquidación)
            "descripcion": (86, 322),     # Columna Descripción
            "cargos": (360, 398),          # Columna Cargos
            "abonos": (422, 458),          # Columna Abonos
            "saldo": (539, 593),           # Columna Saldo
        }
    },
    
    "Santander": {
        "name": "Santander",
        "columns": {
            "fecha": (18, 52),             # Columna Fecha de Operación
            "descripcion": (107, 149),     # Columna Descripción
            "cargos": (465, 492),          # Columna Cargos
            "abonos": (377, 411),          # Columna Abonos
            "saldo": (554, 573),           # Columna Saldo
        }
    },

    "Scotiabank": {
        "name": "Scotiabank",
        "columns": {
            "fecha": (56, 81),             # Columna Fecha de Operación
            "descripcion": (92, 240),     # Columna Descripción
            "cargos": (465, 509),          # Columna Cargos
            "abonos": (385, 437),          # Columna Abonos
            "saldo": (532, 584),           # Columna Saldo
        }
    },

    "Inbursa": {
        "name": "Inbursa",
        "columns": {
            "fecha": (11, 40),             # Columna Fecha de Operación
            "descripcion": (145, 369),     # Columna Descripción
            "cargos": (400, 441),          # Columna Cargos
            "abonos": (475, 510),          # Columna Abonos
            "saldo": (525, 563),           # Columna Saldo
        }
    },

    "Konfio": {
        "name": "Konfio",
        "columns": {
            "fecha": (48, 102),             # Columna Fecha de Operación
            "descripcion": (120, 240),     # Columna Descripción
            "cargos": (340, 435),          # Columna Cargos
            "abonos": (510, 565),          # Columna Abonos
        }
    },

    "Clara": {
        "name": "Clara",
        "columns": {
            "fecha": (35, 60),             # Columna Fecha de Operación
            "descripcion": (60, 450),       # Columna Descripción (ampliado para capturar todas las palabras de descripción)
            "cargos": (450, 480),          # Columna Cargos
            "abonos": (520, 576),          # Columna Abonos
        }
    },

    "Banregio": {
        "name": "Banregio",
        "columns": {
            "fecha": (35, 42),             # Columna Fecha de Operación
            "descripcion": (53, 310),     # Columna Descripción
            "cargos": (380, 418),          # Columna Cargos
            "abonos": (460, 498),          # Columna Abonos
            "saldo": (530, 573),           # Columna Saldo
        }
    },
    
     "Banorte": {
        "name": "Banorte",
        "columns": {
            "fecha": (54, 85),             # Columna Fecha de Operación
            "descripcion": (87, 167),     # Columna Descripción
            "cargos": (450, 489),          # Columna Cargos
            "abonos": (380, 420),          # Columna Abonos
            "saldo": (533, 560),           # Columna Saldo
        }
    },

     "Banbajío": {
        "name": "Banbajío",
        "columns": {
            "fecha": (21, 41),             # Columna Fecha de Operación
            "descripcion": (87, 362),     # Columna Descripción
            "cargos": (490, 525),          # Columna Cargos
            "abonos": (415, 451),          # Columna Abonos
            "saldo": (550, 585),           # Columna Saldo
        }
    },
    
    "Banamex": {
        "name": "Banamex",
        "movements_start": "DETALLE DE OPERACIONES",  # String que marca el inicio de la sección de movimientos
        "movements_end": "SALDO MINIMO REQUERIDO",    # String que marca el fin de la sección de movimientos
        "columns": {
            "fecha": (17, 45),             # Columna Fecha de Operación
            "descripcion": (55, 260),      # Columna Descripción (ampliado para capturar mejor)
            "cargos": (275, 316),          # Columna Cargos (ampliado ligeramente)
            "abonos": (345, 395),          # Columna Abonos (ampliado ligeramente)
            "saldo": (425, 472),           # Columna Saldo (ampliado ligeramente)
        }
    },

    "HSBC": {
        "name": "HSBC",
        "movements_start": "Día",  # String que marca el inicio de la sección de movimientos (encabezado de la tabla)
        "movements_end": "CoDi",                      # String que marca el fin de la sección de movimientos
        "columns": {
            "fecha": (87, 103),             # Columna Fecha de Operación
            "descripcion": (124, 505),      # Columna Descripción (ampliado para capturar mejor)
            "cargos": (710, 800),          # Columna Cargos (ampliado ligeramente)
            "abonos": (865, 950),          # Columna Abonos (ampliado ligeramente)
            "saldo": (1050, 1130),           # Columna Saldo (ampliado ligeramente)
        }
    },
    "Base": {
        "name": "Base",
        "columns": {
            "fecha": (41, 78),             # Columna Fecha de Operación
            "descripcion": (108, 342),      # Columna Descripción (ampliado para capturar mejor)
            "cargos": (375, 415),          # Columna Cargos (ampliado ligeramente)
            "abonos": (440, 485),          # Columna Abonos (ampliado ligeramente)
            "saldo": (520, 560),           # Columna Saldo (ampliado ligeramente)
        }
    },

    # Add more banks here as needed
}

DEFAULT_BANK = "BBVA"

# Bank detection keywords (case insensitive)
# Order matters: more specific patterns should come first
BANK_KEYWORDS = {
    
    "BBVA": [
        r"\bBBVA\b",
        r"\bBBVA\s+BANCOMER\b",
        r"ESTADO\s+DE\s+CUENTA\s+MAESTRA\s+PYME\s+BBVA",
        r"BBVA\s+ADELANTE",
    ],
    "Banamex": [
        r"\bDIGITEM\b",  # Palabra muy específica de Banamex
        r"\bBANAMEX\b",
        r"\bCITIBANAMEX\b",
        r"\bCITI\s+BANAMEX\b",
    ],
    "Banbajío": [
        r"\bBANBAJ[IÍ]O\b",
        r"\bBANCO\s+DEL\s+BAJ[IÍ]O\b",
    ],
    "Banorte": [
        r"\bBANORTE\b",
        r"\bBANCO\s+MACRO\s+BANORTE\b",
    ],
    "Banregio": [
        r"\bBANREGIO\b",
        r"\bBANCO\s+REGIONAL\b",
    ],
    "Clara": [
        r"\bCLARA\b",
        r"\bBANCO\s+CLARA\b",
    ],
    "Konfio": [
        r"\bKONFIO\b",
        r"\bBANCO\s+KONFIO\b",
    ],
    "Inbursa": [
        r"\bINBURSA\b",
        r"\bBANCO\s+INBURSA\b",
    ],
    "Santander": [
        r"\bSANTANDER\b",
        r"\bBANCO\s+SANTANDER\b",
        r"\bSANTANDER\s+MEXICANO\b",
    ],
    "Scotiabank": [
        r"\bSCOTIABANK\b",
        r"\bSCOTIA\s+BANK\b",
        r"\bBANCO\s+SCOTIABANK\b",
    ],
    "Hey": [
        r"\bHEY\s+BANCO\b",
        r"\bHEY\b",
    ],
    "HSBC": [
        r"\bHSBC\s+BANCO\b",
        r"\bHSBC\s+M[EE]XICO\b",
    ],
    "Base": [
        r"\bBASE\b",
        r"\bBANCO\s+BASE\b",
    ],

}

# Decimal / thousands amount regex (module-level so helpers can use it)
DEC_AMOUNT_RE = re.compile(r"\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")

# Amount normalization function
def normalize_amount_str(amount_str):
    """Normalize amount string by removing commas, spaces, and converting to float."""
    if not amount_str or pd.isna(amount_str):
        return 0.0
    # Remove commas, spaces, and extract numeric value
    cleaned = str(amount_str).replace(',', '').replace(' ', '').replace('$', '')
    try:
        return float(cleaned)
    except (ValueError, AttributeError):
        return 0.0

# Amount normalization function
def normalize_amount_str(amount_str):
    """Normalize amount string by removing commas, spaces, and converting to float."""
    if not amount_str or pd.isna(amount_str):
        return 0.0
    # Remove commas, spaces, and extract numeric value
    cleaned = str(amount_str).replace(',', '').replace(' ', '').replace('$', '')
    try:
        return float(cleaned)
    except (ValueError, AttributeError):
        return 0.0


def fix_duplicated_chars(text_str):
    """Fix duplicated characters in text (e.g., 'PPaaggoo' -> 'Pagos').
    This happens with some PDF encodings where each character is duplicated.
    IMPORTANT: Do NOT apply to amounts/money values unless they have the specific
    duplication pattern with two dots (###..###).
    """
    if not text_str:
        return text_str
    
    # Check for the specific duplication pattern: ###..### (number with duplications + two dots)
    # This pattern indicates duplicated characters in amounts (e.g., "9977,,000000..0000")
    duplication_pattern = re.compile(r'\d+[,\d]*\.\.\d+')
    has_duplication_pattern = duplication_pattern.search(text_str)
    
    # Check if this looks like a money/amount value
    # Pattern: contains $, numbers, commas, dots (e.g., "$50,000.00", "50,000.00", "50.00")
    amount_pattern = re.compile(r'[\$]?\s*\d+[,\d]*\.?\d*')
    is_amount = amount_pattern.fullmatch(text_str.strip())
    
    # If it's an amount but doesn't have the duplication pattern, don't fix it
    if is_amount and not has_duplication_pattern:
        # This is a normal amount (e.g., "$50,000.00"), don't fix duplicated chars
        return text_str
    
    # Apply normal duplication fix for text (or amounts with duplication pattern)
    result = []
    i = 0
    while i < len(text_str):
        # Check if current char and next char are the same (case-insensitive)
        if i + 1 < len(text_str) and text_str[i].lower() == text_str[i + 1].lower():
            # Add only one character
            result.append(text_str[i])
            i += 2  # Skip both characters
        else:
            result.append(text_str[i])
            i += 1
    return ''.join(result)


def find_column_coordinates(pdf_path: str, page_number: int = 1):
    """Extract all words from a page and show their coordinates.
    Helps user find exact X ranges for columns.
    Detecta automáticamente si debe usar OCR para PDFs ilegibles.
    """
    try:
        # PASO 1: Detectar si el PDF es ilegible
        is_illegible, cid_ratio, ascii_ratio = is_pdf_text_illegible(pdf_path)
        
        if is_illegible and TESSERACT_AVAILABLE:
            print(f"[INFO] PDF detectado como ilegible (CID ratio: {cid_ratio:.2%}, ASCII ratio: {ascii_ratio:.2%})")
            print(f"[INFO] Usando OCR para analizar coordenadas...")
            
            # Extraer con OCR (ahora usa image_to_data() con coordenadas reales)
            pages_data = extract_text_with_tesseract_ocr(pdf_path)
            
            # Detectar banco desde texto OCR
            all_text = '\n'.join([p.get('content', '') for p in pages_data[:3]])
            detected_bank = detect_bank_from_text(all_text)
            
            print(f"🏦 Banco detectado: {detected_bank}")
            print(f"📄 Mostrando coordenadas extraídas con OCR (coordenadas reales)...")
            
            # Filtrar página si es necesario
            if page_number:
                pages_data = [p for p in pages_data if p['page'] == page_number]
            
            # Mostrar palabras con coordenadas reales del OCR
            print(f"\n📄 Página {page_number} del PDF (OCR): {pdf_path}")
            print("=" * 120)
            print(f"{'Y (top)':<8} {'X0':<8} {'X1':<8} {'X_center':<10} {'Conf':<6} {'Texto':<40}")
            print("-" * 120)
            
            # Agrupar palabras por Y (filas)
            rows = {}
            for page_data in pages_data:
                words = page_data.get('words', [])
                for word in words:
                    top = int(round(word.get('top', 0)))
                    if top not in rows:
                        rows[top] = []
                    rows[top].append(word)
            
            # Print words sorted by Y then X
            for top in sorted(rows.keys()):
                row_words = sorted(rows[top], key=lambda w: w.get('x0', 0))
                for word in row_words:
                    x0 = word.get('x0', 0)
                    x1 = word.get('x1', 0)
                    x_center = (x0 + x1) / 2
                    conf = word.get('conf', 0)
                    text = word.get('text', '')
                    print(f"{top:<8} {x0:<8.1f} {x1:<8.1f} {x_center:<10.1f} {conf:<6.1f} {text:<40}")
            
            print("\n" + "=" * 120)
            print("\nRangos aproximados de columnas (X0 a X1) - Coordenadas reales del OCR:")
            print("-" * 120)
            
            # Find column boundaries
            all_x0 = [w.get('x0', 0) for word_list in rows.values() for w in word_list]
            all_x1 = [w.get('x1', 0) for word_list in rows.values() for w in word_list]
            
            if all_x0 and all_x1:
                min_x = min(all_x0)
                max_x = max(all_x1)
                print(f"Rango X total: {min_x:.1f} a {max_x:.1f}")
                print("\nAnaliza la salida anterior y proporciona estos valores en BANK_CONFIGS:")
            else:
                print("⚠️  No se encontraron palabras con coordenadas en esta página")
            
            return
        
        # PASO 2: PDF legible - comportamiento normal (pdfplumber)
        # Detect bank to apply Konfio-specific fixes
        detected_bank = detect_bank_from_pdf(pdf_path)
        is_konfio = (detected_bank == "Konfio")
        
        with pdfplumber.open(pdf_path) as pdf:
            if page_number > len(pdf.pages):
                # print(f"❌ El PDF solo tiene {len(pdf.pages)} páginas")
                return
            
            page = pdf.pages[page_number - 1]
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            
            if not words:
                # print("❌ No se encontraron palabras en la página")
                return
            
            # For Konfio, fix duplicated characters in word texts
            if is_konfio:
                for word in words:
                    if 'text' in word and word['text']:
                        word['text'] = fix_duplicated_chars(word['text'])
            
            # Group by approximate Y coordinate (rows)
            rows = {}
            for word in words:
                top = int(round(word['top']))
                if top not in rows:
                    rows[top] = []
                rows[top].append(word)
            
            print(f"\n📄 Página {page_number} del PDF: {pdf_path}")
            print("=" * 120)
            print(f"{'Y (top)':<8} {'X0':<8} {'X1':<8} {'X_center':<10} {'Texto':<40}")
            print("-" * 120)
            
            # Print words sorted by Y then X
            for top in sorted(rows.keys()):
                row_words = sorted(rows[top], key=lambda w: w['x0'])
                for i, word in enumerate(row_words):
                    x_center = (word['x0'] + word['x1']) / 2
                    print(f"{top:<8} {word['x0']:<8.1f} {word['x1']:<8.1f} {x_center:<10.1f} {word['text']:<40}")
            
            print("\n" + "=" * 120)
            print("\nRangos aproximados de columnas (X0 a X1):")
            print("-" * 120)
            
            # Find column boundaries
            all_x0 = [w['x0'] for word_list in rows.values() for w in word_list]
            all_x1 = [w['x1'] for word_list in rows.values() for w in word_list]
            
            min_x = min(all_x0)
            max_x = max(all_x1)
            
            print(f"Rango X total: {min_x:.1f} a {max_x:.1f}")
            print("\nAnaliza la salida anterior y proporciona estos valores en BANK_CONFIGS:")
            # print("""
            # Ejemplo:
            # BANK_CONFIGS = {
            #     "BBVA": {
            #         "name": "BBVA",
            #         "columns": {
            #             "fecha": (x_min, x_max),           # Columna Fecha de Operación
            #             "liq": (x_min, x_max),              # Columna LIQ (Liquidación)
            #             "descripcion": (x_min, x_max),     # Columna Descripción
            #             "cargos": (x_min, x_max),          # Columna Cargos
            #             "abonos": (x_min, x_max),          # Columna Abonos
            #             "saldo": (x_min, x_max),           # Columna Saldo
            #         }
            #     },
            # }
            # """)
    
    except Exception as e:
        pass
        print(f"❌ Error: {e}")


def detect_bank_from_text(text: str) -> str:
    """
    Detect the bank from extracted text content.
    Returns the bank name if detected, otherwise returns DEFAULT_BANK.
    """
    if not text:
        return DEFAULT_BANK
    
    # Split into lines and check each line
    lines = text.split('\n')
    for line in lines:
        line_clean = line.strip()
        if not line_clean:
            continue
        
        # Check each bank's keywords
        for bank_name, keywords in BANK_KEYWORDS.items():
            for keyword_pattern in keywords:
                if re.search(keyword_pattern, line_clean, re.I):
                    print(f"🏦 Banco detectado: {bank_name}")
                    return bank_name
        
        # Also check if line contains bank name directly (case insensitive)
        line_upper = line_clean.upper()
        for bank_name in BANK_KEYWORDS.keys():
            # Check for exact bank name match (as whole word)
            if re.search(rf'\b{re.escape(bank_name.upper())}\b', line_upper):
                print(f"🏦 Banco detectado: {bank_name}")
                return bank_name
    
    # If no bank detected, return default
    return DEFAULT_BANK


def detect_bank_from_pdf(pdf_path: str) -> str:
    """
    Detect the bank from PDF content by reading line by line.
    Returns the bank name if detected, otherwise returns DEFAULT_BANK.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Read first few pages (usually bank name appears early)
            max_pages_to_check = min(3, len(pdf.pages))
            
            all_text = ""
            for page_num in range(max_pages_to_check):
                page = pdf.pages[page_num]
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
            
            if all_text:
                return detect_bank_from_text(all_text)
    
    except Exception as e:
        pass
        # print(f"⚠️  Error al detectar banco: {e}")
    
    # If no bank detected, return default
    #print(f"⚠️  No se pudo detectar el banco, usando: {DEFAULT_BANK}")
    return DEFAULT_BANK


def is_pdf_text_illegible(pdf_path: str, cid_threshold: float = 0.05) -> tuple:
    """
    Detecta si un PDF tiene texto ilegible (caracteres CID).
    
    Args:
        pdf_path: Ruta al archivo PDF
        cid_threshold: Ratio mínimo de caracteres CID para considerar ilegible (default: 5%)
    
    Returns:
        Tuple: (is_illegible: bool, cid_ratio: float, ascii_ratio: float)
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                return False, 0.0, 0.0
            
            # Analizar primera página (suficiente para detectar problema)
            first_page = pdf.pages[0]
            text = first_page.extract_text() or ""
            
            if not text or len(text) < 50:
                # Si no hay texto o es muy corto, puede ser ilegible
                return True, 1.0, 0.0
            
            # Contar caracteres CID
            cid_count = text.count('(cid:')
            total_chars = len(text)
            cid_ratio = cid_count / total_chars if total_chars > 0 else 0.0
            
            # Contar caracteres ASCII
            ascii_count = sum(1 for c in text if ord(c) < 128)
            ascii_ratio = ascii_count / total_chars if total_chars > 0 else 0.0
            
            # Considerar ilegible si:
            # - Ratio CID > threshold, O
            # - Ratio ASCII < 0.7 (menos del 70% es ASCII)
            is_illegible = (cid_ratio > cid_threshold) or (ascii_ratio < 0.7)
            
            return is_illegible, cid_ratio, ascii_ratio
            
    except Exception as e:
        # Si hay error leyendo con pdfplumber, asumir que puede ser ilegible
        print(f"[ADVERTENCIA] Error analizando PDF: {e}")
        return True, 1.0, 0.0


def extract_text_from_ocr_data(ocr_data: dict) -> str:
    """
    Extrae texto plano de los datos OCR para mantener compatibilidad con 'content'.
    
    Args:
        ocr_data: Diccionario con datos de pytesseract.image_to_data()
    
    Returns:
        Texto plano extraído del OCR
    """
    text_parts = []
    n_items = len(ocr_data.get('text', []))
    
    for i in range(n_items):
        level = ocr_data.get('level', [])[i] if i < len(ocr_data.get('level', [])) else 0
        text = ocr_data.get('text', [])[i] if i < len(ocr_data.get('text', [])) else ''
        conf = float(ocr_data.get('conf', [])[i]) if i < len(ocr_data.get('conf', [])) else 0
        
        # Solo agregar palabras (level == 5) con texto y confianza > 0
        if level == 5 and text and conf > 0:
            text_parts.append(text)
    
    return ' '.join(text_parts)


def convert_ocr_data_to_words_format(ocr_data: dict) -> list:
    """
    Convierte datos OCR de pytesseract.image_to_data() a formato de palabras con coordenadas reales.
    Compatible con el formato esperado por el resto del código.
    
    Args:
        ocr_data: Diccionario con datos de pytesseract.image_to_data()
    
    Returns:
        Lista de diccionarios con formato: 
        [{'text': str, 'x0': float, 'top': float, 'x1': float, 'bottom': float, 'conf': float, 'line_num': int}, ...]
    """
    words = []
    n_items = len(ocr_data.get('text', []))
    
    for i in range(n_items):
        level = ocr_data.get('level', [])[i] if i < len(ocr_data.get('level', [])) else 0
        text = ocr_data.get('text', [])[i] if i < len(ocr_data.get('text', [])) else ''
        conf = float(ocr_data.get('conf', [])[i]) if i < len(ocr_data.get('conf', [])) else 0
        
        # Solo procesar palabras (level == 5) con texto y confianza razonable
        if level == 5 and text and conf > 0:
            left = float(ocr_data.get('left', [])[i]) if i < len(ocr_data.get('left', [])) else 0
            top = float(ocr_data.get('top', [])[i]) if i < len(ocr_data.get('top', [])) else 0
            width = float(ocr_data.get('width', [])[i]) if i < len(ocr_data.get('width', [])) else 0
            height = float(ocr_data.get('height', [])[i]) if i < len(ocr_data.get('height', [])) else 0
            line_num = int(ocr_data.get('line_num', [])[i]) if i < len(ocr_data.get('line_num', [])) else 0
            
            words.append({
                'text': text.strip(),
                'x0': left,
                'top': top,
                'x1': left + width,
                'bottom': top + height,
                'conf': conf,
                'line_num': line_num
            })
    
    return words


def fix_ocr_amount_errors(word_text: str, word_x0: float, columns_config: dict, bank_name: str = None) -> str:
    """
    Corrige errores comunes del OCR en montos, usando coordenadas X para validar.
    
    Errores corregidos:
    - "$" leído como "3" al inicio de montos (ej: "3728.00" → "$728.00")
    
    Args:
        word_text: Texto de la palabra extraída por OCR
        word_x0: Coordenada X izquierda de la palabra
        columns_config: Configuración de columnas del banco (con rangos X)
        bank_name: Nombre del banco (solo aplica para bancos específicos)
    
    Returns:
        Texto corregido o texto original si no necesita corrección
    """
    if not word_text or not columns_config:
        return word_text
    
    # Solo aplicar correcciones para bancos específicos
    if bank_name != 'HSBC':
        return word_text
    
    # Patrón: número con 2 decimales que empieza con "3" pero no tiene "$"
    # Ejemplos válidos: "3728.00", "3.78", "3,000.00"
    # No válidos: "30,000.00" (empieza con 30), "3728" (sin decimales)
    # El patrón captura cualquier cantidad de dígitos después del "3"
    amount_pattern = re.compile(r'^3(\d+(?:[,\s]\d{3})*\.\d{2})$')
    match = amount_pattern.match(word_text.strip())
    
    if not match:
        return word_text
    
    # Validación adicional: verificar que esté en rango de columnas de montos
    # Esto reduce falsos positivos (números que realmente empiezan con "3")
    cargos_range = columns_config.get('cargos', (0, 0))
    abonos_range = columns_config.get('abonos', (0, 0))
    saldo_range = columns_config.get('saldo', (0, 0))
    
    # Verificar si la palabra está en alguna columna de montos
    in_amount_column = (
        (cargos_range[0] <= word_x0 <= cargos_range[1]) or
        (abonos_range[0] <= word_x0 <= abonos_range[1]) or
        (saldo_range[0] <= word_x0 <= saldo_range[1])
    )
    
    if in_amount_column:
        # Corregir: reemplazar "3" inicial por "$"
        corrected = f"${match.group(1)}"
        return corrected
    
    return word_text


def convert_ocr_text_to_words_format(ocr_text: str, page_number: int = 1) -> list:
    """
    [DEPRECATED] Convierte texto OCR a formato de palabras con coordenadas aproximadas.
    Esta función está deprecada. Usar convert_ocr_data_to_words_format() en su lugar.
    
    Args:
        ocr_text: Texto extraído por OCR
        page_number: Número de página
    
    Returns:
        Lista de diccionarios con formato: 
        [{'text': str, 'x0': float, 'top': float, 'x1': float, 'bottom': float}, ...]
    """
    words = []
    lines = ocr_text.split('\n')
    
    y_pos = 100  # Posición Y inicial
    
    for line in lines:
        if not line.strip():
            y_pos += 20  # Espacio entre líneas
            continue
        
        # Dividir línea en palabras
        line_words = line.split()
        x_pos = 50  # Posición X inicial
        
        for word_text in line_words:
            if word_text.strip():
                # Aproximación de ancho basado en longitud del texto
                word_width = len(word_text) * 8
                
                words.append({
                    'text': word_text,
                    'x0': x_pos,
                    'top': y_pos,
                    'x1': x_pos + word_width,
                    'bottom': y_pos + 15
                })
                
                x_pos += word_width + 10  # Espacio entre palabras
        
        y_pos += 25  # Altura de línea
    
    return words


def extract_text_with_tesseract_ocr(pdf_path: str, lang: str = 'spa+eng') -> list:
    """
    Extrae texto de PDF usando Tesseract OCR local.
    100% privado - no envía datos a servidores externos.
    Usa PyMuPDF para convertir PDF a imágenes (no requiere Poppler).
    
    Args:
        pdf_path: Ruta al archivo PDF
        lang: Idioma para OCR (default: 'spa+eng' para español+inglés)
    
    Returns:
        Lista de diccionarios con formato: [{"page": int, "content": str, "words": list}, ...]
        Compatible con el formato de retorno de extract_text_from_pdf.
    
    Raises:
        Exception: Si Tesseract no está disponible o hay error
    """
    if not TESSERACT_AVAILABLE:
        raise Exception("Tesseract OCR no está disponible. Instala: pip install pytesseract pymupdf pillow")
    
    # Configurar Tesseract si es necesario
    if not configure_tesseract():
        raise Exception("Tesseract OCR no encontrado. Instala Tesseract desde: https://github.com/UB-Mannheim/tesseract/wiki")
    
    print("[INFO] Extrayendo texto con Tesseract OCR local (100% privado)...")
    
    extracted_data = []
    
    try:
        # Usar PyMuPDF para convertir PDF a imágenes (no requiere Poppler)
        doc = fitz.open(pdf_path)
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            print(f"[INFO] Procesando página {page_num + 1}/{len(doc)} con OCR...")
            
            # Convertir página a imagen (alta resolución)
            mat = fitz.Matrix(2.0, 2.0)  # Zoom 2x para mejor calidad OCR
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            
            # Convertir a PIL Image
            from io import BytesIO
            img = Image.open(BytesIO(img_data))
            
            # Hacer OCR con coordenadas reales
            ocr_data = pytesseract.image_to_data(img, lang=lang, output_type=pytesseract.Output.DICT)
            
            # Extraer texto plano para mantener compatibilidad con 'content'
            text = extract_text_from_ocr_data(ocr_data)
            
            # Convertir datos OCR a formato de palabras con coordenadas reales
            words = convert_ocr_data_to_words_format(ocr_data)
            
            extracted_data.append({
                "page": page_num + 1,
                "content": text,
                "words": words
            })
        
        doc.close()
        
        print(f"[OK] OCR completado. Páginas procesadas: {len(extracted_data)}")
        return extracted_data
        
    except Exception as e:
        raise Exception(f"Error en Tesseract OCR: {e}")


def filter_hsbc_movements_section(pages_data: list, start_string: str, end_string: str) -> list:
    """
    Filtra palabras que están entre start_string y end_string para HSBC.
    El string de inicio puede estar dividido en múltiples palabras, así que busca todas las palabras
    que forman el string de inicio.
    
    Args:
        pages_data: Lista de diccionarios con formato [{"page": int, "words": list}, ...]
        start_string: String que marca el inicio de la sección (ej: "DETALLE MOVIMIENTOS HSBC")
        end_string: String que marca el fin de la sección (ej: "CoDi")
    
    Returns:
        Lista de palabras filtradas que están en la sección de movimientos
    """
    filtered_words = []
    in_section = False
    
    # Normalizar strings para búsqueda (case-insensitive)
    start_string_normalized = start_string.upper().strip()
    start_words_list = start_string_normalized.split()
    end_string_normalized = end_string.upper().strip()
    
    for page_data in pages_data:
        page_num = page_data.get('page', 0)
        words = page_data.get('words', [])
        
        # Buscar inicio: el string puede estar dividido en múltiples palabras
        if not in_section:
            # Construir texto de la página para búsqueda
            page_text = ' '.join([w.get('text', '').strip() for w in words]).upper()
            
            # Buscar si el string de inicio está en el texto de la página (puede estar dividido)
            if start_string_normalized in page_text:
                in_section = True
                # Encontrar la primera palabra después del inicio
                # Contar palabras hasta llegar al inicio
                word_count = 0
                current_text = ''
                start_word_index = None
                for idx, word in enumerate(words):
                    word_text = word.get('text', '').strip()
                    current_text += word_text + ' '
                    if start_string_normalized in current_text.upper():
                        # Ya pasamos el inicio, encontrar el índice de la última palabra del inicio
                        # Buscar cuántas palabras forman el inicio
                        start_word_index = idx + 1  # Empezar desde la siguiente palabra después del inicio
                        break
                
                # Si encontramos el inicio, procesar las palabras restantes de esta página
                if start_word_index is not None and start_word_index < len(words):
                    # Procesar palabras desde después del inicio hasta el final de la página
                    for word in words[start_word_index:]:
                        word_text = word.get('text', '').strip()
                        word_text_upper = word_text.upper()
                        
                        # Detectar fin (búsqueda case-insensitive y parcial)
                        if end_string_normalized in word_text_upper:
                            # Detener completamente (no procesar más páginas)
                            return filtered_words
                        
                        # Agregar información de página a la palabra si no la tiene
                        if 'page' not in word:
                            word['page'] = page_num
                        
                        # Agregar palabra a la sección
                        filtered_words.append(word)
                    
                    # Continuar con la siguiente página
                    continue
        
        # Si ya estamos en la sección, procesar todas las palabras de esta página
        if in_section:
            for word in words:
                word_text = word.get('text', '').strip()
                word_text_upper = word_text.upper()
                
                # Detectar fin (búsqueda case-insensitive y parcial)
                if end_string_normalized in word_text_upper:
                    # Detener completamente (no procesar más páginas)
                    return filtered_words
                
                # Agregar información de página a la palabra si no la tiene
                if 'page' not in word:
                    word['page'] = page_num
                
                # Agregar palabra a la sección
                filtered_words.append(word)
        else:
            # Buscar inicio palabra por palabra (método alternativo si el método anterior no funcionó)
            # Construir texto acumulativo de palabras consecutivas
            for i in range(len(words) - len(start_words_list) + 1):
                # Tomar un grupo de palabras consecutivas del tamaño del string de inicio
                consecutive_words = words[i:i+len(start_words_list)]
                consecutive_text = ' '.join([w.get('text', '').strip() for w in consecutive_words]).upper()
                
                # Verificar si este grupo de palabras coincide con el string de inicio
                if consecutive_text == start_string_normalized or start_string_normalized in consecutive_text:
                    in_section = True
                    # Continuar desde la siguiente palabra después del inicio
                    break
    
    return filtered_words


def extract_hsbc_movements_from_ocr_text(pages_data: list, columns_config: dict = None) -> list:
    """
    Extrae movimientos de HSBC del texto OCR usando coordenadas X reales.
    Asigna montos a columnas (Cargos/Abonos/Saldo) basándose en posición X, no en palabras clave.
    
    IMPORTANTE: Solo extrae información después de "ISR Retenido en el año" y antes de "CoDi"
    
    Args:
        pages_data: Lista de diccionarios con formato [{"page": int, "content": str, "words": list}, ...]
        columns_config: Diccionario con rangos X de columnas desde BANK_CONFIGS.
                       Formato: {"cargos": (x_min, x_max), "abonos": (x_min, x_max), "saldo": (x_min, x_max), ...}
    
    Returns:
        Lista de diccionarios con movimientos: [{"fecha": str, "descripcion": str, "cargos": str, "abonos": str, "saldo": str}, ...]
        - fecha: Solo el día (ej: "03", "13")
        - descripcion: Descripción completa del movimiento
        - cargos: Monto si es retiro/cargo (ej: "$9,500.00") o vacío
        - abonos: Monto si es depósito/abono (ej: "$278,400.00") o vacío
        - saldo: Saldo después del movimiento (ej: "$466,722.66")
    """
    movements = []
    all_text = '\n'.join([p.get('content', '') for p in pages_data])
    lines = all_text.split('\n')
    
    # Buscar sección de movimientos después de "ISR Retenido en el año" y antes de "CoDi"
    in_movements_section = False
    start_found = False
    
    # Patrón para detectar inicio de sección de movimientos
    start_pattern = re.compile(r'ISR\s+Retenido\s+en\s+el\s+a[ñn]o', re.IGNORECASE)
    
    # Patrón para detectar fin de sección de movimientos
    end_pattern = re.compile(r'\bCoDi\b', re.IGNORECASE)
    
    # Patrón para detectar día al inicio de línea (01-31)
    day_pattern = re.compile(r'^(\d{1,2})\s+')
    
    # Patrón para detectar montos ($X,XXX.XX) - incluye el símbolo $ en el resultado
    # Mejorado para capturar montos con espacios internos (ej: "$721 588.28" -> "$721,588.28")
    # Captura montos completos incluyendo espacios internos: $ seguido de dígitos con espacios/comas, y opcionalmente .XX
    # El patrón captura: "$721 588.28", "$721,588.28", "$ 278,400.00", etc.
    # Usa un patrón que captura desde $ hasta encontrar un espacio seguido de texto no numérico o fin de línea
    # Patrón mejorado que captura montos completos con espacios internos
    # Captura: $ seguido de dígitos con espacios/comas entre grupos, y opcionalmente .XX
    # El patrón debe capturar "$721 588.28" completo, no solo "$721"
    # Usa un patrón más flexible que captura desde $ hasta encontrar un espacio seguido de texto no numérico
    # o hasta el fin de línea, pero incluyendo espacios internos entre dígitos
    amount_pattern = re.compile(r'(\$\s*\d{1,3}(?:[,\s]\d{3})*(?:\s+\d{3})*(?:\.\d{2})?)')
    
    # Validar que columns_config tenga los rangos necesarios
    if not columns_config:
        raise ValueError("columns_config es requerido. Debe contener rangos X para 'cargos', 'abonos' y 'saldo'.")
    
    required_columns = ['cargos', 'abonos', 'saldo']
    missing_columns = [col for col in required_columns if col not in columns_config]
    if missing_columns:
        raise ValueError(f"columns_config debe contener rangos para: {missing_columns}")
    
    # Obtener rangos X de columnas
    cargos_range = columns_config.get('cargos')
    abonos_range = columns_config.get('abonos')
    saldo_range = columns_config.get('saldo')
    
    line_count = 0
    for line in lines:
        line_count += 1
        line = line.strip()
        
        # Detectar inicio de sección de movimientos
        if not start_found:
            if start_pattern.search(line):
                start_found = True
                in_movements_section = True
                continue
        
        # Detectar fin de sección de movimientos
        if in_movements_section and end_pattern.search(line):
            break  # Detener completamente la extracción
        
        if not in_movements_section:
            continue
        
        # Buscar línea que empieza con día (01-31)
        day_match = day_pattern.match(line)
        if not day_match:
            continue
        
        # Extraer todos los montos en la línea
        # Usar un método más robusto que capture montos completos con espacios internos
        # Buscar todas las posiciones de $ en la línea
        dollar_positions = [i for i, char in enumerate(line) if char == '$']
        amounts_raw = []
        
        for pos in dollar_positions:
            # Capturar desde $ hasta el siguiente $ o hasta espacio seguido de letra (no dígito)
            remaining = line[pos:]
            # Buscar el monto completo: desde $ hasta encontrar espacio seguido de letra o siguiente $
            # Patrón: $ seguido de dígitos, espacios, comas, puntos
            match = re.match(r'(\$\s*[\d\s,\.]+)', remaining)
            if match:
                monto_candidato = match.group(1)
                # Encontrar dónde termina realmente el monto
                # Buscar el siguiente $ después del primero
                next_dollar = remaining.find('$', 1)
                # Buscar espacio seguido de letra (no dígito)
                next_letter_space = re.search(r'\s+[A-Za-z]', remaining[1:])
                
                if next_dollar != -1 and (next_letter_space is None or next_dollar < next_letter_space.start() + 1):
                    # El siguiente $ está antes del espacio con letra, cortar ahí
                    monto_candidato = remaining[:next_dollar].strip()
                elif next_letter_space:
                    # Hay un espacio seguido de letra, cortar antes de ese espacio
                    monto_candidato = remaining[:next_letter_space.start() + 1].strip()
                else:
                    # No hay siguiente $ ni espacio con letra, tomar hasta el fin de línea
                    monto_candidato = remaining.strip()
                
                amounts_raw.append(monto_candidato)
        
        if len(amounts_raw) < 1:
            continue  # No hay montos, no es un movimiento válido
        
        # Limpiar montos: reemplazar espacios internos con comas para formato consistente
        # Ejemplo: "$721 588.28" -> "$721,588.28" o "$ 278,400.00" -> "$278,400.00"
        amounts = []
        for amt in amounts_raw:
            # Primero, quitar espacios después del $ si los hay
            cleaned = amt.strip()
            # Reemplazar espacios entre dígitos con comas
            # Patrón: buscar dígito seguido de espacio(s) seguido de dígito
            # Esto convierte "$721 588.28" en "$721,588.28"
            # Iterar hasta que no haya más espacios entre dígitos
            while re.search(r'(\d)\s+(\d)', cleaned):
                cleaned = re.sub(r'(\d)\s+(\d)', r'\1,\2', cleaned)
            # Eliminar espacios restantes (como el espacio después del $)
            cleaned = re.sub(r'\s+', '', cleaned)
            # Asegurar formato correcto: si hay múltiples comas consecutivas, dejar solo una
            cleaned = re.sub(r',+', ',', cleaned)
            # Asegurar que después del $ no haya espacio
            cleaned = re.sub(r'\$\s+', '$', cleaned)
            amounts.append(cleaned)
        
        day = day_match.group(1)
        
        # Extraer descripción y referencia (todo hasta el primer $)
        first_dollar = line.find('$')
        if first_dollar == -1:
            continue
        
        desc_and_ref = line[day_match.end():first_dollar].strip()
        
        # Limpiar descripción (normalizar espacios)
        desc_and_ref = re.sub(r'\s+', ' ', desc_and_ref)
        
        # Construir movimiento - fecha es solo el día (ej: "03", "13")
        movement = {
            'fecha': day,  # Solo el día, sin mes/año
            'descripcion': desc_and_ref,
            'cargos': '',
            'abonos': '',
            'saldo': ''
        }
        
        # Asignar montos según coordenadas X (no palabras clave)
        # Calcular posición Y aproximada de la línea (para buscar coordenadas)
        # Usar el índice de la línea en el texto para estimar Y
        line_index = lines.index(line)
        line_y_approx = 100 + (line_index * 25)  # Estimación: 25 píxeles por línea
        
        # Para cada monto, encontrar sus coordenadas X y asignar a columna
        amounts_with_columns = []
        for amount in amounts:
            amount_clean = amount.strip()
            
            # Encontrar coordenadas X reales del monto
            x0, x1 = find_amount_coordinates(amount_clean, line, line_y_approx, pages_data, y_tolerance=20)
            
            if x0 is not None and x1 is not None:
                # Calcular centro del monto para debug
                x_center = (x0 + x1) / 2
                
                # Usar assign_word_to_column() para determinar la columna
                col_name = assign_word_to_column(x0, x1, columns_config)
                
                if col_name:
                    amounts_with_columns.append({
                        'amount': amount_clean,
                        'column': col_name,
                        'x0': x0,
                        'x1': x1,
                        'x_center': x_center
                    })
                else:
                    amounts_with_columns.append({
                        'amount': amount_clean,
                        'column': None,
                        'x0': x0,
                        'x1': x1
                    })
            else:
                amounts_with_columns.append({
                    'amount': amount_clean,
                    'column': None,
                    'x0': None,
                    'x1': None
                })
        
        # Asignar montos a columnas según resultados
        # Si hay ambigüedad (monto asignado a cargos pero podría ser abonos), usar contexto
        # Estrategia: Si hay 2 montos y ambos están en rango de cargos, el más a la izquierda es cargos, el otro es saldo
        # Si un monto está en rango de cargos pero está más cerca del centro de abonos, reconsiderar
        
        # Primero, identificar montos que podrían estar mal asignados
        # Si un monto está en rango de cargos pero su centro está más cerca del centro de abonos, reconsiderar
        for amt_data in amounts_with_columns:
            if amt_data['column'] == 'cargos' and amt_data.get('x_center') is not None:
                x_center = amt_data['x_center']
                # Calcular distancia a centros de ambos rangos
                if 'cargos' in columns_config and 'abonos' in columns_config:
                    cargos_min, cargos_max = columns_config['cargos']
                    abonos_min, abonos_max = columns_config['abonos']
                    if cargos_min > cargos_max:
                        cargos_min, cargos_max = cargos_max, cargos_min
                    if abonos_min > abonos_max:
                        abonos_min, abonos_max = abonos_max, abonos_min
                    
                    cargos_center = (cargos_min + cargos_max) / 2
                    abonos_center = (abonos_min + abonos_max) / 2
                    
                    dist_cargos = abs(x_center - cargos_center)
                    dist_abonos = abs(x_center - abonos_center)
                    
                    # Si está significativamente más cerca de abonos (diferencia > 30 píxeles), reconsiderar
                    if dist_abonos < dist_cargos and (dist_cargos - dist_abonos) > 30:
                        # Verificar si está dentro o cerca del rango de abonos (expandir rango 50 píxeles)
                        abonos_range_expanded_min = abonos_min - 50
                        abonos_range_expanded_max = abonos_max + 50
                        if abonos_range_expanded_min <= x_center <= abonos_range_expanded_max:
                            amt_data['column'] = 'abonos'
        
        # Ahora asignar montos a columnas
        for amt_data in amounts_with_columns:
            amt = amt_data['amount']
            col = amt_data['column']
            
            # Solo asignar si la columna no está ya ocupada
            if col == 'saldo' and not movement.get('saldo'):
                movement['saldo'] = amt
            elif col == 'cargos' and not movement.get('cargos'):
                movement['cargos'] = amt
            elif col == 'abonos' and not movement.get('abonos'):
                movement['abonos'] = amt
        
        # Asegurar que saldo esté asignado cuando hay 2 montos
        # Si hay 2 montos y saldo no está asignado, el segundo monto es saldo (estructura típica HSBC)
        if len(amounts) == 2 and not movement.get('saldo'):
            # Verificar si el segundo monto está asignado a otra columna
            second_amt_data = amounts_with_columns[1] if len(amounts_with_columns) > 1 else None
            if second_amt_data and second_amt_data.get('column') and second_amt_data['column'] != 'saldo':
                # El segundo monto está asignado a otra columna (cargos o abonos), pero debería ser saldo
                # Reasignar el segundo monto a saldo
                movement['saldo'] = amounts[1].strip()
            elif not second_amt_data or not second_amt_data.get('column'):
                # El segundo monto no está asignado, asignarlo a saldo
                movement['saldo'] = amounts[1].strip()
        
        # Fallback: si hay montos sin asignar, usar lógica por posición
        unassigned = [a for a in amounts_with_columns if a['column'] is None]
        if unassigned:
            # Si hay 2 montos
            if len(amounts) == 2:
                assigned_cols = [a['column'] for a in amounts_with_columns if a['column']]
                
                # Si saldo aún no está asignado, el segundo monto es saldo
                if 'saldo' not in assigned_cols and not movement.get('saldo'):
                    movement['saldo'] = amounts[1].strip()
                
                # Si cargos y abonos no están asignados, el primer monto es cargos (por estructura HSBC)
                if 'cargos' not in assigned_cols and 'abonos' not in assigned_cols and not movement.get('cargos'):
                    # En HSBC, cuando hay 2 montos, el primero es generalmente Cargos (Retiro/Cargo)
                    # Solo asignar a Abonos si hay evidencia clara (coordenadas X muy cerca de rango abonos)
                    first_amt = amounts_with_columns[0]
                    if first_amt['x0'] is not None:
                        # Calcular distancia a cada rango
                        x_center = (first_amt['x0'] + first_amt['x1']) / 2
                        cargos_center = (cargos_range[0] + cargos_range[1]) / 2
                        abonos_center = (abonos_range[0] + abonos_range[1]) / 2
                        
                        dist_cargos = abs(x_center - cargos_center)
                        dist_abonos = abs(x_center - abonos_center)
                        
                        # Si está significativamente más cerca de abonos (diferencia > 50 píxeles), asignar a abonos
                        # De lo contrario, asignar a cargos (estructura típica de HSBC)
                        if dist_abonos < dist_cargos and (dist_cargos - dist_abonos) > 50:
                            movement['abonos'] = first_amt['amount']
                        else:
                            movement['cargos'] = first_amt['amount']
                    else:
                        # Sin coordenadas, usar estructura típica: primer monto = cargos
                        movement['cargos'] = first_amt['amount']
            
            # Si hay 3+ montos, asignar por posición: primero=cargos, segundo=abonos, último=saldo
            elif len(amounts) >= 3:
                if not movement.get('cargos'):
                    movement['cargos'] = amounts[0].strip()
                if not movement.get('abonos'):
                    movement['abonos'] = amounts[1].strip()
                if not movement.get('saldo'):
                    movement['saldo'] = amounts[-1].strip()
        
        # Verificación final: Asegurar que saldo esté asignado cuando hay 2+ montos
        # Esta es una verificación de seguridad adicional
        if len(amounts) >= 2 and not movement.get('saldo'):
            # Si hay 2+ montos y saldo no está asignado, asignar el último monto a saldo
            movement['saldo'] = amounts[-1].strip()
        
        if movement['descripcion'] and (movement['cargos'] or movement['abonos'] or movement['saldo']):
            movements.append(movement)
    print(f"    Total de líneas procesadas: {line_count}")
    print(f"    Sección de movimientos encontrada: {start_found}")
    print(f"    Total de movimientos extraídos: {len(movements)}")
    if movements:
        print(f"    Primer movimiento: fecha='{movements[0]['fecha']}', descripcion='{movements[0]['descripcion'][:50]}'")
        print(f"    Último movimiento: fecha='{movements[-1]['fecha']}', descripcion='{movements[-1]['descripcion'][:50]}'")
    else:
        print(f"    ⚠️  NO SE EXTRAJERON MOVIMIENTOS")
    
    return movements


def find_amount_coordinates(amount_text: str, line_text: str, line_y_approx: float, pages_data: list, y_tolerance: float = 15) -> tuple:
    """
    Encuentra las coordenadas X reales de un monto en pages_data['words'] usando coordenadas reales del OCR.
    
    Args:
        amount_text: Texto del monto (ej: "$30,022.54")
        line_text: Texto completo de la línea donde aparece el monto
        line_y_approx: Posición Y aproximada de la línea (para filtrar palabras)
        pages_data: Lista de diccionarios con palabras y coordenadas
        y_tolerance: Tolerancia en píxeles para buscar en la misma línea Y (reducida a 15 para coordenadas reales)
    
    Returns:
        Tupla (x0, x1) con coordenadas X del monto, o (None, None) si no se encuentra
    """
    # Normalizar amount_text para comparación (quitar espacios, normalizar formato)
    amount_normalized = amount_text.replace(' ', '').replace(',', '').replace('$', '')
    
    # Debug: mostrar información de búsqueda
    print(f"    [DEBUG find_amount_coordinates] Buscando monto: '{amount_text}' (normalizado: '{amount_normalized}')")
    print(f"    [DEBUG find_amount_coordinates] Línea Y aproximada: {line_y_approx:.1f}, tolerancia: {y_tolerance}")
    print(f"    [DEBUG find_amount_coordinates] Línea completa: '{line_text[:100]}'")
    
    # Buscar en todas las páginas
    words_checked = 0
    words_in_y_range = 0
    for page_data in pages_data:
        words = page_data.get('words', [])
        
        # Agrupar palabras por line_num si está disponible (más preciso)
        words_by_line = {}
        for word in words:
            line_num = word.get('line_num', None)
            if line_num is not None:
                if line_num not in words_by_line:
                    words_by_line[line_num] = []
                words_by_line[line_num].append(word)
        
        for word in words:
            words_checked += 1
            word_text = word.get('text', '').strip()
            word_y = word.get('top', 0)
            word_x0 = word.get('x0', 0)
            word_x1 = word.get('x1', 0)
            word_line_num = word.get('line_num', None)
            word_conf = word.get('conf', 100)
            
            # Filtrar palabras con baja confianza (opcional, pero útil)
            if word_conf < 30:
                continue
            
            # Verificar si está en la misma línea Y (con tolerancia reducida para coordenadas reales)
            if abs(word_y - line_y_approx) > y_tolerance:
                continue
            
            words_in_y_range += 1
            
            # Verificar si esta palabra contiene el monto
            # Normalizar palabra para comparación
            word_normalized = word_text.replace(' ', '').replace(',', '').replace('$', '')
            
            # Buscar si el monto está contenido en esta palabra o en palabras adyacentes
            if amount_normalized in word_normalized or word_normalized in amount_normalized:
                print(f"    [DEBUG find_amount_coordinates] ✓ Encontrado en palabra: '{word_text}' (Y: {word_y:.1f}, X: {word_x0:.1f}-{word_x1:.1f}, conf: {word_conf:.1f})")
                return (word_x0, word_x1)
            
            # También buscar si el monto aparece como parte de la línea
            # (puede estar dividido en múltiples palabras: "$30" "022.54")
            if '$' in word_text:
                # Esta palabra tiene $, puede ser el inicio del monto
                # Si tenemos line_num, buscar solo en la misma línea
                if word_line_num is not None and word_line_num in words_by_line:
                    # Buscar en palabras de la misma línea (más preciso)
                    line_words = words_by_line[word_line_num]
                    word_idx = next((i for i, w in enumerate(line_words) if w == word), -1)
                    
                    if word_idx >= 0:
                        combined_text = word_text
                        combined_x0 = word_x0
                        combined_x1 = word_x1
                        
                        # Buscar palabras siguientes en la misma línea
                        for next_word in line_words[word_idx + 1:]:
                            next_word_text = next_word.get('text', '').strip()
                            
                            # Agregar palabra siguiente
                            combined_text += next_word_text
                            combined_x1 = next_word.get('x1', combined_x1)
                            
                            # Normalizar y comparar
                            combined_normalized = combined_text.replace(' ', '').replace(',', '').replace('$', '')
                            if amount_normalized in combined_normalized:
                                print(f"    [DEBUG find_amount_coordinates] ✓ Encontrado en palabras combinadas (line_num={word_line_num}): '{combined_text}' (X: {combined_x0:.1f}-{combined_x1:.1f})")
                                return (combined_x0, combined_x1)
                            
                            # Si ya tenemos suficientes caracteres, detener
                            if len(combined_normalized) >= len(amount_normalized) + 5:
                                break
                else:
                    # Fallback: buscar en todas las palabras (método anterior)
                    word_idx = words.index(word)
                    combined_text = word_text
                    combined_x0 = word_x0
                    combined_x1 = word_x1
                    
                    # Buscar palabras siguientes en la misma línea Y
                    for next_word in words[word_idx + 1:]:
                        next_word_y = next_word.get('top', 0)
                        next_word_text = next_word.get('text', '').strip()
                        
                        if abs(next_word_y - line_y_approx) > y_tolerance:
                            break  # Ya no está en la misma línea
                        
                        # Agregar palabra siguiente
                        combined_text += next_word_text
                        combined_x1 = next_word.get('x1', combined_x1)
                        
                        # Normalizar y comparar
                        combined_normalized = combined_text.replace(' ', '').replace(',', '').replace('$', '')
                        if amount_normalized in combined_normalized:
                            print(f"    [DEBUG find_amount_coordinates] ✓ Encontrado en palabras combinadas: '{combined_text}' (X: {combined_x0:.1f}-{combined_x1:.1f})")
                            return (combined_x0, combined_x1)
                        
                        # Si ya tenemos suficientes caracteres, detener
                        if len(combined_normalized) >= len(amount_normalized) + 5:
                            break
    
    # Debug: mostrar estadísticas si no se encontró
    print(f"    [DEBUG find_amount_coordinates] ✗ NO encontrado. Palabras revisadas: {words_checked}, en rango Y: {words_in_y_range}")
    print(f"    [DEBUG find_amount_coordinates] Posibles causas:")
    print(f"        - line_y_approx ({line_y_approx:.1f}) no coincide con Y real de las palabras")
    print(f"        - El monto está dividido en múltiples palabras y no se combinó correctamente")
    print(f"        - El formato del monto en words no coincide con amount_text")
    
    return (None, None)


def extract_hsbc_summary_from_ocr_text(pages_data: list) -> dict:
    """
    Extrae información de resumen de HSBC desde el texto OCR de la página 1.
    
    Busca:
    - Total Abonos: "Depósitos/Abonos $X,XXX.XX"
    - Total Cargos: "Retiros/Cargos $X,XXX.XX"
    - Saldo Final: "Saldo Final del Periodo $X,XXX.XX"
    
    Args:
        pages_data: Lista de diccionarios con formato [{"page": int, "content": str, "words": list}, ...]
    
    Returns:
        Diccionario con información de resumen: {
            'total_abonos': str o None,
            'total_cargos': str o None,
            'saldo_final': str o None,
            'total_depositos': str o None,
            'total_retiros': str o None
        }
    """
    summary_data = {
        'total_depositos': None,
        'total_retiros': None,
        'total_cargos': None,
        'total_abonos': None,
        'saldo_final': None,
        'total_movimientos': None,
        'saldo_anterior': None
    }
    
    # Obtener texto de la página 1 (índice 0)
    if not pages_data or len(pages_data) == 0:
        return summary_data
    
    page_1_data = pages_data[0]  # Primera página
    page_1_text = page_1_data.get('content', '')
    
    if not page_1_text:
        return summary_data
    
    lines = page_1_text.split('\n')
    
    
    # Patrones más flexibles para buscar los valores (permitir espacios variables, $ opcional, y caracteres especiales)
    # Buscar "Depósitos" seguido de "!" o "/Abonos" y luego $ y monto
    # Ejemplo real: "Depósitos! $ 308,422.54"
    # El patrón debe buscar "Depósitos" seguido de "!" o ">" y luego el monto
    depositos_pattern = re.compile(r'Dep[oó]sitos?\s*[!>]\s*\$?\s*([\d,\.]+)', re.IGNORECASE)
    
    # Buscar "Retiros/Cargos" seguido de $ y monto
    retiros_pattern = re.compile(r'Retiros?/Cargos?\s*\$?\s*([\d,\.]+)', re.IGNORECASE)
    
    # Buscar "Saldo Final del" (con o sin "Periodo") seguido de $ y monto
    # Ejemplo real: "Saldo Final del $ 671,749.84"
    # El patrón debe permitir texto antes de "Saldo Final del"
    saldo_final_pattern = re.compile(r'Saldo\s+Final\s+del\s+(?:Periodo\s+)?\$?\s*([\d,\.]+)', re.IGNORECASE)
    
    # También buscar variaciones de "Saldo Final" sin "del" (más flexible, permite texto antes)
    saldo_final_pattern2 = re.compile(r'Saldo\s+Final\s+del\s*\$?\s*([\d,\.]+)', re.IGNORECASE)
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        
        # Buscar Depósitos/Abonos - buscar "Depósitos" seguido de "!" o ">" y luego monto
        # Ejemplo: "Depósitos! $ 308,422.54"
        if not summary_data['total_abonos']:
            match = depositos_pattern.search(line)
            if match:
                amount_str = match.group(1)
                # Convertir a número usando normalize_amount_str (igual que extract_summary_from_pdf)
                amount = normalize_amount_str(amount_str)
                if amount > 0:
                    summary_data['total_abonos'] = amount
                    summary_data['total_depositos'] = amount
        
        # Buscar Retiros/Cargos
        if not summary_data['total_cargos']:
            match = retiros_pattern.search(line)
            if match:
                amount_str = match.group(1)
                # Convertir a número usando normalize_amount_str (igual que extract_summary_from_pdf)
                amount = normalize_amount_str(amount_str)
                if amount > 0:
                    summary_data['total_cargos'] = amount
                    summary_data['total_retiros'] = amount
        
        # Buscar Saldo Final del Periodo (patrón completo)
        if not summary_data['saldo_final']:
            match = saldo_final_pattern.search(line)
            if not match:
                # Intentar patrón alternativo: "Saldo Final del" sin "Periodo"
                match = saldo_final_pattern2.search(line)
            if match:
                amount_str = match.group(1)
                # Convertir a número usando normalize_amount_str (igual que extract_summary_from_pdf)
                amount = normalize_amount_str(amount_str)
                if amount > 0:
                    summary_data['saldo_final'] = amount
    
    
    return summary_data


def extract_summary_from_pdf(pdf_path: str) -> dict:
    """
    Extract summary information from PDF (totals, deposits, withdrawals, balance, movement count).
    Uses bank-specific patterns to extract summary data accurately.
    Returns a dictionary with extracted values or None if not found.
    """
    summary_data = {
        'total_depositos': None,
        'total_retiros': None,
        'total_cargos': None,
        'total_abonos': None,
        'saldo_final': None,
        'total_movimientos': None,
        'saldo_anterior': None
    }
    
    try:
        # First, detect the bank
        bank_name = detect_bank_from_pdf(pdf_path)
        # print(f"🏦 Extrayendo resumen para banco: {bank_name}")
        
        with pdfplumber.open(pdf_path) as pdf:
            # Check first few pages and last page for summary information
            # For Banregio, check all pages to find "Total" line which can be on any page
            pages_to_check = min(3, len(pdf.pages))
            all_text = ""
            all_lines = []
            
            # For Konfio, read page 2 for summary information and all pages for "Subtotal" line
            if bank_name == "Konfio":
                # Read page 2 for summary information (Pagos, Devoluciones, Compras y cargos, Saldo total al corte)
                if len(pdf.pages) >= 2:
                    page = pdf.pages[1]  # Page 2 (0-indexed)
                    text = page.extract_text()
                    if text:
                        # Fix duplicated characters issue (e.g., "PPaaggoo" -> "Pagos")
                        # This happens with some PDF encodings where each character is duplicated
                        text = fix_duplicated_chars(text)
                        all_text = text + "\n"
                        all_lines = text.split('\n')
                else:
                    all_lines = []  # Not enough pages
                
                # Also read all pages to find "Subtotal" line (usually at the end of movements)
                # This line contains: "Subtotal $ X,XXX.XX $ Y,YYY.YY" where first is cargos, second is abonos
                for page_num in range(len(pdf.pages)):
                    page = pdf.pages[page_num]
                    text = page.extract_text()
                    if text:
                        text = fix_duplicated_chars(text)
                        all_lines.extend(text.split('\n'))
            # For Banregio, collect text from all pages to find "Total" line
            elif bank_name == "Banregio":
                for page_num in range(len(pdf.pages)):
                    page = pdf.pages[page_num]
                    text = page.extract_text()
                    if text:
                        all_text += text + "\n"
                        all_lines.extend(text.split('\n'))
            else:
                # Collect text from first pages
                for page_num in range(pages_to_check):
                    page = pdf.pages[page_num]
                    text = page.extract_text()
                    if text:
                        all_text += text + "\n"
                        all_lines.extend(text.split('\n'))
                
                # Also check last page for Santander and BanRegio
                if len(pdf.pages) > pages_to_check:
                    last_page = pdf.pages[-1]
                    last_text = last_page.extract_text()
                    if last_text:
                        all_lines.extend(last_text.split('\n'))
            
            # Bank-specific extraction
            if bank_name == "BBVA":
                # BBVA: "Depósitos / Abonos (+) 1 25,000.00" - el último número es el total
                # "Retiros / Cargos (-) 25 53,877.37"
                # "Saldo Final (+) 166,301.83"
                #print(f"🔍 Buscando patrones BBVA en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    # Try multiple patterns for BBVA depósitos
                    if not summary_data['total_abonos']:
                        patterns_bbva = [
                            r'Dep[oó]sitos\s*/\s*Abonos\s*\(\+\)\s+\d+\s+([\d,\.]+)',  # Original pattern
                            r'Dep[oó]sitos\s*/\s*Abonos\s*\(\+\).*?([\d,\.]+)',  # More flexible
                            r'Dep[oó]sitos.*?Abonos.*?\(\+\).*?([\d,\.]+)',  # Even more flexible
                        ]
                        for pattern in patterns_bbva:
                            match = re.search(pattern, line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    #print(f"✅ BBVA: Encontrado depósitos/abonos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                    summary_data['total_abonos'] = amount
                                    summary_data['total_depositos'] = amount
                                    break
                    
                    # Search for Retiros / Cargos
                    if not summary_data['total_cargos']:
                        patterns_retiros = [
                            r'Retiros\s*/\s*Cargos\s*\(\-\)\s+\d+\s+([\d,\.]+)',
                            r'Retiros\s*/\s*Cargos\s*\(\-\).*?([\d,\.]+)',
                            r'Retiros.*?Cargos.*?\(\-\).*?([\d,\.]+)',
                        ]
                        for pattern in patterns_retiros:
                            match = re.search(pattern, line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    print(f"✅ BBVA: Encontrado retiros/cargos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                    summary_data['total_cargos'] = amount
                                    summary_data['total_retiros'] = amount
                                    break
                    
                    # Search for Saldo Final
                    if not summary_data['saldo_final']:
                        patterns_saldo = [
                            r'Saldo\s+Final\s*\(\+\)\s+([\d,\.]+)',
                            r'Saldo\s+Final.*?([\d,\.]+)',
                        ]
                        for pattern in patterns_saldo:
                            match = re.search(pattern, line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    print(f"✅ BBVA: Encontrado saldo final: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                    summary_data['saldo_final'] = amount
                                    break
            
            elif bank_name == "Inbursa":
                # Inbursa: "ABONOS 9,375.49" - el número después de ABONOS
                # "CARGOS 58,927.68"
                # "SALDO ACTUAL 546,409.22"
                # "SALDO ANTERIOR 595,961.41"
                #print(f"🔍 Buscando patrones Inbursa en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    if not summary_data['total_abonos']:
                        match = re.search(r'ABONOS\s+([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Inbursa: Encontrado abonos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_abonos'] = amount
                                summary_data['total_depositos'] = amount
                    
                    if not summary_data['total_cargos']:
                        match = re.search(r'CARGOS\s+([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Inbursa: Encontrado cargos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_cargos'] = amount
                                summary_data['total_retiros'] = amount
                    
                    if not summary_data['saldo_final']:
                        match = re.search(r'SALDO\s+ACTUAL\s+([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Inbursa: Encontrado saldo actual: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_final'] = amount
                    
                    if not summary_data['saldo_anterior']:
                        match = re.search(r'SALDO\s+ANTERIOR\s+([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Inbursa: Encontrado saldo anterior: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_anterior'] = amount
            
            elif bank_name == "Santander":
                # Santander: Extract from first page, section between "CUENTA DE CHEQUES" and "GRAFICO CUENTA DE CHEQUES"
                # Patterns:
                # "+ DEPOSITOS 821,646.20" -> Total de Abonos
                # "- RETIROS 820,238.73" -> Total de Retiros
                # "SALDO ACTUAL 1,417.18" -> Saldo Final
                #print(f"🔍 Buscando patrones Santander en primera página...")
                
                # Get first page text for more reliable extraction
                first_page_text = ""
                if len(pdf.pages) > 0:
                    first_page = pdf.pages[0]
                    first_page_text = first_page.extract_text() or ""
                
                # Find the section between "CUENTA DE CHEQUES" and "GRAFICO CUENTA DE CHEQUES"
                cuenta_match = re.search(r'CUENTA\s+DE\s+CHEQUES', first_page_text, re.I)
                grafico_match = re.search(r'GRAFICO\s+CUENTA\s+DE\s+CHEQUES', first_page_text, re.I)
                
                if cuenta_match and grafico_match:
                    # Extract the section text
                    section_start = cuenta_match.start()
                    section_end = grafico_match.start()
                    section_text = first_page_text[section_start:section_end]
                    
                    # Pattern: "+ DEPOSITOS 821,646.20" or "+DEPOSITOS 821,646.20"
                    if not summary_data['total_depositos'] or not summary_data['total_abonos']:
                        match = re.search(r'[+\s]+DEPOSITOS\s+([\d,\.]+)', section_text, re.I)
                        if match:
                            depositos = normalize_amount_str(match.group(1))
                            #print(f"✅ Santander: Encontrado DEPOSITOS: ${depositos:,.2f}")
                            if depositos > 0:
                                summary_data['total_depositos'] = depositos
                                summary_data['total_abonos'] = depositos
                    
                    # Pattern: "- RETIROS 820,238.73" or "-RETIROS 820,238.73"
                    if not summary_data['total_retiros'] or not summary_data['total_cargos']:
                        match = re.search(r'[-\s]+RETIROS\s+([\d,\.]+)', section_text, re.I)
                        if match:
                            retiros = normalize_amount_str(match.group(1))
                            print(f"✅ Santander: Encontrado RETIROS: ${retiros:,.2f}")
                            if retiros > 0:
                                summary_data['total_retiros'] = retiros
                                summary_data['total_cargos'] = retiros
                    
                    # Pattern: "SALDO ACTUAL 1,417.18" or "= SALDO ACTUAL 1,417.18"
                    if not summary_data['saldo_final']:
                        match = re.search(r'(?:=\s*)?SALDO\s+ACTUAL\s+([\d,\.]+)', section_text, re.I)
                        if match:
                            saldo = normalize_amount_str(match.group(1))
                            #print(f"✅ Santander: Encontrado SALDO ACTUAL: ${saldo:,.2f}")
                            if saldo > 0:
                                summary_data['saldo_final'] = saldo
                else:
                    # Fallback: search in all text if section markers not found
                    #print(f"⚠️  No se encontró sección CUENTA DE CHEQUES, buscando en todo el texto...")
                    full_text = all_text
                    
                    # Pattern: "+ DEPOSITOS 821,646.20"
                    if not summary_data['total_depositos'] or not summary_data['total_abonos']:
                        match = re.search(r'[+\s]+DEPOSITOS\s+([\d,\.]+)', full_text, re.I)
                        if match:
                            depositos = normalize_amount_str(match.group(1))
                            if depositos > 0:
                                summary_data['total_depositos'] = depositos
                                summary_data['total_abonos'] = depositos
                    
                    # Pattern: "- RETIROS 820,238.73"
                    if not summary_data['total_retiros'] or not summary_data['total_cargos']:
                        match = re.search(r'[-\s]+RETIROS\s+([\d,\.]+)', full_text, re.I)
                        if match:
                            retiros = normalize_amount_str(match.group(1))
                            if retiros > 0:
                                summary_data['total_retiros'] = retiros
                                summary_data['total_cargos'] = retiros
                    
                    # Pattern: "SALDO ACTUAL 1,417.18"
                    if not summary_data['saldo_final']:
                        match = re.search(r'(?:=\s*)?SALDO\s+ACTUAL\s+([\d,\.]+)', full_text, re.I)
                        if match:
                            saldo = normalize_amount_str(match.group(1))
                            if saldo > 0:
                                summary_data['saldo_final'] = saldo
            
            elif bank_name == "Banorte":
                # Banorte: 
                # "Saldo inicial del periodo $ 2,284.38"
                # "+ Total de depósitos $ 38,396.00"
                # "- Total de retiros $ 36,805.40"
                # "Saldo actual $ 3,347.18"
                #print(f"🔍 Buscando patrones Banorte en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    # Saldo inicial
                    if not summary_data['saldo_anterior']:
                        match = re.search(r'Saldo\s+inicial\s+del\s+periodo\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Banorte: Encontrado saldo inicial: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_anterior'] = amount
                    
                    # Depósitos
                    if not summary_data['total_depositos']:
                        match = re.search(r'\+\s*Total\s+de\s+dep[oó]sitos\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Banorte: Encontrado depósitos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_depositos'] = amount
                                summary_data['total_abonos'] = amount
                    
                    # Retiros
                    if not summary_data['total_retiros']:
                        match = re.search(r'-\s*Total\s+de\s+retiros\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Banorte: Encontrado retiros: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_retiros'] = amount
                                summary_data['total_cargos'] = amount
                    
                    # Saldo actual
                    if not summary_data['saldo_final']:
                        match = re.search(r'Saldo\s+actual\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Banorte: Encontrado saldo actual: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_final'] = amount
            
            elif bank_name == "Banamex":
                # Banamex: 
                # "Saldo Anterior $5,297.64"
                # "( + ) 8 Depósitos $344,527.26"
                # "( - ) 16 Retiros $254,072.38"
                # "SALDO AL 31 DE ENERO DE 2020 $95,752.52"
                #print(f"🔍 Buscando patrones Banamex en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    # Saldo Anterior
                    if not summary_data['saldo_anterior']:
                        match = re.search(r'Saldo\s+Anterior\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Banamex: Encontrado saldo anterior: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_anterior'] = amount
                    
                    # Depósitos
                    if not summary_data['total_depositos']:
                        match = re.search(r'\(\s*\+\s*\)\s+\d+\s+Dep[oó]sitos\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Banamex: Encontrado depósitos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_depositos'] = amount
                                summary_data['total_abonos'] = amount
                    
                    # Retiros
                    if not summary_data['total_retiros']:
                        match = re.search(r'\(\s*-\s*\)\s+\d+\s+Retiros\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Banamex: Encontrado retiros: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_retiros'] = amount
                                summary_data['total_cargos'] = amount
                    
                    # Saldo Final
                    if not summary_data['saldo_final']:
                        match = re.search(r'SALDO\s+AL\s+\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4}\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Banamex: Encontrado saldo final: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_final'] = amount
                        # Also try simpler pattern
                        if not summary_data['saldo_final']:
                            match = re.search(r'SALDO\s+AL.*?\$\s*([\d,\.]+)', line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    #print(f"✅ Banamex: Encontrado saldo final (patrón simple): ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                    summary_data['saldo_final'] = amount
            
            elif bank_name == "Banbajío":
                # BanBajío: tabla con "SALDO ANTERIOR (+) DEPOSITOS (-) CARGOS SALDO ACTUAL" y valores en la siguiente línea
                # Pattern: "$ 5,280.55 $ 1,441,951.06 $ 1,350,565.02 $ 96,666.59"
                #print(f"🔍 Buscando patrones BanBajío en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    if re.search(r'SALDO\s+ANTERIOR.*DEPOSITOS.*CARGOS.*SALDO\s+ACTUAL', line, re.I):
                        #print(f"✅ BanBajío: Encontrada tabla en línea {i+1}: {line[:100]}")
                        # Next line should have the values
                        if i + 1 < len(all_lines):
                            values_line = all_lines[i + 1]
                            #print(f"   Línea de valores: {values_line[:100]}")
                            # Extract all amounts from the line
                            amounts = re.findall(r'\$\s*([\d,\.]+)', values_line)
                            if len(amounts) >= 4:
                                summary_data['saldo_anterior'] = normalize_amount_str(amounts[0])
                                summary_data['total_depositos'] = normalize_amount_str(amounts[1])
                                summary_data['total_abonos'] = normalize_amount_str(amounts[1])
                                summary_data['total_cargos'] = normalize_amount_str(amounts[2])
                                summary_data['total_retiros'] = normalize_amount_str(amounts[2])
                                summary_data['saldo_final'] = normalize_amount_str(amounts[3])
                                #print(f"   Extraídos: Saldo anterior=${summary_data['saldo_anterior']:,.2f}, Depósitos=${summary_data['total_depositos']:,.2f}, Cargos=${summary_data['total_cargos']:,.2f}, Saldo final=${summary_data['saldo_final']:,.2f}")
                            else:
                                pass
                                # print(f"   ⚠️  Solo se encontraron {len(amounts)} valores, se esperaban 4")
                            break
            
            elif bank_name == "Banregio":
                # BanRegio: 
                # "Saldo Inicial $903.18"
                # "+ Abonos $49,675.60"
                # "- Retiros $7,000.00"
                # "- Comisiones Efectivamente Cobradas $320.00"
                # "- Otros Cargos $38,678.00"
                # "= Saldo Final $4,580.78"
                # Also: "Total 45,998.00 49,675.60 4,580.78" (Cargos, Abonos, Saldo)
                #print(f"🔍 Buscando patrones BanRegio en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    # First, try to extract from "Total" line (most reliable)
                    # Pattern: "Total 45,998.00 49,675.60 4,580.78" (Cargos, Abonos, Saldo)
                    # Make pattern more flexible to handle variations in spacing and allow for optional leading text
                    total_match = re.search(r'\bTotal\s+([\d,\.]+)\s+([\d,\.]+)\s+([\d,\.]+)', line, re.I)
                    if total_match:
                        cargos = normalize_amount_str(total_match.group(1))
                        abonos = normalize_amount_str(total_match.group(2))
                        saldo = normalize_amount_str(total_match.group(3))
                        if cargos > 0:
                            summary_data['total_cargos'] = cargos
                            summary_data['total_retiros'] = cargos
                        if abonos > 0:
                            summary_data['total_abonos'] = abonos
                            summary_data['total_depositos'] = abonos
                        if saldo > 0:
                            summary_data['saldo_final'] = saldo
                        # If we found all values from Total line, we can break early
                        if summary_data['total_cargos'] and summary_data['total_abonos'] and summary_data['saldo_final']:
                            break
                        # Continue to next iteration to avoid checking fallback patterns if we found Total line
                        continue
                    
                    # Saldo Inicial
                    if not summary_data['saldo_anterior']:
                        match = re.search(r'Saldo\s+Inicial\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ BanRegio: Encontrado saldo inicial: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_anterior'] = amount
                    
                    # Abonos (fallback if not found in Total line)
                    if not summary_data['total_abonos']:
                        match = re.search(r'\+\s*Abonos\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ BanRegio: Encontrado abonos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_abonos'] = amount
                                summary_data['total_depositos'] = amount
                    
                    # Retiros (fallback if not found in Total line)
                    if not summary_data['total_retiros']:
                        # First check if line has "Retiros" and next line has amount
                        if re.search(r'-\s*Retiros', line, re.I) and i + 1 < len(all_lines):
                            next_line = all_lines[i + 1]
                            match = re.search(r'\$\s*([\d,\.]+)', next_line)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    #print(f"✅ BanRegio: Encontrado retiros: ${amount:,.2f} en línea {i+2}: {next_line[:80]}")
                                    summary_data['total_retiros'] = amount
                                    summary_data['total_cargos'] = amount
                    
                    # Saldo Final (fallback if not found in Total line)
                    if not summary_data['saldo_final']:
                        # First check if line has "Saldo Final" and next line has amount
                        if re.search(r'=\s*Saldo\s+Final', line, re.I) and i + 1 < len(all_lines):
                            next_line = all_lines[i + 1]
                            match = re.search(r'\$\s*([\d,\.]+)', next_line)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    #print(f"✅ BanRegio: Encontrado saldo final: ${amount:,.2f} en línea {i+2}: {next_line[:80]}")
                                    summary_data['saldo_final'] = amount
            
            elif bank_name == "Clara":
                # Clara:
                # "+ Saldo anterior 3,305.40"
                # "- Pagos -3,305.40"
                # "+ Compras y cargos del periodo 3,115.30"
                # "Saldo al corte 3,115.30"
                #print(f"🔍 Buscando patrones Clara en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    # Saldo anterior
                    if not summary_data['saldo_anterior']:
                        match = re.search(r'\+\s*Saldo\s+anterior\s+([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Clara: Encontrado saldo anterior: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_anterior'] = amount
                    
                    # Compras y cargos (esto son los cargos)
                    if not summary_data['total_cargos']:
                        match = re.search(r'\+\s*Compras\s+y\s+cargos\s+del\s+periodo\s+([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Clara: Encontrado cargos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_cargos'] = amount
                                summary_data['total_retiros'] = amount
                    
                    # Saldo al corte
                    if not summary_data['saldo_final']:
                        match = re.search(r'Saldo\s+al\s+corte\s+([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Clara: Encontrado saldo al corte: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_final'] = amount
                    
                    # Total MXN [cargos] MXN [abonos] - el segundo monto es el total de abonos
                    # Ejemplo: "Total MXN 3,115.30 MXN -3,305.40"
                    if not summary_data['total_abonos']:
                        # Pattern: "Total MXN [amount1] MXN [amount2]"
                        # Try multiple flexible patterns to handle variations
                        total_patterns = [
                            r'Total\s+MXN\s+([\d,\.\-]+)\s+MXN\s+([\d,\.\-]+)',  # "Total MXN 3,115.30 MXN -3,305.40"
                            r'Total\s+([\d,\.\-]+)\s+([\d,\.\-]+)',  # "Total 3,115.30 -3,305.40" (fallback)
                        ]
                        for pattern in total_patterns:
                            match = re.search(pattern, line, re.I)
                            if match:
                                # First amount is cargos, second amount is abonos
                                cargos_amount = normalize_amount_str(match.group(1))
                                abonos_amount = normalize_amount_str(match.group(2))
                                
                                # Set total_cargos if not already set
                                if not summary_data['total_cargos'] and cargos_amount != 0:
                                    summary_data['total_cargos'] = abs(cargos_amount)  # Use absolute value
                                    summary_data['total_retiros'] = abs(cargos_amount)
                                
                                # Set total_abonos (can be negative)
                                if abonos_amount != 0:
                                    summary_data['total_abonos'] = abonos_amount
                                    summary_data['total_depositos'] = abonos_amount
                                    #print(f"✅ Clara: Encontrado total abonos: ${abonos_amount:,.2f} en línea {i+1}: {line[:80]}")
                                    break
            
            elif bank_name == "Konfio":
                # Konfio: Extract from page 2
                # "Pagos - $ 97,000.00" -> part of Total Abonos (positive)
                # "Devoluciones y ajustes - $ -115.66" -> part of Total Abonos (convert to positive)
                # "Compras y cargos $ 56,176.79" -> Total Cargos
                # "Saldo total al corte $ 312,227.05" -> Saldo Final
                pagos_amount = None
                devoluciones_amount = None
                subtotal_found = False  # Flag to track if Subtotal line was found
                
                for i, line in enumerate(all_lines):
                    # Subtotal line: "Subtotal $ X,XXX.XX $ Y,YYY.YY" (first is cargos, second is abonos)
                    # This takes priority over "Compras y cargos" and "Pagos + Devoluciones"
                    # Always check for Subtotal first, as it has the highest priority
                    if not subtotal_found:
                        match = re.search(r'Subtotal\s+\$\s*([\d,\.]+)\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            cargos_amount = normalize_amount_str(match.group(1))
                            abonos_amount = normalize_amount_str(match.group(2))
                            if cargos_amount > 0:
                                summary_data['total_cargos'] = cargos_amount
                                summary_data['total_retiros'] = cargos_amount
                            if abonos_amount > 0:
                                summary_data['total_abonos'] = abonos_amount
                                summary_data['total_depositos'] = abonos_amount
                            subtotal_found = True
                            continue  # Skip other patterns if Subtotal was found
                    
                    # Pagos (positive value) - only if Subtotal not found
                    if not subtotal_found and pagos_amount is None:
                        match = re.search(r'Pagos\s*-\s*\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            pagos_amount = normalize_amount_str(match.group(1))
                    
                    # Devoluciones y ajustes (negative value, convert to positive) - only if Subtotal not found
                    if not subtotal_found and devoluciones_amount is None:
                        match = re.search(r'Devoluciones\s+y\s+ajustes\s*-\s*\$\s*-?\s*([\d,\.]+)', line, re.I)
                        if match:
                            devoluciones_amount = normalize_amount_str(match.group(1))
                            # Convert to positive (take absolute value)
                            devoluciones_amount = abs(devoluciones_amount)
                    
                    # Compras y cargos (only if Subtotal not found)
                    if not subtotal_found and not summary_data['total_cargos']:
                        match = re.search(r'Compras\s+y\s+cargos\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                summary_data['total_cargos'] = amount
                                summary_data['total_retiros'] = amount
                    
                    # Saldo total al corte
                    if not summary_data['saldo_final']:
                        match = re.search(r'Saldo\s+total\s+al\s+corte\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                summary_data['saldo_final'] = amount
                
                # Calculate Total Abonos = Pagos + Devoluciones (both as positive)
                # Only if Subtotal was NOT found (Subtotal takes priority)
                if not subtotal_found and not summary_data['total_abonos']:
                    if pagos_amount is not None and devoluciones_amount is not None:
                        total_abonos = pagos_amount + devoluciones_amount
                        summary_data['total_abonos'] = total_abonos
                        summary_data['total_depositos'] = total_abonos
                    elif pagos_amount is not None:
                        # Only Pagos found
                        summary_data['total_abonos'] = pagos_amount
                        summary_data['total_depositos'] = pagos_amount
                elif subtotal_found:
                    pass  # Total Abonos already set from Subtotal
                elif devoluciones_amount is not None:
                    # Only Devoluciones found
                    summary_data['total_abonos'] = devoluciones_amount
                    summary_data['total_depositos'] = devoluciones_amount
                    print(f"✅ Konfio: Total Abonos (solo Devoluciones): ${devoluciones_amount:,.2f}")
                else:
                    print(f"⚠️ Konfio: No se encontraron Pagos ni Devoluciones")
            
            elif bank_name == "Base":
                # Base:
                # "Saldo al Corte $ 733,809.84" -> Saldo Final
                # "Depósitos/Abonos ( + ) $ 356,742.33" -> Total Abonos
                # "Retiros/Cargos ( - ) $ 102,609.46" -> Total Cargos
                #print(f"🔍 Buscando patrones Base en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    # Saldo al Corte -> Saldo Final
                    if not summary_data['saldo_final']:
                        match = re.search(r'Saldo\s+al\s+Corte\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Base: Encontrado saldo al corte: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_final'] = amount
                    
                    # Depósitos/Abonos -> Total Abonos
                    if not summary_data['total_abonos']:
                        match = re.search(r'Dep[oó]sitos\s*/\s*Abonos\s*\(\s*\+\s*\)\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Base: Encontrado depósitos/abonos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_abonos'] = amount
                                summary_data['total_depositos'] = amount
                    
                    # Retiros/Cargos -> Total Cargos
                    if not summary_data['total_cargos']:
                        match = re.search(r'Retiros\s*/\s*Cargos\s*\(\s*-\s*\)\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Base: Encontrado retiros/cargos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_cargos'] = amount
                                summary_data['total_retiros'] = amount
            
            elif bank_name == "Scotiabank":
                # Scotiabank:
                # "Saldo inicial $1,031,652.97"
                # "(+) Depósitos $35,461,511.04"
                # "(-) Retiros $33,018,203.16"
                # "(=) Saldo final de la cuenta $3,473,941.21"
                #print(f"🔍 Buscando patrones Scotiabank en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    # Saldo inicial
                    #if not summary_data['saldo_anterior']:
                     #   match = re.search(r'Saldo\s+inicial\s+\$\s*([\d,\.]+)', line, re.I)
                      #  if match:
                       #     amount = normalize_amount_str(match.group(1))
                        #    if amount > 0:
                                #print(f"✅ Scotiabank: Encontrado saldo inicial: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                         #       summary_data['saldo_anterior'] = amount
                    
                    # Depósitos
                    if not summary_data['total_depositos']:
                        match = re.search(r'\(\+\s*\)\s*Dep[oóOÓ]sitos\s*\$\s*([\d,\.]+)', line, re.I)
                        #print(f"✅ Scotiabank: Encontrado depósitos: {match}")
                        if match:
                            amount = normalize_amount_str(match.group(1))
                         #   print(f"✅ Scotiabank: Encontrado depósitos: {amount}")
                            if amount > 0:
                          #      print(f"✅ Scotiabank: Encontrado depósitos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_depositos'] = amount
                                summary_data['total_abonos'] = amount
                    
                    # Retiros
                    if not summary_data['total_retiros']:
                        match = re.search(r'\(-\s*\)\s+Retiros\s+\$\s*([\d,\.]+)', line, re.I)
                        #print(f"✅ Scotiabank: Encontrado retiros: {match}")
                        if not match:
                            # Try alternative pattern without space after minus
                            match = re.search(r'\(-\s*\)\s*Retiros\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Scotiabank: Encontrado retiros: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_retiros'] = amount
                                summary_data['total_cargos'] = amount
                    
                    # Saldo final
                    if not summary_data['saldo_final']:
                        # Try multiple patterns to handle variations
                        match = re.search(r'\(=\s*\)\s*Saldo\s+final\s+de\s+la\s+cuenta\s*\$\s*([\d,\.]+)', line, re.I)
                        if not match:
                            # Try with space before =
                            match = re.search(r'\(\s*=\s*\)\s*Saldo\s+final\s+de\s+la\s+cuenta\s*\$\s*([\d,\.]+)', line, re.I)
                        if not match:
                            # Try more flexible pattern
                            match = re.search(r'\(=\s*\)\s*Saldo.*?final.*?de.*?la.*?cuenta\s*\$\s*([\d,\.]+)', line, re.I)
                        # If not found in single line, try with next line (in case text is split)
                        if not match and i + 1 < len(all_lines):
                            combined_line = line + " " + all_lines[i + 1]
                            match = re.search(r'\(=\s*\)\s*Saldo\s+final\s+de\s+la\s+cuenta\s*\$\s*([\d,\.]+)', combined_line, re.I)
                            if not match:
                                match = re.search(r'\(\s*=\s*\)\s*Saldo\s+final\s+de\s+la\s+cuenta\s*\$\s*([\d,\.]+)', combined_line, re.I)
                            if not match:
                                match = re.search(r'\(=\s*\)\s*Saldo.*?final.*?de.*?la.*?cuenta\s*\$\s*([\d,\.]+)', combined_line, re.I)
                        #print(f"✅ Scotiabank: Encontrado Saldo Final: {match}")
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            #print(f"✅ Scotiabank: Encontrado Saldo Final: {amount}")
                            if amount > 0:
                                #print(f"✅ Scotiabank: Encontrado saldo final: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_final'] = amount
            
            else:
                # Generic patterns for other banks
                patterns = {
                    'depositos': [
                        r'(?:dep[oó]sitos?|abonos?)\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                        r'(?:\+\s*)?\d+\s+dep[oó]sitos?\s+([\d,\.\s]+)',
                        r'total\s+dep[oó]sitos?\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                    ],
                    'retiros': [
                        r'(?:retiros?|cargos?)\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                        r'(?:\-\s*)?\d+\s+retiros?\s+([\d,\.\s]+)',
                        r'total\s+retiros?\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                        r'total\s+cargos?\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                    ],
                    'abonos': [
                        r'total\s+abonos?\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                        r'abonos?\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                    ],
                    'cargos': [
                        r'total\s+cargos?\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                        r'cargos?\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                    ],
                    'saldo_final': [
                        r'saldo\s+(?:al|final|al\s+\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                        r'saldo\s+(?:final|total)\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                        r'saldo\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                    ],
                    'saldo_anterior': [
                        r'saldo\s+anterior\s*[:\-]?\s*\$?\s*([\d,\.\s]+)',
                    ],
                    'movimientos': [
                        r'total\s+de\s+movimientos?\s*[:\-]?\s*(\d+)',
                        r'(\d+)\s+movimientos?',
                    ]
                }
                
                # Search for patterns in all lines
                for line in all_lines:
                    # Look for depositos/abonos
                    if not summary_data['total_depositos'] and not summary_data['total_abonos']:
                        for pattern in patterns['depositos']:
                            match = re.search(pattern, line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    summary_data['total_depositos'] = amount
                                    summary_data['total_abonos'] = amount
                                    break
                    
                    # Look for retiros/cargos
                    if not summary_data['total_retiros'] and not summary_data['total_cargos']:
                        for pattern in patterns['retiros']:
                            match = re.search(pattern, line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    summary_data['total_retiros'] = amount
                                    summary_data['total_cargos'] = amount
                                    break
                    
                    # Look for abonos specifically
                    if not summary_data['total_abonos']:
                        for pattern in patterns['abonos']:
                            match = re.search(pattern, line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    summary_data['total_abonos'] = amount
                                    break
                    
                    # Look for cargos specifically
                    if not summary_data['total_cargos']:
                        for pattern in patterns['cargos']:
                            match = re.search(pattern, line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    summary_data['total_cargos'] = amount
                                    break
                    
                    # Look for saldo final
                    if not summary_data['saldo_final']:
                        for pattern in patterns['saldo_final']:
                            match = re.search(pattern, line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    summary_data['saldo_final'] = amount
                                    break
                    
                    # Look for saldo anterior
                    if not summary_data['saldo_anterior']:
                        for pattern in patterns['saldo_anterior']:
                            match = re.search(pattern, line, re.I)
                            if match:
                                amount = normalize_amount_str(match.group(1))
                                if amount > 0:
                                    summary_data['saldo_anterior'] = amount
                                    break
                    
                    # Look for total movimientos
                    if not summary_data['total_movimientos']:
                        for pattern in patterns['movimientos']:
                            match = re.search(pattern, line, re.I)
                            if match:
                                count = int(match.group(1))
                                if count > 0:
                                    summary_data['total_movimientos'] = count
                                    break
    
    except Exception as e:
        #print(f"⚠️  Error al extraer resumen del PDF: {e}")
        import traceback
        traceback.print_exc()
    
    # Print what was found
    found_items = [k for k, v in summary_data.items() if v is not None]
    if found_items:
        pass
        # print(f"✅ Valores encontrados en PDF: {', '.join(found_items)}")
    else:
        pass
        # print("⚠️  No se encontraron valores de resumen en el PDF")
    
    return summary_data


def calculate_extracted_totals(df_mov: pd.DataFrame, bank_name: str) -> dict:
    """
    Calculate totals from extracted movements DataFrame.
    Returns a dictionary with calculated totals.
    """
    totals = {
        'total_abonos': 0.0,
        'total_cargos': 0.0,
        'total_retiros': 0.0,
        'total_depositos': 0.0,
        'saldo_final': 0.0,
        'total_movimientos': len(df_mov)
    }
    
    # For Base, exclude "PAGO DE INTERES" from totals calculation
    # because it's not included in the summary totals on the PDF cover page
    df_for_totals = df_mov.copy()
    if bank_name == 'Base' and 'Descripción' in df_for_totals.columns:
        # Filter out rows where description contains "PAGO DE INTERES"
        df_for_totals = df_for_totals[
            ~df_for_totals['Descripción'].astype(str).str.contains('PAGO DE INTERES', case=False, na=False)
        ]
    
    # Calculate based on available columns
    if 'Abonos' in df_for_totals.columns:
        totals['total_abonos'] = df_for_totals['Abonos'].apply(normalize_amount_str).sum()
        totals['total_depositos'] = totals['total_abonos']
    
    if 'Cargos' in df_for_totals.columns:
        totals['total_cargos'] = df_for_totals['Cargos'].apply(normalize_amount_str).sum()
        totals['total_retiros'] = totals['total_cargos']
    
    # Get final balance (last row's saldo if available)
    # Use the last non-empty value from the "Saldo" column in Movements tab
    # This is called BEFORE adding the "Total" row, so we can safely get the last value
    # For BBVA, use the last value from "LIQUIDACIÓN" column instead of "Saldo"
    if bank_name == 'BBVA' and 'Liquidación' in df_mov.columns:
        liquidacion_col = df_mov['Liquidación']
        # Get last non-empty liquidacion value (iterate from end to beginning)
        for idx in range(len(liquidacion_col) - 1, -1, -1):
            val = liquidacion_col.iloc[idx]
            if val and pd.notna(val) and str(val).strip() and str(val).strip() != '':
                totals['saldo_final'] = normalize_amount_str(val)
                break
    elif 'Saldo' in df_mov.columns:
        saldo_col = df_mov['Saldo']
        # Get last non-empty saldo value (iterate from end to beginning)
        for idx in range(len(saldo_col) - 1, -1, -1):
            val = saldo_col.iloc[idx]
            if val and pd.notna(val) and str(val).strip() and str(val).strip() != '':
                totals['saldo_final'] = normalize_amount_str(val)
                #print(f"✅ Saldo final extraído de Movements: ${totals['saldo_final']:,.2f} (fila {idx+1} de {len(saldo_col)})")
                break
    
    return totals


def create_validation_sheet(pdf_summary: dict, extracted_totals: dict, has_saldo_column: bool = True) -> pd.DataFrame:
    """
    Create a validation DataFrame comparing PDF summary with extracted totals.
    
    Args:
        pdf_summary: Dictionary with summary data extracted from PDF
        extracted_totals: Dictionary with totals calculated from extracted movements
        has_saldo_column: Whether the bank has a "Saldo" column in Movements (default: True)
    """
    validation_data = []
    
    # Tolerance for floating point comparison (0.01 for cents)
    tolerance = 0.01
    
    # Compare Abonos/Depositos
    pdf_abonos = pdf_summary.get('total_abonos') or pdf_summary.get('total_depositos')
    ext_abonos = extracted_totals.get('total_abonos', 0.0)
    abonos_match = pdf_abonos is None or abs(pdf_abonos - ext_abonos) < tolerance
    validation_data.append({
        'Concepto': 'Total Abonos / Depósitos',
        'Valor en PDF': f"${pdf_abonos:,.2f}" if pdf_abonos else "No encontrado",
        'Valor Extraído': f"${ext_abonos:,.2f}",
        'Diferencia': f"${abs(pdf_abonos - ext_abonos):,.2f}" if pdf_abonos else "N/A",
        'Estado': '✓' if abonos_match else '✗'
    })
    
    # Compare Cargos/Retiros
    pdf_cargos = pdf_summary.get('total_cargos') or pdf_summary.get('total_retiros')
    ext_cargos = extracted_totals.get('total_cargos', 0.0)
    cargos_match = pdf_cargos is None or abs(pdf_cargos - ext_cargos) < tolerance
    validation_data.append({
        'Concepto': 'Total Cargos / Retiros',
        'Valor en PDF': f"${pdf_cargos:,.2f}" if pdf_cargos else "No encontrado",
        'Valor Extraído': f"${ext_cargos:,.2f}",
        'Diferencia': f"${abs(pdf_cargos - ext_cargos):,.2f}" if pdf_cargos else "N/A",
        'Estado': '✓' if cargos_match else '✗'
    })
    
    # Compare Saldo Final - only if bank has Saldo column in Movements
    saldo_match = True  # Initialize to True (will be set to actual value if has_saldo_column)
    if has_saldo_column:
        pdf_saldo = pdf_summary.get('saldo_final')
        ext_saldo = extracted_totals.get('saldo_final', 0.0)
        saldo_match = pdf_saldo is None or abs(pdf_saldo - ext_saldo) < tolerance
        validation_data.append({
            'Concepto': 'Saldo Final',
            'Valor en PDF': f"${pdf_saldo:,.2f}" if pdf_saldo else "No encontrado",
            'Valor Extraído': f"${ext_saldo:,.2f}",
            'Diferencia': f"${abs(pdf_saldo - ext_saldo):,.2f}" if pdf_saldo else "N/A",
            'Estado': '✓' if saldo_match else '✗'
        })
    
    # Compare Total Movimientos
    pdf_mov = pdf_summary.get('total_movimientos')
    ext_mov = extracted_totals.get('total_movimientos', 0)
    mov_match = pdf_mov is None or pdf_mov == ext_mov
    validation_data.append({
        'Concepto': 'Total de Movimientos',
        'Valor en PDF': str(pdf_mov) if pdf_mov else "No encontrado",
        'Valor Extraído': str(ext_mov),
        'Diferencia': str(abs(pdf_mov - ext_mov)) if pdf_mov else "N/A",
        'Estado': '✓' if mov_match else '✗'
    })
    
    # Overall status
    all_match = abonos_match and cargos_match and saldo_match and mov_match
    validation_data.append({
        'Concepto': 'VALIDACIÓN GENERAL',
        'Valor en PDF': '',
        'Valor Extraído': '',
        'Diferencia': '',
        'Estado': '✓ TODO CORRECTO' if all_match else '✗ HAY DISCREPANCIAS'
    })
    
    return pd.DataFrame(validation_data)


def print_validation_summary(pdf_summary: dict, extracted_totals: dict, validation_df: pd.DataFrame):
    """
    Print validation summary to console with checkmarks or X marks.
    """
    # print("\n" + "=" * 80)
    # print("📊 VALIDACIÓN DE DATOS")
    # print("=" * 80)
    
    # Check overall status
    overall_status = validation_df[validation_df['Concepto'] == 'VALIDACIÓN GENERAL']['Estado'].values[0]
    
    if '✓' in overall_status:
        print("✅ VALIDACIÓN: TODO CORRECTO")
    else:
        print("❌ VALIDACIÓN: HAY DISCREPANCIAS")
        pass
    
    # Print VALIDACIÓN GENERAL
    for _, row in validation_df.iterrows():
        if row['Concepto'] == 'VALIDACIÓN GENERAL':
            print(f"VALIDACIÓN GENERAL")
            break
    
    # print("\nComparación de Totales:")
    # print("-" * 80)
    
    # for _, row in validation_df.iterrows():
    #     if row['Concepto'] == 'VALIDACIÓN GENERAL':
    #         print(f"\n{row['Concepto']}: {row['Estado']}")
    #     else:
    #         status_icon = "✅" if row['Estado'] == '✓' else "❌"
    #         print(f"{status_icon} {row['Concepto']}")
    #         print(f"   PDF: {row['Valor en PDF']}")
    #         print(f"   Extraído: {row['Valor Extraído']}")
    #         if row['Diferencia'] != "N/A":
    #             print(f"   Diferencia: {row['Diferencia']}")
    
    # print("=" * 80 + "\n")


def extract_digitem_section(pdf_path: str, columns_config: dict) -> pd.DataFrame:
    """
    Extract DIGITEM section from Banamex PDF using the same coordinate-based extraction as Movements.
    Section starts with "DIGITEM" and ends with "TRANSFERENCIA ELECTRONICA DE FONDOS".
    Returns a DataFrame with columns: Fecha, Descripción, Importe
    """
    digitem_rows = []
    
    try:
        # Use the same extraction method as Movements
        extracted_data = extract_text_from_pdf(pdf_path)
        # Pattern for dates: supports multiple formats:
        # - "DIA MES" (01 ABR)
        # - "MES DIA" (ABR 01)
        # - "DIA MES AÑO" (06 mar 2023) - for Konfio
        date_pattern = re.compile(r"\b(?:(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4})\b", re.I)
        dec_amount_re = re.compile(r"\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")
        
        in_digitem_section = False
        extraction_stopped = False
        
        for page_data in extracted_data:
            if extraction_stopped:
                break
                
            page_num = page_data['page']
            words = page_data.get('words', [])
            text = page_data.get('content', '')
            
            if not words and not text:
                continue
            
            # DIGITEM section cannot be on the first page - skip page 1
            if page_num == 1:
                continue
            
            # Check if we're entering DIGITEM section (starting from page 2)
            if not in_digitem_section:
                # Check both text and words for "DIGITEM"
                if re.search(r'\bDIGITEM\b', text, re.I):
                    in_digitem_section = True
                    #print(f"📄 Sección DIGITEM encontrada en página {page_num}")
                    # Skip the header line "DETALLE DE OPERACIONES" that comes after DIGITEM
                    skip_next_line = True
                else:
                    # Also check in words (in case text extraction missed it)
                    all_words_text = ' '.join([w.get('text', '') for w in words])
                    if re.search(r'\bDIGITEM\b', all_words_text, re.I):
                        in_digitem_section = True
                        # print(f"📄 Sección DIGITEM encontrada en página {page_num} (desde words)")
                        skip_next_line = True
                    else:
                        continue
            else:
                skip_next_line = False
            
            # Extract rows using coordinate-based method (same as Movements)
            if in_digitem_section and words:
                # Group words by row
                word_rows = group_words_by_row(words)
                
                for row_words in word_rows:
                    if not row_words:
                        continue
                    
                    # Check if we're leaving DIGITEM section (check each row)
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    if re.search(r'TRANSFERENCIA\s+ELECTRONICA\s+DE\s+FONDOS', all_row_text, re.I):
                        # print(f"📄 Fin de sección DIGITEM encontrado en página {page_num}")
                        extraction_stopped = True
                        break
                    
                    # Skip header line "DETALLE DE OPERACIONES" that comes right after DIGITEM
                    if skip_next_line:
                        if re.search(r'DETALLE\s+DE\s+OPERACIONES', all_row_text, re.I):
                            # print(f"   ⏭️  Saltando línea de encabezado: DETALLE DE OPERACIONES")
                            skip_next_line = False
                            continue
                        else:
                            skip_next_line = False  # Reset if we didn't find it
                    
                    if extraction_stopped:
                        break
                    
                    # Extract structured row using coordinates (same as Movements)
                    # Pass date_pattern to enable date/description separation
                    row_data = extract_movement_row(row_words, columns_config, None, date_pattern)
                    
                    # Check if this row has a date (same logic as Movements)
                    fecha_val = str(row_data.get('fecha') or '')
                    has_date = bool(date_pattern.search(fecha_val))
                    
                    # Check if description contains "EMP" (required for DIGITEM rows)
                    desc_val = str(row_data.get('descripcion') or '')
                    has_emp = 'EMP' in desc_val.upper()
                    
                    # Debug: print row info for rows that might be DIGITEM
                    # if has_date or has_emp or any(row_data.get(k) for k in ['cargos', 'abonos', 'saldo'] if row_data.get(k)):
                    #     print(f"   🔍 Fila potencial DIGITEM - Fecha: {fecha_val[:20] if fecha_val else 'N/A'}, has_date: {has_date}, has_EMP: {has_emp}, Desc: {desc_val[:50] if desc_val else 'N/A'}")
                    
                    if has_date and has_emp:
                        # This is a valid DIGITEM row (has date and EMP in description)
                        row_data['page'] = page_num
                        digitem_rows.append(row_data)
                        # print(f"   ✅ Fila DIGITEM agregada (total: {len(digitem_rows)})")
                    else:
                        # Continuation row: append to previous DIGITEM row if exists and has EMP
                        if digitem_rows:
                            prev = digitem_rows[-1]
                            
                            # Check if previous row has EMP (to ensure we're continuing a DIGITEM row)
                            prev_desc = str(prev.get('descripcion') or '')
                            if 'EMP' in prev_desc.upper():
                                # Merge amounts from continuation row
                                cont_amounts = row_data.get('_amounts', [])
                                if cont_amounts and columns_config:
                                    # Get description range
                                    descripcion_range = None
                                    if 'descripcion' in columns_config:
                                        x0, x1 = columns_config['descripcion']
                                        descripcion_range = (x0, x1)
                                    
                                    # Get column ranges for numeric columns
                                    col_ranges = {}
                                    for col in ('cargos', 'abonos', 'saldo'):
                                        if col in columns_config:
                                            x0, x1 = columns_config[col]
                                            col_ranges[col] = (x0, x1)
                                    
                                    # Assign amounts from continuation row
                                    tolerance = 10
                                    for amt_text, center in cont_amounts:
                                        if descripcion_range and descripcion_range[0] <= center <= descripcion_range[1]:
                                            continue
                                        
                                        assigned = False
                                        for col in ('cargos', 'abonos', 'saldo'):
                                            if col in col_ranges:
                                                x0, x1 = col_ranges[col]
                                                if (x0 - tolerance) <= center <= (x1 + tolerance):
                                                    existing = prev.get(col) or ''
                                                    if not existing or amt_text not in existing:
                                                        if existing:
                                                            prev[col] = (existing + ' ' + amt_text).strip()
                                                        else:
                                                            prev[col] = amt_text
                                                    assigned = True
                                                    break
                                
                                # Merge amounts list
                                prev_amounts = prev.get('_amounts', [])
                                prev['_amounts'] = prev_amounts + cont_amounts
                                
                                # Collect text pieces from continuation row
                                cont_parts = []
                                for k in ('descripcion', 'fecha'):
                                    v = row_data.get(k)
                                    if v:
                                        cont_parts.append(str(v))
                                
                                cont_text = ' '.join(cont_parts)
                                cont_text = dec_amount_re.sub('', cont_text)
                                cont_text = ' '.join(cont_text.split()).strip()
                                
                                if cont_text:
                                    if prev.get('descripcion'):
                                        prev['descripcion'] = (prev.get('descripcion') or '') + ' ' + cont_text
                                    else:
                                        prev['descripcion'] = cont_text
        
        # Debug: print summary
        # print(f"🔍 Total de filas DIGITEM encontradas: {len(digitem_rows)}")
        # if digitem_rows:
        #     print(f"   📋 Primeras 3 filas:")
        #     for i, r in enumerate(digitem_rows[:3]):
        #         print(f"      Fila {i+1}: fecha={r.get('fecha', 'N/A')[:20]}, desc={str(r.get('descripcion', 'N/A'))[:50]}")
        
        # Process digitem_rows similar to how movements are processed
        if digitem_rows:
            # Reassign amounts to appropriate columns (same logic as Movements)
            col_centers = {}
            col_ranges = {}
            descripcion_range = None
            
            if 'descripcion' in columns_config:
                x0, x1 = columns_config['descripcion']
                descripcion_range = (x0, x1)
            
            for col in ('cargos', 'abonos', 'saldo'):
                if col in columns_config:
                    x0, x1 = columns_config[col]
                    col_centers[col] = (x0 + x1) / 2
                    col_ranges[col] = (x0, x1)
            
            for r in digitem_rows:
                amounts = r.get('_amounts', [])
                if not amounts:
                    continue
                
                for amt_text, center in amounts:
                    tolerance = 10
                    assigned = False
                    
                    # Check numeric columns first
                    for col in ('cargos', 'abonos', 'saldo'):
                        if col in col_ranges:
                            x0, x1 = col_ranges[col]
                            if (x0 - tolerance) <= center <= (x1 + tolerance):
                                existing = r.get(col, '').strip()
                                if not existing:
                                    r[col] = amt_text
                                    assigned = True
                                    break
                    
                    # Remove amount from description
                    if not assigned and descripcion_range:
                        if descripcion_range[0] <= center <= descripcion_range[1]:
                            continue
                
                # Remove amount tokens from descripcion
                if r.get('descripcion'):
                    r['descripcion'] = DEC_AMOUNT_RE.sub('', r.get('descripcion'))
                
                # Cleanup helper key
                if '_amounts' in r:
                    del r['_amounts']
            
            # Convert to DataFrame and format columns
            df_digitem = pd.DataFrame(digitem_rows)
            
            # Extract dates and format columns
            if 'fecha' in df_digitem.columns:
                dates = df_digitem['fecha'].astype(str).apply(lambda txt: _extract_two_dates(txt) if txt else (None, None))
                df_digitem['Fecha'] = dates.apply(lambda t: t[0])
            else:
                df_digitem['Fecha'] = ''
            
            # Build description
            def _build_digitem_description(row):
                parts = []
                if 'descripcion' in row and row.get('descripcion'):
                    parts.append(str(row.get('descripcion')))
                text = ' '.join(parts)
                text = dec_amount_re.sub('', text)
                text = ' '.join(text.split()).strip()
                return text if text else ''
            
            df_digitem['Descripción'] = df_digitem.apply(_build_digitem_description, axis=1)
            
            # Get Importe from cargos, abonos, or saldo (whichever has value)
            def _get_importe(row):
                for col in ['cargos', 'abonos', 'saldo']:
                    if col in row and row.get(col):
                        return str(row.get(col)).strip()
                return ''
            
            df_digitem['Importe'] = df_digitem.apply(_get_importe, axis=1)
            
            # Keep only needed columns
            df_digitem = df_digitem[['Fecha', 'Descripción', 'Importe']]
            
            #print(f"✅ Se extrajeron {len(df_digitem)} registros de DIGITEM del PDF")
            return df_digitem
        else:
            pass
            # print("ℹ️  No se encontró sección DIGITEM en el PDF")
            return pd.DataFrame(columns=['Fecha', 'Descripción', 'Importe'])
    
    except Exception as e:
        pass
        # print(f"⚠️  Error al extraer DIGITEM del PDF: {e}")
        # import traceback
        # traceback.print_exc()
        return pd.DataFrame(columns=['Fecha', 'Descripción', 'Importe'])


def _extract_two_dates(txt):
    """Helper function to extract dates (same as in main function)"""
    if not txt or not isinstance(txt, str):
        return (None, None)
    # Pattern for dates: supports multiple formats:
    # - "DIA MES" (01 ABR)
    # - "MES DIA" (ABR 01)
    # - "DIA MES AÑO" (06 mar 2023) - for Konfio
    date_pattern = re.compile(r"(?:(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4})", re.I)
    found = date_pattern.findall(txt)
    if not found:
        return (None, None)
    if len(found) == 1:
        return (found[0], None)
    return (found[0], found[1])


def extract_transferencia_section(pdf_path: str) -> pd.DataFrame:
    """
    Extract TRANSFERENCIA ELECTRONICA DE FONDOS section from Banamex PDF.
    Section starts with "TRANSFERENCIA ELECTRONICA DE FONDOS" and ends with "TOTALES:".
    Returns a DataFrame with columns: Fecha, Descripción, Importe, Comisiones, I.V.A, Total
    """
    transferencia_rows = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            in_transferencia_section = False
            
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if not text:
                    continue
                
                lines = text.split('\n')
                
                for line in lines:
                    line_clean = line.strip()
                    if not line_clean:
                        continue
                    
                    # Check if we're entering TRANSFERENCIA section
                    if re.search(r'TRANSFERENCIA\s+ELECTRONICA\s+DE\s+FONDOS', line_clean, re.I):
                        in_transferencia_section = True
                        #print(f"📄 Sección TRANSFERENCIA encontrada en página {page_num}")
                        continue
                    
                    # Check if we're leaving TRANSFERENCIA section
                    if in_transferencia_section and re.search(r'^TOTALES:', line_clean, re.I):
                        #print(f"📄 Fin de sección TRANSFERENCIA encontrado en página {page_num}")
                        break
                    
                    # Extract rows from TRANSFERENCIA section
                    if in_transferencia_section:
                        # Try to extract date (DD MMM format)
                        date_match = re.search(r'(\d{1,2})\s+([A-Z]{3})', line_clean)
                        if date_match:
                            fecha = f"{date_match.group(1)} {date_match.group(2)}"
                            
                            # Extract all amounts in the line
                            amounts = DEC_AMOUNT_RE.findall(line_clean)
                            
                            # Typically: Importe, Comisiones, I.V.A, Total
                            importe = amounts[0] if len(amounts) > 0 else ''
                            comisiones = amounts[1] if len(amounts) > 1 else ''
                            iva = amounts[2] if len(amounts) > 2 else ''
                            total = amounts[3] if len(amounts) > 3 else (amounts[-1] if len(amounts) > 0 else '')
                            
                            # Extract description (everything between date and first amount)
                            desc_start = date_match.end()
                            if amounts:
                                desc_end = line_clean.find(amounts[0])
                                descripcion = line_clean[desc_start:desc_end].strip()
                            else:
                                descripcion = line_clean[desc_start:].strip()
                            
                            if fecha and descripcion:
                                transferencia_rows.append({
                                    'Fecha': fecha,
                                    'Descripción': descripcion,
                                    'Importe': importe,
                                    'Comisiones': comisiones,
                                    'I.V.A': iva,
                                    'Total': total
                                })
                        else:
                            # Multi-line entry - append to previous row's description
                            if transferencia_rows and line_clean:
                                # Check if line has amounts
                                amounts = DEC_AMOUNT_RE.findall(line_clean)
                                if amounts:
                                    # Update amounts in previous row
                                    if len(amounts) >= 1 and not transferencia_rows[-1]['Importe']:
                                        transferencia_rows[-1]['Importe'] = amounts[0]
                                    if len(amounts) >= 2 and not transferencia_rows[-1]['Comisiones']:
                                        transferencia_rows[-1]['Comisiones'] = amounts[1]
                                    if len(amounts) >= 3 and not transferencia_rows[-1]['I.V.A']:
                                        transferencia_rows[-1]['I.V.A'] = amounts[2]
                                    if len(amounts) >= 4 and not transferencia_rows[-1]['Total']:
                                        transferencia_rows[-1]['Total'] = amounts[3]
                                else:
                                    # Just description continuation
                                    if transferencia_rows[-1]['Descripción']:
                                        transferencia_rows[-1]['Descripción'] += ' ' + line_clean
                                    else:
                                        transferencia_rows[-1]['Descripción'] = line_clean
        
        if transferencia_rows:
            df_transferencia = pd.DataFrame(transferencia_rows)
            #print(f"✅ Se extrajeron {len(df_transferencia)} registros de TRANSFERENCIA del PDF")
            return df_transferencia
        else:
            pass
            # print("ℹ️  No se encontró sección TRANSFERENCIA en el PDF")
            return pd.DataFrame(columns=['Fecha', 'Descripción', 'Importe', 'Comisiones', 'I.V.A', 'Total'])
    
    except Exception as e:
        #print(f"⚠️  Error al extraer TRANSFERENCIA del PDF: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame(columns=['Fecha', 'Descripción', 'Importe', 'Comisiones', 'I.V.A', 'Total'])


def extract_text_from_pdf(pdf_path: str) -> list:
    """
    Extract text and word positions from each page of a PDF.
    Returns a list of dictionaries (page_number, text, words).
    
    Si detecta texto ilegible (caracteres CID), usa Tesseract OCR como fallback.
    """
    # PASO 1: Detectar si el PDF tiene texto ilegible
    is_illegible, cid_ratio, ascii_ratio = is_pdf_text_illegible(pdf_path)
    
    if is_illegible and TESSERACT_AVAILABLE:
        print(f"[INFO] PDF detectado como ilegible (CID ratio: {cid_ratio:.2%}, ASCII ratio: {ascii_ratio:.2%})")
        print(f"[INFO] Usando Tesseract OCR local como fallback...")
        print(f"[INFO] El banco se detectará después de procesar con OCR...")
        try:
            # Usar Tesseract OCR
            extracted_data = extract_text_with_tesseract_ocr(pdf_path)
            # Marcar que se usó OCR
            for page_data in extracted_data:
                page_data['_used_ocr'] = True
            return extracted_data
        except Exception as e:
            print(f"[ADVERTENCIA] Error con Tesseract OCR: {e}")
            print(f"[INFO] Continuando con extracción normal (puede tener caracteres ilegibles)...")
            # Continuar con extracción normal si OCR falla
    
    # PASO 2: Extracción normal con pdfplumber (código existente - SIN CAMBIOS)
    extracted_data = []
    
    # Detect bank to apply Konfio-specific fixes
    detected_bank = detect_bank_from_pdf(pdf_path)
    is_konfio = (detected_bank == "Konfio")

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            # also extract words with positions for coordinate-based column detection
            try:
                words = page.extract_words(x_tolerance=3, y_tolerance=3)
            except Exception:
                words = []
            
            # For Konfio, fix duplicated characters in text and words
            if is_konfio:
                if text:
                    text = fix_duplicated_chars(text)
                # Fix duplicated characters in word texts
                for word in words:
                    if 'text' in word and word['text']:
                        word['text'] = fix_duplicated_chars(word['text'])
            
            extracted_data.append({
                "page": page_number,
                "content": text if text else "",
                "words": words,
                "_used_ocr": False  # Flag para indicar que NO viene de OCR
            })

    return extracted_data


def export_to_excel(data: list, output_path: str):
    """
    Export extracted PDF content to an Excel file.
    """
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)
    print(f"✅ Excel file created -> {output_path}")


def split_pages_into_lines(pages: list) -> list:
    """Return list of page dicts with lines: [{'page': n, 'lines': [...]}, ...]"""
    pages_lines = []
    for p in pages:
        content = (p.get('content') or '')
        # normalize NBSPs
        content = content.replace('\u00A0', ' ').replace('\u202F', ' ')
        lines = [" ".join(l.split()) for l in content.splitlines() if l and l.strip()]
        pages_lines.append({'page': p.get('page'), 'lines': lines})
    return pages_lines


def group_entries_from_lines(lines):
    """Group lines into transaction entries: a line starting with a date begins a new entry."""
    # Pattern for dates: supports multiple formats including "DIA MES AÑO" (06 mar 2023)
    day_re = re.compile(r"\b(?:(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4})\b", re.I)
    entries = []
    for line in lines:
        if day_re.search(line):
            entries.append(line)
        else:
            if entries:
                entries[-1] = entries[-1] + ' ' + line
            else:
                entries.append(line)
    return entries


def group_words_by_row(words, y_tolerance=5):
    """Group words by Y-coordinate (rows) to extract table rows."""
    if not words:
        return []
    
    # Sort by page first (if available), then by top coordinate
    # This ensures words from page 1 come before words from page 2, etc.
    sorted_words = sorted(words, key=lambda w: (w.get('page', 0), w.get('top', 0)))
    
    rows = []
    current_row = []
    current_y = None
    current_page = None
    
    for word in sorted_words:
        word_y = word.get('top', 0)
        word_page = word.get('page', 0)
        
        if current_y is None:
            current_y = word_y
            current_page = word_page
        
        # If word is on the same page and within y_tolerance of current row, add it
        if word_page == current_page and abs(word_y - current_y) <= y_tolerance:
            current_row.append(word)
        else:
            # Start a new row (either new page or different Y position)
            if current_row:
                rows.append(current_row)
            current_row = [word]
            current_y = word_y
            current_page = word_page
    
    # Don't forget the last row
    if current_row:
        rows.append(current_row)
    
    return rows


def assign_word_to_column(word_x0, word_x1, columns):
    """Assign a word (with x0, x1 coordinates) to a column based on X-ranges.
    Returns column name or None if not in any range.
    Prioritizes numeric columns (cargos, abonos, saldo) over description when there's overlap.
    When a word falls in multiple overlapping ranges, assigns to the range whose center is closest.
    """
    word_center = (word_x0 + word_x1) / 2
    
    # First, check numeric columns (cargos, abonos, saldo) to prioritize them
    # This fixes the issue where cargos (360-398) overlaps with descripcion (160-400)
    numeric_cols = ['cargos', 'abonos', 'saldo']
    
    # Collect all matching columns and their distances to center
    matching_cols = []
    for col_name in numeric_cols:
        if col_name in columns:
            x_min, x_max = columns[col_name]
            
            # Validar y corregir rangos invertidos
            if x_min > x_max:
                print(f"[ADVERTENCIA] Rango invertido para '{col_name}': ({x_min}, {x_max}). Corrigiendo a ({x_max}, {x_min})")
                x_min, x_max = x_max, x_min  # Intercambiar valores
            
            if x_min <= word_center <= x_max:
                # Calcular distancia del centro del monto al centro del rango
                range_center = (x_min + x_max) / 2
                distance = abs(word_center - range_center)
                matching_cols.append((col_name, distance, range_center))
    
    # Si hay múltiples coincidencias, retornar la más cercana al centro del rango
    if matching_cols:
        # Ordenar por distancia (más cercana primero)
        matching_cols.sort(key=lambda x: x[1])
        return matching_cols[0][0]  # Retornar la columna más cercana
    
    # Then check other columns (fecha, liq, descripcion, etc.)
    for col_name, (x_min, x_max) in columns.items():
        if col_name not in numeric_cols:  # Skip numeric cols, already checked
            # Validar y corregir rangos invertidos
            if x_min > x_max:
                x_min, x_max = x_max, x_min  # Intercambiar valores
            
            if x_min <= word_center <= x_max:
                return col_name
    
    return None


def is_transaction_row(row_data, bank_name=None):
    """Check if a row is an actual bank transaction (not a header or empty row).
    A transaction must have:
    - A date in 'fecha' column
    - At least one amount in cargos, abonos, or saldo
    """
    fecha = (row_data.get('fecha') or '').strip()
    cargos = (row_data.get('cargos') or '').strip()
    abonos = (row_data.get('abonos') or '').strip()
    saldo = (row_data.get('saldo') or '').strip()
    
    # Must have a date matching DD/MMM pattern
    # For HSBC: date is only 2 digits (01-31)
    # For Banregio: date is only 2 digits (01-31)
    # For Base: date format is DD/MM/YYYY (e.g., "30/04/2024")
    # For other banks: supports both "DIA MES" (01 ABR) and "MES DIA" (ABR 01) formats
    # Pattern for dates: supports multiple formats including "DIA MES AÑO" (06 mar 2023)
    has_date = False
    if bank_name == 'HSBC':
        # For HSBC, date is only 2 digits (01-31)
        hsbc_date_re = re.compile(r'^(0[1-9]|[12][0-9]|3[01])$')
        has_date = bool(hsbc_date_re.match(fecha))
    elif bank_name == 'Banregio':
        # For Banregio, date is only 2 digits (01-31)
        banregio_date_re = re.compile(r'^(0[1-9]|[12][0-9]|3[01])$')
        has_date = bool(banregio_date_re.match(fecha))
        print("Debug: has_date", has_date, banregio_date_re)
    elif bank_name == 'Base':
        # For Base, date format is DD/MM/YYYY (e.g., "30/04/2024")
        base_date_re = re.compile(r'^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(\d{4})$')
        has_date = bool(base_date_re.match(fecha))
    else:
        # For other banks, use the general date pattern
        day_re = re.compile(r"\b(?:(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4})\b", re.I)
        has_date = bool(day_re.search(fecha))
    
    # Must have at least one numeric amount
    has_amount = bool(cargos or abonos or saldo)
    
    return has_date and has_amount


def extract_movement_row(words, columns, bank_name=None, date_pattern=None):
    """Extract a structured movement row from grouped words using coordinate-based column assignment."""
    row_data = {col: '' for col in columns.keys()}
    amounts = []
    # Show detailed debug for Banamex (disabled)
    show_detailed_debug = False
    
    # Pattern to detect dates (for separating date from description)
    if date_pattern is None:
        if bank_name == 'HSBC':
            # For HSBC, date is only 2 digits (01-31) at the start of the text
            # Pattern should match "03" or "03 RETIRO ..." but not "003" or "30"
            date_pattern = re.compile(r"^(0[1-9]|[12][0-9]|3[01])(?=\s|$)")
        elif bank_name == 'Banregio':
            # For Banregio, date is only 2 digits (01-31) at the start of the text
            # Pattern should match "04" or "04 TRA ..." but not "004" or "40"
            date_pattern = re.compile(r"^(0[1-9]|[12][0-9]|3[01])(?=\s|$)")
        elif bank_name == 'Base':
            # For Base, date format is DD/MM/YYYY (e.g., "30/04/2024")
            date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(\d{4})\b')
        else:
            date_pattern = re.compile(r"\b(?:(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4})\b", re.I)
    
    # Sort words by X coordinate within the row
    sorted_words = sorted(words, key=lambda w: w.get('x0', 0))
    
    # For HSBC, first try to extract date from fecha column range
    if bank_name == 'HSBC' and 'fecha' in columns:
        fecha_x0, fecha_x1 = columns['fecha']
        # Collect all words that are in the fecha column range
        fecha_words = []
        for word in sorted_words:
            x0 = word.get('x0', 0)
            x1 = word.get('x1', 0)
            center = (x0 + x1) / 2
            if fecha_x0 <= center <= fecha_x1:
                fecha_words.append(word)
        
        if len(fecha_words) > 0:
            # Get all text from fecha words, preserving order
            fecha_texts = [w.get('text', '').strip() for w in fecha_words]
            fecha_text = ' '.join(fecha_texts)
            
            # For HSBC, date is only 2 digits (01-31)
            hsbc_date_match = date_pattern.search(fecha_text)
            if hsbc_date_match:
                date_text = hsbc_date_match.group(1)  # Just the 2 digits (01-31)
                row_data['fecha'] = date_text
                # Remove these words from sorted_words so they're not processed again
                sorted_words = [w for w in sorted_words if w not in fecha_words]
    
    # For Konfio, first try to reconstruct dates that might be split across multiple words
    # Konfio dates can be split like "31" and "mar 2023" in separate words
    if bank_name == 'Konfio' and 'fecha' in columns:
        fecha_x0, fecha_x1 = columns['fecha']
        # Collect all words that are in the fecha column range
        fecha_words = []
        for word in sorted_words:
            x0 = word.get('x0', 0)
            x1 = word.get('x1', 0)
            center = (x0 + x1) / 2
            if fecha_x0 <= center <= fecha_x1:
                fecha_words.append(word)
        
        # If we have words in fecha column, try to reconstruct the date
        if len(fecha_words) > 0:
            # Get all text from fecha words, preserving order
            fecha_texts = [w.get('text', '').strip() for w in fecha_words]
            fecha_text = ' '.join(fecha_texts)
            
            # For Konfio, dates can be split like:
            # 1. "31" and "mar 2023" in separate words -> "31 mar 2023"
            # 2. "31 mar 2023" in one word -> "31 mar 2023"
            # 3. "31 mar" and "2023" in separate words -> "31 mar 2023"
            
            # Try to match full date pattern first
            konfio_date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4}\b', re.I)
            full_match = konfio_date_pattern.search(fecha_text)
            if full_match:
                date_text = full_match.group()
                row_data['fecha'] = date_text
                # Remove these words from sorted_words so they're not processed again
                sorted_words = [w for w in sorted_words if w not in fecha_words]
            else:
                # Try to reconstruct: look for day (1-2 digits) + month (3 letters) + year (2-4 digits)
                # Pattern: day might be separate, month and year might be together or separate
                day_match = re.search(r'\b(0[1-9]|[12][0-9]|3[01])\b', fecha_text)
                month_year_match = re.search(r'\b([A-Za-z]{3})\s+(\d{2,4})\b', fecha_text, re.I)
                if day_match and month_year_match:
                    day = day_match.group(1)
                    month = month_year_match.group(1)
                    year = month_year_match.group(2)
                    date_text = f"{day} {month} {year}"
                    row_data['fecha'] = date_text
                    # Remove these words from sorted_words so they're not processed again
                    sorted_words = [w for w in sorted_words if w not in fecha_words]
    
    # For Clara, first try to reconstruct dates that might be split across multiple words
    if bank_name == 'Clara' and 'fecha' in columns:
        fecha_x0, fecha_x1 = columns['fecha']
        # Collect all words that are in the fecha column range
        fecha_words = []
        for word in sorted_words:
            x0 = word.get('x0', 0)
            x1 = word.get('x1', 0)
            center = (x0 + x1) / 2
            if fecha_x0 <= center <= fecha_x1:
                fecha_words.append(word)
        
        # If we have words in fecha column, try to reconstruct the date
        if len(fecha_words) > 0:
            # Get all text from fecha words, preserving order
            fecha_texts = [w.get('text', '').strip() for w in fecha_words]
            fecha_text = ' '.join(fecha_texts)
            
            # For Clara, dates can be split in multiple ways:
            # 1. "0 2 E N E" -> "02 ENE" (digits and letters all separated)
            # 2. "02 E N E" -> "02 ENE" (day together, month letters separated)
            # 3. "02 ENE" -> "02 ENE" (standard format)
            
            # First, try to reconstruct: join consecutive digits and consecutive letters
            reconstructed_parts = []
            current_part = ''
            for text in fecha_texts:
                if text.isdigit():
                    # Digit: append to current part if it's also digits, otherwise start new
                    if current_part.isdigit() or not current_part:
                        current_part += text
                    else:
                        reconstructed_parts.append(current_part)
                        current_part = text
                elif text.isalpha() and len(text) == 1:
                    # Single letter: append to current part if it's also letters, otherwise start new
                    if current_part.isalpha() or not current_part:
                        current_part += text
                    else:
                        reconstructed_parts.append(current_part)
                        current_part = text
                else:
                    # Other text: finish current part and add this
                    if current_part:
                        reconstructed_parts.append(current_part)
                    current_part = text
            
            # Add the last part
            if current_part:
                reconstructed_parts.append(current_part)
            
            # Now try to match date patterns with reconstructed parts
            reconstructed_text = ' '.join(reconstructed_parts)
            
            # Try pattern: day (1-2 digits) followed by month (3 letters, possibly separated)
            # Pattern 1: "02 E N E" or "0 2 E N E" -> "02 ENE"
            clara_date_match = re.search(r'(\d{1,2})\s+([A-Z])\s*([A-Z])\s*([A-Z])', reconstructed_text, re.I)
            if clara_date_match:
                day = clara_date_match.group(1)
                month_letters = clara_date_match.group(2) + clara_date_match.group(3) + clara_date_match.group(4)
                date_text = f"{day} {month_letters.upper()}"
                row_data['fecha'] = date_text
                # Remove these words from sorted_words so they're not processed again
                sorted_words = [w for w in sorted_words if w not in fecha_words]
            else:
                # Try pattern: "02 ENE" (standard format)
                clara_date_match = re.search(r'(\d{1,2}\s+[A-Z]{3})', reconstructed_text, re.I)
                if clara_date_match:
                    date_text = clara_date_match.group(1).strip()
                    row_data['fecha'] = date_text
                    # Remove these words from sorted_words so they're not processed again
                    sorted_words = [w for w in sorted_words if w not in fecha_words]
                elif len(reconstructed_parts) >= 2:
                    # Fallback: if we have at least 2 parts, try to construct date manually
                    # First part should be digits (day), rest should be letters (month)
                    day_part = reconstructed_parts[0]
                    month_part = ''.join(reconstructed_parts[1:]) if len(reconstructed_parts) > 1 else ''
                    
                    if day_part.isdigit() and len(day_part) <= 2 and month_part.isalpha() and len(month_part) == 3:
                        date_text = f"{day_part} {month_part.upper()}"
                        row_data['fecha'] = date_text
                        # Remove these words from sorted_words so they're not processed again
                        sorted_words = [w for w in sorted_words if w not in fecha_words]
    
    for word in sorted_words:
        text = word.get('text', '')
        x0 = word.get('x0', 0)
        x1 = word.get('x1', 0)
        center = (x0 + x1) / 2


        # Aplicar corrección de errores OCR (solo para HSBC)
        if bank_name == 'HSBC' and columns:
            original_text = text
            text = fix_ocr_amount_errors(text, x0, columns, bank_name)
            # Actualizar el texto en el diccionario word para que se use en descripciones
            if text != original_text:
                word['text'] = text

        # Check if word is in description range BEFORE detecting amounts
        # For Banamex: if word is in description range, don't detect amounts within it
        # This prevents values like "865.60" from "IVA $2865.60" being extracted as separate amounts
        word_in_description_range = False
        if bank_name == 'Banamex' and 'descripcion' in columns:
            desc_x0, desc_x1 = columns['descripcion']
            if desc_x0 <= center <= desc_x1:
                word_in_description_range = True
        
        # detect amount tokens inside the word
        # For Konfio, only detect amounts that have currency format ($) or decimal format (with .XX or ,XX)
        # This prevents account numbers like "3817" from being detected as amounts
        if bank_name == 'Konfio':
            # Pattern for Konfio: must have $ or decimal part (.,XX or .,XX)
            konfio_amount_pattern = re.compile(r'\$\s*\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})|\d{1,3}(?:[\.,\s]\d{3})*[\.,]\d{2}')
            m = konfio_amount_pattern.search(text)
        else:
            # Para HSBC, buscar todos los montos en la palabra (puede haber múltiples)
            if bank_name == 'HSBC':
                # Buscar todos los montos (con o sin $) usando findall para capturar múltiples
                # But skip if word is in description range (for Banamex logic, though HSBC doesn't use this)
                if not word_in_description_range:
                    amount_matches = DEC_AMOUNT_RE.findall(text)
                    if amount_matches:
                        for amt_match in amount_matches:
                            amounts.append((amt_match, center))
            else:
                # For Banamex: don't detect amounts if word is in description range
                if not (bank_name == 'Banamex' and word_in_description_range):
                    m = DEC_AMOUNT_RE.search(text)
                    if m:
                        amounts.append((m.group(), center))

        # Check if word contains a date followed by description text
        # For Banregio: date is only 2 digits at the start, e.g., "04 TRA ..."
        # For Banorte: date format is "DIA-MES-AÑO", e.g., "12-ENE-23EST EPIGMENIO"
        date_match = date_pattern.search(text)
        #print("Debug: date_match", date_match, "text", text)
        if date_match and 'fecha' in columns and 'descripcion' in columns:
            date_text = date_match.group()
            date_end_pos = date_match.end()
            # For HSBC, ensure we split the 2-digit date from the rest of the text
            if bank_name == 'HSBC':
                # The date should be exactly 2 digits (01-31) at the start
                # Extract the 2 digits and everything after as description
                hsbc_date_match = re.match(r'^(0[1-9]|[12][0-9]|3[01])(\s+.*)?', text)
                if hsbc_date_match:
                    date_text = hsbc_date_match.group(1)  # Just the 2 digits (01-31)
                    description_text = hsbc_date_match.group(2)  # Everything after (with leading space)
                    if description_text:
                        description_text = description_text.strip()  # Remove leading space
                    else:
                        description_text = ''
                    
                    # Assign date to fecha column (only if not already assigned)
                    if not row_data.get('fecha'):
                        row_data['fecha'] = date_text
                    
                    # Assign description to descripcion column
                    if description_text:
                        if row_data['descripcion']:
                            row_data['descripcion'] += ' ' + description_text
                        else:
                            row_data['descripcion'] = description_text
            # For Banregio, ensure we split the 2-digit date from the rest of the text
            elif bank_name == 'Banregio':
                # The date should be exactly 2 digits (01-31) at the start
                # Extract the 2 digits and everything after as description
                banregio_date_match = re.match(r'^(0[1-9]|[12][0-9]|3[01])(\s+.*)?', text)
                if banregio_date_match:
                    date_text = banregio_date_match.group(1)  # Just the 2 digits (01-31)
                    description_text = banregio_date_match.group(2)  # Everything after (with leading space)
                    if description_text:
                        description_text = description_text.strip()  # Remove leading space
                    else:
                        description_text = ''
                    
                    # Assign date to fecha column
                    # For Banregio, always replace (don't concatenate) since date is only 2 digits
                    row_data['fecha'] = date_text
                    
                    # Assign description to descripcion column
                    if description_text:
                        if row_data['descripcion']:
                            row_data['descripcion'] += ' ' + description_text
                        else:
                            row_data['descripcion'] = description_text
                    
                    continue  # Skip normal assignment for this word
            
            # For Banorte format "DIA-MES-AÑO", check if the date pattern captured the full date
            # Sometimes the pattern might only capture "30-ENE" and miss "-23"
            # Try to find a more complete date match by looking for the full pattern
            if bank_name == 'Banorte':
                # Pattern specifically for Banorte: DIA-MES-AÑO (e.g., "30-ENE-23")
                banorte_date_pattern = re.compile(r'(\d{1,2}-[A-Z]{3}-\d{2,4})', re.I)
                banorte_match = banorte_date_pattern.search(text)
                if banorte_match:
                    date_text = banorte_match.group(1)  # Full date including year
                    date_end_pos = banorte_match.end()
            
            # If there's text after the date, split it
            if date_end_pos < len(text):
                description_text = text[date_end_pos:].strip()
                
                # Remove leading hyphen if it's part of the date (e.g., "-23" from "30-ENE-23")
                # This happens when the date pattern didn't capture the full date
                if description_text.startswith('-') and len(description_text) > 1:
                    # Check if the next characters are digits (likely part of year)
                    next_chars = description_text[1:4]  # Check up to 3 digits
                    if next_chars.isdigit():
                        # This is likely part of the date, not description
                        # Try to reconstruct the full date
                        potential_year = description_text[1:3] if len(description_text) > 3 else description_text[1:]
                        if potential_year.isdigit():
                            # Check if we can find a date pattern that includes this
                            # Don't use word boundaries since the date might be at the start of text
                            full_date_pattern = re.compile(r'(\d{1,2}-[A-Z]{3}-\d{2,4})', re.I)
                            full_match = full_date_pattern.search(text)
                            if full_match:
                                date_text = full_match.group(1)
                                date_end_pos = full_match.end()
                                description_text = text[date_end_pos:].strip()
                
                # Check which column this word's center belongs to
                fecha_col_center = None
                descripcion_col_center = None
                if 'fecha' in columns:
                    fecha_x0, fecha_x1 = columns['fecha']
                    fecha_col_center = (fecha_x0 + fecha_x1) / 2
                if 'descripcion' in columns:
                    descripcion_x0, descripcion_x1 = columns['descripcion']
                    descripcion_col_center = (descripcion_x0 + descripcion_x1) / 2
                
                # Always split: assign date to fecha column and description to descripcion column
                if row_data['fecha']:
                    row_data['fecha'] += ' ' + date_text
                else:
                    row_data['fecha'] = date_text
                
                if description_text:
                    if row_data['descripcion']:
                        row_data['descripcion'] += ' ' + description_text
                    else:
                        row_data['descripcion'] = description_text
                
                continue  # Skip normal assignment for this word
        
        # Normal column assignment
        col_name = assign_word_to_column(x0, x1, columns)
        if col_name:
            # For Banamex, prevent non-amount text from being assigned to cargos/abonos/saldo
            # Only assign to these columns if the word is actually an amount
            if bank_name == 'Banamex' and col_name in ('cargos', 'abonos', 'saldo'):
                # Check if this is actually an amount (matches DEC_AMOUNT_RE pattern)
                is_amount = bool(DEC_AMOUNT_RE.search(text))
                if not is_amount:
                    # This is not an amount, assign to description instead
                    if row_data.get('descripcion'):
                        row_data['descripcion'] += ' ' + text
                    else:
                        row_data['descripcion'] = text
                else:
                    # It's an amount, assign normally
                    if row_data[col_name]:
                        row_data[col_name] += ' ' + text
                    else:
                        row_data[col_name] = text
            # For Konfio, prevent account numbers (like "3817") from being assigned to cargos/abonos
            # Account numbers are typically 4-digit numbers without $ or decimal format
            elif bank_name == 'Konfio' and col_name in ('cargos', 'abonos'):
                # Check if this looks like an account number (4 digits, no $, no decimal part)
                account_number_pattern = re.compile(r'^\d{4}$')
                if account_number_pattern.match(text.strip()):
                    # This is likely an account number, assign to description instead
                    if row_data.get('descripcion'):
                        row_data['descripcion'] += ' ' + text
                    else:
                        row_data['descripcion'] = text
                else:
                    # Normal assignment
                    if row_data[col_name]:
                        row_data[col_name] += ' ' + text
                    else:
                        row_data[col_name] = text
            else:
                # Normal assignment for other banks or other columns
                if row_data[col_name]:
                    row_data[col_name] += ' ' + text
                else:
                    row_data[col_name] = text
                
        # For Konfio, if word doesn't match any column but is in description area, add it to description
        elif bank_name == 'Konfio' and 'descripcion' in columns:
            desc_x0, desc_x1 = columns['descripcion']
            word_center = (x0 + x1) / 2
            # If word is between fecha and cargos columns, it's likely part of description
            # Also check if it's slightly outside the descripcion range but still in the middle area
            if 'fecha' in columns and 'cargos' in columns:
                fecha_x0, fecha_x1 = columns['fecha']
                cargos_x0, cargos_x1 = columns['cargos']
                # If word is after fecha and before cargos, it's likely description
                if word_center > fecha_x1 and word_center < cargos_x0:
                    # Skip "FÍSICA" and "DIGITAL" words
                    if text.strip().upper() not in ('FÍSICA', 'DIGITAL'):
                        if row_data.get('descripcion'):
                            row_data['descripcion'] += ' ' + text
                        else:
                            row_data['descripcion'] = text

    # attach detected amounts for later disambiguation
    row_data['_amounts'] = amounts
    
    # Debug for Banamex: print final state before returning
    if bank_name == 'Banamex':
        # Check if footer pattern is in any column
        banamex_footer_pattern = re.compile(r'\d+\.\w+\.OD\.\d+\.\d+', re.I)
        all_row_text = ' '.join([
            str(row_data.get('fecha', '')),
            str(row_data.get('descripcion', '')),
            str(row_data.get('cargos', '')),
            str(row_data.get('abonos', '')),
            str(row_data.get('saldo', ''))
        ])
    
    return row_data


def split_row_if_multiple_movements(row_words, columns_config, date_pattern, bank_name=None):
    """
    Detect if a row contains multiple movements and split it into separate rows.
    Returns a list of row_word lists (each representing one movement).
    """
    if not row_words or not columns_config:
        return [row_words]
    
    # Find fecha column range if available
    fecha_range = None
    if 'fecha' in columns_config:
        fecha_x0, fecha_x1 = columns_config['fecha']
        fecha_range = (fecha_x0, fecha_x1)
    
    # Pattern to exclude hours (like "17:47:53") from being detected as dates
    # Hours have format HH:MM:SS or HH:MM
    hour_pattern = re.compile(r'\b\d{1,2}:\d{2}(?::\d{2})?\b')
    
    # Find all words that contain dates (either in fecha column or anywhere if no fecha column)
    date_words = []
    for word in row_words:
        text = word.get('text', '')
        x0 = word.get('x0', 0)
        x1 = word.get('x1', 0)
        center = (x0 + x1) / 2
        
        # Skip if this looks like a time (hour:minute:second)
        if hour_pattern.search(text):
            continue
        
        # Check if word contains a date pattern
        # For Banorte, also check if the word contains multiple dates (like "17-ENE-2317-ENE-23")
        date_matches = date_pattern.findall(text)
        if date_matches:
            # If fecha column is defined, only consider dates in that column
            if fecha_range:
                if fecha_range[0] <= center <= fecha_range[1]:
                    # If multiple dates found in same word, create separate entries for each
                    if len(date_matches) > 1:
                        for i, date_match in enumerate(date_matches):
                            # Create a virtual word entry for each date
                            date_words.append({
                                'text': date_match,
                                'x0': x0,
                                'x1': x1,
                                'top': word.get('top', 0) + (i * 0.1),  # Slight offset to distinguish
                                'original_word': word
                            })
                    else:
                        date_words.append(word)
            else:
                # No fecha column defined, consider any date
                if len(date_matches) > 1:
                    for i, date_match in enumerate(date_matches):
                        date_words.append({
                            'text': date_match,
                            'x0': x0,
                            'x1': x1,
                            'top': word.get('top', 0) + (i * 0.1),
                            'original_word': word
                        })
                else:
                    date_words.append(word)
    
    # If we found multiple dates, split the row
    if len(date_words) > 1:
        # Sort date words by Y coordinate
        date_words.sort(key=lambda w: w.get('top', 0))
        
        # Get Y positions of all dates (except first)
        split_y_positions = [w.get('top', 0) for w in date_words[1:]]
        
        # Split row_words into multiple rows based on date positions
        split_rows = []
        current_split = []
        
        # Sort all words by Y coordinate
        sorted_words = sorted(row_words, key=lambda w: w.get('top', 0))
        
        for word in sorted_words:
            word_y = word.get('top', 0)
            
            # Check if this word starts a new movement (after a date position)
            if split_y_positions and word_y >= split_y_positions[0] - 3:
                if current_split:
                    split_rows.append(current_split)
                current_split = [word]
                # Remove the first split position as we've used it
                split_y_positions.pop(0)
            else:
                if current_split is None:
                    current_split = []
                current_split.append(word)
        
        if current_split:
            split_rows.append(current_split)
        
        return split_rows if len(split_rows) > 1 else [row_words]
    
    # Check if there are multiple amounts in numeric columns (cargos, abonos, saldo)
    # This indicates multiple movements even if there's only one date
    numeric_cols = ['cargos', 'abonos', 'saldo']
    numeric_ranges = {}
    for col in numeric_cols:
        if col in columns_config:
            x0, x1 = columns_config[col]
            numeric_ranges[col] = (x0, x1)
    
    if numeric_ranges:
        # Count amounts in each numeric column
        amounts_per_col = {col: [] for col in numeric_ranges.keys()}
        
        for word in row_words:
            text = word.get('text', '')
            x0 = word.get('x0', 0)
            x1 = word.get('x1', 0)
            center = (x0 + x1) / 2
            
            # Check if word contains an amount
            if DEC_AMOUNT_RE.search(text):
                # Find which numeric column this amount belongs to
                for col, (col_x0, col_x1) in numeric_ranges.items():
                    if col_x0 <= center <= col_x1:
                        amounts_per_col[col].append((word, word.get('top', 0)))
                        break
        
        # Check if any column has multiple amounts at different Y positions
        # This suggests multiple movements
        y_positions = set()
        words_with_multiple_amounts = []
        
        for col, amounts in amounts_per_col.items():
            if len(amounts) > 1:
                # Check if amounts are at different Y positions (more than 2 pixels apart)
                amount_ys = sorted([y for _, y in amounts])
                for i in range(len(amount_ys) - 1):
                    if abs(amount_ys[i] - amount_ys[i + 1]) > 2:
                        y_positions.add(amount_ys[i + 1])
                # For Banorte, be more aggressive: if we have multiple amounts in same column,
                # it's likely multiple movements even if Y positions are close
                if bank_name == 'Banorte' and len(amounts) > 1:
                    # Check if amounts are in different words (not just different Y positions)
                    amount_words = [w for w, _ in amounts]
                    unique_words = len(set(id(w) for w in amount_words))
                    if unique_words > 1:
                        # Multiple amounts in different words - likely multiple movements
                        # Use the Y position of the second amount as split point
                        if len(amount_ys) > 1:
                            y_positions.add(amount_ys[1])
        
        # Also check if a single word contains multiple amounts (like "$2,572.02 0.00 125,000.00")
        # This is a strong indicator of multiple movements combined
        for word in row_words:
            text = word.get('text', '')
            # Count how many distinct amounts are in this word
            amount_matches = DEC_AMOUNT_RE.findall(text)
            if len(amount_matches) > 1:
                # Multiple amounts in one word - this suggests multiple movements
                words_with_multiple_amounts.append(word)
                # Use the word's Y position as a split point
                y_positions.add(word.get('top', 0))
        
        # For Banorte, also check if we have multiple pairs of amounts (Abonos + Saldo)
        # This is a strong indicator of multiple movements
        if bank_name == 'Banorte' and 'abonos' in numeric_ranges and 'saldo' in numeric_ranges:
            abonos_amounts = amounts_per_col.get('abonos', [])
            saldo_amounts = amounts_per_col.get('saldo', [])
            # If we have multiple abonos and multiple saldos, likely multiple movements
            if len(abonos_amounts) > 1 and len(saldo_amounts) > 1:
                # Use Y positions of second and subsequent amounts as split points
                abonos_ys = sorted([y for _, y in abonos_amounts])
                saldo_ys = sorted([y for _, y in saldo_amounts])
                for y_list in [abonos_ys, saldo_ys]:
                    if len(y_list) > 1:
                        for i in range(1, len(y_list)):
                            y_positions.add(y_list[i])
        
        # If we found multiple Y positions with amounts OR words with multiple amounts, split the row
        if len(y_positions) > 0 or len(words_with_multiple_amounts) > 0:
            split_y_positions = sorted(list(y_positions))
            split_rows = []
            current_split = []
            
            sorted_words = sorted(row_words, key=lambda w: w.get('top', 0))
            
            for word in sorted_words:
                word_y = word.get('top', 0)
                
                # Check if this word starts a new movement
                if split_y_positions and word_y >= split_y_positions[0] - 3:
                    if current_split:
                        split_rows.append(current_split)
                    current_split = [word]
                    split_y_positions.pop(0)
                else:
                    if current_split is None:
                        current_split = []
                    current_split.append(word)
            
            if current_split:
                split_rows.append(current_split)
            
            return split_rows if len(split_rows) > 1 else [row_words]
        
        # If we found words with multiple amounts but couldn't split by Y position,
        # we still need to handle this case. The amounts will be extracted separately
        # in extract_movement_row and assigned to columns, but we should try to create
        # separate rows if possible based on the structure of the data.
        # For now, if we can't split by Y, we'll keep the row as is and let the
        # amount assignment logic handle it (though this may result in multiple amounts
        # in the same column, which is what we're trying to avoid)
        if len(words_with_multiple_amounts) > 0:
            # Try to split based on the positions of words with multiple amounts
            # Sort words by Y and try to find natural break points
            sorted_all_words = sorted(row_words, key=lambda w: w.get('top', 0))
            multi_amount_ys = sorted([w.get('top', 0) for w in words_with_multiple_amounts])
            
            if len(multi_amount_ys) > 0:
                # Use the first multi-amount word position as a split point
                split_y = multi_amount_ys[0]
                split_rows = []
                current_split = []
                
                for word in sorted_all_words:
                    word_y = word.get('top', 0)
                    if word_y >= split_y - 2 and current_split:
                        # Start a new row
                        split_rows.append(current_split)
                        current_split = [word]
                    else:
                        if current_split is None:
                            current_split = []
                        current_split.append(word)
                
                if current_split:
                    split_rows.append(current_split)
                
                if len(split_rows) > 1:
                    return split_rows
    
    # No splitting needed
    return [row_words]


def main():
    # Validate input
    if len(sys.argv) < 2:
        #print("Usage:")
        #print("  python main2.py <input.pdf>              # Parse PDF and create Excel")
        #print("  python main2.py <input.pdf> --find <page> # Find column coordinates on page N")
        #print("\nExample:")
        #print("  python main2.py BBVA.pdf")
        #print("  python main2.py BBVA.pdf --find 2")
        sys.exit(1)

    pdf_path = sys.argv[1]
    
    # Check for --find mode
    if len(sys.argv) >= 3 and sys.argv[2] == '--find':
        page_num = int(sys.argv[3]) if len(sys.argv) > 3 else 1
        print(f"🔍 Buscando coordenadas en página {page_num}...")
        find_column_coordinates(pdf_path, page_num)
        return

    if not os.path.isfile(pdf_path):
        #print("❌ File not found.")
        sys.exit(1)

    if not pdf_path.lower().endswith(".pdf"):
        #print("❌ Only PDF files are supported.")
        sys.exit(1)

    output_excel = os.path.splitext(pdf_path)[0] + ".xlsx"

    print("Reading PDF...")
    
    # Now extract full data
    extracted_data = extract_text_from_pdf(pdf_path)
    
    # Detectar si se usó OCR
    used_ocr = any(p.get('_used_ocr', False) for p in extracted_data)
    
    # Detect bank: from extracted text if OCR was used, otherwise from PDF
    if used_ocr:
        # Si se usó OCR, detectar banco desde el texto extraído
        print(f"[INFO] Detectando banco desde texto extraído con OCR...")
        all_text = '\n'.join([p.get('content', '') for p in extracted_data[:3]])  # Primeras 3 páginas
        detected_bank = detect_bank_from_text(all_text)
    else:
        # Si no se usó OCR, detectar banco desde PDF (método normal)
        detected_bank = detect_bank_from_pdf(pdf_path)
    
    is_hsbc = (detected_bank == "HSBC")
    
    # split pages into lines (necesario para el procesamiento posterior)
    pages_lines = split_pages_into_lines(extracted_data)
    
    # Get bank config based on detected bank
    bank_config = BANK_CONFIGS.get(detected_bank)
    if not bank_config:
        #print(f"⚠️  Bank config not found for {detected_bank}, using generic fallback")
        # Create a generic config for non-BBVA banks
        # They will use raw text extraction instead of coordinate-based
        bank_config = {
            "name": detected_bank,
            "columns": {}  # Empty columns means will use raw extraction
        }
    else:
        # Ensure the name matches the detected bank
        bank_config = bank_config.copy()
        bank_config["name"] = detected_bank
    
    columns_config = bank_config.get("columns", {})

    # find where movements start (first line anywhere that matches a date or contains header keywords)
    # Pattern for dates: supports multiple formats:
    # - "DIA MES" (01 ABR)
    # - "MES DIA" (ABR 01)
    # - "DIA MES AÑO" (06 mar 2023) - for Konfio
    # - "DIA-MES-AÑO" (12-ENE-23) - for Banorte
    # Common month abbreviations: ENE, FEB, MAR, ABR, MAY, JUN, JUL, AGO, SEP, OCT, NOV, DIC
    # Updated pattern to also match "DIA-MES-AÑO" format with hyphens (e.g., "12-ENE-23")
    day_re = re.compile(r"\b(?:(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4})\b", re.I)
    # match lines that contain both 'fecha' AND 'descripcion', OR lines that contain 'concepto'
    # Implemented with lookahead for the AND case, and an alternation for 'concepto'
    header_keywords_re = re.compile(r"(?:(?=.*\bfecha\b)(?=.*\bdescripcion\b))|(?:\bconcepto\b)", re.I)
    # ensure a reusable date pattern is available for later checks
    # For Banregio, date is only 2 digits (01-31) at the start
    if bank_config['name'] == 'Banregio':
        date_pattern = re.compile(r"^(0[1-9]|[12][0-9]|3[01])(?=\s|$)")
    elif bank_config['name'] == 'Base':
        # For Base, date format is DD/MM/YYYY (e.g., "30/04/2024")
        date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(\d{4})\b')
    elif bank_config['name'] == 'Banbajío':
        # For BanBajío, date format is "DIA MES" (e.g., "3 ENE") without year
        date_pattern = re.compile(r'\b(0?[1-9]|[12][0-9]|3[01])\s+[A-Z]{3}\b', re.I)
    else:
        date_pattern = day_re
    movement_start_found = False
    movement_start_page = None
    movement_start_index = None
    movements_lines = []
    
    # For Inbursa, movements start after "DETALLE DE MOVIMIENTOS" and the header line
    # For Banorte, movements start after "DETALLE DE MOVIMIENTOS (PESOS)"
    # For Banbajío, movements start after the header line "FECHA NO. REF. / DOCTO DESCRIPCION DE LA OPERACION DEPOSITOS RETIROS SALDO"
    # For Banregio, movements start after the header line "DIA CONCEPTO CARGOS ABONOS SALDO"
    inbursa_header_pattern = None
    banorte_detalle_pattern = None
    banbajio_start_pattern = None
    banbajio_header_pattern = None
    banregio_header_pattern = None
    scotiabank_header_pattern = None
    konfio_start_pattern = None
    clara_start_pattern = None
    base_start_pattern = None
    detalle_found = False
    header_line_skipped = False
    if bank_config['name'] == 'Inbursa':
        # Pattern to detect the header line: "FECHA REFERENCIA CONCEPTO CARGOS ABONOS SALDO"
        # Make it more flexible to handle variations in spacing
        inbursa_header_pattern = re.compile(r'FECHA.*?REFERENCIA.*?CONCEPTO.*?CARGOS.*?ABONOS.*?SALDO', re.I)
    elif bank_config['name'] == 'Banorte':
        banorte_detalle_pattern = re.compile(r'DETALLE\s+DE\s+MOVIMIENTOS\s*\(PESOS\)', re.I)
    elif bank_config['name'] == 'Banbajío':
        # Pattern to detect the start of movements section: "DETALLE DE LA CUENTA: CUENTA"
        banbajio_start_pattern = re.compile(r'DETALLE\s+DE\s+LA\s+CUENTA\s*:\s*CUENTA', re.I)
        # Pattern to detect the header line: "FECHA NO. REF. / DOCTO DESCRIPCION DE LA OPERACION DEPOSITOS RETIROS SALDO"
        # Make it flexible to handle variations in spacing and line breaks
        banbajio_header_pattern = re.compile(r'FECHA.*?NO\.?\s*REF\.?.*?DOCTO.*?DESCRIPCION.*?OPERACION.*?DEPOSITOS.*?RETIROS.*?SALDO', re.I)
    elif bank_config['name'] == 'Banregio':
        # Pattern to detect the header line: "DIA CONCEPTO CARGOS ABONOS SALDO"
        # Make it flexible to handle variations in spacing
        banregio_header_pattern = re.compile(r'DIA.*?CONCEPTO.*?CARGOS.*?ABONOS.*?SALDO', re.I)
    elif bank_config['name'] == 'Scotiabank':
        # Pattern to detect the header line: "Fecha Concepto Origen / Referencia Depósito Retiro Saldo"
        # Make it flexible to handle variations in spacing
        scotiabank_header_pattern = re.compile(r'Fecha.*?Concepto.*?Origen.*?Referencia.*?Dep[oó]sito.*?Retiro.*?Saldo', re.I)
    elif bank_config['name'] == 'Konfio':
        # Pattern to detect the start of movements section: "Historial de movimientos del titular"
        konfio_start_pattern = re.compile(r'Historial\s+de\s+movimientos\s+del\s+titular', re.I)
    elif bank_config['name'] == 'Clara':
        # Pattern to detect the start of movements section: "Movimientos"
        clara_start_pattern = re.compile(r'\bMovimientos\b', re.I)
    elif bank_config['name'] == 'Base':
        # Pattern to detect the start of movements section: "DETALLE DE OPERACIONES"
        base_start_pattern = re.compile(r'DETALLE\s+DE\s+OPERACIONES', re.I)
    
    # Generic movement_start_pattern from BANK_CONFIGS (for Banamex, Base, Clara, Konfio, etc.)
    movement_start_pattern = None
    movement_start_string = bank_config.get('movements_start')
    if movement_start_string:
        # Create pattern from movements_start string (escape special chars)
        movement_start_pattern = re.compile(re.escape(movement_start_string), re.I)
    
    for p in pages_lines:
        if not movement_start_found:
            for i, ln in enumerate(p['lines']):
                # For Konfio, find the start pattern "Historial de movimientos del titular"
                if konfio_start_pattern and not detalle_found:
                    if konfio_start_pattern.search(ln):
                        detalle_found = True
                        continue  # Skip the start pattern line itself
                
                # For Clara, find the start pattern "Movimientos"
                if clara_start_pattern and not detalle_found:
                    if clara_start_pattern.search(ln):
                        detalle_found = True
                        continue  # Skip the start pattern line itself
                
                # For Base, find the start pattern "DETALLE DE OPERACIONES"
                if base_start_pattern and not detalle_found:
                    if base_start_pattern.search(ln):
                        detalle_found = True
                        continue  # Skip the start pattern line itself
                
                # Generic movement_start_pattern (for Banamex, and other banks with movements_start in BANK_CONFIGS)
                if movement_start_pattern and not detalle_found:
                    if movement_start_pattern.search(ln):
                        detalle_found = True
                        continue  # Skip the start pattern line itself
                
                # For Inbursa, find the header line "FECHA REFERENCIA CONCEPTO CARGOS ABONOS SALDO"
                if inbursa_header_pattern and not header_line_skipped:
                    if inbursa_header_pattern.search(ln):
                        header_line_skipped = True
                        detalle_found = True
                        continue  # Skip the header line
                
                # For Banorte, find "DETALLE DE MOVIMIENTOS (PESOS)"
                if banorte_detalle_pattern and not detalle_found:
                    if banorte_detalle_pattern.search(ln):
                        detalle_found = True
                        continue  # Skip the "DETALLE DE MOVIMIENTOS (PESOS)" line itself
                
                # For Banbajío, find the start pattern "DETALLE DE LA CUENTA: CUENTA"
                if banbajio_start_pattern and not detalle_found:
                    if banbajio_start_pattern.search(ln):
                        detalle_found = True
                        #print(f"✅ BanBajío: Encontrado 'DETALLE DE LA CUENTA: CUENTA' en página {p['page']}, línea {i+1}: {ln[:100]}")
                        continue  # Continue to next iteration to find the first date or header
                
                # For Banbajío, find the header line "FECHA NO. REF. / DOCTO DESCRIPCION DE LA OPERACION DEPOSITOS RETIROS SALDO"
                # This should come right after "DETALLE DE LA CUENTA: CUENTA"
                if banbajio_header_pattern and detalle_found and not header_line_skipped:
                    if banbajio_header_pattern.search(ln):
                        header_line_skipped = True
                        print(f"✅ BanBajío: Header encontrado en página {p['page']}, línea {i+1}: {ln[:100]}")
                        # After header, start extraction from next line (could be "SALDO INICIAL" or first movement)
                        if i + 1 < len(p['lines']):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i + 1  # Start from line after header
                            next_line = p['lines'][i + 1] if i + 1 < len(p['lines']) else ""
                            #print(f"✅ BanBajío: Inicio de extracción detectado en página {p['page']}, línea {i+2}: {next_line[:100]}")
                            # collect from line after header onward in this page
                            movements_lines.extend(p['lines'][i+1:])
                            break
                        continue  # Skip the header line
                
                # For Banregio, find the header line "DIA CONCEPTO CARGOS ABONOS SALDO"
                if banregio_header_pattern and not header_line_skipped:
                    if banregio_header_pattern.search(ln):
                        header_line_skipped = True
                        detalle_found = True
                        continue  # Skip the header line
                
                # For Scotiabank, find the header line "Fecha Concepto Origen / Referencia Depósito Retiro Saldo"
                if scotiabank_header_pattern and not header_line_skipped:
                    if scotiabank_header_pattern.search(ln):
                        header_line_skipped = True
                        detalle_found = True
                        continue  # Skip the header line
                
                # After finding header for Inbursa, or for Banorte, or for Banbajío, or for Banregio, or for Scotiabank, or for Konfio, or for Clara, or for other banks, look for date/header
                if (inbursa_header_pattern and detalle_found and header_line_skipped) or (banorte_detalle_pattern and detalle_found) or (banbajio_start_pattern and detalle_found) or (banbajio_header_pattern and detalle_found and header_line_skipped) or (banregio_header_pattern and detalle_found and header_line_skipped) or (scotiabank_header_pattern and detalle_found and header_line_skipped) or (konfio_start_pattern and detalle_found) or (clara_start_pattern and detalle_found) or (base_start_pattern and detalle_found) or (not inbursa_header_pattern and not banorte_detalle_pattern and not banbajio_start_pattern and not banbajio_header_pattern and not banregio_header_pattern and not scotiabank_header_pattern and not konfio_start_pattern and not clara_start_pattern and not base_start_pattern):
                    # For Inbursa, only look for dates (not headers, as we already skipped the header line)
                    if inbursa_header_pattern:
                        # For Inbursa, only start when we find a date (actual movement row) after the header
                        if day_re.search(ln):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            # collect from this line onward in this page
                            movements_lines.extend(p['lines'][i:])
                            break
                    elif banorte_detalle_pattern:
                        # For Banorte, start when we find a date (actual movement row)
                        if day_re.search(ln):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            # collect from this line onward in this page
                            movements_lines.extend(p['lines'][i:])
                            break
                    elif banbajio_start_pattern:
                        # For Banbajío, after finding "DETALLE DE LA CUENTA: CUENTA", we should have already found the header
                        # If we reach here, it means we're looking for the first movement row
                        # Accept either a date or "SALDO INICIAL" as valid start
                        if day_re.search(ln) or re.search(r'SALDO\s+INICIAL', ln, re.I):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            print(f"✅ BanBajío: Inicio de extracción detectado en página {p['page']}, línea {i+1}: {ln[:100]}")
                            # collect from this line onward in this page
                            movements_lines.extend(p['lines'][i:])
                            break
                    elif banbajio_header_pattern:
                        # This should not happen if logic is correct, but keep as fallback
                        # For Banbajío, start when we find a date or "SALDO INICIAL" (actual movement row) after the header
                        if day_re.search(ln) or re.search(r'SALDO\s+INICIAL', ln, re.I):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            print(f"✅ BanBajío: Inicio de extracción detectado en página {p['page']}, línea {i+1}: {ln[:100]}")
                            # collect from this line onward in this page
                            movements_lines.extend(p['lines'][i:])
                            break
                    elif banregio_header_pattern:
                        # For Banregio, start when we find a date (actual movement row) after the header
                        if day_re.search(ln):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            # collect from this line onward in this page
                            movements_lines.extend(p['lines'][i:])
                            break
                    elif scotiabank_header_pattern:
                        # For Scotiabank, start when we find a date (actual movement row) after the header
                        if day_re.search(ln):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            # collect from this line onward in this page
                            movements_lines.extend(p['lines'][i:])
                            break
                    elif konfio_start_pattern:
                        # For Konfio, start when we find a date (actual movement row) after the start pattern
                        # But verify it's a valid date format for Konfio (DIA MES AÑO, e.g., "14 mar 2023")
                        # This prevents false positives like "14 IVA" from being detected as a date
                        konfio_date_in_line = date_pattern.search(ln)
                        if konfio_date_in_line:
                            # Verify it's a valid Konfio date format (not just a number)
                            date_match = konfio_date_in_line.group()
                            # Check if it matches the full Konfio date pattern (DIA MES AÑO)
                            konfio_full_date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4}\b', re.I)
                            if konfio_full_date_pattern.search(ln):
                                movement_start_found = True
                                movement_start_page = p['page']
                                movement_start_index = i
                                # collect from this line onward in this page
                                movements_lines.extend(p['lines'][i:])
                                break
                    elif clara_start_pattern:
                        # For Clara, start when we find a date (actual movement row) after the start pattern
                        if day_re.search(ln):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            # collect from this line onward in this page
                            movements_lines.extend(p['lines'][i:])
                            break
                    else:
                        # For other banks, look for date or header
                        if day_re.search(ln) or header_keywords_re.search(ln):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            # collect from this line onward in this page
                            movements_lines.extend(p['lines'][i:])
                            break
        else:
            # Already found movement start, collect all lines from this page
            # For Banbajío, Banregio, Inbursa, and Scotiabank, filter out the header line if it appears again on subsequent pages
            if bank_config['name'] == 'Banbajío' and banbajio_header_pattern:
                filtered_lines = [ln for ln in p['lines'] if not banbajio_header_pattern.search(ln)]
                movements_lines.extend(filtered_lines)
            elif bank_config['name'] == 'Banregio' and banregio_header_pattern:
                # Filter out header line and rows starting with "del 01 al"
                filtered_lines = [ln for ln in p['lines'] if not banregio_header_pattern.search(ln) and not re.search(r'^del\s+01\s+al', ln, re.I)]
                movements_lines.extend(filtered_lines)
            elif bank_config['name'] == 'Inbursa' and inbursa_header_pattern:
                filtered_lines = [ln for ln in p['lines'] if not inbursa_header_pattern.search(ln)]
                movements_lines.extend(filtered_lines)
            elif bank_config['name'] == 'Scotiabank' and scotiabank_header_pattern:
                filtered_lines = [ln for ln in p['lines'] if not scotiabank_header_pattern.search(ln)]
                movements_lines.extend(filtered_lines)
            elif bank_config['name'] == 'Base' and base_start_pattern:
                # Filter out the start pattern line
                filtered_lines = [ln for ln in p['lines'] if not base_start_pattern.search(ln)]
                movements_lines.extend(filtered_lines)
            elif bank_config['name'] == 'Clara' and clara_start_pattern:
                # For Clara, filter out "Movimientos" header line if it appears again on subsequent pages
                filtered_lines = [ln for ln in p['lines'] if not clara_start_pattern.search(ln)]
                movements_lines.extend(filtered_lines)
            else:
                movements_lines.extend(p['lines'])

    # build summary from first page (lines before movements start if movements begin on page 1)
    if pages_lines:
        first_page = pages_lines[0]
        if movement_start_found and movement_start_page == first_page['page'] and movement_start_index is not None:
            summary_lines = first_page['lines'][:movement_start_index]
        else:
            summary_lines = first_page['lines']
    else:
        summary_lines = []

    # Extract movements using coordinate-based column detection
    # For all banks including Konfio, use coordinate-based extraction
    movement_rows = []
    df_mov = None  # Initialize to avoid UnboundLocalError
    pdf_summary = None  # Initialize to avoid UnboundLocalError
    
    # Si es HSBC y se usó OCR, usar la misma lógica que otros bancos
    if is_hsbc and used_ocr:
        # Obtener columns_config desde BANK_CONFIGS
        columns_config = bank_config.get("columns", {})
        
        if not columns_config:
            print("[ADVERTENCIA] BANK_CONFIGS para HSBC no tiene 'columns'. Usando valores por defecto.")
            # Valores por defecto (deberían calibrarse con --find)
            columns_config = {
                "fecha": (87, 103),
                "descripcion": (124, 505),
                "cargos": (710, 800),
                "abonos": (865, 950),
                "saldo": (1050, 1130),
            }
        
        # Obtener strings de inicio/fin desde BANK_CONFIGS
        start_string = bank_config.get('movements_start', 'ISR Retenido en el año')
        end_string = bank_config.get('movements_end', 'CoDi')
        
        # Filtrar palabras en la sección de movimientos
        filtered_words = filter_hsbc_movements_section(extracted_data, start_string, end_string)
        
        if not filtered_words:
            df_mov = pd.DataFrame(columns=['fecha', 'descripcion', 'cargos', 'abonos', 'saldo'])
        else:
            # Agrupar palabras por filas (igual que otros bancos)
            # Para HSBC, usar tolerancia Y más amplia para capturar montos que pueden estar ligeramente desalineados
            word_rows = group_words_by_row(filtered_words, y_tolerance=5)
            
            # Patrón de fecha para HSBC (solo día: 01-31)
            date_pattern = re.compile(r"^(0[1-9]|[12][0-9]|3[01])(?=\s|$)")
            
            # Extraer movimientos usando la misma lógica que otros bancos
            for row_idx, row_words in enumerate(word_rows):
                if not row_words:
                    continue
                
                # Extraer movimiento usando BANK_CONFIGS
                row_data = extract_movement_row(row_words, columns_config, 'HSBC', date_pattern)
                
                # Procesar montos detectados (_amounts) y asignarlos a columnas para HSBC
                amounts_list = row_data.get('_amounts', [])
                if amounts_list and len(amounts_list) >= 1:
                    # Si hay montos pero no están asignados a columnas, asignarlos basándose en coordenadas X
                    if not row_data.get('cargos') and not row_data.get('abonos') and not row_data.get('saldo'):
                        # Asignar montos según coordenadas X
                        for amt_text, center in amounts_list:
                            col_name = assign_word_to_column(center - 50, center + 50, columns_config)
                            if col_name and col_name in ('cargos', 'abonos', 'saldo'):
                                if not row_data.get(col_name):
                                    row_data[col_name] = amt_text
                    
                    # Asegurar que cuando hay 2 montos, el segundo sea saldo (estructura típica HSBC)
                    if len(amounts_list) == 2:
                        if not row_data.get('saldo'):
                            # El segundo monto es siempre saldo en HSBC
                            second_amt = amounts_list[1][0]
                            row_data['saldo'] = second_amt
                        # Si el primer monto no está asignado, asignarlo a cargos o abonos según coordenadas
                        if not row_data.get('cargos') and not row_data.get('abonos'):
                            first_amt, first_center = amounts_list[0]
                            col_name = assign_word_to_column(first_center - 50, first_center + 50, columns_config)
                            if col_name and col_name in ('cargos', 'abonos'):
                                row_data[col_name] = first_amt
                            else:
                                # Por defecto, si no se puede determinar, asignar a cargos
                                row_data['cargos'] = first_amt
                
                # Validar si es una transacción válida
                is_valid = is_transaction_row(row_data, 'HSBC')
                
                if is_valid:
                    movement_rows.append(row_data)
        
        df_mov = pd.DataFrame(movement_rows) if movement_rows else pd.DataFrame(columns=['fecha', 'descripcion', 'cargos', 'abonos', 'saldo'])
        
        # Extraer resumen desde texto OCR de la página 1
        pdf_summary = extract_hsbc_summary_from_ocr_text(extracted_data)
    else:
        # Flujo normal: procesar con coordenadas o texto plano
        # regex to detect decimal-like amounts (used to strip amounts from descriptions and detect amounts)
        dec_amount_re = re.compile(r"\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")
        # Pattern to detect end of movements table for specific banks
        # First, try to use movements_end from BANK_CONFIGS (for Banamex, HSBC, etc.)
        movement_end_pattern = None
        movement_end_string = bank_config.get('movements_end')
        if movement_end_string:
            # Create pattern from movements_end string (escape special chars)
            movement_end_pattern = re.compile(re.escape(movement_end_string), re.I)
        elif bank_config['name'] == 'Santander':
            # Santander: "TOTAL 821,646.20 820,238.73 1,417.18" - indicates end of movements table
            movement_end_pattern = re.compile(r'^TOTAL\s+[\d,\.]+\s+[\d,\.]+\s+[\d,\.]+', re.I)
        elif bank_config['name'] == 'Banorte':
            # Banorte: "INVERSION ENLACE NEGOCIOS" - indicates end of movements table
            movement_end_pattern = re.compile(r'INVERSION\s+ENLACE\s+NEGOCIOS', re.I)
        elif bank_config['name'] == 'Banregio':
            # Banregio: "Total" - indicates end of movements table
            movement_end_pattern = re.compile(r'^Total\b', re.I)
        elif bank_config['name'] == 'Scotiabank':
            # Scotiabank: "LAS TASAS DE INTERES ESTAN EXPRESADAS EN TERMINOS ANUALES SIMPLES." - indicates end of movements table
            # Use a more flexible pattern to handle variations in spacing and formatting
            movement_end_pattern = re.compile(r'LAS\s+TASAS\s+DE\s+INTERES\s+ESTAN\s+EXPRESADAS\s+EN\s+TERMINOS\s+ANUALES\s+SIMPLES\.?', re.I)
        elif bank_config['name'] == 'Inbursa':
            # Inbursa: "Si desea recibir pagos a través de transferencias bancarias electrónicas" - indicates end of movements table
            # Use a flexible pattern to handle variations in spacing and formatting
            movement_end_pattern = re.compile(r'Si\s+desea\s+recibir\s+pagos\s+a\s+trav[eé]s\s+de\s+transferencias\s+bancarias\s+electr[oó]nicas', re.I)
        elif bank_config['name'] == 'Konfio':
            # Konfio: "Subtotal" - indicates end of movements table
            movement_end_pattern = re.compile(r'^Subtotal\b', re.I)
        elif bank_config['name'] == 'Clara':
            # Clara: "Total" followed by 2 amounts - indicates end of movements table
            # Pattern: "Total" followed by space and two amounts (numbers with optional commas and decimals)
            movement_end_pattern = re.compile(r'\bTotal\b\s+[\d,\.]+\s+[\d,\.]+', re.I)
        elif bank_config['name'] == 'Base':
            # Base: "[SALDO INICIAL DE" - indicates end of movements table
            movement_end_pattern = re.compile(r'\[SALDO\s+INICIAL\s+DE', re.I)
        elif bank_config['name'] == 'Banbajío':
            # Banbajío: "SALDO TOTAL*" - indicates end of movements table
            movement_end_pattern = re.compile(r'SALDO\s+TOTAL\*?', re.I)
        
        extraction_stopped = False
        # For Banregio, initialize flag to track when we're in the commission zone
        in_comision_zone = False
        # For BanBajío, track detected rows for debugging
        banbajio_detected_rows = 0
        for page_data in extracted_data:
            if extraction_stopped:
                break
            
            page_num = page_data['page']
            words = page_data.get('words', [])
            
            if not words:
                continue
            
            # For banks with movement_start_pattern (Banamex, Base, etc.), check if start pattern is required but not yet found
            if movement_start_pattern and not movement_start_found:
                # Check if this page contains the start pattern
                page_text = ' '.join([w.get('text', '') for w in words])
                if not movement_start_pattern.search(page_text):
                    # Start pattern not found on this page, skip it
                    continue
            
            # Check if this page contains movements (page >= movement_start_page if found)
            if movement_start_found and page_num < movement_start_page:
                continue
            
            # Group words by row
            # For Konfio, use a larger y_tolerance to capture multi-line descriptions
            # Use y_tolerance=3 for all banks to avoid grouping multiple movements into one row
            # The split_row_if_multiple_movements function will handle cases where movements are still grouped
            y_tol = 8 if bank_config['name'] == 'Konfio' else 3
            word_rows = group_words_by_row(words, y_tolerance=y_tol)
            
            # Check if grouped rows contain multiple movements and split them
            # This applies to all banks to ensure each movement is in its own row
            if columns_config:
                split_rows = []
                for row_words in word_rows:
                    if not row_words:
                        continue
                    
                    # Use the generic function to split rows with multiple movements
                    # Pass bank_name to enable bank-specific logic
                    split_result = split_row_if_multiple_movements(row_words, columns_config, date_pattern, bank_config['name'])
                    split_rows.extend(split_result)
                
                word_rows = split_rows
            
            for row_idx, row_words in enumerate(word_rows):
                if not row_words or extraction_stopped:
                    continue

                # For Banbajío, Banregio, Inbursa, and Scotiabank, skip the header line if it appears on subsequent pages
                if bank_config['name'] == 'Banbajío' and banbajio_header_pattern:
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    if banbajio_header_pattern.search(all_row_text):
                        continue  # Skip the header line
                
                if bank_config['name'] == 'Banregio' and banregio_header_pattern:
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    if banregio_header_pattern.search(all_row_text):
                        continue  # Skip the header line
                
                if bank_config['name'] == 'Inbursa' and inbursa_header_pattern:
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    if inbursa_header_pattern.search(all_row_text):
                        continue  # Skip the header line
                
                # For Scotiabank, skip the header line if it appears during coordinate-based extraction
                if bank_config['name'] == 'Scotiabank' and scotiabank_header_pattern:
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    if scotiabank_header_pattern.search(all_row_text):
                        continue  # Skip the header line
                
                # For Base, skip the start pattern line if it appears during coordinate-based extraction
                if bank_config['name'] == 'Base' and base_start_pattern:
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    if base_start_pattern.search(all_row_text):
                        continue  # Skip the start pattern line
                
                # For banks with movement_start_pattern (Banamex, Base, etc.), skip the start pattern line if it appears during coordinate-based extraction
                if movement_start_pattern:
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    if movement_start_pattern.search(all_row_text):
                        continue  # Skip the start pattern line
                
                # For Clara, skip the "Movimientos" header line if it appears during coordinate-based extraction
                if bank_config['name'] == 'Clara' and clara_start_pattern:
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    if clara_start_pattern.search(all_row_text):
                        continue  # Skip the header line
                    
                # For Banregio, skip rows that start with "del 01 al" (irrelevant information)
                if bank_config['name'] == 'Banregio':
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    if re.search(r'^del\s+01\s+al', all_row_text, re.I):
                        continue  # Skip irrelevant information rows

                # Debug for Banamex: print all words in the row before processing
                # Check for end pattern (for Banamex, Santander, Banregio, Scotiabank, Konfio, Clara, etc.)
                if movement_end_pattern:
                    all_text = ' '.join([w.get('text', '') for w in row_words])
                    match_found = False
                    
                    # For Scotiabank, try multiple flexible patterns to handle variations in spacing and formatting
                    if bank_config['name'] == 'Scotiabank':
                        # Try multiple flexible patterns
                        scotiabank_end_patterns = [
                            re.compile(r'LAS\s+TASAS\s+DE\s+INTERES\s+ESTAN\s+EXPRESADAS\s+EN\s+TERMINOS\s+ANUALES\s+SIMPLES\.?', re.I),
                            re.compile(r'LAS\s+TASAS\s+DE\s+INTERES.*?ESTAN.*?EXPRESADAS.*?EN.*?TERMINOS.*?ANUALES.*?SIMPLES\.?', re.I),
                            re.compile(r'LAS.*?TASAS.*?DE.*?INTERES.*?ESTAN.*?EXPRESADAS.*?EN.*?TERMINOS.*?ANUALES.*?SIMPLES', re.I),
                        ]
                        for pattern in scotiabank_end_patterns:
                            if pattern.search(all_text):
                                match_found = True
                                break
                    elif bank_config['name'] == 'Clara':
                        # For Clara, check for "Total" followed by 2 amounts
                        # Try multiple flexible patterns
                        clara_end_patterns = [
                            re.compile(r'\bTotal\b\s+MXN\s+[\d,\.]+\s+MXN\s+[\d,\.]+', re.I),  # "Total MXN ... MXN ..."
                            re.compile(r'\bTotal\b\s+[\d,\.]+\s+[\d,\.]+', re.I),  # "Total" followed by two amounts
                            re.compile(r'^Total\b\s+[\d,\.]+\s+[\d,\.]+', re.I),  # "Total" at start of line
                        ]
                        for pattern in clara_end_patterns:
                            if pattern.search(all_text):
                                match_found = True
                                break
                    else:
                        # For other banks (Banregio, Banamex, Santander, Banorte, Konfio, Banbajío), use the standard pattern
                        if movement_end_pattern.search(all_text):
                            match_found = True
                            #if bank_config['name'] == 'Banbajío':
                                #print(f"🛑 BanBajío: Fin de extracción detectado en página {page_num}, fila {row_idx+1}: {all_row_text[:100]}")
                    
                    if match_found:
                        #print(f"🛑 Fin de tabla de movimientos detectado en página {page_num}")
                        extraction_stopped = True
                        break

                # Extract structured row using coordinates
                # Pass bank_name and date_pattern to enable date/description separation
                
                row_data = extract_movement_row(row_words, columns_config, bank_config['name'], date_pattern)
                
                # Debug for Konfio: print row data after extraction (only first 3 pages)
                if bank_config['name'] == 'Konfio' and page_num <= 3:
                    all_row_text = ' '.join([w.get('text', '') for w in row_words])
                    fecha_val = str(row_data.get('fecha') or '').strip()
                    desc_val = str(row_data.get('descripcion') or '').strip()
                    cargos_val = str(row_data.get('cargos') or '').strip()
                    abonos_val = str(row_data.get('abonos') or '').strip()
                    #print(f"🔍 Konfio: Fila extraída (página {page_num}, fila {row_idx+1}) - Fecha: '{fecha_val}', Desc: '{desc_val[:60]}', Cargos: '{cargos_val}', Abonos: '{abonos_val}'")
                    #print(f"   Texto completo: {all_row_text[:150]}")
                
                # Special debug for "IVA SOBRE COMISIONES E INTERESES" row
                if bank_config['name'] == 'Konfio':
                    desc_val_check = str(row_data.get('descripcion') or '').strip()
                    all_row_text_check = ' '.join([w.get('text', '') for w in row_words])
                    if 'IVA SOBRE COMISIONES E INTERESES' in desc_val_check.upper() or 'IVA SOBRE COMISIONES E INTERESES' in all_row_text_check.upper():
                        fecha_val = str(row_data.get('fecha') or '').strip()

                # Determine if this row starts a new movement (contains a date)
                # If columns_config is empty, check all words for dates
                if not columns_config:
                    # Check all words in the row for dates
                    all_text = ' '.join([w.get('text', '') for w in row_words])
                    has_date = bool(date_pattern.search(all_text))
                    if has_date:
                        # Create a basic row_data structure
                        row_data = {'raw': all_text, '_amounts': row_data.get('_amounts', [])}
                else:
                    # A new movement begins when the 'fecha' column contains a date token.
                    fecha_val = str(row_data.get('fecha') or '')
                    # For Banregio, use match() instead of search() since we want to verify the entire string is the date
                    if bank_config['name'] == 'Banregio':
                        has_date = bool(date_pattern.match(fecha_val))
                    elif bank_config['name'] == 'Clara':
                        # For Clara, check if fecha contains a valid date pattern (e.g., "01 ENE" or "01 E N E")
                        clara_date_pattern = re.compile(r'\d{1,2}\s+[A-Z]{3}|\d{1,2}\s+[A-Z]\s+[A-Z]\s+[A-Z]', re.I)
                        has_date = bool(clara_date_pattern.search(fecha_val))
                    elif bank_config['name'] == 'Banamex':
                        # For Banamex, check if fecha contains a valid date pattern (e.g., "30 ENE" or "10 ENE")
                        banamex_date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])\s+[A-Z]{3}\b', re.I)
                        has_date = bool(banamex_date_pattern.search(fecha_val))
                    else:
                        has_date = bool(date_pattern.search(fecha_val))
                        # For Konfio, if no date in fecha column, check all words in the row
                        # But use strict format: must be "DIA MES AÑO" (e.g., "14 mar 2023"), not just "14"
                        if bank_config['name'] == 'Konfio' and not has_date:
                            all_row_text = ' '.join([w.get('text', '') for w in row_words])
                            # Use strict Konfio date pattern: DIA MES AÑO format
                            konfio_full_date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4}\b', re.I)
                            has_date = bool(konfio_full_date_pattern.search(all_row_text))
                            if has_date:
                                # Extract the date from the row text and assign to fecha
                                date_match = konfio_full_date_pattern.search(all_row_text)
                                if date_match:
                                    row_data['fecha'] = date_match.group()
                        # Debug for Konfio
                        #if bank_config['name'] == 'Konfio' and not has_date and fecha_val:
                            #all_row_text = ' '.join([w.get('text', '') for w in row_words])
                            #print(f"⚠️ Konfio: Fecha no detectada - Valor en columna fecha: '{fecha_val}', Texto completo: '{all_row_text[:100]}', Patrón usado: {date_pattern.pattern}")
                    
                    # Check if row has valid data (date, description, or amounts)
                    has_valid_data = has_date
                    if not has_valid_data:
                        # Check if row has description or amounts
                        desc_val = str(row_data.get('descripcion') or '').strip()
                        has_amounts = len(row_data.get('_amounts', [])) > 0
                        has_cargos_abonos = bool(row_data.get('cargos') or row_data.get('abonos') or row_data.get('saldo'))
                        has_valid_data = bool(desc_val or has_amounts or has_cargos_abonos)
                        
                        # Debug for Konfio
                        #if bank_config['name'] == 'Konfio' and not has_valid_data:
                            #all_row_text = ' '.join([w.get('text', '') for w in row_words])
                            #print(f"⚠️ Konfio: Fila rechazada (sin fecha ni datos válidos) - Fecha: '{fecha_val}', Desc: '{desc_val[:60]}', Montos: {has_amounts}, Cargos/Abonos: {has_cargos_abonos}")
                            #print(f"   Texto completo: {all_row_text[:150]}")
                    
                    # For BanBajío, also accept rows with "SALDO INICIAL" even if they don't have a date
                    if bank_config['name'] == 'Banbajío' and not has_date:
                        desc_val = str(row_data.get('descripcion') or '').strip()
                        if re.search(r'SALDO\s+INICIAL', desc_val, re.I):
                            has_date = True  # Treat "SALDO INICIAL" as a valid row even without date
                            has_valid_data = True
                    

                if has_date:
                    # Only add rows that have date AND (description OR amounts)
                    # This ensures we don't add incomplete rows
                    desc_val = str(row_data.get('descripcion') or '').strip()
                    has_amounts = len(row_data.get('_amounts', [])) > 0
                    has_cargos_abonos = bool(row_data.get('cargos') or row_data.get('abonos') or row_data.get('saldo'))
                    has_description_or_amounts = bool(desc_val or has_amounts or has_cargos_abonos)
                    
                    
                    # For Konfio, be more lenient - if we have a date, add the row even if description/amounts are in continuation rows
                    if bank_config['name'] == 'Konfio':
                        # Verify the date is a valid Konfio date format (DIA MES AÑO), not just a number like "14"
                        fecha_val_check = str(row_data.get('fecha') or '').strip()
                        konfio_full_date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4}\b', re.I)
                        if fecha_val_check and not konfio_full_date_pattern.search(fecha_val_check):
                            # Date is not in valid Konfio format (e.g., "14" instead of "14 mar 2023"), skip this row
                            continue
                        
                        # Check if this is a sub-row pattern (like "SBO 020902DX1 - 02 abr 2023 4316 FÍSICA")
                        # These should be rejected as they are continuation lines, not main movement rows
                        all_row_text = ' '.join([w.get('text', '') for w in row_words])
                        konfio_sub_row_pattern = re.compile(r'^[A-Z]{2,4}\s*[A-Z0-9]{8,15}\s*-\s*\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4}', re.I)
                        # If the row starts with this pattern, it's a sub-row and should be skipped
                        if konfio_sub_row_pattern.match(all_row_text.strip()):
                            # This is a sub-row, skip it (it will be handled as continuation if needed)
                            continue
                        # Add row if it has date, even if description/amounts are empty (they'll be added in continuation rows)
                        movement_rows.append(row_data)
                    elif bank_config['name'] == 'Banamex':
                        # For Banamex, if a row has a date, it ALWAYS starts a new movement
                        # Even if it doesn't have description or amounts yet (they'll come in continuation rows)
                        # This handles multi-line descriptions where date is on first line
                        
                        # Filter out footer information (e.g., "000180.B61CHDA011.OD.0131.01")
                        desc_val_check = str(row_data.get('descripcion') or '').strip()
                        fecha_val_check = str(row_data.get('fecha') or '').strip()
                        all_row_text = desc_val_check + ' ' + fecha_val_check
                        
                        # Pattern to match footer: numbers.letters.OD.numbers.numbers
                        banamex_footer_pattern = re.compile(r'\d+\.\w+\.OD\.\d+\.\d+', re.I)
                        if banamex_footer_pattern.search(all_row_text):
                            # This is footer information, skip it
                            continue
                        
                        #row_data['page'] = page_num
                        
                        movement_rows.append(row_data)
                    elif has_description_or_amounts:
                        # For Banregio, only include rows where description starts with "TRA" or "DOC"
                        if bank_config['name'] == 'Banregio':
                            desc_val_check = str(row_data.get('descripcion') or '').strip()
                            if not (desc_val_check.startswith('TRA') or desc_val_check.startswith('DOC') or desc_val_check.startswith('INT') or desc_val_check.startswith('EFE')):
                                # Skip this row - doesn't start with TRA or DOC or INT or EFE
                                continue
                        
                        # Filter out footer information for Banamex (e.g., "000180.B61CHDA011.OD.0131.01")
                        if bank_config['name'] == 'Banamex':
                            desc_val_check = str(row_data.get('descripcion') or '').strip()
                            fecha_val_check = str(row_data.get('fecha') or '').strip()
                            all_row_text = desc_val_check + ' ' + fecha_val_check
                            
                            # Pattern to match footer: numbers.letters.OD.numbers.numbers
                            banamex_footer_pattern = re.compile(r'\d+\.\w+\.OD\.\d+\.\d+', re.I)
                            if banamex_footer_pattern.search(all_row_text):
                                # This is footer information, skip it
                                continue
                        
                        #row_data['page'] = page_num
                        
                        movement_rows.append(row_data)
                elif has_valid_data and bank_config['name'] == 'Banbajío':
                    # For BanBajío, also add rows without date if they contain "SALDO INICIAL"
                    desc_val = str(row_data.get('descripcion') or '').strip()
                    if re.search(r'SALDO\s+INICIAL', desc_val, re.I):
                        has_amounts = len(row_data.get('_amounts', [])) > 0
                        has_cargos_abonos = bool(row_data.get('cargos') or row_data.get('abonos') or row_data.get('saldo'))
                        if has_amounts or has_cargos_abonos:
                            #row_data['page'] = page_num
                            movement_rows.append(row_data)
                
                # For Banregio, check if this is a monthly commission (last movement) - stop extraction
                if bank_config['name'] == 'Banregio' and has_valid_data:
                    desc_val_check = str(row_data.get('descripcion') or '').strip()
                    cargos_val = str(row_data.get('cargos') or '').strip()
                    saldo_val = str(row_data.get('saldo') or '').strip()
                    # Check if description contains "COMISION MENSUAL" and has cargos and saldo values
                    has_comision = bool(re.search(r'COMISION\s+MENSUAL', desc_val_check, re.I) and cargos_val and saldo_val)
                    
                    if has_comision:
                        # Mark that we're in the commission zone and continue extracting
                        in_comision_zone = True
                    elif in_comision_zone:
                        # We were in commission zone but this row is not a commission - stop extraction
                        extraction_stopped = True
                        break  # Stop extraction - no more monthly commissions
                
                # If row has valid data but no date - treat as continuation or standalone row
                if not has_date and has_valid_data:
                    # Row has valid data but no date - treat as continuation or standalone row
                    
                    # Filter out footer information for Banamex (e.g., "000180.B61CHDA011.OD.0131.01")
                    if bank_config['name'] == 'Banamex':
                        desc_val_check = str(row_data.get('descripcion') or '').strip()
                        cargos_val_check = str(row_data.get('cargos') or '').strip()
                        abonos_val_check = str(row_data.get('abonos') or '').strip()
                        saldo_val_check = str(row_data.get('saldo') or '').strip()
                        
                        # For continuation rows, must have description AND it must not contain footer pattern
                        if not desc_val_check:
                            # No description - skip this continuation row
                            continue
                        
                        # Check footer pattern in description and all row text (including amounts)
                        all_row_text = ' '.join([w.get('text', '') for w in row_words])
                        all_row_text_with_values = desc_val_check + ' ' + cargos_val_check + ' ' + abonos_val_check + ' ' + saldo_val_check
                        
                        # Pattern to match footer: numbers.letters.OD.numbers.numbers
                        # Also check for partial patterns that might be split across words
                        banamex_footer_pattern = re.compile(r'\d+\.\w+\.OD\.\d+\.\d+', re.I)
                        # Also check for patterns that might be split: look for "OD" followed by numbers.numbers
                        banamex_footer_partial = re.compile(r'\.OD\.\d+\.\d+', re.I)
                        # Check if any word contains "OD" followed by numbers (part of footer)
                        has_od_pattern = any('OD' in w.get('text', '') and re.search(r'OD[\.\s]*\d+[\.\s]*\d+', w.get('text', ''), re.I) for w in row_words)
                        
                        if banamex_footer_pattern.search(all_row_text) or banamex_footer_pattern.search(all_row_text_with_values) or banamex_footer_partial.search(all_row_text) or has_od_pattern:
                            # This is footer information, skip it
                            continue
                        
                        # Additional check: if saldo contains a value that looks like part of footer (e.g., "131.01")
                        # and there's no meaningful description, skip it
                        if saldo_val_check and re.match(r'^\d+\.\d{2}$', saldo_val_check) and len(desc_val_check) < 10:
                            # Check if this value appears near footer-related words
                            if any('OD' in w.get('text', '') or re.search(r'\d+\.\w+', w.get('text', ''), re.I) for w in row_words):
                                continue
                    # For BanBajío, filter out informational rows like "1 DE ENERO AL 31 DE ENERO DE 2024 PERIODO:"
                    if bank_config['name'] == 'Banbajío':
                        all_row_text_check = ' '.join([w.get('text', '') for w in row_words])
                        all_row_text_check = all_row_text_check.strip()
                        # Check if row contains period information (e.g., "1 DE ENERO AL 31 DE ENERO DE 2024 PERIODO:")
                        if re.search(r'\bPERIODO\s*:?\s*$', all_row_text_check, re.I) or re.search(r'\d+\s+DE\s+[A-Z]+\s+AL\s+\d+\s+DE\s+[A-Z]+', all_row_text_check, re.I):
                            # Skip this row - it's period information, not a movement
                            continue
                    
                    # For Banregio, if we're in commission zone, check for irrelevant rows and stop extraction
                    if bank_config['name'] == 'Banregio' and in_comision_zone:
                        # Collect all text from the row to check for irrelevant content
                        all_row_text_check = ' '.join([w.get('text', '') for w in row_words])
                        all_row_text_check = all_row_text_check.strip()
                        
                        # Check if row contains "Total" (summary line)
                        if re.search(r'\bTotal\b', all_row_text_check, re.I):
                            extraction_stopped = True
                            break  # Stop extraction - "Total" line detected
                        
                        # Check if text is too long (likely not part of a movement description)
                        if len(all_row_text_check) > 200:
                            extraction_stopped = True
                            break  # Stop extraction - text too long, likely irrelevant
                        
                        # Check if text contains keywords indicating legal/informational content
                        irrelevant_keywords = [
                            'FOLIO FISCAL', 'Sello Digital', 'Institución de Banca', 'Certificado',
                            'Regimen Fiscal', 'Método de Pago', 'Cadena Original', 'Complemento',
                            'Gráfico', 'Transaccional', 'Mensajes', 'Abreviaturas', 'Origen de la Operación',
                            'Sociedades de', 'Inversión', 'Escala de Calificaciones', 'Riesgo',
                            'OFRECEMOS', 'CONSULTE', 'INFORMACION RELEVANTE', 'UNIDAD ESPECIALIZADA',
                            'NOTA', 'CONCEPTO DIA', 'Saldo Inicial', 'Saldo Final'
                        ]
                        all_row_text_upper = all_row_text_check.upper()
                        if any(keyword.upper() in all_row_text_upper for keyword in irrelevant_keywords):
                            extraction_stopped = True
                            break  # Stop extraction - irrelevant information detected
                        
                        # If none of the above, stop extraction anyway (we're past commissions)
                        extraction_stopped = True
                        break  # Stop extraction - no more monthly commissions
                    
                    if movement_rows:
                        # For Banamex: additional validation before merging continuation row
                        if bank_config['name'] == 'Banamex':
                            desc_val_check = str(row_data.get('descripcion') or '').strip()
                            # Must have description to merge continuation row
                            if not desc_val_check:
                                continue
                            
                            # Check if description contains footer pattern
                            all_row_text = ' '.join([w.get('text', '') for w in row_words])
                            banamex_footer_pattern = re.compile(r'\d+\.\w+\.OD\.\d+\.\d+', re.I)
                            banamex_footer_partial = re.compile(r'\.OD\.\d+\.\d+', re.I)
                            has_od_pattern = any('OD' in w.get('text', '') and re.search(r'OD[\.\s]*\d+[\.\s]*\d+', w.get('text', ''), re.I) for w in row_words)
                            
                            if banamex_footer_pattern.search(desc_val_check) or banamex_footer_pattern.search(all_row_text) or banamex_footer_partial.search(desc_val_check) or has_od_pattern:
                                continue
                        
                        # Continuation row: append description-like text and amounts to previous movement
                        prev = movement_rows[-1]
                        
                        # IMPORTANT: First, copy amounts that are already assigned to columns (cargos, abonos, saldo)
                        # These might not be in _amounts if they were assigned directly in extract_movement_row
                        cont_cargos = str(row_data.get('cargos') or '').strip()
                        cont_abonos = str(row_data.get('abonos') or '').strip()
                        cont_saldo = str(row_data.get('saldo') or '').strip()
                        
                        # If amounts are already assigned to columns, copy them directly to previous row
                        if cont_cargos and not prev.get('cargos'):
                            prev['cargos'] = cont_cargos
                        if cont_abonos and not prev.get('abonos'):
                            prev['abonos'] = cont_abonos
                        if cont_saldo and not prev.get('saldo'):
                            prev['saldo'] = cont_saldo
                        
                        # First, capture amounts from continuation row and assign to appropriate columns
                        cont_amounts = row_data.get('_amounts', [])
                        if cont_amounts and columns_config:
                            # Get description range to exclude amounts from it
                            descripcion_range = None
                            if 'descripcion' in columns_config:
                                x0, x1 = columns_config['descripcion']
                                descripcion_range = (x0, x1)
                            
                            # Get column ranges for numeric columns
                            col_ranges = {}
                            for col in ('cargos', 'abonos', 'saldo'):
                                if col in columns_config:
                                    x0, x1 = columns_config[col]
                                    col_ranges[col] = (x0, x1)
                            
                            # Assign amounts from continuation row
                            tolerance = 10
                            for amt_text, center in cont_amounts:
                                # Skip if amount is within description range
                                if descripcion_range and descripcion_range[0] <= center <= descripcion_range[1]:
                                    continue
                                
                                # Find which numeric column this amount belongs to
                                assigned = False
                                for col in ('cargos', 'abonos', 'saldo'):
                                    if col in col_ranges:
                                        x0, x1 = col_ranges[col]
                                        if (x0 - tolerance) <= center <= (x1 + tolerance):
                                            # Only assign if the column is empty or if this is a better match
                                            existing = prev.get(col) or ''
                                            if not existing or amt_text not in existing:
                                                if existing:
                                                    prev[col] = (existing + ' ' + amt_text).strip()
                                                else:
                                                    prev[col] = amt_text
                                            assigned = True
                                            break
                                
                                # If not assigned by range, use proximity as fallback
                                if not assigned and col_ranges:
                                    valid_cols = {}
                                    for col in col_ranges.keys():
                                        x0, x1 = col_ranges[col]
                                        if center >= (x0 - 20) and center <= (x1 + 20):
                                            col_center = (x0 + x1) / 2
                                            valid_cols[col] = abs(center - col_center)
                                    
                                    if valid_cols:
                                        nearest = min(valid_cols.keys(), key=lambda c: valid_cols[c])
                                        if not descripcion_range or not (descripcion_range[0] <= center <= descripcion_range[1]):
                                            existing = prev.get(nearest) or ''
                                            if not existing or amt_text not in existing:
                                                if existing:
                                                    prev[nearest] = (existing + ' ' + amt_text).strip()
                                                else:
                                                    prev[nearest] = amt_text
                        
                        # Also merge amounts list for later processing
                        prev_amounts = prev.get('_amounts', [])
                        prev['_amounts'] = prev_amounts + cont_amounts
                        
                        # Collect possible text pieces from this row (prefer descripcion, then liq, then any other text)
                        cont_parts = []
                        for k in ('descripcion', 'fecha'):
                            v = row_data.get(k)
                            if v:
                                cont_parts.append(str(v))
                        # Also capture any stray text in other columns
                        for k, v in row_data.items():
                            if k in ('descripcion', 'fecha', 'cargos', 'abonos', 'saldo', 'page', '_amounts'):
                                continue
                            if v:
                                cont_parts.append(str(v))

                        cont_text = ' '.join(cont_parts)
                        # Remove decimal amounts (they belong to cargos/abonos/saldo)
                        cont_text = dec_amount_re.sub('', cont_text)
                        cont_text = ' '.join(cont_text.split()).strip()

                        if cont_text:
                            # append to previous 'descripcion' field
                            if prev.get('descripcion'):
                                prev['descripcion'] = (prev.get('descripcion') or '') + ' ' + cont_text
                            else:
                                prev['descripcion'] = cont_text
                        
                    else:
                        # No previous movement and no date - skip this row
                        # Only rows with dates should be added to movements
                        # Rows without dates are only used as continuation of previous rows
                        # For Konfio, also try to append rows with text to previous movement if exists
                        if bank_config['name'] == 'Konfio' and movement_rows:
                            all_row_text = ' '.join([w.get('text', '') for w in row_words]).strip()
                            # If row has any text, treat it as continuation
                            if all_row_text:
                                prev = movement_rows[-1]
                                
                                # Check if this is a sub-row pattern (like "SBO 020902DX1 - 02 abr 2023 4316 FÍSICA")
                                # These should be completely skipped, not added as continuation
                                konfio_sub_row_pattern = re.compile(r'^[A-Z]{2,4}\s*[A-Z0-9]{8,15}\s*-\s*\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4}', re.I)
                                if konfio_sub_row_pattern.match(all_row_text.strip()):
                                    # This is a sub-row, skip it completely (don't add to description)
                                    continue
                                
                                # Check if this row might contain a date (could be a new movement)
                                # If it does, don't treat it as continuation
                                potential_date = date_pattern.search(all_row_text)
                                
                                if not potential_date:
                                    # No date found, treat as continuation
                                    # IMPORTANT: First extract amounts from the continuation row before filtering text
                                    # This ensures amounts like "$2,266.83" are captured even if they're on the same line as "DIGITAL"
                                    
                                    # IMPORTANT: Always re-scan words directly for amounts to ensure we capture them
                                    # This is critical because amounts might not be detected in extract_movement_row if they're in continuation rows
                                    cont_amounts = []
                                    konfio_amount_pattern = re.compile(r'\$\s*\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})|\d{1,3}(?:[\.,\s]\d{3})*[\.,]\d{2}')
                                    for word in row_words:
                                        word_text = word.get('text', '').strip()
                                        word_x0 = word.get('x0', 0)
                                        word_x1 = word.get('x1', 0)
                                        word_center = (word_x0 + word_x1) / 2
                                        amount_match = konfio_amount_pattern.search(word_text)
                                        if amount_match:
                                            cont_amounts.append((amount_match.group(), word_center))
                                    
                                    # Also include amounts from row_data if any (in case they were detected earlier)
                                    existing_amounts = row_data.get('_amounts', [])
                                    if existing_amounts:
                                        # Merge with cont_amounts, avoiding duplicates
                                        existing_centers = {center for _, center in cont_amounts}
                                        for amt_text, center in existing_amounts:
                                            if center not in existing_centers:
                                                cont_amounts.append((amt_text, center))
                                    
                                    # Now filter text for description (but amounts are already captured above)
                                    cont_text_parts = []
                                    for word in row_words:
                                        word_text = word.get('text', '').strip()
                                        # Skip if it's an amount (has $ or decimal format)
                                        if not re.search(r'\$\s*\d|^\d+[\.,]\d{2}$', word_text):
                                            cont_text_parts.append(word_text)
                                    
                                    cont_text = ' '.join(cont_text_parts).strip()
                                    
                                    # For Konfio, filter out unwanted text patterns
                                    if cont_text:
                                        # Remove patterns like "SEB 1108096N7 - 02 abr 2023", "PCS 071026A78 - 02 abr 2023", "CACG551025EK9 - 02 abr 2023", etc.
                                        # Pattern: 2-4 letters + optional space + alphanumeric code (8-15 chars) + dash + date
                                        # Examples: "PCS 071026A78 - 02 abr 2023" (with space), "CACG551025EK9 - 02 abr 2023" (without space)
                                        konfio_second_line_pattern = re.compile(r'[A-Z]{2,4}\s*[A-Z0-9]{8,15}\s*-\s*\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4}', re.I)
                                        cont_text = konfio_second_line_pattern.sub('', cont_text)
                                        # Remove "FÍSICA" and "DIGITAL"
                                        cont_text = re.sub(r'\bFÍSICA\b', '', cont_text, flags=re.I)
                                        cont_text = re.sub(r'\bDIGITAL\b', '', cont_text, flags=re.I)
                                        # Normalize whitespace
                                        cont_text = ' '.join(cont_text.split()).strip()
                                    
                                    if cont_text:
                                        if prev.get('descripcion'):
                                            prev['descripcion'] = (prev.get('descripcion') or '') + ' ' + cont_text
                                        else:
                                            prev['descripcion'] = cont_text
                                    
                                    # Assign amounts from continuation row to cargos/abonos
                                    if cont_amounts and columns_config:
                                        # Get column ranges for numeric columns
                                        col_ranges = {}
                                        for col in ('cargos', 'abonos', 'saldo'):
                                            if col in columns_config:
                                                x0, x1 = columns_config[col]
                                                col_ranges[col] = (x0, x1)
                                        
                                        # Assign amounts from continuation row
                                        tolerance = 20  # Increased tolerance for Konfio to capture amounts that might be slightly misaligned
                                        for amt_text, center in cont_amounts:
                                            assigned = False
                                            for col in ('cargos', 'abonos', 'saldo'):
                                                if col in col_ranges:
                                                    x0, x1 = col_ranges[col]
                                                    if (x0 - tolerance) <= center <= (x1 + tolerance):
                                                        existing = prev.get(col) or ''
                                                        if not existing or amt_text not in existing:
                                                            if existing:
                                                                prev[col] = (existing + ' ' + amt_text).strip()
                                                            else:
                                                                prev[col] = amt_text
                                                        assigned = True
                                                        break
                                            
                                            # If not assigned by range, use proximity as fallback for Konfio
                                            if not assigned and bank_config['name'] == 'Konfio':
                                                # Find the nearest column
                                                nearest_col = None
                                                min_distance = float('inf')
                                                for col in ('cargos', 'abonos', 'saldo'):
                                                    if col in col_ranges:
                                                        x0, x1 = col_ranges[col]
                                                        col_center = (x0 + x1) / 2
                                                        distance = abs(center - col_center)
                                                        if distance < min_distance:
                                                            min_distance = distance
                                                            nearest_col = col
                                                
                                                # If amount is reasonably close to a column (within 50 pixels), assign it
                                                if nearest_col and min_distance < 50:
                                                    existing = prev.get(nearest_col) or ''
                                                    if not existing or amt_text not in existing:
                                                        if existing:
                                                            prev[nearest_col] = (existing + ' ' + amt_text).strip()
                                                        else:
                                                            prev[nearest_col] = amt_text
                                                    assigned = True
                                    # Merge amounts list
                                    prev_amounts = prev.get('_amounts', [])
                                    prev['_amounts'] = prev_amounts + row_data.get('_amounts', [])

    # Process summary lines to format them properly
    def format_summary_line(line):
        """Format a summary line and return list of (titulo, dato) tuples.
        Intelligently separates title:value pairs based on data type changes.
        """
        line = line.strip()
        if not line:
            return []
        
        results = []
        
        # Pattern for "Periodo DEL [date] AL [date]" (case insensitive)
        periodo_pattern = re.compile(r'periodo\s+del\s+(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\s+al\s+(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})', re.I)
        match = periodo_pattern.search(line)
        if match:
            fecha_inicio = match.group(1)
            fecha_fin = match.group(2)
            results.append(('Periodo DEL:', fecha_inicio))
            results.append(('AL:', fecha_fin))
            return results
        
        # Pattern for "No. de Cuenta [number]" or "No. de Cliente [number]" (case insensitive)
        cuenta_pattern = re.compile(r'(no\.?\s*de\s+(?:cuenta|cliente))\s+(.+)', re.I)
        match = cuenta_pattern.search(line)
        if match:
            titulo = match.group(1).strip()
            # Capitalize properly: "No. de Cuenta" or "No. de Cliente"
            if 'cuenta' in titulo.lower():
                titulo = 'No. de Cuenta'
            elif 'cliente' in titulo.lower():
                titulo = 'No. de Cliente'
            dato = match.group(2).strip()
            results.append((titulo + ':', dato))
            return results
        
        # Pattern for lines with "/" that separate multiple concepts (e.g., "Retiros / Cargos (-) 73 1,120,719.64")
        # Look for pattern: "Title1 / Title2 (optional) number1 number2"
        # More flexible: allows for optional parentheses and various number formats
        slash_pattern = re.compile(r'^([A-Za-z][A-Za-z\s]+?)\s*/\s*([A-Za-z][A-Za-z\s]*(?:\([^)]+\))?)\s+(\d+(?:,\d{3})*(?:\.\d{2})?)\s+([\d,\.\s]+)$', re.I)
        match = slash_pattern.search(line)
        if match:
            title1 = match.group(1).strip()
            title2 = match.group(2).strip()
            value1 = match.group(3).strip()
            value2 = match.group(4).strip()
            results.append((title1 + ':', value1))
            results.append((title2 + ':', value2))
            return results
        
        # Check if line is likely a company name/address (all caps, or contains common business terms)
        # Don't split these
        is_company_name = (
            line.isupper() or 
            re.search(r'\b(SA DE CV|S\.A\.|S\.A\. DE C\.V\.|S\.R\.L\.|INC\.|CORP\.|LLC)\b', line, re.I) or
            (len(line.split()) > 3 and not re.search(r'\d', line))  # Long text without numbers
        )
        if is_company_name:
            results.append((line, ''))
            return results
        
        # Pattern for lines that already have a colon
        if ':' in line:
            parts = line.split(':', 1)
            if len(parts) == 2:
                titulo = parts[0].strip()
                dato = parts[1].strip()
                results.append((titulo + ':', dato))
                return results
        
        # Intelligent splitting: detect data type changes
        # Split when we transition from text to number, or number to text
        tokens = line.split()
        if len(tokens) < 2:
            results.append((line, ''))
            return results
        
        # Find split points based on data type changes
        split_points = []
        for i in range(len(tokens) - 1):
            current = tokens[i]
            next_token = tokens[i + 1]
            
            # Check if current is text and next is number (or vice versa)
            # Better number detection: allows for formatted numbers with commas and decimals
            # Pattern: digits with optional thousands separators (commas) and optional decimal part
            num_pattern = re.compile(r'^\d{1,3}(?:,\d{3})*(?:\.\d{2})?$|^\d+(?:\.\d{2})?$')
            current_is_num = bool(num_pattern.match(current))
            next_is_num = bool(num_pattern.match(next_token))
            
            # Also check for date patterns
            current_is_date = bool(re.match(r'^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$', current))
            next_is_date = bool(re.match(r'^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$', next_token))
            
            # Split if: text->number, number->text, or text->date
            if (not current_is_num and not current_is_date and (next_is_num or next_is_date)):
                split_points.append(i + 1)
            elif (current_is_num or current_is_date) and not next_is_num and not next_is_date:
                # Number/date followed by text - might be a new field
                # But be careful, don't split if it's part of a longer description
                if i < len(tokens) - 2:  # Not the last token
                    split_points.append(i + 1)
        
        # If we found split points, use them
        if split_points:
            # Use first split point
            split_idx = split_points[0]
            titulo = ' '.join(tokens[:split_idx]).strip()
            dato = ' '.join(tokens[split_idx:]).strip()
            
            # Clean up título (remove trailing special chars, add colon)
            titulo = re.sub(r'[:\-]+$', '', titulo).strip()
            if not titulo.endswith(':'):
                titulo += ':'
            
            results.append((titulo, dato))
            return results
        
        # Fallback: try to split on multiple spaces or first number
        kv_match = re.match(r'^([A-Za-z][A-Za-z\s]+(?:de|del|la|el|los|las)?)\s{2,}(.+)$', line, re.I)
        if kv_match:
            titulo = kv_match.group(1).strip()
            dato = kv_match.group(2).strip()
            if not titulo.endswith(':'):
                titulo += ':'
            results.append((titulo, dato))
            return results
        
        # Try to split on first number or date
        num_match = re.search(r'(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}|[\d,\.]+)', line)
        if num_match:
            split_pos = num_match.start()
            if split_pos > 0:
                titulo = line[:split_pos].strip()
                dato = line[split_pos:].strip()
                titulo = re.sub(r'[:\-]+$', '', titulo).strip()
                if not titulo.endswith(':'):
                    titulo += ':'
                results.append((titulo, dato))
                return results
        
        # If no pattern matches, return as título with empty dato
        results.append((line, ''))
        return results
    
    # Process all summary lines
    summary_rows = []
    for line in summary_lines:
        formatted = format_summary_line(line)
        summary_rows.extend(formatted)
    
    # Create DataFrame with two columns: Título and Dato
    if summary_rows:
        df_summary = pd.DataFrame(summary_rows, columns=['Título', 'Dato'])
    else:
        df_summary = pd.DataFrame({'Título': [], 'Dato': []})
    
    # Reassign amounts to cargos/abonos/saldo by proximity when needed
    # Process movement_rows for all banks including Konfio (now uses coordinate-based extraction)
    if movement_rows:
        # prepare column centers and ranges
        col_centers = {}
        col_ranges = {}
        descripcion_range = None
        
        # Get description range to exclude amounts from it
        if 'descripcion' in columns_config:
            x0, x1 = columns_config['descripcion']
            descripcion_range = (x0, x1)
        
        for col in ('cargos', 'abonos', 'saldo'):
            if col in columns_config:
                x0, x1 = columns_config[col]
                col_centers[col] = (x0 + x1) / 2
                col_ranges[col] = (x0, x1)

        for row_idx_debug, r in enumerate(movement_rows):
            amounts = r.get('_amounts', [])
            
            # Debug for Banamex: print before processing amounts (only first 5 rows)
            if bank_config['name'] == 'Banamex' and row_idx_debug < 5:
                fecha_val = str(r.get('fecha') or '').strip()
                desc_val = str(r.get('descripcion') or '').strip()[:60]
                cargos_val = str(r.get('cargos') or '').strip()
            if not amounts:
                continue

            # For Konfio, filter out account numbers (like "3817") from amounts list
            # Account numbers are typically 4-digit numbers without $ or decimal format
            if bank_config['name'] == 'Konfio':
                filtered_amounts = []
                for amt_text, center in amounts:
                    # Check if this looks like an account number (4 digits, no $, no decimal part)
                    account_number_pattern = re.compile(r'^\d{4}$')
                    if not account_number_pattern.match(amt_text.strip()):
                        # Not an account number, keep it
                        filtered_amounts.append((amt_text, center))
                    else:
                        # This is an account number, add it to description instead
                        if r.get('descripcion'):
                            r['descripcion'] = r['descripcion'] + ' ' + amt_text.strip()
                        else:
                            r['descripcion'] = amt_text.strip()
                amounts = filtered_amounts

            # Check if columns already have values from initial extraction
            # If they do, we should preserve them unless we find better matches
            existing_cargos = r.get('cargos', '').strip()
            existing_abonos = r.get('abonos', '').strip()
            existing_saldo = r.get('saldo', '').strip()

            # If columns already have numbers, keep them but prefer reassignment
            # We'll assign each detected amount to the appropriate numeric column
            # ONLY if it's within the column's coordinate range
            for amt_text, center in amounts:
                # Find which numeric column this amount belongs to based on coordinate range
                # Use a small tolerance for edge cases (amounts near column boundaries)
                # For Konfio, use larger tolerance to capture amounts that might be slightly misaligned
                # The monto "$2,266.83" has coordinate 410.22, and cargos range is (340, 388)
                # So we need a tolerance of at least 22 to capture it (410.22 - 388 = 22.22)
                tolerance = 30 if bank_config['name'] == 'Konfio' else 10
                assigned = False
                
                # First, check if amount is within any numeric column range
                # This takes priority over description range check
                for col in ('cargos', 'abonos', 'saldo'):
                    if col in col_ranges:
                        x0, x1 = col_ranges[col]
                        # Check with tolerance to handle amounts slightly outside the range
                        # Special debug for "IVA SOBRE COMISIONES E INTERESES" row
                        if bank_config['name'] == 'Konfio':
                            desc_val_check = str(r.get('descripcion') or '').strip()
                            if 'IVA SOBRE COMISIONES E INTERESES' in desc_val_check.upper():
                                in_range = (x0 - tolerance) <= center <= (x1 + tolerance)
                                #print(f"   Verificando columna {col}: rango ({x0}, {x1}), centro={center:.2f}, dentro del rango? {in_range} (rango efectivo: {x0 - tolerance:.2f} a {x1 + tolerance:.2f})")
                        
                        if (x0 - tolerance) <= center <= (x1 + tolerance):
                            # Amount is within a numeric column range - assign it regardless of description range
                            # Amount is within this column's range (with tolerance)
                            existing = r.get(col, '').strip()
                            
                            # Check if this amount is already in the column (to avoid duplicates)
                            if existing and amt_text in existing:
                                assigned = True
                                break
                            # Only assign if column is empty or if this amount is not already there
                            # IMPORTANT: Preserve existing values if they're already valid numbers
                            if not existing:
                                # Column is empty, assign the amount
                                r[col] = amt_text
                                assigned = True
                                break
                            elif amt_text not in existing:
                                # Column has a value but this amount is different
                                # Check if existing is a valid amount (has digits and decimal)
                                if DEC_AMOUNT_RE.search(existing):
                                    # Existing looks like a valid amount - preserve it
                                    # Don't overwrite or append, just keep the existing value
                                    # This preserves values extracted during initial coordinate-based extraction
                                    assigned = True
                                    break
                                else:
                                    # Existing doesn't look like an amount, replace it
                                    r[col] = amt_text
                                    assigned = True
                                    break
                            else:
                                # Amount already in column
                                assigned = True
                                break
                
                # If not assigned by range, use proximity as fallback
                # Only exclude from description range if it's NOT in any numeric column range
                if not assigned and col_centers:
                    # Calculate distances, but only consider columns that are reasonably close
                    # For Konfio, use larger tolerance to capture amounts that might be slightly misaligned
                    proximity_check_tolerance = 50 if bank_config['name'] == 'Konfio' else 20
                    valid_cols = {}
                    for col in col_centers.keys():
                        if col in col_ranges:
                            x0, x1 = col_ranges[col]
                            # Only consider if center is reasonably close to the column
                            if center >= (x0 - proximity_check_tolerance) and center <= (x1 + proximity_check_tolerance):
                                valid_cols[col] = abs(center - col_centers[col])
                    
                    if valid_cols:
                        nearest = min(valid_cols.keys(), key=lambda c: valid_cols[c])
                        nearest_distance = valid_cols[nearest]
                        
                        # Check if amount is in description range AND not in any numeric column
                        # If it's close to a numeric column, assign it even if it's also in description range
                        in_desc_range = descripcion_range and descripcion_range[0] <= center <= descripcion_range[1]
                        in_num_range = False
                        # For Konfio, use larger tolerance for proximity check
                        proximity_tolerance = 50 if bank_config['name'] == 'Konfio' else 20
                        for col in col_ranges.keys():
                            x0, x1 = col_ranges[col]
                            if (x0 - proximity_tolerance) <= center <= (x1 + proximity_tolerance):
                                in_num_range = True
                                break
                        
                        # Only skip if in description range AND NOT in any numeric column range
                        # For Konfio, be more lenient - if amount is close to a numeric column, assign it
                        if bank_config['name'] == 'Konfio' and nearest_distance < 50:
                            # For Konfio, if amount is within 50 pixels of nearest column, assign it
                            existing = r.get(nearest, '').strip()
                            if not existing or amt_text not in existing:
                                if existing:
                                    r[nearest] = (existing + ' ' + amt_text).strip()
                                else:
                                    r[nearest] = amt_text
                                assigned = True
                        elif not in_desc_range or in_num_range:
                            existing = r.get(nearest, '').strip()
                            if existing:
                                # If existing is a valid amount, preserve it
                                if DEC_AMOUNT_RE.search(existing):
                                    # Existing is a valid amount, preserve it
                                    assigned = True
                                    break
                                elif amt_text not in existing:
                                    r[nearest] = (existing + ' ' + amt_text).strip()
                                    assigned = True
                                    break
                            else:
                                r[nearest] = amt_text
                                assigned = True
                                break
                        
                        # For Konfio, if still not assigned and amount is reasonably close, assign it anyway
                        if not assigned and bank_config['name'] == 'Konfio':
                            if nearest_distance < 100:  # Within 100 pixels
                                existing = r.get(nearest, '').strip()
                                if not existing or amt_text not in existing:
                                    if existing:
                                        r[nearest] = (existing + ' ' + amt_text).strip()
                                    else:
                                        r[nearest] = amt_text
                                    assigned = True
                
                # Only skip if amount is in description range AND NOT assigned to any numeric column
                # This prevents amounts in cargos/abonos/saldo from being skipped
                # For Konfio, if amount is not assigned and is close to cargos/abonos columns, assign it anyway
                if not assigned:
                    if bank_config['name'] == 'Konfio':
                        # For Konfio, if amount is reasonably close to cargos or abonos, assign it
                        # The monto "$2,266.83" has coordinate 410.22, which is close to cargos (340-388)
                        for col in ('cargos', 'abonos'):
                            if col in col_ranges:
                                x0, x1 = col_ranges[col]
                                col_center = (x0 + x1) / 2
                                distance = abs(center - col_center)
                                # If within 50 pixels of column center, assign it
                                if distance < 50:
                                    existing = r.get(col, '').strip()
                                    if not existing or amt_text not in existing:
                                        if existing:
                                            r[col] = (existing + ' ' + amt_text).strip()
                                        else:
                                            r[col] = amt_text
                                        
                                        # Special debug for "IVA SOBRE COMISIONES E INTERESES" row
                                        desc_val_check = str(r.get('descripcion') or '').strip()
                                        if 'IVA SOBRE COMISIONES E INTERESES' in desc_val_check.upper():
                                            print(f"   ✅ Monto '{amt_text}' asignado a {col} por proximidad (distancia: {distance:.2f})")
                                        assigned = True
                                        break
                        if assigned:
                            continue
                    
                    if descripcion_range:
                        if descripcion_range[0] <= center <= descripcion_range[1]:
                            # Amount is in description range and wasn't assigned to any numeric column
                            # Skip it to avoid assigning description amounts to numeric columns
                            continue

            # Remove amount tokens from descripcion if present
            # For Banamex: be more careful - only remove amounts that are standalone, not part of text like "IVA $2865.60"
            if r.get('descripcion'):
                if bank_config['name'] == 'Banamex':
                    # For Banamex, don't remove amounts that are part of descriptive text
                    # Only remove amounts that appear to be standalone (at the end, with spaces around them)
                    desc_val = str(r.get('descripcion', '')).strip()
                    # Pattern to match standalone amounts (at end of string, with space before)
                    standalone_amount_pattern = re.compile(r'\s+(\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2}))\s*$')
                    # Only remove if it's a standalone amount at the end
                    desc_val_cleaned = standalone_amount_pattern.sub('', desc_val)
                    if desc_val_cleaned != desc_val:
                        r['descripcion'] = desc_val_cleaned.strip()
                else:
                    # For other banks, use the original logic
                    r['descripcion'] = DEC_AMOUNT_RE.sub('', r.get('descripcion'))

            # For Banamex, when there are 2 amounts, ensure correct assignment based on X coordinates
            # IMPORTANT: This runs AFTER the general _amounts processing, so we should only
            # reassign if values are missing or incorrect, not if they're already correctly assigned
            # Since extract_movement_row already assigns values correctly based on coordinates,
            # this logic is only needed as a fallback if values are missing
            if bank_config['name'] == 'Banamex' and columns_config:
                amounts = r.get('_amounts', [])
                # Get current assignments
                existing_cargos = r.get('cargos', '').strip()
                existing_abonos = r.get('abonos', '').strip()
                existing_saldo = r.get('saldo', '').strip()
                
                # If all values are already assigned correctly, skip reassignment completely
                # This preserves values from extract_movement_row which uses coordinate-based assignment
                if not ((existing_cargos or existing_abonos) and existing_saldo) and len(amounts) == 2:
                    # Only execute if values are missing - this is a fallback
                    
                    # Get X coordinates for each amount
                    first_amt_text, first_amt_center = amounts[0]
                    second_amt_text, second_amt_center = amounts[1]
                    
                    # Assign second amount to saldo if in saldo range and saldo is empty
                    if not existing_saldo and 'saldo' in columns_config:
                        saldo_x0, saldo_x1 = columns_config['saldo']
                        tolerance = 10
                        if (saldo_x0 - tolerance) <= second_amt_center <= (saldo_x1 + tolerance):
                            r['saldo'] = second_amt_text
                            existing_saldo = second_amt_text
                    
                    # Assign first amount based on X coordinate if not already assigned
                    if not existing_cargos and not existing_abonos and 'cargos' in columns_config and 'abonos' in columns_config:
                        cargos_x0, cargos_x1 = columns_config['cargos']
                        abonos_x0, abonos_x1 = columns_config['abonos']
                        tolerance = 10
                        
                        # Check if first amount is in cargos range
                        if (cargos_x0 - tolerance) <= first_amt_center <= (cargos_x1 + tolerance):
                            r['cargos'] = first_amt_text
                        # Check if first amount is in abonos range
                        elif (abonos_x0 - tolerance) <= first_amt_center <= (abonos_x1 + tolerance):
                            r['abonos'] = first_amt_text
                        else:
                            # Fallback: use proximity to determine
                            cargos_center = (cargos_x0 + cargos_x1) / 2
                            abonos_center = (abonos_x0 + abonos_x1) / 2
                            dist_cargos = abs(first_amt_center - cargos_center)
                            dist_abonos = abs(first_amt_center - abonos_center)
                            
                            if dist_cargos < dist_abonos:
                                r['cargos'] = first_amt_text
                            else:
                                r['abonos'] = first_amt_text

            # cleanup helper key
            if '_amounts' in r:
                del r['_amounts']

            # Debug: Print row data for BBVA to see what's happening with cargos
            if bank_config['name'] == 'BBVA' and len([x for x in movement_rows if x.get('cargos')]) <= 5:  # Only print first 5 rows with cargos for debugging
                cargos_val = r.get('cargos', '')
                abonos_val = r.get('abonos', '')
                saldo_val = r.get('saldo', '')
                if cargos_val or abonos_val or saldo_val:
                    pass
                    # print(f"DEBUG BBVA Row: cargos='{cargos_val}', abonos='{abonos_val}', saldo='{saldo_val}'")
        
        # If columns_config was empty, we may have rows but without cargos/abonos/saldo
        # Try to extract them from raw text or _amounts
        if movement_rows and not columns_config and bank_config['name'] != 'BBVA':
            has_saldo = 'saldo' in columns_config
            for r in movement_rows:
                # If row has 'raw', extract from it
                if 'raw' in r and r.get('raw'):
                    raw_text = str(r.get('raw'))
                    amounts = DEC_AMOUNT_RE.findall(raw_text)
                    if len(amounts) >= 3:
                        r['cargos'] = amounts[-3]
                        r['abonos'] = amounts[-2]
                        if has_saldo:
                            r['saldo'] = amounts[-1]
                    elif len(amounts) == 2:
                        r['abonos'] = amounts[0]
                        if has_saldo:
                            r['saldo'] = amounts[1]
                    elif len(amounts) == 1:
                        if has_saldo:
                            r['saldo'] = amounts[0]
                # If row has _amounts but no cargos/abonos/saldo, try to assign them
                elif '_amounts' in r and r.get('_amounts'):
                    amounts_list = [amt for amt, _ in r.get('_amounts', [])]
                    if len(amounts_list) >= 3:
                        r['cargos'] = amounts_list[-3]
                        r['abonos'] = amounts_list[-2]
                        if has_saldo:
                            r['saldo'] = amounts_list[-1]
                    elif len(amounts_list) == 2:
                        r['abonos'] = amounts_list[0]
                        if has_saldo:
                            r['saldo'] = amounts_list[1]
                    elif len(amounts_list) == 1:
                        if has_saldo:
                            r['saldo'] = amounts_list[0]

        # Create df_mov from movement_rows if not already created
        if df_mov is None:
            df_mov = pd.DataFrame(movement_rows) if movement_rows else pd.DataFrame(columns=['fecha', 'descripcion', 'cargos', 'abonos', 'saldo'])                
    else:
        # No coordinate-based extraction available, use raw text extraction
        movement_entries = group_entries_from_lines(movements_lines)
        df_mov = pd.DataFrame({'raw': movement_entries})
        
        # For non-BBVA banks, try to extract Cargos, Abonos, Saldo from raw text
        if bank_config['name'] != 'BBVA' and len(df_mov) > 0:
            has_saldo = 'saldo' in columns_config
            def extract_amounts_from_raw(raw_text):
                """Extract Cargos, Abonos, and Saldo from raw text line."""
                if not raw_text or not isinstance(raw_text, str):
                    return {'cargos': None, 'abonos': None, 'saldo': None}
                
                # Find all amounts in the line
                amounts = DEC_AMOUNT_RE.findall(str(raw_text))
                
                # Common patterns in bank statements:
                # - Usually: Date Description Amount1 Amount2 Amount3
                # - Or: Date Description Cargos Abonos Saldo
                # - Or: Date Description Amount (could be any of the three)
                
                result = {'cargos': None, 'abonos': None, 'saldo': None}
                
                # If we have amounts, try to identify them
                # Typically, the last amount is Saldo, and before that could be Cargos/Abonos
                if len(amounts) >= 3:
                    # Three amounts: likely Cargos, Abonos, Saldo
                    result['cargos'] = amounts[-3]
                    result['abonos'] = amounts[-2]
                    if has_saldo:
                        result['saldo'] = amounts[-1]
                elif len(amounts) == 2:
                    # Two amounts: could be Cargos/Abonos and Saldo, or just two of them
                    # Check if there are keywords in the text
                    text_lower = raw_text.lower()
                    if 'cargo' in text_lower or 'retiro' in text_lower:
                        result['cargos'] = amounts[0]
                        if has_saldo:
                            result['saldo'] = amounts[1]
                    elif 'abono' in text_lower or 'deposito' in text_lower:
                        result['abonos'] = amounts[0]
                        if has_saldo:
                            result['saldo'] = amounts[1]
                    else:
                        # Default: first is Abonos (more common), second is Saldo (if exists)
                        result['abonos'] = amounts[0]
                        if has_saldo:
                            result['saldo'] = amounts[1]
                        else:
                            # If no saldo column, second amount could be cargos
                            result['cargos'] = amounts[1]
                elif len(amounts) == 1:
                    # One amount: could be any of them, check keywords
                    text_lower = raw_text.lower()
                    if 'cargo' in text_lower or 'retiro' in text_lower:
                        result['cargos'] = amounts[0]
                    elif 'abono' in text_lower or 'deposito' in text_lower:
                        result['abonos'] = amounts[0]
                    else:
                        # Default: assume it's Saldo (if exists), otherwise Abonos
                        if has_saldo:
                            result['saldo'] = amounts[0]
                        else:
                            result['abonos'] = amounts[0]
                
                return result
            
            # Extract amounts from raw text
            extracted = df_mov['raw'].apply(extract_amounts_from_raw)
            df_mov['cargos'] = extracted.apply(lambda x: x.get('cargos'))
            df_mov['abonos'] = extracted.apply(lambda x: x.get('abonos'))
            if has_saldo:
                df_mov['saldo'] = extracted.apply(lambda x: x.get('saldo'))

    # Split combined fecha values into two separate columns: Fecha Oper and Fecha Liq
    # Works for coordinate-based extraction (column 'fecha') and for fallback raw lines ('raw').
    # Pattern for dates: supports multiple formats:
    # - "DIA MES" (01 ABR)
    # - "MES DIA" (ABR 01)
    # - "DIA MES AÑO" (06 mar 2023) - for Konfio
    date_pattern = re.compile(r"(?:(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4})", re.I)

    def _extract_two_dates(txt):
        if not txt or not isinstance(txt, str):
            return (None, None)
        found = date_pattern.findall(txt)
        if not found:
            return (None, None)
        if len(found) == 1:
            return (found[0], None)
        # If more than two, take first two
        return (found[0], found[1])

    # Initialize dates variable to avoid UnboundLocalError
    dates = None
    
    # For Banregio, preserve the 'fecha' column as-is (it's already 2 digits)
    if bank_config['name'] == 'Banregio' and 'fecha' in df_mov.columns:
        df_mov['Fecha'] = df_mov['fecha'].astype(str)
        df_mov = df_mov.drop(columns=['fecha'])
    elif bank_config['name'] == 'Base' and 'fecha' in df_mov.columns:
        # For Base, preserve the date format DD/MM/YYYY as-is
        df_mov['Fecha'] = df_mov['fecha'].astype(str)
        df_mov = df_mov.drop(columns=['fecha'])
    elif bank_config['name'] == 'Banbajío' and 'fecha' in df_mov.columns:
        # For BanBajío, preserve the date format "DIA MES" (e.g., "3 ENE") as-is
        df_mov['Fecha'] = df_mov['fecha'].astype(str)
        df_mov = df_mov.drop(columns=['fecha'])
    elif bank_config['name'] == 'BBVA':
        # For BBVA, extract dates from separate 'fecha' and 'liq' columns
        fecha_oper_dates = None
        fecha_liq_dates = None
        
        if 'fecha' in df_mov.columns:
            fecha_oper_dates = df_mov['fecha'].astype(str).apply(_extract_two_dates)
        elif 'raw' in df_mov.columns:
            fecha_oper_dates = df_mov['raw'].astype(str).apply(_extract_two_dates)
        else:
            fecha_oper_dates = pd.Series([(None, None)] * len(df_mov))
        
        if 'liq' in df_mov.columns:
            fecha_liq_dates = df_mov['liq'].astype(str).apply(_extract_two_dates)
        else:
            fecha_liq_dates = pd.Series([(None, None)] * len(df_mov))
        
        # Extract first date from 'fecha' column for Fecha Oper
        df_mov['Fecha Oper'] = fecha_oper_dates.apply(lambda t: t[0])
        # Extract first date from 'liq' column for Fecha Liq
        df_mov['Fecha Liq'] = fecha_liq_dates.apply(lambda t: t[0])
        
        # Remove original 'fecha' column (liq will be removed later when building description)
        if 'fecha' in df_mov.columns:
            df_mov = df_mov.drop(columns=['fecha'])
    elif bank_config['name'] == 'HSBC':
        # For HSBC, preserve 'fecha' column as-is (will be renamed to 'Fecha' later)
        # Don't create 'Fecha Oper' or 'Fecha Liq' for HSBC
        pass
    else:
        # For other banks, use existing logic (search for two dates in 'fecha' column)
        if 'fecha' in df_mov.columns:
            dates = df_mov['fecha'].astype(str).apply(_extract_two_dates)
        elif 'raw' in df_mov.columns:
            dates = df_mov['raw'].astype(str).apply(_extract_two_dates)
        else:
            dates = pd.Series([(None, None)] * len(df_mov))

        if dates is not None:
            df_mov['Fecha Oper'] = dates.apply(lambda t: t[0])
            df_mov['Fecha Liq'] = dates.apply(lambda t: t[1])

    # Remove original 'fecha' if present (only if not already removed)
    # Skip this for HSBC - HSBC will handle 'fecha' to 'Fecha' conversion in its own block
    if 'fecha' in df_mov.columns and bank_config['name'] != 'HSBC':
        df_mov = df_mov.drop(columns=['fecha'])
    
    # For non-BBVA banks, use only 'Fecha' column (based on Fecha Oper) and remove Fecha Liq
    # Skip this for Banregio, Base, BanBajío, and HSBC (HSBC already has 'Fecha' from OCR or will be set below)
    if bank_config['name'] != 'BBVA' and bank_config['name'] != 'Banregio' and bank_config['name'] != 'Base' and bank_config['name'] != 'Banbajío' and bank_config['name'] != 'HSBC':
        if 'Fecha Oper' in df_mov.columns:
            df_mov['Fecha'] = df_mov['Fecha Oper']
            df_mov = df_mov.drop(columns=['Fecha Oper', 'Fecha Liq'])
    
    # For HSBC, ensure 'Fecha' column exists and remove all unwanted columns
    if bank_config['name'] == 'HSBC':
        # Convert 'fecha' to 'Fecha' if needed (DO THIS FIRST before removing columns)
        if 'fecha' in df_mov.columns:
            if 'Fecha' not in df_mov.columns:
                df_mov['Fecha'] = df_mov['fecha']
            # Remove 'fecha' after creating 'Fecha'
            df_mov = df_mov.drop(columns=['fecha'])
        elif 'Fecha Oper' in df_mov.columns and 'Fecha' not in df_mov.columns:
            df_mov['Fecha'] = df_mov['Fecha Oper']
        
        # Remove ALL other unwanted columns for HSBC
        unwanted_cols = ['Fecha Oper', 'Fecha Liq', 'raw', 'liq']
        for col in unwanted_cols:
            if col in df_mov.columns:
                df_mov = df_mov.drop(columns=[col])

    # For BBVA, split 'saldo' column into 'OPERACIÓN' and 'LIQUIDACIÓN'
    if bank_config['name'] == 'BBVA' and 'saldo' in df_mov.columns:
        def _extract_two_amounts(txt):
            """Extract two amounts from the saldo column (OPERACIÓN and LIQUIDACIÓN)."""
            if not txt or not isinstance(txt, str):
                return (None, None)
            # Find all amounts matching the decimal pattern
            amounts = DEC_AMOUNT_RE.findall(str(txt))
            if len(amounts) >= 2:
                # Check if both amounts are the same (normalize by removing separators)
                # Normalize: remove commas, spaces, and compare numeric parts
                def normalize_amount(amt):
                    # Remove commas, spaces, keep only digits and decimal separator
                    normalized = re.sub(r'[,\s]', '', amt)
                    return normalized
                
                amt1_normalized = normalize_amount(amounts[0])
                amt2_normalized = normalize_amount(amounts[1])
                
                if amt1_normalized == amt2_normalized:
                    # If amounts are the same, return the same value for both columns
                    return (amounts[0], amounts[0])
                # Return first two amounts (first is LIQUIDACIÓN, second is OPERACIÓN)
                return (amounts[0], amounts[1])
            elif len(amounts) == 1:
                # Only one amount found - assign it to both columns
                return (amounts[0], amounts[0])
            else:
                return (None, None)
        
        # Extract the two amounts from saldo column
        amounts = df_mov['saldo'].astype(str).apply(_extract_two_amounts)
        df_mov['OPERACIÓN'] = amounts.apply(lambda t: t[1])  # Second amount is OPERACIÓN
        df_mov['LIQUIDACIÓN'] = amounts.apply(lambda t: t[0])  # First amount is LIQUIDACIÓN
        
        # Remove the original 'saldo' column
        df_mov = df_mov.drop(columns=['saldo'])

    # Merge 'liq' and 'descripcion' into a single 'Descripcion' column.
    # Remove any date tokens and decimal amounts from the description text.
    dec_amount_re = re.compile(r"\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")
    # date_pattern already defined above; reuse it

    def _build_description(row):
        parts = []
        # For BBVA, don't include 'liq' in description (it only contains the date which is already extracted)
        if bank_config['name'] == 'BBVA':
            # Only use 'descripcion' column for BBVA
            if 'descripcion' in row and row.get('descripcion'):
                parts.append(str(row.get('descripcion')))
            # fallback to raw if no column-based description
            if not parts and 'raw' in row and row.get('raw'):
                parts = [str(row.get('raw'))]
        elif bank_config['name'] == 'HSBC':
            # For HSBC, only use 'descripcion' column (don't use 'raw' or 'liq')
            if 'descripcion' in row and row.get('descripcion'):
                parts.append(str(row.get('descripcion')))
        else:
            # For other banks, use existing logic
            # prefer explicit columns if present
            if 'liq' in row and row.get('liq'):
                parts.append(str(row.get('liq')))
            if 'descripcion' in row and row.get('descripcion'):
                parts.append(str(row.get('descripcion')))
            # fallback to raw if no column-based description
            if not parts and 'raw' in row and row.get('raw'):
                parts = [str(row.get('raw'))]

        text = ' '.join(parts)
        # remove extracted dates from description (handle both BBVA and non-BBVA formats)
        if 'Fecha Oper' in row:
            fo = (row.get('Fecha Oper') or '')
            if fo:
                text = text.replace(str(fo), '')
        if 'Fecha Liq' in row:
            fl = (row.get('Fecha Liq') or '')
            if fl:
                text = text.replace(str(fl), '')
        if 'Fecha' in row:
            fecha = (row.get('Fecha') or '')
            if fecha:
                text = text.replace(str(fecha), '')

        # Remove decimal amounts (they belong to cargos/abonos/saldo)
        text = dec_amount_re.sub('', text)

        # For Konfio, remove specific unwanted text patterns
        if bank_config['name'] == 'Konfio':
            # Remove patterns like "SEB 1108096N7 - 02 abr 2023", "PCS 071026A78 - 02 abr 2023", "CACG551025EK9 - 02 abr 2023", etc.
            # Pattern: 2-4 letters + optional space + alphanumeric code (8-15 chars) + dash + date
            # Examples: "PCS 071026A78 - 02 abr 2023" (with space), "CACG551025EK9 - 02 abr 2023" (without space)
            konfio_second_line_pattern = re.compile(r'[A-Z]{2,4}\s*[A-Z0-9]{8,15}\s*-\s*\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4}', re.I)
            text = konfio_second_line_pattern.sub('', text)
            # Remove "FÍSICA" and "DIGITAL"
            text = re.sub(r'\bFÍSICA\b', '', text, flags=re.I)
            text = re.sub(r'\bDIGITAL\b', '', text, flags=re.I)

        # Normalize whitespace
        text = ' '.join(text.split()).strip()
        return text if text else None

    # Apply description building
    df_mov['Descripcion'] = df_mov.apply(_build_description, axis=1)
    

    # Drop old columns used to build description
    drop_cols = []
    for c in ('liq', 'descripcion'):
        if c in df_mov.columns:
            drop_cols.append(c)
    # For HSBC, also drop 'raw' column
    if bank_config['name'] == 'HSBC' and 'raw' in df_mov.columns:
        drop_cols.append('raw')
    if drop_cols:
        df_mov = df_mov.drop(columns=drop_cols)


    # Rename columns to match desired format
    column_rename = {}
    if bank_config['name'] == 'BBVA':
        if 'Descripcion' in df_mov.columns:
            column_rename['Descripcion'] = 'Descripción'
        if 'cargos' in df_mov.columns:
            column_rename['cargos'] = 'Cargos'
        if 'abonos' in df_mov.columns:
            column_rename['abonos'] = 'Abonos'
        if 'OPERACIÓN' in df_mov.columns:
            column_rename['OPERACIÓN'] = 'Operación'
        if 'LIQUIDACIÓN' in df_mov.columns:
            column_rename['LIQUIDACIÓN'] = 'Liquidación'
    else:
        # For HSBC and other banks, rename columns
        # For HSBC, 'fecha' was already converted to 'Fecha' earlier, so skip if 'Fecha' already exists
        if 'fecha' in df_mov.columns and 'Fecha' not in df_mov.columns:
            column_rename['fecha'] = 'Fecha'
        if 'Descripcion' in df_mov.columns:
            column_rename['Descripcion'] = 'Descripción'
        if 'descripcion' in df_mov.columns:
            column_rename['descripcion'] = 'Descripción'
        if 'cargos' in df_mov.columns:
            column_rename['cargos'] = 'Cargos'
        if 'abonos' in df_mov.columns:
            column_rename['abonos'] = 'Abonos'
        if 'saldo' in df_mov.columns:
            column_rename['saldo'] = 'Saldo'
    
    if column_rename:
        df_mov = df_mov.rename(columns=column_rename)
        
    # For HSBC, remove 'raw' column if it still exists (after renaming)
    if bank_config['name'] == 'HSBC' and 'raw' in df_mov.columns:
        df_mov = df_mov.drop(columns=['raw'])

    # For Clara, clean "MXN" and other text from Abonos column
    if bank_config['name'] == 'Clara' and 'Abonos' in df_mov.columns:
        def clean_clara_abonos(value):
            """Remove MXN and other non-numeric text from Abonos values for Clara."""
            if pd.isna(value) or not value:
                return ''
            value_str = str(value).strip()
            # Remove "MXN" (case insensitive) and any surrounding spaces
            value_str = re.sub(r'\bMXN\b', '', value_str, flags=re.I)
            # Remove any remaining non-numeric characters except digits, commas, dots, and minus sign
            # Keep only: digits, commas, dots (for decimals), and minus sign (for negative values)
            value_str = re.sub(r'[^\d,\.\-]', '', value_str)
            return value_str.strip()
        
        df_mov['Abonos'] = df_mov['Abonos'].apply(clean_clara_abonos)
    
    # For Konfio, clean "$" and other text from Cargos and Abonos columns
    if bank_config['name'] == 'Konfio':
        def clean_konfio_amounts(value):
            """Remove $ and other non-numeric text from amount values for Konfio."""
            if pd.isna(value):
                return ''
            value_str = str(value).strip()
            if not value_str or value_str == '':
                return ''
            # Remove "$" and any surrounding spaces
            value_str = re.sub(r'\$\s*', '', value_str)
            # Remove any remaining non-numeric characters except digits, commas, dots, and minus sign
            # Keep only: digits, commas, dots (for decimals), and minus sign (for negative values)
            value_str = re.sub(r'[^\d,\.\-]', '', value_str)
            cleaned = value_str.strip()
            # Return empty string if nothing remains after cleaning
            return cleaned if cleaned else ''
        
        if 'Cargos' in df_mov.columns:
            df_mov['Cargos'] = df_mov['Cargos'].apply(clean_konfio_amounts)
        
        if 'Abonos' in df_mov.columns:
            df_mov['Abonos'] = df_mov['Abonos'].apply(clean_konfio_amounts)

    # Rename "Fecha Liq" to "Fecha Liq." for BBVA if needed
    if bank_config['name'] == 'BBVA' and 'Fecha Liq' in df_mov.columns:
        df_mov = df_mov.rename(columns={'Fecha Liq': 'Fecha Liq.'})

    # Reorder columns according to bank type
    if bank_config['name'] == 'BBVA':
        # For BBVA: Fecha Oper, Fecha Liq., Descripción, Abonos, Cargos, Operación, Liquidación
        desired_order = ['Fecha Oper', 'Fecha Liq.', 'Descripción', 'Abonos', 'Cargos', 'Operación', 'Liquidación']
        # Filter to only include columns that exist in the dataframe
        desired_order = [col for col in desired_order if col in df_mov.columns]
        # Add any remaining columns that are not in desired_order
        other_cols = [c for c in df_mov.columns if c not in desired_order]
        df_mov = df_mov[desired_order + other_cols]
    elif bank_config['name'] == 'HSBC':
        # For HSBC: ONLY Fecha, Descripción, Abonos, Cargos, Saldo
        # Ensure 'Fecha' column exists (should already exist from earlier processing)
        if 'Fecha' not in df_mov.columns and 'fecha' in df_mov.columns:
            df_mov['Fecha'] = df_mov['fecha']
            df_mov = df_mov.drop(columns=['fecha'])
        
        desired_order = ['Fecha', 'Descripción', 'Abonos', 'Cargos', 'Saldo']
        # Filter to only include columns that exist in the dataframe
        desired_order = [col for col in desired_order if col in df_mov.columns]
        # Remove ALL other columns that are not in desired_order
        columns_to_remove = [c for c in df_mov.columns if c not in desired_order]
        if columns_to_remove:
            df_mov = df_mov.drop(columns=columns_to_remove)
        # Reorder to match desired_order exactly
        if desired_order:  # Only reorder if we have columns
            df_mov = df_mov[desired_order]
    else:
        # For other banks: Fecha, Descripción, Abonos, Cargos, Saldo (if available)
        # Build desired_order based on what columns are configured for this bank
        desired_order = ['Fecha', 'Descripción', 'Abonos', 'Cargos']
        
        # For other banks, only include Saldo if it's in the bank's column configuration
        if 'saldo' in columns_config:
            desired_order.append('Saldo')
        
        # Remove 'saldo' column if it exists but shouldn't (e.g., for Konfio)
        if 'saldo' in df_mov.columns and 'saldo' not in columns_config:
            df_mov = df_mov.drop(columns=['saldo'])
        if 'Saldo' in df_mov.columns and 'saldo' not in columns_config:
            df_mov = df_mov.drop(columns=['Saldo'])
        
        # Filter to only include columns that exist in the dataframe
        desired_order = [col for col in desired_order if col in df_mov.columns]
        # Only keep the desired columns, remove all others
        df_mov = df_mov[desired_order]

    print("📊 Exporting to Excel...")
    
    # Filter summary/info rows from Movements for Banamex
    # These should not appear in Movements: "Saldo mínimo requerido", "COMISIONES COBRADAS"
    if bank_config['name'] == 'Banamex':
        #print("🔍 Filtrando filas de información (Saldo mínimo requerido, COMISIONES COBRADAS) de Movements...")
        info_rows_to_remove = []
        
        # Check each row in df_mov for summary/info rows
        for idx, row in df_mov.iterrows():
            # Get description from various possible column names
            desc_col = None
            for col in ['Descripción', 'Descripcion', 'descripcion', 'raw']:
                if col in df_mov.columns:
                    desc_col = col
                    break
            
            if desc_col:
                desc_text = str(row.get(desc_col, '')).upper()
                
                # Check if this is an info row that should be filtered
                if 'SALDO MINIMO REQUERIDO' in desc_text or 'SALDO MÍNIMO REQUERIDO' in desc_text:
                    info_rows_to_remove.append(idx)
                    #print(f"   ✅ Fila filtrada (Saldo mínimo requerido): {str(row.get(desc_col, ''))[:60]}...")
                elif 'COMISIONES COBRADAS' in desc_text:
                    info_rows_to_remove.append(idx)
                    #print(f"   ✅ Fila filtrada (Comisiones cobradas): {str(row.get(desc_col, ''))[:60]}...")
        
        # Remove info rows from Movements
        if info_rows_to_remove:
            #print(f"   📝 Removiendo {len(info_rows_to_remove)} filas de información de Movements...")
            df_mov = df_mov.drop(index=info_rows_to_remove).reset_index(drop=True)
            #print(f"   ✅ Filas removidas de Movements")
    
    # Filter info rows from Movements for Banregio
    # These should not appear in Movements: rows starting with "del 01 al"
    if bank_config['name'] == 'Banregio':
        info_rows_to_remove = []
        
        # Check each row in df_mov for info rows
        for idx, row in df_mov.iterrows():
            # Get description from various possible column names
            desc_col = None
            for col in ['Descripción', 'Descripcion', 'descripcion', 'raw', 'Fecha']:
                if col in df_mov.columns:
                    desc_col = col
                    break
            
            if desc_col:
                desc_text = str(row.get(desc_col, '')).strip()
                
                # Check if this row starts with "del 01 al" (irrelevant information)
                if re.search(r'^del\s+01\s+al', desc_text, re.I):
                    info_rows_to_remove.append(idx)
        
        # Remove info rows from Movements
        if info_rows_to_remove:
            df_mov = df_mov.drop(index=info_rows_to_remove).reset_index(drop=True)
    
    # Extract DIGITEM and Transferencias sections directly from PDF for Banamex
    # This must be done BEFORE calculating totals for validation
    df_transferencias = None
    df_digitem = None
    if bank_config['name'] == 'Banamex':
        # print("🔍 Extrayendo secciones DIGITEM y TRANSFERENCIA directamente del PDF...")
        
        # Extract DIGITEM section from PDF using same coordinate-based extraction as Movements
        df_digitem = extract_digitem_section(pdf_path, columns_config)
        
        # Extract Transferencias section from PDF
        df_transferencias = extract_transferencia_section(pdf_path)
        
        # Add total row for DIGITEM if there are rows
        if df_digitem is not None and not df_digitem.empty and len(df_digitem) > 0:
            # print("📊 Agregando fila de totales para DIGITEM...")
            total_row_digitem = {
                'Fecha': 'Total',
                'Descripción': '',
                'Importe': ''
            }
            
            # Calculate total for Importe column
            try:
                importe_values = df_digitem['Importe'].apply(lambda x: normalize_amount_str(x) if pd.notna(x) and str(x).strip() else 0.0)
                total_importe = importe_values.sum()
                if total_importe > 0:
                    total_row_digitem['Importe'] = f"{total_importe:,.2f}"
            except Exception as e:
                pass
                # print(f"⚠️  Error al calcular total de Importe en DIGITEM: {e}")
            
            # Append total row
            total_df_digitem = pd.DataFrame([total_row_digitem])
            df_digitem = pd.concat([df_digitem, total_df_digitem], ignore_index=True)
            # print(f"✅ Fila de totales agregada a DIGITEM")
        
        # Add total row for Transferencias if there are rows
        if df_transferencias is not None and not df_transferencias.empty and len(df_transferencias) > 0:
            #print("📊 Agregando fila de totales para Transferencias...")
            total_row_transferencia = {
                'Fecha': 'Total',
                'Descripción': '',
                'Importe': '',
                'Comisiones': '',
                'I.V.A': '',
                'Total': ''
            }
            
            # Calculate totals for all numeric columns
            try:
                if 'Importe' in df_transferencias.columns:
                    importe_values = df_transferencias['Importe'].apply(lambda x: normalize_amount_str(x) if pd.notna(x) and str(x).strip() else 0.0)
                    total_importe = importe_values.sum()
                    if total_importe > 0:
                        total_row_transferencia['Importe'] = f"{total_importe:,.2f}"
                
                if 'Comisiones' in df_transferencias.columns:
                    comisiones_values = df_transferencias['Comisiones'].apply(lambda x: normalize_amount_str(x) if pd.notna(x) and str(x).strip() else 0.0)
                    total_comisiones = comisiones_values.sum()
                    if total_comisiones > 0:
                        total_row_transferencia['Comisiones'] = f"{total_comisiones:,.2f}"
                
                if 'I.V.A' in df_transferencias.columns:
                    iva_values = df_transferencias['I.V.A'].apply(lambda x: normalize_amount_str(x) if pd.notna(x) and str(x).strip() else 0.0)
                    total_iva = iva_values.sum()
                    if total_iva > 0:
                        total_row_transferencia['I.V.A'] = f"{total_iva:,.2f}"
                
                if 'Total' in df_transferencias.columns:
                    total_values = df_transferencias['Total'].apply(lambda x: normalize_amount_str(x) if pd.notna(x) and str(x).strip() else 0.0)
                    total_total = total_values.sum()
                    if total_total > 0:
                        total_row_transferencia['Total'] = f"{total_total:,.2f}"
            except Exception as e:
                pass
                # print(f"⚠️  Error al calcular totales en Transferencias: {e}")
            
            # Append total row
            total_df_transferencia = pd.DataFrame([total_row_transferencia])
            df_transferencias = pd.concat([df_transferencias, total_df_transferencia], ignore_index=True)
            #print(f"✅ Fila de totales agregada a Transferencias")
    
    # Extract summary from PDF and calculate totals for validation
    # IMPORTANT: Calculate totals AFTER removing DIGITEM rows and BEFORE adding the "Total" row
    #print("🔍 Extrayendo información de resumen del PDF para validación...")
    # Para HSBC con OCR, el resumen ya fue extraído desde el texto OCR arriba
    # Para otros casos, extraer desde PDF
    if not (is_hsbc and used_ocr):
        pdf_summary = extract_summary_from_pdf(pdf_path)
    # Si es HSBC con OCR, pdf_summary ya fue extraído arriba en extract_hsbc_summary_from_ocr_text
    extracted_totals = calculate_extracted_totals(df_mov, bank_config['name'])
    
    # Add a "Total" row at the end summing only "Abonos" and "Cargos" columns
    # This is done AFTER calculating totals for validation
    #print("📊 Agregando fila de totales...")
    total_row = {}
    
    # For each column, only calculate sum for "Abonos" and "Cargos"
    for col in df_mov.columns:
        if col in ['Fecha', 'Fecha Oper', 'Fecha Liq.', 'Descripción']:
            # Text columns: put "Total" in the first column, empty in others
            if col == df_mov.columns[0]:
                total_row[col] = 'Total'
            else:
                total_row[col] = ''
        elif col in ['Abonos', 'Cargos']:
            # Only sum Abonos and Cargos columns
            try:
                # Try to convert to numeric and sum
                # For Konfio, ensure we handle cleaned values properly
                def safe_normalize(val):
                    """Safely normalize amount value, handling empty strings and None."""
                    if pd.isna(val):
                        return 0.0
                    val_str = str(val).strip()
                    if not val_str or val_str == '':
                        return 0.0
                    normalized = normalize_amount_str(val_str)
                    return normalized if normalized is not None else 0.0
                
                numeric_values = df_mov[col].apply(safe_normalize)
                total = numeric_values.sum()
                
                # Debug for Konfio
                if bank_config['name'] == 'Konfio':
                    non_zero_count = (numeric_values > 0).sum()
                    #print(f"🔍 Konfio: Columna {col} - Total: {total:,.2f}, Filas con valores > 0: {non_zero_count} de {len(numeric_values)}")
                
                # For Clara, always show total for Abonos (even if negative, zero, or positive)
                if bank_config['name'] == 'Clara' and col == 'Abonos':
                    # Format as currency with 2 decimals (always show, even if negative or zero)
                    total_row[col] = f"{total:,.2f}"
                elif total > 0:
                    # Format as currency with 2 decimals
                    total_row[col] = f"{total:,.2f}"
                else:
                    total_row[col] = ''
            except Exception as e:
                # If conversion fails, leave empty
                if bank_config['name'] == 'Konfio':
                    print(f"❌ Error calculating total for {col}: {e}")
                total_row[col] = ''
        else:
            # All other columns (Saldo, Operación, Liquidación, etc.) - leave empty
            total_row[col] = ''
    
    # Append the total row to the dataframe
    total_df = pd.DataFrame([total_row])
    df_mov = pd.concat([df_mov, total_df], ignore_index=True)
    #print(f"✅ Fila de totales agregada (solo Abonos y Cargos)")
    
    # Create validation sheet
    #print("📋 Creando pestaña de validación...")
    # Check if bank has Saldo column in Movements
    has_saldo_column = 'Saldo' in df_mov.columns or 'saldo' in df_mov.columns
    df_validation = create_validation_sheet(pdf_summary, extracted_totals, has_saldo_column=has_saldo_column)
    #print(f"✅ DataFrame de validación creado con {len(df_validation)} filas")
    #print(f"   Columnas: {list(df_validation.columns)}")
    
    # Print validation summary to console
    print_validation_summary(pdf_summary, extracted_totals, df_validation)
    
    # Determine number of sheets to write
    num_sheets = 3  # Summary, Movements, Data Validation
    if df_transferencias is not None and not df_transferencias.empty:
        num_sheets += 1  # Add Transferencias sheet
    if df_digitem is not None and not df_digitem.empty:
        num_sheets += 1  # Add DIGITEM sheet
    
    # write sheets: summary, movements, validation, and optionally Transferencias and DIGITEM
    try:
        sheet_names = "Summary, Movements, Data Validation"
        if df_transferencias is not None and not df_transferencias.empty:
            sheet_names += ", Transferencias"
        if df_digitem is not None and not df_digitem.empty:
            sheet_names += ", DIGITEM"
        #print(f"📝 Escribiendo Excel con {num_sheets} pestañas: {sheet_names}")
        # Clean amount columns for Banamex: extract only numeric amounts from mixed text
        if bank_config['name'] == 'Banamex':
            for col in ['Cargos', 'Abonos', 'Saldo']:
                if col in df_mov.columns:
                    def extract_amount(value):
                        """Extract only the numeric amount from a string that may contain text."""
                        if pd.isna(value) or value == '':
                            return ''
                        value_str = str(value).strip()
                        # Find all amounts matching DEC_AMOUNT_RE pattern
                        amounts = DEC_AMOUNT_RE.findall(value_str)
                        if amounts:
                            # Return the last amount found (usually the correct one)
                            # Format: ensure it has proper decimal places
                            amount = amounts[-1]
                            # Normalize: ensure it has 2 decimal places
                            if '.' not in amount:
                                amount = amount + '.00'
                            elif amount.count('.') == 1:
                                parts = amount.split('.')
                                if len(parts[1]) == 1:
                                    amount = amount + '0'
                            return amount
                        return ''
                    
                    # Apply cleaning to the column
                    df_mov[col] = df_mov[col].apply(extract_amount)
        
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            #print("   - Escribiendo pestaña 'Summary'...")
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # Special debug for "IVA SOBRE COMISIONES E INTERESES" row before exporting to Excel
            #print("   - Escribiendo pestaña 'Movements'...")
            df_mov.to_excel(writer, sheet_name='Movements', index=False)
            
            # Write Transferencias sheet if available
            if df_transferencias is not None and not df_transferencias.empty:
                #print("   - Escribiendo pestaña 'Transferencias'...")
                df_transferencias.to_excel(writer, sheet_name='Transferencias', index=False)
                #print(f"   ✅ Pestaña 'Transferencias' creada exitosamente con {len(df_transferencias)} filas")
            
            # Write DIGITEM sheet if available
            if df_digitem is not None and not df_digitem.empty:
                # print("   - Escribiendo pestaña 'DIGITEM'...")
                df_digitem.to_excel(writer, sheet_name='DIGITEM', index=False)
                # print(f"   ✅ Pestaña 'DIGITEM' creada exitosamente con {len(df_digitem)} filas")
            
            # Ensure validation DataFrame exists and is not empty
            #print("   - Escribiendo pestaña 'Data Validation'...")
            if df_validation is not None and not df_validation.empty:
                try:
                    df_validation.to_excel(writer, sheet_name='Data Validation', index=False)
                    #print(f"   ✅ Pestaña 'Data Validation' creada exitosamente con {len(df_validation)} filas")
                except Exception as e:
                    pass
                    # print(f"   ❌ Error al escribir pestaña 'Data Validation': {e}")
                    # Try with a simpler name
                    try:
                        df_validation.to_excel(writer, sheet_name='Validation', index=False)
                        pass
                        # print(f"   ✅ Pestaña 'Validation' creada exitosamente (nombre alternativo)")
                    except Exception as e2:
                        pass
                        # print(f"   ❌ Error también con nombre alternativo: {e2}")
            else:
                #print("   ⚠️  DataFrame de validación está vacío o es None")
                # Create a minimal validation sheet even if empty
                empty_validation = pd.DataFrame({
                    'Concepto': ['No se pudo crear la validación'],
                    'Valor en PDF': [''],
                    'Valor Extraído': [''],
                    'Diferencia': [''],
                    'Estado': ['⚠️']
                })
                try:
                    empty_validation.to_excel(writer, sheet_name='Data Validation', index=False)
                    #print("   ✅ Pestaña 'Data Validation' creada con datos mínimos")
                except Exception as e:
                    pass
                    # print(f"   ❌ Error al crear pestaña mínima: {e}")
        
        print(f"✅ Excel file created -> {output_excel}")
    except Exception as e:
        print(f'❌ Error writing Excel: {e}')
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
