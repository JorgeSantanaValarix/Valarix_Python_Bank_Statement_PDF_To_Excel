import sys
import os
import re
import pdfplumber
import pandas as pd

# NEW IMPORTS for Tesseract OCR (minimal):
try:
    import pytesseract
    import fitz  # PyMuPDF - to convert PDF to images
    from PIL import Image, ImageEnhance, ImageFilter
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False
    print("[WARNING] Tesseract OCR not available. Install: pip install pytesseract pymupdf pillow")

# Configure UTF-8 encoding for Windows (improves compatibility with Windows Server)
if sys.platform == 'win32':
    import io
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    except AttributeError:
        # If already configured, ignore
        pass

# Bank configurations with column coordinate ranges (X-axis)
# Use find_coordinates.py to get the exact ranges for your PDF
# python find_coordinates.py <pdf_path> <page_number>

def configure_tesseract():
    """
    Configure the path to Tesseract OCR if it's not in PATH.
    Automatically detects common location on Windows.
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
    
    # If not found, try to use the one in PATH
    try:
        pytesseract.get_tesseract_version()
        return True
    except:
        return False

BANK_CONFIGS = {
    "BBVA": {
        "name": "BBVA",
        "movements_start": "Detalle de Movimientos Realizados",
        "movements_end": "Total de Movimientos",
        "columns": {
            "fecha": (9, 38),              # Operation Date column
            "liq": (52, 81),                # LIQ (Liquidation) column
            "descripcion": (86, 322),     # Description column
            "cargos": (360, 398),          # Charges column
            "abonos": (422, 458),          # Credits column
            "saldo": (539, 593),           # Balance column
        }
    },
    
    "Santander": {
        "name": "Santander",
        "movements_start": "DETALLEDEMOVIMIENTOSCUENTADECHEQUES",
        # Only this duplicated-letter full phrase (no loose patterns: no "Detalles de movimientos", no d+e+ alone, no tertiary header)
        "movements_start_secondary": r'D+e+t+a+l+l+e+\s+d+e+\s+m+o+v+i+m+i+e+n+t+o+s+\s+c+u+e+n+t+a+',
        "movements_end": "TOTAL",
        "metas_start": "DETALLE DE MOVIMIENTOS MIS METAS SANTANDER",
        "metas_end": "INFORMACION FISCAL",
        "columns": {
            "fecha": (18, 52),             # Operation Date column
            "descripcion": (107, 376),     # Description column (was 149; widened so full text e.g. "IVA POR COMISION MEMBRESIA" is captured)
            "cargos": (465, 492),          # Charges column
            "abonos": (377, 411),          # Credits column
            "saldo": (554, 573),           # Balance column
        }
    },

    "Scotiabank": {
        "name": "Scotiabank",
        "movements_start": "Detalledetusmovimientos",
        "movements_end": "LAS TASAS DE INTERES",
        "movements_end_patterns": [  # Patrones adicionales como fallback
            r'LAS\s+TASAS\s+DE\s+INTERES\s+ESTAN\s+EXPRESADAS\s+EN\s+TERMINOS\s+ANUALES\s+SIMPLES\.?',
            r'LAS\s+TASAS\s+DE\s+INTERES.*?ESTAN.*?EXPRESADAS.*?EN.*?TERMINOS.*?ANUALES.*?SIMPLES\.?',
            r'LAS.*?TASAS.*?DE.*?INTERES.*?ESTAN.*?EXPRESADAS.*?EN.*?TERMINOS.*?ANUALES.*?SIMPLES',
        ],
        "columns": {
            "fecha": (56, 81),             # Operation Date column
            "descripcion": (92, 240),     # Description column
            "cargos": (465, 509),          # Charges column
            "abonos": (385, 437),          # Credits column
            "saldo": (532, 584),           # Balance column
        }
    },

    "Inbursa": {
        "name": "Inbursa",
        "movements_start": "FECHA REFERENCIA CONCEPTO CARGOS ABONOS SALDO",
        "movements_end": "Si desea recibir pagos",
        "columns": {
            "fecha": (11, 40),             # Operation Date column
            "descripcion": (145, 369),     # Description column
            "cargos": (400, 441),          # Charges column
            "abonos": (475, 510),          # Credits column
            "saldo": (525, 563),           # Balance column
        }
    },

    "INTERCAM": {
        "name": "INTERCAM",
        "movements_start": "DÍA FOLIO CONCEPTO DEPÓSITOS RETIROS SALDO",
        "movements_end": "Total",
        "columns": {
            "fecha": (40, 45),             # Day only (1-31)
            "descripcion": (78, 390),      # Description column
            "abonos": (430, 450),          # Credits (Depósitos)
            "cargos": (480, 513),          # Charges (Retiros)
            "saldo": (530, 572),           # Balance column
        }
    },

    "Konfio": {
        "name": "Konfio",
        "movements_start": "Historial de movimientos del titular",
        "movements_end": "Subtotal",
        "columns": {
            "fecha": (48, 102),             # Operation Date column
            "descripcion": (120, 240),     # Description column
            "cargos": (340, 435),          # Charges column
            "abonos": (510, 565),          # Credits column
        }
    },

    "Clara": {
        "name": "Clara",
        "movements_start": "Movimientos",
        "movements_end": "Total MXN",
        "columns": {
            "fecha": (35, 60),             # Operation Date column
            "descripcion": (60, 450),       # Description column (expanded to capture all description words)
            "cargos": (450, 480),          # Charges column
            "abonos": (520, 576),          # Credits column
        }
    },

    "Banregio": {
        "name": "Banregio",
        "movements_start": "DIA CONCEPTO CARGOS ABONOS SALDO",
        "movements_end": "Total",
        "columns": {
            "fecha": (35, 42),             # Operation Date column
            "descripcion": (53, 310),     # Description column
            "cargos": (380, 418),          # Charges column
            "abonos": (460, 498),          # Credits column
            "saldo": (530, 573),           # Balance column
        }
    },
    
     "Banorte": {
        "name": "Banorte",
        "movements_start": "DETALLE DE MOVIMIENTOS (PESOS)",
        "movements_start_secondary": "DETALLE DE MOVIMIENTOS (DOLAR AMERICANO)",
        "movements_end": "INVERSION ENLACE NEGOCIOS",
        "movements_end_secondary": "CARGOS OBJETADOS EN EL PERÍODO",
        "columns": {
            "fecha": (54, 85),             # Operation Date column
            "descripcion": (87, 167),     # Description column
            "cargos": (450, 489),          # Charges column
            "abonos": (380, 420),          # Credits column
            "saldo": (533, 560),           # Balance column
        }
    },

     "Banbajío": {
        "name": "Banbajío",
        "movements_start": "DETALLE DE LA CUENTA",
        "movements_end": "SALDO TOTAL",
        "columns": {
            "fecha": (21, 41),             # Operation Date column
            "descripcion": (87, 362),     # Description column
            "cargos": (490, 525),          # Charges column
            "abonos": (415, 451),          # Credits column
            "saldo": (550, 585),           # Balance column
        }
    },
    
    "Banamex": {
        "name": "Banamex",
        "movements_start": "DETALLE DE OPERACIONES",  # String that marks the start of the movements section
        "movements_start_secondary": "DESGLOSE DE MOVIMIENTOS",  # New format (mixed text/image); OCR may output as separate words (--find row Y≈852)
        "movements_start_tertiary": [r"CARGOS,?\s+ABONOS\s+Y\s+COMPRAS\s+REGULARES"],  # Flexible match for OCR (comma optional, whitespace)
        "movements_end": "SALDO PROMEDIO MINIMO REQUERIDO",    # String that marks the end of the movements section
        "movements_end_secondary": "SALDO MINIMO REQUERIDO",    # Alternative end string in some Banamex PDFs
        "movements_end_new_format": "Total cargos +",           # End for new format (do not parse the numeric value after)
        "columns": {
            "fecha": (17, 45),             # Operation Date column
            "descripcion": (55, 260),      # Description column (expanded to capture better)
            "cargos": (275, 316),          # Charges column (slightly expanded)
            "abonos": (345, 395),          # Credits column (slightly expanded)
            "saldo": (425, 472),           # Balance column (slightly expanded)
        },
        # New format: Fecha de cargo, Descripción del movimiento, Monto (- = Abono, + = Cargo). Calibrate with --find if needed.
        "columns_new_format": {
            "fecha": (50, 130),            # Fecha de cargo column (use this date)
            "descripcion": (140, 450),     # Descripción del movimiento
            "monto": (460, 550),           # Monto: "-" prefix = Abono, "+" prefix = Cargo
        }
    },

    "HSBC": {
        "name": "HSBC",
        "movements_start": "ISR Retenido en el Año",  # String that marks the start of the movements section (table header)
        "movements_end": "procesada por CoDi",                      # String that marks the end of the movements section
        "movements_end_secondary": "Enviados durante el periodo del",            # Alternative end string in some HSBC PDFs
        "columns": {
            "fecha": (87, 103),             # Operation Date column
            "descripcion": (124, 505),      # Description column (expanded to capture better)
            "cargos": (710, 800),          # Charges column (slightly expanded)
            "abonos": (865, 950),          # Credits column (slightly expanded)
            "saldo": (1050, 1130),           # Balance column (slightly expanded)
        }
    },
    "Base": {
        "name": "Base",
        "movements_start": "DETALLE DE OPERACIONES",
        "movements_end": "[SALDO INICIAL DE",
        "columns": {
            "fecha": (41, 78),             # Operation Date column
            "descripcion": (108, 342),      # Description column (expanded to capture better)
            "cargos": (375, 415),          # Charges column (slightly expanded)
            "abonos": (440, 485),          # Credits column (slightly expanded)
            "saldo": (520, 560),           # Balance column (slightly expanded)
        }
    },

    "Hey": {
        "name": "Hey",
        "movements_start": "DIA CONCEPTO CARGOS ABONOS SALDO",
        "movements_end": "Total",
        "columns": {
            "fecha": (35, 42),             # Operation Date column (placeholder, calibrate with --find)
            "descripcion": (53, 310),     # Description column (placeholder, calibrate with --find)
            "cargos": (380, 418),          # Charges column (placeholder, calibrate with --find)
            "abonos": (460, 498),          # Credits column (placeholder, calibrate with --find)
            "saldo": (530, 573),           # Balance column (placeholder, calibrate with --find)
        }
    },

    "Mercury": {
        "name": "Mercury",
        "movements_start": "Date (UTC) Description Type Amount End of Day Balance",
        "movements_end": "Total",
        "columns": {
            "fecha": (47, 69),             # Month (3 chars) + day (1-31), e.g. "Jul 01"
            "descripcion": (100, 350),     # Description column
            "cargos": (430, 480),          # Shared with abonos (430-480). Negative amounts (e.g. –$1,199.00) -> Cargos only
            "abonos": (430, 480),          # Shared with cargos (430-480). Positive amounts (e.g. $6,830.00) -> Abonos only
            "saldo": (520, 570),           # End of day balance (can be negative)
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
        r"BBVA\s+MEXICO",
        r"GRUPO\s+FINANCIERO\s+BBVA",
        r"GRUPO\s+FINANCIERO\s+BBVA\s+MEXICO",
    ],
    "Banamex": [
        r"\bDIGITEM\b",  # Very specific word for Banamex
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
    "INTERCAM": [
        r"\bINTERCAM\b",
        r"\bINTERCAM\s+BANCO\b",
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
        r"\bHSBC\s",
        r"\bHSBC.\s",
        r"\bHSBC\s+BANCO\b",
        r"\bHSBC\s+M[EE]XICO\b",
        r"\bCoDi\b",
        r"\bHSBC\s+ADVANCE\b",
    ],
    "Base": [
        r"\bBASE\b",
        r"\bBANCO\s+BASE\b",
    ],
    "Mercury": [
        r"\bMERCURY\b",
        r"Mercury\s+Credit",
        r"Mercury\s+IO",
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
    Automatically detects if OCR should be used for illegible PDFs or Banamex mixed format.
    """
    try:
        # STEP 1: Detect if PDF is illegible or Banamex mixed (needs OCR for coordinates)
        is_illegible, cid_ratio, ascii_ratio = is_pdf_text_illegible(pdf_path)
        detected_bank_early = detect_bank_from_pdf(pdf_path)
        use_ocr_banamex_mixed = (
            detected_bank_early == 'Banamex' and ascii_ratio < 0.99 and TESSERACT_AVAILABLE
        )
        
        if (is_illegible or use_ocr_banamex_mixed) and TESSERACT_AVAILABLE:
            if is_illegible:
                print(f"[INFO] PDF detected as illegible (CID ratio: {cid_ratio:.2%}, ASCII ratio: {ascii_ratio:.2%})", flush=True)
            if use_ocr_banamex_mixed:
                print(f"[INFO] Banamex PDF with mixed content (ASCII ratio: {ascii_ratio:.2%} < 99%). Using OCR for coordinates...", flush=True)
            elif is_illegible:
                print(f"[INFO] Using OCR to analyze coordinates...")
            
            # Extract with OCR (now uses image_to_data() with real coordinates)
            pages_data = extract_text_with_tesseract_ocr(pdf_path)
            
            # Detect bank from OCR text
            all_text = '\n'.join([p.get('content', '') for p in pages_data[:3]])
            detected_bank = detect_bank_from_text(all_text)
            
            print(f"📄 Showing coordinates extracted with OCR (real coordinates)...")
            
            # Filter page if necessary
            if page_number:
                pages_data = [p for p in pages_data if p['page'] == page_number]
            
            # Show words with real OCR coordinates
            print(f"\n📄 Page {page_number} of PDF (OCR): {pdf_path}")
            print("=" * 120)
            print(f"{'Y (top)':<8} {'X0':<8} {'X1':<8} {'X_center':<10} {'Conf':<6} {'Text':<40}")
            print("-" * 120)
            
            # Group words by Y (rows)
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
            print("\nApproximate column ranges (X0 to X1) - Real OCR coordinates:")
            print("-" * 120)
            
            # Find column boundaries
            all_x0 = [w.get('x0', 0) for word_list in rows.values() for w in word_list]
            all_x1 = [w.get('x1', 0) for word_list in rows.values() for w in word_list]
            
            if all_x0 and all_x1:
                min_x = min(all_x0)
                max_x = max(all_x1)
                print(f"Total X range: {min_x:.1f} to {max_x:.1f}")
                print("\nAnalyze the output above and provide these values in BANK_CONFIGS:")
            else:
                print("⚠️  No words with coordinates found on this page")
            
            return
        
        # STEP 2: Legible PDF - normal behavior (pdfplumber)
        # Detect bank to apply Konfio-specific fixes
        detected_bank = detect_bank_from_pdf(pdf_path)
        is_konfio = (detected_bank == "Konfio")
        
        with pdfplumber.open(pdf_path) as pdf:
            if page_number > len(pdf.pages):
                # print(f"❌ PDF only has {len(pdf.pages)} pages")
                return
            
            page = pdf.pages[page_number - 1]
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            
            if not words:
                # print("❌ No words found on page")
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
            
            print(f"\n📄 Page {page_number} of PDF: {pdf_path}")
            print("=" * 120)
            print(f"{'Y (top)':<8} {'X0':<8} {'X1':<8} {'X_center':<10} {'Text':<40}")
            print("-" * 120)
            
            # Print words sorted by Y then X
            for top in sorted(rows.keys()):
                row_words = sorted(rows[top], key=lambda w: w['x0'])
                for i, word in enumerate(row_words):
                    x_center = (word['x0'] + word['x1']) / 2
                    print(f"{top:<8} {word['x0']:<8.1f} {word['x1']:<8.1f} {x_center:<10.1f} {word['text']:<40}")
            
            print("\n" + "=" * 120)
            print("\nApproximate column ranges (X0 to X1):")
            print("-" * 120)
            
            # Find column boundaries
            all_x0 = [w['x0'] for word_list in rows.values() for w in word_list]
            all_x1 = [w['x1'] for word_list in rows.values() for w in word_list]
            
            min_x = min(all_x0)
            max_x = max(all_x1)
            
            print(f"Total X range: {min_x:.1f} to {max_x:.1f}")
            print("\nAnalyze the output above and provide these values in BANK_CONFIGS:")
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


# Combined date pattern for bank detection: matches any bank date format (DD/MM/YYYY, DIA-MES-AÑO, "30 ENE", etc.)
# Used to skip counting BANK_KEYWORDS when the match is on a movement row (line or neighbor has a date).
BANK_DETECTION_DATE_PATTERN = re.compile(
    r'\b(?:'
    r'(?:0[1-9]|[12][0-9]|3[01])/(?:0[1-9]|1[0-2])/\d{4}'  # DD/MM/YYYY
    r'|\d{1,2}-[A-Z]{3}-\d{2,4}'  # DIA-MES-AÑO (e.g. 12-ENE-23)
    r'|(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?'
    r'|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])(?:\s+\d{2,4})?'
    r'|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4}'
    r')\b',
    re.I
)


def detect_bank_from_text(text: str, from_ocr: bool = False) -> str:
    """
    Detect the bank from extracted text content.
    Phase 1: If BANK_KEYWORDS match, return that bank immediately.
      - When from_ocr=True (OCR path): check all lines of the text (full first page).
      - Otherwise: check only the first 30 lines.
    Phase 2: Otherwise, search all lines, count matches per bank (excluding matches near a date),
    and return the bank with the most occurrences.
    When from_ocr=True and no bank is detected (Phase 2 max_count=0), returns "HSBC" as fallback.
    """
    if not text:
        return "HSBC" if from_ocr else DEFAULT_BANK
    
    amount_pattern = re.compile(r"\b\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")
    lines = text.split('\n')
    n_lines = len(lines)
    
    # Phase 1: first match wins. OCR: check full first page; else first 30 lines only
    phase1_lines = lines if from_ocr else lines[:30]
    for line_idx, line in enumerate(phase1_lines):
        line_clean = line.strip()
        if not line_clean:
            continue
        if amount_pattern.search(line_clean):
            continue
        for bank_name, keywords in BANK_KEYWORDS.items():
            for keyword_pattern in keywords:
                if re.search(keyword_pattern, line_clean, re.I):
                    return bank_name
        line_upper = line_clean.upper()
        for bank_name in BANK_KEYWORDS.keys():
            if re.search(rf'\b{re.escape(bank_name.upper())}\b', line_upper):
                return bank_name
    
    # Phase 2: No match in Phase 1 — use new logic on all lines (count + filter near date)
    bank_counts = {bank_name: 0 for bank_name in BANK_KEYWORDS.keys()}
    for i, line in enumerate(lines):
        line_clean = line.strip()
        if not line_clean:
            continue
        if amount_pattern.search(line_clean):
            continue
        
        has_date_nearby = BANK_DETECTION_DATE_PATTERN.search(line_clean)
        if not has_date_nearby:
            for offset in (-2, -1, 1, 2):
                j = i + offset
                if 0 <= j < n_lines:
                    neighbor = lines[j].strip()
                    if neighbor and BANK_DETECTION_DATE_PATTERN.search(neighbor):
                        has_date_nearby = True
                        break
        if has_date_nearby:
            continue
        
        for bank_name, keywords in BANK_KEYWORDS.items():
            matched = False
            for keyword_pattern in keywords:
                if re.search(keyword_pattern, line_clean, re.I):
                    matched = True
                    break
            if not matched:
                line_upper = line_clean.upper()
                if re.search(rf'\b{re.escape(bank_name.upper())}\b', line_upper):
                    matched = True
            if matched:
                bank_counts[bank_name] += 1
    
    max_count = max(bank_counts.values()) if bank_counts else 0
    if max_count == 0:
        return "HSBC" if from_ocr else DEFAULT_BANK
    return max(bank_counts, key=bank_counts.get)


def detect_bank_from_pdf(pdf_path: str) -> str:
    """
    Detect the bank from PDF content by reading line by line.
    Returns the bank name if detected, otherwise returns DEFAULT_BANK.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Read all pages so BANK_KEYWORDS are checked across the entire PDF
            all_text = ""
            for page_num in range(len(pdf.pages)):
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
    Detects if a PDF has illegible text (CID characters).
    Analyzes first and second page (if available) to determine if PDF is illegible.
    Uses majority strategy: PDF is considered illegible only if BOTH pages are illegible.
    For single-page PDFs, uses the result of that page directly.
    
    Args:
        pdf_path: Path to PDF file
        cid_threshold: Minimum ratio of CID characters to consider illegible (default: 5%)
    
    Returns:
        Tuple: (is_illegible: bool, cid_ratio: float, ascii_ratio: float)
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                return False, 0.0, 0.0
            
            # Analyze first and second page (if available)
            # Strategy: PDF is illegible only if BOTH pages are illegible (majority strategy)
            pages_to_check = min(2, len(pdf.pages))
            page_results = []
            
            for i in range(pages_to_check):
                page = pdf.pages[i]
                page_text = page.extract_text() or ""
                
                if not page_text or len(page_text) < 50:
                    # If page has no text or very short, consider it illegible
                    page_results.append({
                        'page_num': i + 1,
                        'is_illegible': True,
                        'cid_ratio': 1.0,
                        'ascii_ratio': 0.0
                    })
                    continue
                
                # Count CID characters for this page
                page_cid_count = page_text.count('(cid:')
                page_total_chars = len(page_text)
                page_cid_ratio = page_cid_count / page_total_chars if page_total_chars > 0 else 0.0
                
                # Count ASCII characters for this page
                page_ascii_count = sum(1 for c in page_text if ord(c) < 128)
                page_ascii_ratio = page_ascii_count / page_total_chars if page_total_chars > 0 else 0.0
                
                # Check if this page is illegible
                page_is_illegible = (page_cid_ratio > cid_threshold) or (page_ascii_ratio < 0.7)
                
                page_results.append({
                    'page_num': i + 1,
                    'is_illegible': page_is_illegible,
                    'cid_ratio': page_cid_ratio,
                    'ascii_ratio': page_ascii_ratio
                })
            
            # Determine overall illegibility using majority strategy
            if pages_to_check == 1:
                # Only one page: use its result directly
                overall_is_illegible = page_results[0]['is_illegible']
                combined_cid_ratio = page_results[0]['cid_ratio']
                combined_ascii_ratio = page_results[0]['ascii_ratio']
            else:
                # Two pages: PDF is illegible only if BOTH are illegible
                page1_illegible = page_results[0]['is_illegible']
                page2_illegible = page_results[1]['is_illegible']
                overall_is_illegible = page1_illegible and page2_illegible
                
                # Calculate combined ratios (average of both pages)
                combined_cid_ratio = (page_results[0]['cid_ratio'] + page_results[1]['cid_ratio']) / 2.0
                combined_ascii_ratio = (page_results[0]['ascii_ratio'] + page_results[1]['ascii_ratio']) / 2.0
            
            return overall_is_illegible, combined_cid_ratio, combined_ascii_ratio
            
    except Exception as e:
        # If error reading with pdfplumber, assume it may be illegible
        print(f"[WARNING] Error analyzing PDF: {e}")
        return True, 1.0, 0.0


def extract_text_from_ocr_data(ocr_data: dict) -> str:
    """
    Extracts plain text from OCR data to maintain compatibility with 'content'.
    
    Args:
        ocr_data: Dictionary with data from pytesseract.image_to_data()
    
    Returns:
        Plain text extracted from OCR
    """
    text_parts = []
    n_items = len(ocr_data.get('text', []))
    
    for i in range(n_items):
        level = ocr_data.get('level', [])[i] if i < len(ocr_data.get('level', [])) else 0
        text = ocr_data.get('text', [])[i] if i < len(ocr_data.get('text', [])) else ''
        conf = float(ocr_data.get('conf', [])[i]) if i < len(ocr_data.get('conf', [])) else 0
        
        # Only add words (level == 5) with text and confidence > 0
        if level == 5 and text and conf > 0:
            text_parts.append(text)
    
    return ' '.join(text_parts)


def convert_ocr_data_to_words_format(ocr_data: dict, zoom_normalization_factor: float = 1.0) -> list:
    """
    Converts OCR data from pytesseract.image_to_data() to word format with real coordinates.
    Compatible with the format expected by the rest of the code.
    
    Args:
        ocr_data: Dictionary with data from pytesseract.image_to_data()
        zoom_normalization_factor: Factor to normalize coordinates (e.g.: 3.0/2.0 = 1.5 if using 3.0x zoom
                                   but ranges are calibrated for 2.0x). Default: 1.0 (no normalization)
    
    Returns:
        List of dictionaries with format: 
        [{'text': str, 'x0': float, 'top': float, 'x1': float, 'bottom': float, 'conf': float, 'line_num': int}, ...]
    """
    words = []
    n_items = len(ocr_data.get('text', []))
    
    for i in range(n_items):
        level = ocr_data.get('level', [])[i] if i < len(ocr_data.get('level', [])) else 0
        text = ocr_data.get('text', [])[i] if i < len(ocr_data.get('text', [])) else ''
        conf = float(ocr_data.get('conf', [])[i]) if i < len(ocr_data.get('conf', [])) else 0
        
        # Only process words (level == 5) with text and reasonable confidence
        if level == 5 and text and conf > 0:
            left = float(ocr_data.get('left', [])[i]) if i < len(ocr_data.get('left', [])) else 0
            top = float(ocr_data.get('top', [])[i]) if i < len(ocr_data.get('top', [])) else 0
            width = float(ocr_data.get('width', [])[i]) if i < len(ocr_data.get('width', [])) else 0
            height = float(ocr_data.get('height', [])[i]) if i < len(ocr_data.get('height', [])) else 0
            line_num = int(ocr_data.get('line_num', [])[i]) if i < len(ocr_data.get('line_num', [])) else 0
            
            # Normalize coordinates if necessary (to maintain compatibility with calibrated ranges)
            if zoom_normalization_factor != 1.0:
                left = left / zoom_normalization_factor
                top = top / zoom_normalization_factor
                width = width / zoom_normalization_factor
                height = height / zoom_normalization_factor
            
            # Clean OCR errors: replace underscores with spaces (common OCR error for spaces)
            # This helps with cases like "30_ PAGO" → "30  PAGO" → "30 PAGO" after normalization
            # Examples: "30_ PAGO SERVICIO" → "30 PAGO SERVICIO", "06_1V.A." → "06 1V.A."
            text_cleaned = text.replace('_', ' ').strip()
            
            words.append({
                'text': text_cleaned,
                'x0': left,
                'top': top,
                'x1': left + width,
                'bottom': top + height,
                'conf': conf,
                'line_num': line_num
            })
    
    return words


def fix_ocr_date_errors(date_text: str, bank_name: str = None) -> str:
    """
    Corrige errores comunes del OCR en fechas, especialmente para HSBC donde las fechas son solo 2 dígitos (01-31).
    
    Errores corregidos:
    - "og" → "09" (0 confundido con o, 9 confundido con g)
    - "o1" → "01" (0 confundido con o)
    - "o2" → "02", "o3" → "03", etc.
    - "1g" → "19" (9 confundido con g)
    - "2g" → "29", "3g" → "39" (aunque 39 no es válido para días, se corrige por consistencia)
    
    Args:
        date_text: Texto de la fecha extraída por OCR (puede ser solo la fecha o texto con fecha al inicio)
        bank_name: Nombre del banco (solo aplica para bancos específicos)
    
    Returns:
        Texto corregido o texto original si no necesita corrección
    """
    if not date_text:
        return date_text
    
    # Solo aplicar correcciones para HSBC (fechas de 2 dígitos)
    if bank_name != 'HSBC':
        return date_text
    
    original_text = date_text.strip()
    
    # Common OCR corrections for digits in dates
    # Search and replace common patterns at the start of text (where the date is)
    
    # Pattern 1: "og" at start → "09" (very common, specific user case)
    # Can be just "og" or "og " followed by more text
    if original_text.lower().startswith('og'):
        # Replace "og" at start with "09"
        corrected = re.sub(r'^og\b', '09', original_text, flags=re.IGNORECASE)
        if corrected != original_text:
            return corrected
    
    # Pattern 2: "o" followed by digit at start (0 confused with o)
    # Examples: "o1" → "01", "o2" → "02", etc.
    o_digit_pattern = re.compile(r'^o([0-9])(\s|$)', re.IGNORECASE)
    match = o_digit_pattern.match(original_text)
    if match:
        digit = match.group(1)
        # Only correct if the result is a valid day (01-09)
        if 1 <= int(digit) <= 9:
            return re.sub(r'^o([0-9])', f'0\\1', original_text, flags=re.IGNORECASE)
    
    # Pattern 3: Digit followed by "g" at start (9 confused with g)
    # Examples: "1g" → "19", "2g" → "29", etc.
    digit_g_pattern = re.compile(r'^([0-3])g(\s|$)', re.IGNORECASE)
    match = digit_g_pattern.match(original_text)
    if match:
        digit = match.group(1)
        # Only correct if the result is a valid day (19, 29)
        corrected_day = int(f"{digit}9")
        if 1 <= corrected_day <= 31:
            return re.sub(r'^([0-3])g', f'\\19', original_text, flags=re.IGNORECASE)
    
    # Pattern 4: Digit followed by "/" at start (7 confused with /)
    # Examples: "2/" → "27", "1/" → "17", "3/" → "37" (aunque 37 no es válido, se corrige por consistencia)
    # This is a common OCR error where "7" is read as "/"
    # The "/" can be followed by space, end of string, or any character (like "2/ PAGO" or "2/PAGO")
    digit_slash_pattern = re.compile(r'^([1-3])/', re.IGNORECASE)
    match = digit_slash_pattern.match(original_text)
    if match:
        digit = match.group(1)
        # Only correct if the result is a valid day (17, 27)
        corrected_day = int(f"{digit}7")
        if 1 <= corrected_day <= 31:
            # Replace "X/" with "X7" at the start
            # Use \g<1> to reference the captured group, followed by literal "7"
            return re.sub(r'^([1-3])/', r'\g<1>7', original_text, flags=re.IGNORECASE)
    
    # Si el texto completo es solo "og" o variaciones
    if original_text.lower() in ['og', 'og.', 'og,', 'og:']:
        return '09'

    # HSBC: if fecha is 1-2 digits and > 31 (invalid day), normalize to 01-31 (OCR often reads 2 as 7, etc.)
    digit_only = re.match(r'^(\d{1,2})$', original_text.strip())
    if digit_only:
        day = int(digit_only.group(1))
        if day > 31:
            if 32 <= day <= 39:
                day = 20 + (day % 10)   # 32 -> 22, 33 -> 23, ..., 39 -> 29
            elif 70 <= day <= 79:
                day = 20 + (day % 10)   # 73 -> 23 (OCR reads 2 as 7), 70 -> 20, ..., 79 -> 29
            else:
                day = 10 + (day % 10)   # 40 -> 10, 81 -> 11, 99 -> 19
            day = min(31, max(1, day))
            return str(day).zfill(2)
    
    return date_text


# DISABLED: This function was causing problems converting "30.40" to "$0.40"
# No longer necessary after improving OCR resolution
# def fix_ocr_amount_errors(word_text: str, word_x0: float, columns_config: dict, bank_name: str = None) -> str:
#     """
#     Corrige errores comunes del OCR en montos, usando coordenadas X para validar.
#     
#     Errores corregidos:
#     - "$" leído como "3" al inicio de montos (ej: "3728.00" → "$728.00")
#     
#     Args:
#         word_text: Texto de la palabra extraída por OCR
#         word_x0: Coordenada X izquierda de la palabra
#         columns_config: Configuración de columnas del banco (con rangos X)
#         bank_name: Nombre del banco (solo aplica para bancos específicos)
#     
#     Returns:
#         Texto corregido o texto original si no necesita corrección
#     """
#     if not word_text or not columns_config:
#         return word_text
#     
#     # Solo aplicar correcciones para bancos específicos
#     if bank_name != 'HSBC':
#         return word_text
#     
#     # Patrón: número con 2 decimales que empieza con "3" pero no tiene "$"
#     # Ejemplos válidos: "3728.00", "3.78", "3,000.00"
#     # No válidos: "30,000.00" (empieza con 30), "3728" (sin decimales)
#     # El patrón captura cualquier cantidad de dígitos después del "3"
#     amount_pattern = re.compile(r'^3(\d+(?:[,\s]\d{3})*\.\d{2})$')
#     match = amount_pattern.match(word_text.strip())
#     
#     if not match:
#         return word_text
#     
#     # Validación adicional: verificar que esté en rango de columnas de montos
#     # Esto reduce falsos positivos (números que realmente empiezan con "3")
#     cargos_range = columns_config.get('cargos', (0, 0))
#     abonos_range = columns_config.get('abonos', (0, 0))
#     saldo_range = columns_config.get('saldo', (0, 0))
#     
#     # Verificar si la palabra está en alguna columna de montos
#     in_amount_column = (
#         (cargos_range[0] <= word_x0 <= cargos_range[1]) or
#         (abonos_range[0] <= word_x0 <= abonos_range[1]) or
#         (saldo_range[0] <= word_x0 <= saldo_range[1])
#     )
#     
#     if in_amount_column:
#         # Corregir: reemplazar "3" inicial por "$"
#         corrected = f"${match.group(1)}"
#         return corrected
#     
#     return word_text


def convert_ocr_text_to_words_format(ocr_text: str, page_number: int = 1) -> list:
    """
    [DEPRECATED] Converts OCR text to word format with approximate coordinates.
    This function is deprecated. Use convert_ocr_data_to_words_format() instead.
    
    Args:
        ocr_text: Text extracted by OCR
        page_number: Page number
    
    Returns:
        List of dictionaries with format: 
        [{'text': str, 'x0': float, 'top': float, 'x1': float, 'bottom': float}, ...]
    """
    words = []
    lines = ocr_text.split('\n')
    
    y_pos = 100  # Initial Y position
    
    for line in lines:
        if not line.strip():
            y_pos += 20  # Space between lines
            continue
        
        # Split line into words
        line_words = line.split()
        x_pos = 50  # Initial X position
        
        for word_text in line_words:
            if word_text.strip():
                # Approximate width based on text length
                word_width = len(word_text) * 8
                
                words.append({
                    'text': word_text,
                    'x0': x_pos,
                    'top': y_pos,
                    'x1': x_pos + word_width,
                    'bottom': y_pos + 15
                })
                
                x_pos += word_width + 10  # Espacio entre palabras
        
        y_pos += 25  # Line height
    
    return words


def extract_text_with_tesseract_ocr(pdf_path: str, lang: str = 'spa+eng') -> list:
    """
    Extracts text from PDF using local Tesseract OCR.
    100% private - does not send data to external servers.
    Uses PyMuPDF to convert PDF to images (does not require Poppler).
    
    Args:
        pdf_path: Path to PDF file
        lang: Language for OCR (default: 'spa+eng' for Spanish+English)
    
    Returns:
        List of dictionaries with format: [{"page": int, "content": str, "words": list}, ...]
        Compatible with the return format of extract_text_from_pdf.
    
    Raises:
        Exception: If Tesseract is not available or there is an error
    """
    if not TESSERACT_AVAILABLE:
        raise Exception("Tesseract OCR is not available. Install: pip install pytesseract pymupdf pillow")
    
    # Configure Tesseract if necessary
    if not configure_tesseract():
        raise Exception("Tesseract OCR not found. Install Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki")
    
    print("[INFO] Extracting text with local Tesseract OCR (100% private)...", flush=True)
    
    extracted_data = []
    
    try:
        # Use PyMuPDF to convert PDF to images (does not require Poppler)
        doc = fitz.open(pdf_path)
        
        # Get page dimensions for DPI calculation (use first page as reference)
        first_page = doc[0]
        page_width_pts = first_page.rect.width
        page_height_pts = first_page.rect.height
        page_width_inches = page_width_pts / 72.0
        page_height_inches = page_height_pts / 72.0
        
        # Use fixed zoom factor for optimal OCR quality (7.0x = 504 DPI)
        # Target DPI: 504 DPI (7.0x zoom factor)
        target_dpi = 504
        zoom_factor = 7.0  # Fixed zoom for consistent OCR quality
        
        # Calculate effective DPI with fixed zoom
        effective_dpi = 72.0 * zoom_factor  # 504 DPI
        
        # Calculate resulting image dimensions in pixels
        img_width_px = int(page_width_pts * zoom_factor)
        img_height_px = int(page_height_pts * zoom_factor)
        
        # Show DPI calculation info (only for first page to avoid spam) — commented to reduce console noise
        # print(f"[INFO] Page dimensions: {page_width_pts:.1f} x {page_height_pts:.1f} pts ({page_width_inches:.2f} x {page_height_inches:.2f} inches)")
        # print(f"[INFO] Target DPI: {target_dpi} DPI (zoom: {zoom_factor:.1f}x)")
        # print(f"[INFO] Using zoom: {zoom_factor:.1f}x → Effective DPI: {effective_dpi:.0f} DPI")
        # print(f"[INFO] Image size: {img_width_px} x {img_height_px} pixels")
        # if effective_dpi < 300:
        #     print(f"[WARNING] DPI ({effective_dpi:.0f}) is below recommended 300 DPI for OCR")
        # elif effective_dpi > 500:
        #     print(f"[INFO] DPI ({effective_dpi:.0f}) is high - may have diminishing returns beyond 400-500 DPI")
        # else:
        #     print(f"[INFO] DPI ({effective_dpi:.0f}) is within optimal range for OCR accuracy")
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            print(f"[INFO] Processing page {page_num + 1}/{len(doc)} with OCR...", flush=True)
            
            # Convert page to image (high resolution)
            # Coordinates will be normalized later to maintain compatibility with column ranges calibrated for 2.0x
            mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            
            # Convert to PIL Image
            from io import BytesIO
            img = Image.open(BytesIO(img_data))
            
            # Preprocess image for better OCR quality (Opción A1: PIL only improvements)
            # Pipeline: Grayscale → Contrast Enhancement → Sharpening
            # Step 1: Convert to grayscale (simplifies data, improves thresholding)
            img_gray = img.convert('L')
            
            # Step 2: Enhance contrast (helps text stand out from background)
            enhancer = ImageEnhance.Contrast(img_gray)
            img_enhanced = enhancer.enhance(2.0)  # Increase contrast by 2x
            
            # Step 3: Sharpen edges (improves definition of small characters and dots like "I.V.A.")
            # Kernel: [[0, -1, 0], [-1, 5, -1], [0, -1, 0]] - standard sharpening kernel
            sharpening_kernel = ImageFilter.Kernel((3, 3), 
                [0, -1, 0, -1, 5, -1, 0, -1, 0], scale=1)
            img_sharpened = img_enhanced.filter(sharpening_kernel)
            
            # Use preprocessed image for OCR (grayscale + contrast + sharpening)
            img_for_ocr = img_sharpened
            
            # Perform OCR with real coordinates and improved configuration
            # PSM 6: Single uniform block of text (good for multi-line content)
            # Note: PSM 7 was tested but caused issues with word extraction and coordinates
            # OEM 1: LSTM engine (better accuracy than legacy engine)
            tesseract_config = r'--oem 1 --psm 6'
            ocr_data = pytesseract.image_to_data(img_for_ocr, lang=lang, output_type=pytesseract.Output.DICT, config=tesseract_config)
            
            # Extract plain text to maintain compatibility with 'content'
            text = extract_text_from_ocr_data(ocr_data)
            
            # Convert OCR data to word format with real coordinates
            # IMPORTANT: Normalize coordinates by dividing by zoom_factor to maintain compatibility
            # with column ranges calibrated for 2.0x
            words = convert_ocr_data_to_words_format(ocr_data, zoom_normalization_factor=zoom_factor / 2.0)
            
            extracted_data.append({
                "page": page_num + 1,
                "content": text,
                "words": words
            })
        
        doc.close()
        
        print(f"[OK] OCR completed. Pages processed: {len(extracted_data)}", flush=True)
        return extracted_data
        
    except Exception as e:
        raise Exception(f"Error en Tesseract OCR: {e}")


def filter_hsbc_movements_section(pages_data: list, start_string: str, end_string: str, end_strings_also: list = None) -> list:
    """
    Filtra palabras que están entre start_string y end_string para HSBC.
    El string de inicio puede estar dividido en múltiples palabras, así que busca todas las palabras
    que forman el string de inicio.
    
    Args:
        pages_data: List of dictionaries with format [{"page": int, "words": list}, ...]
        start_string: String that marks the start of the section (e.g.: "DETALLE MOVIMIENTOS HSBC")
        end_string: String that marks the end of the section (e.g.: "Información CoDi")
        end_strings_also: Optional list of alternative end strings (e.g.: ["Información SPEI"])
    
    Returns:
        List of filtered words that are in the movements section
    """
    filtered_words = []
    in_section = False
    
    # Normalize strings for search (case-insensitive)
    start_string_normalized = start_string.upper().strip()
    start_words_list = start_string_normalized.split()
    end_list = [end_string.upper().strip()]
    if end_strings_also:
        end_list.extend([s.upper().strip() for s in end_strings_also if s])
    
    for page_data in pages_data:
        page_num = page_data.get('page', 0)
        words = page_data.get('words', [])
        
        # Search for start: the string may be split across multiple words
        if not in_section:
            # Build page text for search
            page_text = ' '.join([w.get('text', '').strip() for w in words]).upper()
            
            # Search if the start string is in the page text (may be split)
            if start_string_normalized in page_text:
                in_section = True
                # Find the first word after the start
                # Contar palabras hasta llegar al inicio
                word_count = 0
                current_text = ''
                start_word_index = None
                for idx, word in enumerate(words):
                    word_text = word.get('text', '').strip()
                    current_text += word_text + ' '
                    if start_string_normalized in current_text.upper():
                        # We've passed the start, find the index of the last word of the start
                        # Search how many words form the start
                        start_word_index = idx + 1  # Start from the next word after the start
                        break
                
                # If we found the start, process the remaining words of this page
                if start_word_index is not None and start_word_index < len(words):
                    # Process words from after the start until the end of the page
                    for word in words[start_word_index:]:
                        word_text = word.get('text', '').strip()
                        word_text_upper = word_text.upper()
                        
                        # Detect end (case-insensitive and partial search; any of end_list)
                        if any(e in word_text_upper for e in end_list):
                            # Stop completely (do not process more pages)
                            return filtered_words
                        
                        # Add page information to word if it doesn't have it
                        if 'page' not in word:
                            word['page'] = page_num
                        
                        # Add word to section
                        filtered_words.append(word)
                    
                    # Continue with next page
                    continue
        
        # If we're already in the section, process all words of this page
        if in_section:
            for word in words:
                word_text = word.get('text', '').strip()
                word_text_upper = word_text.upper()
                
                # Detect end (case-insensitive and partial search; any of end_list)
                if any(e in word_text_upper for e in end_list):
                    # Stop completely (do not process more pages)
                    return filtered_words
                
                # Add page information to word if it doesn't have it
                if 'page' not in word:
                    word['page'] = page_num
                
                # Add word to section
                filtered_words.append(word)
        else:
            # Search for start word by word (alternative method if previous method didn't work)
            # Build cumulative text of consecutive words
            for i in range(len(words) - len(start_words_list) + 1):
                # Take a group of consecutive words of the size of the start string
                consecutive_words = words[i:i+len(start_words_list)]
                consecutive_text = ' '.join([w.get('text', '').strip() for w in consecutive_words]).upper()
                
                # Check if this group of words matches the start string
                if consecutive_text == start_string_normalized or start_string_normalized in consecutive_text:
                    in_section = True
                    # Continue from the next word after the start
                    break
    
    return filtered_words


def extract_hsbc_movements_from_ocr_text(pages_data: list, columns_config: dict = None) -> list:
    """
    Extracts HSBC movements from OCR text using real X coordinates.
    Assigns amounts to columns (Charges/Credits/Balance) based on X position, not keywords.
    
    IMPORTANT: Only extracts information after "ISR Retenido en el año" and before "CoDi"
    
    Args:
        pages_data: List of dictionaries with format [{"page": int, "content": str, "words": list}, ...]
        columns_config: Dictionary with X ranges for columns from BANK_CONFIGS.
                       Format: {"cargos": (x_min, x_max), "abonos": (x_min, x_max), "saldo": (x_min, x_max), ...}
    
    Returns:
        List of dictionaries with movements: [{"fecha": str, "descripcion": str, "cargos": str, "abonos": str, "saldo": str}, ...]
        - fecha: Only the day (e.g.: "03", "13")
        - descripcion: Complete movement description
        - cargos: Amount if withdrawal/charge (e.g.: "$9,500.00") or empty
        - abonos: Amount if deposit/credit (e.g.: "$278,400.00") or empty
        - saldo: Balance after movement (e.g.: "$466,722.66")
    """
    movements = []
    all_text = '\n'.join([p.get('content', '') for p in pages_data])
    lines = all_text.split('\n')
    
    # Search for movements section after "ISR Retenido en el año" and before "CoDi"
    in_movements_section = False
    start_found = False
    
    # Pattern to detect start of movements section
    start_pattern = re.compile(r'ISR\s+Retenido\s+en\s+el\s+a[ñn]o', re.IGNORECASE)
    
    # Pattern to detect end of movements section
    end_pattern = re.compile(r'\bCoDi\b', re.IGNORECASE)
    
    # Pattern to detect day at start of line (01-31)
    day_pattern = re.compile(r'^(\d{1,2})\s+')
    
    # Pattern to detect amounts ($X,XXX.XX) - includes $ symbol in result
    # Improved to capture amounts with internal spaces (e.g.: "$721 588.28" -> "$721,588.28")
    # Captures complete amounts including internal spaces: $ followed by digits with spaces/commas, and optionally .XX
    # Pattern captures: "$721 588.28", "$721,588.28", "$ 278,400.00", etc.
    # Uses a pattern that captures from $ until finding a space followed by non-numeric text or end of line
    # Improved pattern that captures complete amounts with internal spaces
    # Captures: $ followed by digits with spaces/commas between groups, and optionally .XX
    # Pattern must capture "$721 588.28" completely, not just "$721"
    # Uses a more flexible pattern that captures from $ until finding a space followed by non-numeric text
    # or until end of line, but including internal spaces between digits
    amount_pattern = re.compile(r'(\$\s*\d{1,3}(?:[,\s]\d{3})*(?:\s+\d{3})*(?:\.\d{2})?)')
    
    # Validate that columns_config has required ranges
    if not columns_config:
        raise ValueError("columns_config is required. Must contain X ranges for 'cargos', 'abonos' and 'saldo'.")
    
    required_columns = ['cargos', 'abonos', 'saldo']
    missing_columns = [col for col in required_columns if col not in columns_config]
    if missing_columns:
        raise ValueError(f"columns_config must contain ranges for: {missing_columns}")
    
    # Get X ranges for columns
    cargos_range = columns_config.get('cargos')
    abonos_range = columns_config.get('abonos')
    saldo_range = columns_config.get('saldo')
    
    line_count = 0
    for line in lines:
        line_count += 1
        line = line.strip()
        
        # Detect start of movements section
        if not start_found:
            if start_pattern.search(line):
                start_found = True
                in_movements_section = True
                continue
        
        # Detect end of movements section
        if in_movements_section and end_pattern.search(line):
            break  # Stop extraction completely
        
        if not in_movements_section:
            continue
        
        # Search for line starting with day (01-31)
        day_match = day_pattern.match(line)
        if not day_match:
            continue
        
        # Only print debug if line contains "I.V.A." or "IVA"
        contains_iva = 'I.V.A.' in line or 'IVA' in line or '1VA' in line
        
        # Extract all amounts in the line
        # Use a more robust method that captures complete amounts with internal spaces
        # Find all $ positions in the line
        dollar_positions = [i for i, char in enumerate(line) if char == '$']
        amounts_raw = []
        
        for pos in dollar_positions:
            # Capture from $ until next $ or until space followed by letter (not digit)
            remaining = line[pos:]
            # Search for complete amount: from $ until finding space followed by letter or next $
            # Pattern: $ followed by digits, spaces, commas, dots
            match = re.match(r'(\$\s*[\d\s,\.]+)', remaining)
            if match:
                monto_candidato = match.group(1)
                # Find where the amount really ends
                # Search for next $ after the first one
                next_dollar = remaining.find('$', 1)
                # Search for space followed by letter (not digit)
                next_letter_space = re.search(r'\s+[A-Za-z]', remaining[1:])
                
                if next_dollar != -1 and (next_letter_space is None or next_dollar < next_letter_space.start() + 1):
                    # Next $ is before space with letter, cut there
                    monto_candidato = remaining[:next_dollar].strip()
                elif next_letter_space:
                    # There is a space followed by letter, cut before that space
                    monto_candidato = remaining[:next_letter_space.start() + 1].strip()
                else:
                    # No next $ or space with letter, take until end of line
                    monto_candidato = remaining.strip()
                
                amounts_raw.append(monto_candidato)
        
        if len(amounts_raw) < 1:
            continue  # No amounts, not a valid movement
        
        # Clean amounts: replace internal spaces with commas for consistent format
        # Example: "$721 588.28" -> "$721,588.28" or "$ 278,400.00" -> "$278,400.00"
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
        
        # Extract description and reference (everything until the first $)
        first_dollar = line.find('$')
        if first_dollar == -1:
            continue
        
        desc_and_ref = line[day_match.end():first_dollar].strip()
        
        # Clean description (normalize spaces)
        desc_and_ref = re.sub(r'\s+', ' ', desc_and_ref)
        
        # Build movement - date is only the day (e.g.: "03", "13")
        movement = {
            'fecha': day,  # Only the day, without month/year
            'descripcion': desc_and_ref,
            'cargos': '',
            'abonos': '',
            'saldo': ''
        }
        
        # Assign amounts according to X coordinates (not keywords)
        # Calculate approximate Y position of the line (to search for coordinates)
        # Use the line index in the text to estimate Y
        line_index = lines.index(line)
        line_y_approx = 100 + (line_index * 25)  # Estimate: 25 pixels per line
        
        # For each amount, find its X coordinates and assign to column
        amounts_with_columns = []
        for amount in amounts:
            amount_clean = amount.strip()
            
            # Find real X coordinates of the amount
            x0, x1 = find_amount_coordinates(amount_clean, line, line_y_approx, pages_data, y_tolerance=20)
            
            if x0 is not None and x1 is not None:
                # Calculate center of amount for debug
                x_center = (x0 + x1) / 2
                
                # Use assign_word_to_column() to determine the column
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
        
        # Assign amounts to columns according to results
        # If there is ambiguity (amount assigned to cargos but could be abonos), use context
        # Strategy: If there are 2 amounts and both are in cargos range, the leftmost is cargos, the other is saldo
        # If an amount is in cargos range but is closer to abonos center, reconsider
        
        # First, identify amounts that might be incorrectly assigned
        # If an amount is in cargos range but its center is closer to abonos center, reconsider
        for amt_data in amounts_with_columns:
            if amt_data['column'] == 'cargos' and amt_data.get('x_center') is not None:
                x_center = amt_data['x_center']
                # Calculate distance to centers of both ranges
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
                    
                    # If significantly closer to abonos (difference > 30 pixels), reconsider
                    if dist_abonos < dist_cargos and (dist_cargos - dist_abonos) > 30:
                        # Check if it's within or near abonos range (expand range 50 pixels)
                        abonos_range_expanded_min = abonos_min - 50
                        abonos_range_expanded_max = abonos_max + 50
                        if abonos_range_expanded_min <= x_center <= abonos_range_expanded_max:
                            amt_data['column'] = 'abonos'
        
        # Now assign amounts to columns
        for amt_data in amounts_with_columns:
            amt = amt_data['amount']
            col = amt_data['column']
            
            # Only assign if the column is not already occupied
            if col == 'saldo' and not movement.get('saldo'):
                movement['saldo'] = amt
            elif col == 'cargos' and not movement.get('cargos'):
                movement['cargos'] = amt
            elif col == 'abonos' and not movement.get('abonos'):
                movement['abonos'] = amt
        
        # Ensure saldo is assigned when there are 2 amounts
        # If there are 2 amounts and saldo is not assigned, the second amount is saldo (typical HSBC structure)
        if len(amounts) == 2 and not movement.get('saldo'):
            # Check if the second amount is assigned to another column
            second_amt_data = amounts_with_columns[1] if len(amounts_with_columns) > 1 else None
            if second_amt_data and second_amt_data.get('column') and second_amt_data['column'] != 'saldo':
                # The second amount is assigned to another column (cargos or abonos), but should be saldo
                # Reassign the second amount to saldo
                movement['saldo'] = amounts[1].strip()
            elif not second_amt_data or not second_amt_data.get('column'):
                # The second amount is not assigned, assign it to saldo
                movement['saldo'] = amounts[1].strip()
        
        # Fallback: si hay montos sin asignar, usar lógica por posición
        unassigned = [a for a in amounts_with_columns if a['column'] is None]
        if unassigned:
            # Si hay 2 montos
            if len(amounts) == 2:
                assigned_cols = [a['column'] for a in amounts_with_columns if a['column']]
                
                # If saldo is still not assigned, the second amount is saldo
                if 'saldo' not in assigned_cols and not movement.get('saldo'):
                    movement['saldo'] = amounts[1].strip()
                
                # If cargos and abonos are not assigned, the first amount is cargos (by HSBC structure)
                if 'cargos' not in assigned_cols and 'abonos' not in assigned_cols and not movement.get('cargos'):
                    # In HSBC, when there are 2 amounts, the first is generally Cargos (Withdrawal/Charge)
                    # Only assign to Abonos if there is clear evidence (X coordinates very close to abonos range)
                    first_amt = amounts_with_columns[0]
                    if first_amt['x0'] is not None:
                        # Calculate distance to each range
                        x_center = (first_amt['x0'] + first_amt['x1']) / 2
                        cargos_center = (cargos_range[0] + cargos_range[1]) / 2
                        abonos_center = (abonos_range[0] + abonos_range[1]) / 2
                        
                        dist_cargos = abs(x_center - cargos_center)
                        dist_abonos = abs(x_center - abonos_center)
                        
                        # If significantly closer to abonos (difference > 50 pixels), assign to abonos
                        # Otherwise, assign to cargos (typical HSBC structure)
                        if dist_abonos < dist_cargos and (dist_cargos - dist_abonos) > 50:
                            movement['abonos'] = first_amt['amount']
                        else:
                            movement['cargos'] = first_amt['amount']
                    else:
                        # Without coordinates, use typical structure: first amount = cargos
                        movement['cargos'] = first_amt['amount']
            
            # If there are 3+ amounts, assign by position: first=cargos, second=abonos, last=saldo
            elif len(amounts) >= 3:
                if not movement.get('cargos'):
                    movement['cargos'] = amounts[0].strip()
                if not movement.get('abonos'):
                    movement['abonos'] = amounts[1].strip()
                if not movement.get('saldo'):
                    movement['saldo'] = amounts[-1].strip()
        
        # Final verification: Ensure saldo is assigned when there are 2+ amounts
        # This is an additional safety check
        if len(amounts) >= 2 and not movement.get('saldo'):
            # If there are 2+ amounts and saldo is not assigned, assign the last amount to saldo
            movement['saldo'] = amounts[-1].strip()
        
        
        if movement['descripcion'] and (movement['cargos'] or movement['abonos'] or movement['saldo']):
            movements.append(movement)
    print(f"    Total lines processed: {line_count}")
    print(f"    Movements section found: {start_found}")
    print(f"    Total movements extracted: {len(movements)}")
    if movements:
        print(f"    First movement: fecha='{movements[0]['fecha']}', descripcion='{movements[0]['descripcion'][:50]}'")
        print(f"    Last movement: fecha='{movements[-1]['fecha']}', descripcion='{movements[-1]['descripcion'][:50]}'")
    else:
        print(f"    ⚠️  NO MOVEMENTS EXTRACTED")
    
    return movements


def find_amount_coordinates(amount_text: str, line_text: str, line_y_approx: float, pages_data: list, y_tolerance: float = 15) -> tuple:
    """
    Finds the real X coordinates of an amount in pages_data['words'] using real OCR coordinates.
    
    Args:
        amount_text: Amount text (e.g.: "$30,022.54")
        line_text: Complete text of the line where the amount appears
        line_y_approx: Approximate Y position of the line (to filter words)
        pages_data: List of dictionaries with words and coordinates
        y_tolerance: Tolerance in pixels to search on the same Y line (reduced to 15 for real coordinates)
    
    Returns:
        Tuple (x0, x1) with X coordinates of the amount, or (None, None) if not found
    """
    # Normalize amount_text for comparison (remove spaces, normalize format)
    amount_normalized = amount_text.replace(' ', '').replace(',', '').replace('$', '')
    
    # Debug: show search information
    print(f"    [DEBUG find_amount_coordinates] Searching for amount: '{amount_text}' (normalized: '{amount_normalized}')")
    print(f"    [DEBUG find_amount_coordinates] Approximate Y line: {line_y_approx:.1f}, tolerance: {y_tolerance}")
    print(f"    [DEBUG find_amount_coordinates] Complete line: '{line_text[:100]}'")
    
    # Search in all pages
    words_checked = 0
    words_in_y_range = 0
    for page_data in pages_data:
        words = page_data.get('words', [])
        
        # Group words by line_num if available (more precise)
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
            
            # Filter words with low confidence (optional, but useful)
            if word_conf < 30:
                continue
            
            # Check if it's on the same Y line (with reduced tolerance for real coordinates)
            if abs(word_y - line_y_approx) > y_tolerance:
                continue
            
            words_in_y_range += 1
            
            # Check if this word contains the amount
            # Normalize word for comparison
            word_normalized = word_text.replace(' ', '').replace(',', '').replace('$', '')
            
            # Search if the amount is contained in this word or in adjacent words
            if amount_normalized in word_normalized or word_normalized in amount_normalized:
                print(f"    [DEBUG find_amount_coordinates] ✓ Found in word: '{word_text}' (Y: {word_y:.1f}, X: {word_x0:.1f}-{word_x1:.1f}, conf: {word_conf:.1f})")
                return (word_x0, word_x1)
            
            # Also search if the amount appears as part of the line
            # (may be split across multiple words: "$30" "022.54")
            if '$' in word_text:
                # This word has $, may be the start of the amount
                # If we have line_num, search only in the same line
                if word_line_num is not None and word_line_num in words_by_line:
                    # Search in words of the same line (more precise)
                    line_words = words_by_line[word_line_num]
                    word_idx = next((i for i, w in enumerate(line_words) if w == word), -1)
                    
                    if word_idx >= 0:
                        combined_text = word_text
                        combined_x0 = word_x0
                        combined_x1 = word_x1
                        
                        # Search for following words in the same line
                        for next_word in line_words[word_idx + 1:]:
                            next_word_text = next_word.get('text', '').strip()
                            
                            # Add next word
                            combined_text += next_word_text
                            combined_x1 = next_word.get('x1', combined_x1)
                            
                            # Normalize and compare
                            combined_normalized = combined_text.replace(' ', '').replace(',', '').replace('$', '')
                            if amount_normalized in combined_normalized:
                                print(f"    [DEBUG find_amount_coordinates] ✓ Found in combined words (line_num={word_line_num}): '{combined_text}' (X: {combined_x0:.1f}-{combined_x1:.1f})")
                                return (combined_x0, combined_x1)
                            
                            # If we already have enough characters, stop
                            if len(combined_normalized) >= len(amount_normalized) + 5:
                                break
                else:
                    # Fallback: search in all words (previous method)
                    word_idx = words.index(word)
                    combined_text = word_text
                    combined_x0 = word_x0
                    combined_x1 = word_x1
                    
                    # Search for following words on the same Y line
                    for next_word in words[word_idx + 1:]:
                        next_word_y = next_word.get('top', 0)
                        next_word_text = next_word.get('text', '').strip()
                        
                        if abs(next_word_y - line_y_approx) > y_tolerance:
                            break  # No longer on the same line
                        
                        # Add next word
                        combined_text += next_word_text
                        combined_x1 = next_word.get('x1', combined_x1)
                        
                        # Normalize and compare
                        combined_normalized = combined_text.replace(' ', '').replace(',', '').replace('$', '')
                        if amount_normalized in combined_normalized:
                            print(f"    [DEBUG find_amount_coordinates] ✓ Found in combined words: '{combined_text}' (X: {combined_x0:.1f}-{combined_x1:.1f})")
                            return (combined_x0, combined_x1)
                        
                        # Si ya tenemos suficientes caracteres, detener
                        if len(combined_normalized) >= len(amount_normalized) + 5:
                            break
    
    # Debug: show statistics if not found
    print(f"    [DEBUG find_amount_coordinates] ✗ NOT found. Words checked: {words_checked}, in Y range: {words_in_y_range}")
    print(f"    [DEBUG find_amount_coordinates] Possible causes:")
    print(f"        - line_y_approx ({line_y_approx:.1f}) does not match real Y of words")
    print(f"        - The amount is split across multiple words and was not combined correctly")
    print(f"        - The amount format in words does not match amount_text")
    
    return (None, None)


def extract_hsbc_summary_from_ocr_text(pages_data: list) -> dict:
    """
    Extrae información de resumen de HSBC desde el texto OCR de las primeras dos páginas.
    
    Busca:
    - Total Abonos: "Depósitos/" seguido de un monto
    - Total Cargos: "Retiros/Cargos" seguido de un monto
    - Saldo Final: "Saldo Final del Periodo" seguido de un monto
    
    Estos valores normalmente están en la primera o segunda página del PDF, antes del inicio de la sección de movimientos.
    
    Args:
        pages_data: Lista de diccionarios con formato [{"page": int, "content": str, "words": list}, ...]
    
    Returns:
        Diccionario con información de resumen: {
            'total_abonos': float o None,
            'total_cargos': float o None,
            'saldo_final': float o None,
            'total_depositos': float o None,
            'total_retiros': float o None
        }
    """
    summary_data = {
        'total_depositos': None,
        'total_retiros': None,
        'total_cargos': None,
        'total_abonos': None,
        'saldo_final': None,
        'total_movimientos': None,
        'saldo_anterior': None,
        'rfc': None,
        'name': None,
        'period_text': None,
    }
    
    if not pages_data or len(pages_data) == 0:
        return summary_data
    
    # Extract RFC, name, period from ALL pages (HSBC): search until we find "RFC" and a valid RFC value
    full_text_all_pages = '\n'.join(pages_data[i].get('content', '') or '' for i in range(len(pages_data)))
    if full_text_all_pages:
        summary_data['period_text'] = extract_period_text_from_text(full_text_all_pages)
        rfc_val, name_val = extract_rfc_and_name_from_text(full_text_all_pages, detected_bank='HSBC')
        summary_data['rfc'] = rfc_val
        summary_data['name'] = name_val
    
    # Search in the first two pages (before the start of movements)
    pages_to_check = min(2, len(pages_data))
    
    # Patterns to search for values according to user specification
    # Pattern 1: "Depósitos/" or "Depósitos\" followed by an amount (Total Abonos)
    depositos_slash_pattern = re.compile(
        r'Dep[oó]sitos?\s*[/\\]\s*\$?\s*([\d,\.\s]+)',
        re.IGNORECASE
    )
    
    # Pattern 2: "Retiros/Cargos" followed by an amount (Total Cargos)
    retiros_cargos_pattern = re.compile(
        r'Retiros?\s*[/]\s*Cargos?\s+\$?\s*([\d,\.\s]+)',
        re.IGNORECASE
    )
    
    # Pattern 3: "Saldo Final del Periodo" followed by an amount
    saldo_final_pattern = re.compile(
        r'Saldo\s+Final\s+del\s+Periodo\s+\$?\s*([\d,\.\s]+)',
        re.IGNORECASE
    )
    # Alternative pattern: "Saldo Final del" (without "Periodo") followed by amount
    saldo_final_pattern_alt = re.compile(
        r'Saldo\s+Final\s+del\s+\$?\s*([\d,\.\s]+)',
        re.IGNORECASE
    )
    
    # Search in the first two pages
    for page_idx in range(pages_to_check):
        page_data = pages_data[page_idx]
        page_text = page_data.get('content', '')
        
        if not page_text:
            continue
        
        lines = page_text.split('\n')
        
        for line_num, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            
            # FIRST: Search for Depósitos/ (Total Abonos)
            if not summary_data['total_abonos']:
                match = depositos_slash_pattern.search(line)
                if match:
                    amount_str = match.group(1).strip()
                    amount_str = re.sub(r'\s+', '', amount_str)
                    amount = normalize_amount_str(amount_str)
                    if amount and amount > 0:
                        summary_data['total_abonos'] = amount
                        summary_data['total_depositos'] = amount
            
            # SECOND: Search for Retiros/Cargos (Total Cargos)
            if not summary_data['total_cargos']:
                match = retiros_cargos_pattern.search(line)
                if match:
                    amount_str = match.group(1).strip()
                    amount_str = re.sub(r'\s+', '', amount_str)
                    amount = normalize_amount_str(amount_str)
                    if amount and amount > 0:
                        summary_data['total_cargos'] = amount
                        summary_data['total_retiros'] = amount
            
            # THIRD: Search for Saldo Final del Periodo
            if not summary_data['saldo_final']:
                match = saldo_final_pattern.search(line)
                if match:
                    amount_str = match.group(1).strip()
                    amount_str = re.sub(r'\s+', '', amount_str)
                    amount = normalize_amount_str(amount_str)
                    if amount and amount > 0:
                        summary_data['saldo_final'] = amount
                else:
                    # Intentar patrón alternativo
                    match = saldo_final_pattern_alt.search(line)
                    if match:
                        amount_str = match.group(1).strip()
                        amount_str = re.sub(r'\s+', '', amount_str)
                        amount = normalize_amount_str(amount_str)
                        if amount and amount > 0:
                            summary_data['saldo_final'] = amount
    
    return summary_data


def extract_period_text_from_text(full_text: str):
    """
    Extract period string from PDF text (e.g. "Corte al Día 29 - 29 Días", "Del 01 Oct. 2024 al 31 Oct. 2024").
    Returns the first full match as shown in the PDF, or None.
    """
    if not full_text or not full_text.strip():
        return None
    # Santander: "PERIODO DEL 01-ENE-2026 AL 31-ENE-2026" -> return only "01-ENE-2026 AL 31-ENE-2026"
    santander_periodo = re.search(
        r'(?i)PERIODO\s+DEL\s+(\d{1,2}-[A-Z]{3}-\d{2,4})\s+AL\s+(\d{1,2}-[A-Z]{3}-\d{2,4})',
        full_text,
    )
    if santander_periodo:
        return f"{santander_periodo.group(1)} AL {santander_periodo.group(2)}".strip()

    patterns = [
        (r'Corte al D[ií]a \d+ - \d+ D[ií]as', 0),
        (r'Del \d{1,2} \w+\.? \d{4} al \d{1,2} \w+\.? \d{4}', 0),
        (r'DEL \d{1,2}/\w+/\d{4} AL \d{1,2}/\w+/\d{4}', 0),
        (r'\d{1,2}/\d{1,2}\d{0,2}/\d{2,4}\s+al\s+\d{1,2}/\d{1,2}/\d{2,4}', 0),
        (r'Periodo del \d{1,2} \w{3} \d{4} al \d{1,2} \w{3} \d{4}', 0),
        (r'Periodo \d{1,2}-\w{3}-\d{2}/\d{1,2}-\w{3}-\d{2}', 0),
        (r'\d{1,2} AL \d{1,2} DE \w+ DE \d{4}', 0),
        # RESUMEN DEL: 01/ENE/2020 AL 31/ENE/2020
        (r'RESUMEN DEL:\s*\d{1,2}/\w{3}/\d{4}\s+AL\s+\d{1,2}/\w{3}/\d{4}', 0),
        # Periodo 01/01/2023 - 31/01/2023
        (r'Periodo\s+\d{1,2}/\d{1,2}/\d{4}\s*-\s*\d{1,2}/\d{1,2}/\d{4}', 0),
        # Periodo DEL 01/06/2025 AL 30/06/2025
        (r'Periodo\s+DEL\s+\d{1,2}/\d{1,2}/\d{4}\s+AL\s+\d{1,2}/\d{1,2}/\d{4}', 0),
        # PERIODO: 1 DE ENERO AL 31 DE ENERO DE 2024
        (r'PERIODO:\s*\d{1,2}\s+DE\s+\w+\s+AL\s+\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4}', 0),
        # PERIODO: 01  DE  ENE  DE  2025  AL  31  DE  ENE  DE  2025 (flexible spaces)
        (r'PERIODO:\s*\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4}\s+AL\s+\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4}', 0),
        (r'Periodo\s*.+', 0),
        (r'(?:Del|DEL)\s+.+?\s+al\s+.+', 0),
    ]
    for pattern, _ in patterns:
        m = re.search(pattern, full_text, re.IGNORECASE | re.DOTALL)
        if m:
            return m.group(0).strip()
    return None


def extract_rfc_and_name_from_text(full_text: str, detected_bank=None):
    """
    Extract RFC and name (razón social) from PDF text.
    Returns (rfc, name). RFC is normalized (spaces removed). Name excludes lines that contain
    both a company suffix (SA DE CV, etc.) and the detected bank's BANK_KEYWORDS.
    """
    rfc = None
    name = None
    if not full_text or not full_text.strip():
        return (rfc, name)
    lines = full_text.split('\n')
    bank_keywords = BANK_KEYWORDS.get(detected_bank, []) if detected_bank else []

    # RFC: try multiple patterns (RFC:, R.F.C., R.F.C , Registro Federal de Contribuyentes:, or line with only RFC value)
    rfc_value_re = r'([A-ZÑ]{3,4}\s*\d{6}\s*[A-Z0-9]{2,3})'
    rfc_patterns = [
        re.compile(r'Registro\s+Federal\s+de\s+Contribuyentes\s*:\s*' + rfc_value_re, re.IGNORECASE),
        re.compile(r'(?:R\.F\.C\.?|RFC)\s*:?\s*(?:Cliente\s+)?' + rfc_value_re, re.IGNORECASE),
    ]
    only_rfc_pattern = re.compile(r'^([A-ZÑ]{3,4}\s*\d{6}\s*[A-Z0-9]{2,3})$', re.IGNORECASE)
    # Line is "RFC" or "R.F.C." label only (no value on same line) - e.g. HSBC where value is on next line
    rfc_label_only = re.compile(r'\b(?:R\.F\.C\.?|RFC)\s*$', re.IGNORECASE)

    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue
        if rfc is None:
            for pat in rfc_patterns:
                m = pat.search(line_stripped)
                if m:
                    rfc = re.sub(r'\s+', '', m.group(1)).upper()
                    break
            if rfc is not None:
                break
    # HSBC only: line contains "RFC" (anywhere), then value on same line or 1-2 lines after. Match first value with RFC structure.
    # RFC structure (SAT): Persona física 13 chars (4 letters + 6 digits yymmdd + 3 homoclave), persona moral 12 chars (3 letters + 6 digits + 3 homoclave).
    if rfc is None and detected_bank == 'HSBC':
        # RFC label anywhere in line (not only at end), e.g. "... RFC > Plaza 26 XAXX010101000 ..."
        hsbc_rfc_label_anywhere = re.compile(r'\b(?:R\.F\.C\.?|RFC)\b', re.IGNORECASE)
        # Pattern for RFC substring: 3-4 letters, 6 digits, 3 alphanumeric (homoclave); optional spaces for OCR.
        hsbc_rfc_substring = re.compile(r'\b([A-ZÑ]{3,4}\s*\d{6}\s*[A-Z0-9]{3})\b', re.IGNORECASE)
        for i, line in enumerate(lines):
            line_stripped = line.strip()
            label_match = hsbc_rfc_label_anywhere.search(line_stripped) if line_stripped else None
            rfc_label_found = label_match is not None
            if not line_stripped:
                continue
            if rfc_label_found:
                # First try same line: search for RFC value after the "RFC" label
                after_label = line_stripped[label_match.end():] if label_match else line_stripped
                m = hsbc_rfc_substring.search(after_label)
                if m:
                    rfc = re.sub(r'\s+', '', m.group(1)).upper()
                    break
                # Else search in the next 2 lines
                for j in range(i + 1, min(i + 3, len(lines))):
                    candidate = lines[j].strip()
                    if not candidate:
                        continue
                    m = hsbc_rfc_substring.search(candidate)
                    if m:
                        rfc = re.sub(r'\s+', '', m.group(1)).upper()
                        break
                if rfc is not None:
                    break
                break
    # Fallback: whole line is just RFC value (no prefix)
    if rfc is None:
        for line in lines:
            line_stripped = line.strip()
            if 10 <= len(line_stripped) <= 14 and only_rfc_pattern.match(line_stripped):
                rfc = re.sub(r'\s+', '', line_stripped).upper()
                break

    # Name: lines with company suffixes (SA DE CV, S.A. DE C.V., etc.), excluding bank's own name
    company_suffixes = [
        'SA DE CV', 'S.A. DE C.V.', 'S.A. DE C.V', 'S. DE R.L. DE C.V.', 'S. DE R.L. DE C.V',
        'S.R.L.', 'S. DE R.L.', 'DE C.V.', 'DE CV',
    ]
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped or len(line_stripped) < 5:
            continue
        line_upper = line_stripped.upper()
        has_suffix = any(s in line_upper for s in company_suffixes)
        if not has_suffix:
            continue
        # Exclude if line matches the detected bank's keywords (e.g. "HSBC México S.A. DE C.V.")
        if bank_keywords:
            if any(re.search(pat, line_stripped, re.IGNORECASE) for pat in bank_keywords):
                continue
        name = line_stripped
        break

    # HSBC-specific fallback: personal-name style line before address (e.g. "JUSEOG AN" followed by address)
    if name is None and detected_bank == 'HSBC':
        address_keywords = ['CLL', 'CALLE', 'SECCION', 'COL', 'PROV', 'AV.', 'AV ', 'NO ', 'NUM', 'C.P', 'CP ']
        for idx, line in enumerate(lines):
            line_stripped = line.strip()
            if not line_stripped or len(line_stripped) < 3:
                continue
            line_upper = line_stripped.upper()
            # Skip lines with digits (likely not pure name) or obvious bank/statement text
            if any(ch.isdigit() for ch in line_upper):
                continue
            if 'HSBC' in line_upper or 'ESTADO DE CUENTA' in line_upper:
                continue
            # Exclude if matches bank keywords
            if bank_keywords and any(re.search(pat, line_stripped, re.IGNORECASE) for pat in bank_keywords):
                continue
            # Check next few lines for address-like content: has digits and address keywords
            has_address_below = False
            for j in range(idx + 1, min(idx + 4, len(lines))):
                nxt = lines[j].strip()
                if not nxt:
                    continue
                nxt_upper = nxt.upper()
                if any(k in nxt_upper for k in address_keywords) and any(ch.isdigit() for ch in nxt_upper):
                    has_address_below = True
                    break
            if has_address_below:
                name = line_stripped
                break
        # HSBC: name between "Estado de Cuenta" and "Subtotal" in same line (e.g. "Estado de Cuenta 283879 2 JUSEOG AN Subtotal:")
        if name is None and detected_bank == 'HSBC':
            hsbc_estado_subtotal = re.compile(
                r'Estado\s+de\s+Cuenta\s*(?:\d+\s*)*\s*([A-Za-zÑñ\s]+?)\s*Subtotal',
                re.IGNORECASE
            )
            for idx, line in enumerate(lines):
                line_stripped = line.strip()
                if not line_stripped or 'Estado de Cuenta' not in line_stripped or 'Subtotal' not in line_stripped:
                    continue
                m = hsbc_estado_subtotal.search(line_stripped)
                if m:
                    candidate = m.group(1).strip()
                    if 2 <= len(candidate) <= 80 and not any(ch.isdigit() for ch in candidate) and 1 <= len(candidate.split()) <= 6:
                        name = candidate
                        break
        # HSBC: name (company with DE CV / S.A. DE C.V.) after "Estado de Cuenta" in same line (e.g. "... Estado de Cuenta ... HK DASA DE CV. ...")
        if name is None and detected_bank == 'HSBC':
            hsbc_estado_company = re.compile(
                r'Estado\s+de\s+Cuenta\s.*?\b([A-Za-zÑñ][A-Za-zÑñ\s]*(?:DE\s+CV|S\.A\.\s+DE\s+C\.V\.?))\s*\.?',
                re.IGNORECASE
            )
            for idx, line in enumerate(lines):
                line_stripped = line.strip()
                if not line_stripped or 'Estado de Cuenta' not in line_stripped:
                    continue
                m = hsbc_estado_company.search(line_stripped)
                if m:
                    candidate = m.group(1).strip().rstrip('.')
                    if 3 <= len(candidate) <= 80 and 'HSBC' not in candidate.upper():
                        if bank_keywords and any(re.search(pat, candidate, re.IGNORECASE) for pat in bank_keywords):
                            continue
                        name = candidate
                        break

    return (rfc, name)


def extract_summary_from_pdf(pdf_path: str, movement_start_page: int = None) -> dict:
    """
    Extract summary information from PDF (totals, deposits, withdrawals, balance, movement count).
    Uses bank-specific patterns to extract summary data accurately.
    Returns a dictionary with extracted values or None if not found.
    For INTERCAM, when movement_start_page is provided, Saldo Final is taken only from that page.
    """
    summary_data = {
        'total_depositos': None,
        'total_retiros': None,
        'total_cargos': None,
        'total_abonos': None,
        'saldo_final': None,
        'total_movimientos': None,
        'saldo_anterior': None,
        'rfc': None,
        'name': None,
        'period_text': None,
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
                # For INTERCAM, keep lines per page so Saldo Final can be taken from movements_start page only
                lines_by_page = {} if bank_name == "INTERCAM" else None
                for page_num in range(pages_to_check):
                    page = pdf.pages[page_num]
                    text = page.extract_text()
                    if text:
                        all_text += text + "\n"
                        page_lines = text.split('\n')
                        all_lines.extend(page_lines)
                        if lines_by_page is not None:
                            lines_by_page[page_num + 1] = page_lines  # 1-based page number
                
                # Also check last page for Santander and BanRegio
                if len(pdf.pages) > pages_to_check:
                    last_page = pdf.pages[-1]
                    last_text = last_page.extract_text()
                    if last_text:
                        last_lines = last_text.split('\n')
                        all_lines.extend(last_lines)
                        if lines_by_page is not None:
                            lines_by_page[len(pdf.pages)] = last_lines  # 1-based last page number
                
                # For INTERCAM, ensure we have the movements_start page for Saldo Final (may be beyond first pages)
                if lines_by_page is not None and movement_start_page is not None:
                    if 1 <= movement_start_page <= len(pdf.pages) and movement_start_page not in lines_by_page:
                        page = pdf.pages[movement_start_page - 1]
                        text = page.extract_text()
                        if text:
                            lines_by_page[movement_start_page] = text.split('\n')
            
            # Extract RFC, name, period from full text (all banks)
            full_text = all_text if all_text else '\n'.join(all_lines)
            if full_text:
                summary_data['period_text'] = extract_period_text_from_text(full_text)
                rfc_val, name_val = extract_rfc_and_name_from_text(full_text, detected_bank=bank_name)
                summary_data['rfc'] = rfc_val
                summary_data['name'] = name_val
            
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
                            #print(f"✅ Santander: Found WITHDRAWALS: ${retiros:,.2f}")
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
            
            elif bank_name == "Mercury":
                # Mercury: Total Cargos from line containing "Total withdrawals" (absolute value, e.g. -$9,292.00 -> $9,292.00)
                for line in all_lines:
                    if not summary_data['total_cargos'] and 'Total withdrawals' in line.replace('\r', ' '):
                        # Capture amount: optional minus before/after optional $ (e.g. -$9,292.00 or $-9,292.00 or $9,292.00)
                        match = re.search(r'Total\s+withdrawals\s+([\-]?\s*\$?\s*[\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount is not None and amount != 0:
                                summary_data['total_cargos'] = abs(amount)
                                summary_data['total_retiros'] = summary_data['total_cargos']
                    if not summary_data['total_abonos'] and 'Total deposits' in line.replace('\r', ' '):
                        match = re.search(r'Total\s+deposits\s+\$?\s*([\d,\.\-]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount is not None and amount != 0:
                                summary_data['total_abonos'] = abs(amount)
                                summary_data['total_depositos'] = summary_data['total_abonos']
                    if summary_data.get('saldo_final') is None and 'Statement balance' in line.replace('\r', ' '):
                        match = re.search(r'Statement\s+balance\s+\$?\s*([\d,\.\-]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount is not None:
                                summary_data['saldo_final'] = amount
            
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
            
            elif bank_name == "INTERCAM":
                # INTERCAM: Total Abonos from "+ Depósitos", Total Cargos from "- Retiros", Saldo Final from "Saldo Final"
                # For Saldo Final, use only the page where movements_start was found (when movement_start_page is provided)
                saldo_lines = all_lines
                if movement_start_page is not None and lines_by_page is not None:
                    saldo_lines = lines_by_page.get(movement_start_page, all_lines)
                for i, line in enumerate(all_lines):
                    # Total Abonos: line with "+ Depósitos" (or "+ Depositos") followed by amount
                    if not summary_data['total_abonos']:
                        if re.search(r'\+\s*Dep[oó]sitos', line, re.I):
                            match = re.search(r'\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})', line)
                            if match:
                                amount = normalize_amount_str(match.group(0))
                                if amount > 0:
                                    summary_data['total_abonos'] = amount
                                    summary_data['total_depositos'] = amount
                    # Total Cargos: line with "- Retiros" followed by amount
                    if not summary_data['total_cargos']:
                        if re.search(r'-\s*Retiros', line, re.I):
                            match = re.search(r'\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})', line)
                            if match:
                                amount = normalize_amount_str(match.group(0))
                                if amount > 0:
                                    summary_data['total_cargos'] = amount
                                    summary_data['total_retiros'] = amount
                # Saldo Final: only from lines on the movements_start page (when provided), else all lines
                for i, line in enumerate(saldo_lines):
                    if not summary_data['saldo_final'] and re.search(r'Saldo\s+Final', line, re.I):
                        match = re.search(r'\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})', line)
                        if match:
                            amount = normalize_amount_str(match.group(0))
                            summary_data['saldo_final'] = amount
                            break
            
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
                        # Try patterns from BANK_CONFIGS if available
                        clara_config = BANK_CONFIGS.get('Clara', {})
                        total_patterns = clara_config.get('summary_total_patterns', [])
                        if not total_patterns:
                            # Fallback to default patterns if not in config
                            total_patterns = [
                                r'Total\s+MXN\s+([\d,\.\-]+)\s+MXN\s+([\d,\.\-]+)',  # "Total MXN 3,115.30 MXN -3,305.40"
                                r'Total\s+([\d,\.\-]+)\s+([\d,\.\-]+)',  # "Total 3,115.30 -3,305.40" (fallback)
                            ]
                        for pattern_str in total_patterns:
                            pattern = re.compile(pattern_str, re.I)
                            match = pattern.search(line)
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
            
            elif bank_name == "Mercury":
                # Mercury: Total Cargos from line containing "Total withdrawals" (absolute value, e.g. -$9,292.00 -> $9,292.00)
                for i, line in enumerate(all_lines):
                    if not summary_data['total_cargos'] and 'Total withdrawals' in line.replace('\r', ' '):
                        match = re.search(r'Total\s+withdrawals\s+([\-]?\s*\$?\s*[\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount is not None:
                                summary_data['total_cargos'] = abs(amount)
                                summary_data['total_retiros'] = summary_data['total_cargos']
                    if not summary_data['total_abonos'] and 'Total deposits' in line.replace('\r', ' '):
                        match = re.search(r'Total\s+deposits\s+\$?\s*([\d,\.\-]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount is not None:
                                summary_data['total_abonos'] = abs(amount)
                                summary_data['total_depositos'] = summary_data['total_abonos']
                    if not summary_data['saldo_final'] and 'Statement balance' in line.replace('\r', ' '):
                        match = re.search(r'Statement\s+balance\s+\$?\s*([\d,\.\-]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount is not None:
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
    
    # Initialize df_for_abonos and df_for_cargos (will be set for specific banks)
    df_for_abonos = df_for_totals  # Default: use same DataFrame for Abonos
    df_for_cargos = df_for_totals  # Default: use same DataFrame for Cargos
    
    # For Scotiabank, exclude "COBRO DE COMISION" and "IVA POR COMISIONES" 
    # from Cargos calculation only (not from Abonos)
    # because these are not included in the summary totals on the PDF cover page
    if bank_name == 'Scotiabank' and 'Descripción' in df_for_totals.columns:
        # Create a separate DataFrame for Cargos calculation (excluding commission rows)
        df_for_cargos = df_for_totals[
            ~df_for_totals['Descripción'].astype(str).str.contains('COBRO DE COMISION', case=False, na=False) &
            ~df_for_totals['Descripción'].astype(str).str.contains('IVA POR COMISIONES', case=False, na=False)
        ]
    elif bank_name == 'Banorte' and 'Descripción' in df_for_totals.columns:
        # For Banorte, exclude commission and IVA rows from Cargos calculation only (not from Abonos)
        # because these are not included in the summary totals on the PDF cover page
        # Exclude rows that contain ANY of these strings (can be multiple rows):
        # - "COMISION ORDEN DE PAGO SPE"
        # - "I.V.A. ORDEN DE PAGO SPEI"
        # - "COMISION POR NO"
        # - "I.V.A. LIQ"
        # - "COM. DISPERSION NOMINA"
        # - "INTERESES EXENTO TASA"
        # - "IVA" + "REGISTROS EXCEDIDOS" (e.g. "IVA 00010 REGISTROS EXCEDIDOS")
        df_for_cargos = df_for_totals[
            ~(df_for_totals['Descripción'].astype(str).str.contains('COMISION ORDEN DE PAGO SPEI', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('I.V.A. ORDEN DE PAGO SPEI', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('COMISION POR NO', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('I.V.A. LIQ', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('I.V.A COM', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('COMISION ENV', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('IVA COM MEMB P.M', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('IVA MEMB P.M', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('COM MEMB', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('COM. DISPERSION NOMINA', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('INTERESES EXENTO TASA', case=False, na=False) |
              (df_for_totals['Descripción'].astype(str).str.contains('IVA', case=False, na=False) &
               df_for_totals['Descripción'].astype(str).str.contains('REGISTROS EXCEDIDOS', case=False, na=False)))
        ]
    elif bank_name == 'HSBC' and 'Descripción' in df_for_totals.columns:
        # For HSBC, exclude specific rows from Cargos calculation only (not from Abonos)
        # because these are not included in the summary totals on the PDF cover page
        # Exclude rows that contain ANY of these strings:
        # - "S.R. RETENIDO"
        hsbc_cargos_excluded_mask = df_for_totals['Descripción'].astype(str).str.contains('S.R. RETENIDO', case=False, na=False)
        df_for_cargos = df_for_totals[~hsbc_cargos_excluded_mask]
        # For HSBC, exclude "PAGO DE INTERES NOMINAL", "COMISION", and "REV" from Abonos calculation
        # because these are not included in the summary totals on the PDF cover page
        df_for_abonos = df_for_totals[
            ~(df_for_totals['Descripción'].astype(str).str.contains('PAGO DE INTERES NOMINAL', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('COMISION', case=False, na=False) |
              df_for_totals['Descripción'].astype(str).str.contains('REV', case=False, na=False))
        ]
    
    # Calculate based on available columns
    if 'Abonos' in df_for_totals.columns:
        # For HSBC, use df_for_abonos (which excludes "PAGO DE INTERES NOMINAL", "COMISION", and "REV")
        # For other banks, df_for_abonos will be the same as df_for_totals
        
        totals['total_abonos'] = df_for_abonos['Abonos'].apply(normalize_amount_str).sum()
        totals['total_depositos'] = totals['total_abonos']
    
    if 'Cargos' in df_for_totals.columns:
        # For Scotiabank, Banorte, and HSBC, use df_for_cargos (which excludes commission rows)
        # For other banks, df_for_cargos will be the same as df_for_totals
        totals['total_cargos'] = df_for_cargos['Cargos'].apply(normalize_amount_str).sum()
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
    valor_en_pdf = f"${pdf_abonos:,.2f}" if pdf_abonos else "Not found"
    
    validation_row = {
        'Concepto': 'Total Abonos / Depósitos',
        'Valor en PDF': valor_en_pdf,
        'Valor Extraído': f"${ext_abonos:,.2f}",
        'Diferencia': f"${abs(pdf_abonos - ext_abonos):,.2f}" if pdf_abonos else "N/A",
        'Estado': '✓' if abonos_match else '✗'
    }
    validation_data.append(validation_row)
    
    # Compare Cargos/Retiros
    pdf_cargos = pdf_summary.get('total_cargos') or pdf_summary.get('total_retiros')
    ext_cargos = extracted_totals.get('total_cargos', 0.0)
    cargos_match = pdf_cargos is None or abs(pdf_cargos - ext_cargos) < tolerance
    validation_data.append({
        'Concepto': 'Total Cargos / Retiros',
        'Valor en PDF': f"${pdf_cargos:,.2f}" if pdf_cargos else "Not found",
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
            'Valor en PDF': f"${pdf_saldo:,.2f}" if pdf_saldo else "Not found",
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
        'Valor en PDF': str(pdf_mov) if pdf_mov else "Not found",
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


def has_numeric_values_in_movements(df_mov: pd.DataFrame) -> bool:
    """
    Verifica si df_mov tiene al menos un valor numérico en columnas numéricas.
    
    Args:
        df_mov: DataFrame con los movimientos del banco
    
    Returns:
        True si hay al menos un valor numérico > 0, False en caso contrario
    """
    # Columnas numéricas a verificar (case-insensitive)
    numeric_columns = ['Cargos', 'Abonos', 'Saldo', 'Liquidación']
    
    # Buscar columnas que coincidan (case-insensitive)
    for col in df_mov.columns:
        col_lower = col.lower()
        if col_lower in [nc.lower() for nc in numeric_columns]:
            # Intentar convertir valores a numéricos
            for val in df_mov[col]:
                if pd.notna(val) and str(val).strip():
                    # Excluir la fila "Total" si existe
                    if str(val).strip().lower() == 'total':
                        continue
                    normalized = normalize_amount_str(str(val).strip())
                    if normalized is not None and normalized > 0:
                        return True
    return False


def all_differences_are_na(validation_df: pd.DataFrame) -> bool:
    """
    Verifica si todas las filas de 'Diferencia' (excepto VALIDACIÓN GENERAL) son "N/A".
    
    Args:
        validation_df: DataFrame de validación
    
    Returns:
        True si todas las diferencias son "N/A", False en caso contrario
    """
    # Filtrar filas excluyendo 'VALIDACIÓN GENERAL'
    filtered_df = validation_df[validation_df['Concepto'] != 'VALIDACIÓN GENERAL']
    
    if len(filtered_df) == 0:
        return False  # No hay filas para verificar
    
    # Verificar que todas las diferencias sean "N/A"
    all_na = (filtered_df['Diferencia'] == "N/A").all()
    return all_na


def print_validation_summary(pdf_summary: dict, extracted_totals: dict, validation_df: pd.DataFrame, df_mov: pd.DataFrame):
    """
    Print validation summary to console with checkmarks or X marks.
    """
    # print("\n" + "=" * 80)
    # print("📊 VALIDACIÓN DE DATOS")
    # print("=" * 80)
    
    # Verificar casos especiales ANTES de evaluar overall_status
    force_error = False
    if not has_numeric_values_in_movements(df_mov):
        force_error = True
    elif all_differences_are_na(validation_df):
        force_error = True
    
    # Check overall status
    overall_status = validation_df[validation_df['Concepto'] == 'VALIDACIÓN GENERAL']['Estado'].values[0]
    
    # Reusar código existente: si force_error es True, forzar que caiga en else
    if '✓' in overall_status and not force_error:
        print("✅ VALIDATION: ALL CORRECT", flush=True)
    else:
        print("❌ VALIDATION: THERE ARE DIFFERENCES")
        pass
    
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


def extract_digitem_section(pdf_path: str, columns_config: dict, extracted_data: list = None) -> pd.DataFrame:
    """
    Extract DIGITEM section from Banamex PDF using the same coordinate-based extraction as Movements.
    Section starts with "DIGITEM" and ends with "TRANSFERENCIA ELECTRONICA DE FONDOS".
    Returns a DataFrame with columns: Fecha, Descripción, Importe

    If extracted_data is provided (e.g. from main flow), it is used to avoid re-opening/re-OCR of the PDF.
    """
    digitem_rows = []
    
    try:
        # Use provided extracted data to avoid second PDF extraction (and second OCR for Banamex mixed)
        if extracted_data is None:
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
                    #print(f"📄 DIGITEM section found on page {page_num}")
                    # Skip the header line "DETALLE DE OPERACIONES" that comes after DIGITEM
                    skip_next_line = True
                else:
                    # Also check in words (in case text extraction missed it)
                    all_words_text = ' '.join([w.get('text', '') for w in words])
                    if re.search(r'\bDIGITEM\b', all_words_text, re.I):
                        in_digitem_section = True
                        # print(f"📄 DIGITEM section found on page {page_num} (from words)")
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
            
            #print(f"✅ Extracted {len(df_digitem)} DIGITEM records from PDF")
            return df_digitem
        else:
            pass
            # print("ℹ️  DIGITEM section not found in PDF")
            return pd.DataFrame(columns=['Fecha', 'Descripción', 'Importe'])
    
    except Exception as e:
        pass
        # print(f"⚠️  Error extracting DIGITEM from PDF: {e}")
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
        #print(f"⚠️  Error extracting TRANSFERENCIA from PDF: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame(columns=['Fecha', 'Descripción', 'Importe', 'Comisiones', 'I.V.A', 'Total'])


def extract_text_from_pdf(pdf_path: str) -> list:
    """
    Extract text and word positions from each page of a PDF.
    Returns a list of dictionaries (page_number, text, words).
    
    If it detects illegible text (CID characters), uses Tesseract OCR as fallback.
    """
    # STEP 1: Detect if PDF has illegible text
    is_illegible, cid_ratio, ascii_ratio = is_pdf_text_illegible(pdf_path)
    if '--debug' in sys.argv:
        print(f"[DEBUG] is_pdf_text_illegible: is_illegible={is_illegible}, cid_ratio={cid_ratio:.2%}, ascii_ratio={ascii_ratio:.2%}", flush=True)
    
    # Banamex mixed format: if bank is Banamex and ascii_ratio < 99%, use OCR for movements (some content is embedded as images)
    detected_bank_early = detect_bank_from_pdf(pdf_path)
    use_ocr_banamex_mixed = (
        detected_bank_early == 'Banamex' and ascii_ratio < 0.99 and TESSERACT_AVAILABLE
    )
    if use_ocr_banamex_mixed:
        print(f"[INFO] Banamex PDF with mixed content (ASCII ratio: {ascii_ratio:.2%} < 99%). Using OCR for movements...", flush=True)
    
    if (is_illegible or use_ocr_banamex_mixed) and TESSERACT_AVAILABLE:
        if is_illegible:
            print(f"[INFO] PDF detected as illegible (CID ratio: {cid_ratio:.2%}, ASCII ratio: {ascii_ratio:.2%})", flush=True)
        print(f"[INFO] Using local Tesseract OCR as fallback...", flush=True)
        print(f"[INFO] Bank will be detected after processing with OCR...", flush=True)
        try:
            # Use Tesseract OCR
            extracted_data = extract_text_with_tesseract_ocr(pdf_path)
            # Mark that OCR was used
            for page_data in extracted_data:
                page_data['_used_ocr'] = True
            # When OCR was triggered for Banamex mixed, keep Banamex as bank (OCR text may not detect it or may match HSBC fallback)
            if use_ocr_banamex_mixed and extracted_data:
                extracted_data[0]['_force_bank'] = 'Banamex'
            return extracted_data
        except Exception as e:
            print(f"[WARNING] Error with Tesseract OCR: {e}")
            print(f"[INFO] Continuing with normal extraction (may have illegible characters)...")
            # Continue with normal extraction if OCR fails
    
    # STEP 2: Normal extraction with pdfplumber (existing code - NO CHANGES)
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
    print(f"✅ Excel file created -> {output_path}", flush=True)


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


def _santander_deduplicate_string(s: str) -> str:
    """If s has every character repeated (e.g. '0066--EENNEE' -> '06-ENE'), return deduplicated string; else return s unchanged."""
    if not s or len(s) < 2:
        return s
    n = len(s)
    if n % 2 != 0:
        return s
    if all(s[i] == s[i + 1] for i in range(0, n - 1, 2)):
        return s[0::2]
    return s


def _santander_sanitize_row_words_if_duplicated(row_words: list) -> list:
    """For Santander: if any word text has every-char-repeated pattern, return new row_words with sanitized text; else return original."""
    if not row_words:
        return row_words
    out = []
    any_changed = False
    for w in row_words:
        text = (w.get('text') or '').strip()
        dedup = _santander_deduplicate_string(text)
        if dedup != text:
            any_changed = True
        tw = dict(w)
        tw['text'] = dedup
        out.append(tw)
    return out if any_changed else row_words


def extract_santander_metas_from_pdf(extracted_data, columns_config, metas_start, metas_end):
    """Extract Santander METAS section (Mis Metas) using same coordinate-based logic as Movements.
    Section starts with metas_start and ends with a line containing metas_end (e.g. TOTAL).
    Returns a DataFrame with columns: Fecha, Descripción, Abonos, Cargos, Saldo (same structure as main movements)."""
    if not extracted_data or not columns_config or not metas_start or not metas_end:
        return None
    date_pattern = re.compile(r"\b(?:(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4})\b", re.I)
    start_norm = re.sub(r'\s+', '', metas_start).upper()
    end_word = metas_end.strip().upper()
    metas_rows = []
    in_metas = False
    extraction_stopped = False
    skip_next_line = False
    for page_data in extracted_data:
        if extraction_stopped:
            break
        page_num = page_data.get('page', 0)
        words = page_data.get('words', [])
        text = page_data.get('content', '') or ''
        if not words and not text:
            continue
        if not in_metas:
            text_norm = re.sub(r'\s+', '', text).upper()
            if start_norm in text_norm:
                in_metas = True
                skip_next_line = True
            else:
                continue
        if in_metas and words:
            word_rows = group_words_by_row(words, y_tolerance=3)
            for row_words in word_rows:
                if not row_words:
                    continue
                all_row_text = ' '.join([w.get('text', '') for w in row_words])
                if skip_next_line:
                    if start_norm in re.sub(r'\s+', '', all_row_text).upper() or re.search(r'DETALLE\s+DE\s+MOVIMIENTOS\s+MIS\s+METAS', all_row_text, re.I):
                        skip_next_line = False
                    continue
                if re.search(r'\b' + re.escape(end_word) + r'\b', all_row_text, re.I):
                    extraction_stopped = True
                    break
                row_words_sanitized = _santander_sanitize_row_words_if_duplicated(row_words)
                row_data = extract_movement_row(row_words_sanitized, columns_config, 'Santander', date_pattern)
                fecha_val = str(row_data.get('fecha') or '').strip()
                has_date = bool(date_pattern.search(fecha_val))
                has_amount = any(row_data.get(k) for k in ('cargos', 'abonos', 'saldo') if row_data.get(k))
                if has_date or has_amount:
                    metas_rows.append(row_data)
    if not metas_rows:
        return None
    df = pd.DataFrame([
        {
            'Fecha': str(r.get('fecha') or ''),
            'Descripción': str(r.get('descripcion') or ''),
            'Abonos': str(r.get('abonos') or ''),
            'Cargos': str(r.get('cargos') or ''),
            'Saldo': str(r.get('saldo') or ''),
        }
        for r in metas_rows
    ])
    return df


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
            
            # Validate and correct inverted ranges
            if x_min > x_max:
                print(f"[WARNING] Inverted range for '{col_name}': ({x_min}, {x_max}). Correcting to ({x_max}, {x_min})")
                x_min, x_max = x_max, x_min  # Swap values
            
            if x_min <= word_center <= x_max:
                # Calculate distance from amount center to range center
                range_center = (x_min + x_max) / 2
                distance = abs(word_center - range_center)
                matching_cols.append((col_name, distance, range_center))
    
    # If there are multiple matches, return the closest to the range center
    if matching_cols:
        # Sort by distance (closest first)
        matching_cols.sort(key=lambda x: x[1])
        return matching_cols[0][0]  # Return the closest column
    
    # Then check other columns (fecha, liq, descripcion, etc.)
    for col_name, (x_min, x_max) in columns.items():
        if col_name not in numeric_cols:  # Skip numeric cols, already checked
            # Validate and correct inverted ranges
            if x_min > x_max:
                x_min, x_max = x_max, x_min  # Swap values
            
            if x_min <= word_center <= x_max:
                return col_name
    
    return None


def is_transaction_row(row_data, bank_name=None, debug_only_if_contains_iva=False):
    """Check if a row is an actual bank transaction (not a header or empty row).
    A transaction must have:
    - A date in 'fecha' column
    - At least one VALID amount in cargos, abonos, or saldo (must match amount pattern)
    - A non-empty description (for most banks)
    """
    fecha = (row_data.get('fecha') or '').strip()
    cargos = (row_data.get('cargos') or '').strip()
    abonos = (row_data.get('abonos') or '').strip()
    saldo = (row_data.get('saldo') or '').strip()
    descripcion = (row_data.get('descripcion') or row_data.get('Descripción') or '').strip()
    
    # Must have a date matching DD/MMM pattern
    # For HSBC: date is only 2 digits (01-31)
    # For Banregio: date is only 2 digits (01-31)
    # For Base: date format is DD/MM/YYYY (e.g., "30/04/2024")
    # For other banks: supports both "DIA MES" (01 ABR) and "MES DIA" (ABR 01) formats
    # Pattern for dates: supports multiple formats including "DIA MES AÑO" (06 mar 2023)
    has_date = False
    if bank_name == 'HSBC':
        # For HSBC, date is only 2 digits (01-31)
        # Clean the date: remove dots, spaces and other non-numeric characters at the end
        fecha_clean = fecha.strip().rstrip('.,;:')
        hsbc_date_re = re.compile(r'^(0[1-9]|[12][0-9]|3[01])$')
        has_date = bool(hsbc_date_re.match(fecha_clean))
    elif bank_name == 'Banregio':
        # For Banregio, date is only 2 digits (01-31)
        banregio_date_re = re.compile(r'^(0[1-9]|[12][0-9]|3[01])$')
        has_date = bool(banregio_date_re.match(fecha))
        print("Debug: has_date", has_date, banregio_date_re)
    elif bank_name == 'Base':
        # For Base, date format is DD/MM/YYYY (e.g., "30/04/2024")
        base_date_re = re.compile(r'^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(\d{4})$')
        has_date = bool(base_date_re.match(fecha))
    elif bank_name == 'Inbursa':
        # For Inbursa: full "MES. DD" / "MES DD" or month-only when date is split across 2 lines (e.g. "Nov.", "NOV.")
        inbursa_full_date_re = re.compile(r'^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\.?\s+(0[1-9]|[12][0-9]|3[01])\b', re.I)
        inbursa_month_only_re = re.compile(r'^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\.?$', re.I)
        has_date = bool(inbursa_full_date_re.match(fecha.strip()) or inbursa_month_only_re.match(fecha.strip()))
    elif bank_name == 'Mercury':
        # Mercury: "Jul 01" - 3-letter month (Jan, Feb, Mar, ...) + space + day (1-31)
        mercury_date_re = re.compile(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(0?[1-9]|[12][0-9]|3[01])$', re.I)
        has_date = bool(mercury_date_re.match(fecha.strip()))
    else:
        # For other banks, use the general date pattern
        day_re = re.compile(r"\b(?:(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}(?:[\/\-\s]\d{2,4})?|[A-Za-z]{3}(?:[\/\-\s])(?:0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4})\b", re.I)
        has_date = bool(day_re.search(fecha))
    
    # Must have at least one VALID numeric amount (must match DEC_AMOUNT_RE pattern)
    # This ensures we don't accept text like "15 de mayo de" or "al 15 de" as amounts
    has_valid_amount = False
    has_valid_cargos = False
    has_valid_abonos = False
    has_valid_saldo = False
    
    # Pattern to validate amounts: must match DEC_AMOUNT_RE (digits with optional commas/periods)
    # Also accept amounts with $ sign at the start
    amount_validation_pattern = re.compile(r'^[\$]?\s*\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})?$')
    
    if cargos:
        # Check if cargos is a valid amount (not just any text)
        # Clean the amount: remove $, spaces and commas for validation
        cargos_clean = cargos.replace(' ', '').replace(',', '').replace('$', '').strip()
        if amount_validation_pattern.match(cargos_clean):
            has_valid_amount = True
            has_valid_cargos = True
        elif DEC_AMOUNT_RE.search(cargos):
            has_valid_amount = True
            has_valid_cargos = True
        if bank_name == 'HSBC' and not has_valid_amount:
            pass  # Debug removido
    
    if abonos:
        abonos_clean = abonos.replace(' ', '').replace(',', '').replace('$', '').strip()
        if amount_validation_pattern.match(abonos_clean):
            has_valid_amount = True
            has_valid_abonos = True
        elif DEC_AMOUNT_RE.search(abonos):
            has_valid_amount = True
            has_valid_abonos = True
    
    if saldo:
        # Check if saldo is a valid amount (not just any text)
        saldo_clean = saldo.replace(' ', '').replace(',', '').replace('$', '').strip()
        if amount_validation_pattern.match(saldo_clean):
            has_valid_amount = True
            has_valid_saldo = True
        elif DEC_AMOUNT_RE.search(saldo):
            has_valid_amount = True
            has_valid_saldo = True
    
    # For HSBC, require non-empty description only when other columns are not enough to validate the row
    has_valid_description = True
    if bank_name == 'HSBC':
        # Allow empty/short description when Fecha and at least one of Cargos/Abonos/Saldo are valid
        if not descripcion or len(descripcion.strip()) < 3:
            if has_date and has_valid_amount:
                has_valid_description = True  # other columns OK, accept row
            else:
                has_valid_description = False
        # For HSBC, we don't reject descriptions that only contain numbers/spaces
        # because if they're in the description column coordinates, they are valid descriptions
        # (e.g., reference numbers, transaction IDs, etc.)
    
    # Normal case: date, amount and description must be valid
    if has_date and has_valid_amount and has_valid_description:
        return True
    # HSBC only: if only fecha is failing (not 01-31) but all other columns are valid, accept the row
    if bank_name == 'HSBC' and not has_date and has_valid_amount and has_valid_description and (fecha or '').strip():
        return True
    return False


def extract_movement_row(words, columns, bank_name=None, date_pattern=None, debug_only_if_contains_iva=False):
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
        elif bank_name == 'INTERCAM':
            # INTERCAM: fecha column is day only (1-31)
            date_pattern = re.compile(r"^(0?[1-9]|[12][0-9]|3[01])$")
        elif bank_name == 'Base':
            # For Base, date format is DD/MM/YYYY (e.g., "30/04/2024")
            date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(\d{4})\b')
        elif bank_name == 'Banorte':
            # Banorte: DD/MM/YYYY (e.g. 31/12/2024) and DIA-MES-AÑO (e.g. 12-ENE-23) — hyphen form is DD-MMM-YY only
            date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(\d{4})\b|\b(\d{1,2}-[A-Z]{3}-\d{2})\b', re.I)
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
            
            # Apply OCR error correction in dates before searching for pattern
            fecha_text_corrected = fix_ocr_date_errors(fecha_text, bank_name)
            if fecha_text_corrected != fecha_text:
                fecha_text = fecha_text_corrected
            
            # For HSBC, date is only 2 digits (01-31)
            hsbc_date_match = date_pattern.search(fecha_text)
            if hsbc_date_match:
                date_text = hsbc_date_match.group(1)  # Just the 2 digits (01-31)
                # Clean the date: remove dots, spaces and other non-numeric characters
                date_text = date_text.strip().rstrip('.,;:')
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
    
    # For INTERCAM, extract day (1-31) from fecha column range using same coordinate logic as other banks
    if bank_name == 'INTERCAM' and 'fecha' in columns:
        fecha_x0, fecha_x1 = columns['fecha']
        fecha_words = [w for w in sorted_words if fecha_x0 <= (w.get('x0', 0) + w.get('x1', 0)) / 2 <= fecha_x1]
        if fecha_words:
            fecha_text = ' '.join([w.get('text', '').strip() for w in fecha_words])
            intercam_day_match = re.match(r'^(0?[1-9]|[12][0-9]|3[01])\b', fecha_text.strip())
            if intercam_day_match:
                row_data['fecha'] = intercam_day_match.group(1)
                sorted_words = [w for w in sorted_words if w not in fecha_words]
    
    # For Banorte, first try to extract date (DD/MM/YYYY or DIA-MES-AÑO) from fecha column or from full row text
    if bank_name == 'Banorte' and 'fecha' in columns and not row_data.get('fecha'):
        banorte_date_re = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(\d{4})\b|\b(\d{1,2}-[A-Z]{3}-\d{2})\b', re.I)
        # Try fecha column range first
        fecha_x0, fecha_x1 = columns['fecha']
        fecha_words = [w for w in sorted_words if fecha_x0 <= (w.get('x0', 0) + w.get('x1', 0)) / 2 <= fecha_x1]
        if fecha_words:
            fecha_text = ' '.join([w.get('text', '').strip() for w in fecha_words])
            m = banorte_date_re.search(fecha_text)
            if m:
                row_data['fecha'] = m.group(0)
                sorted_words = [w for w in sorted_words if w not in fecha_words]
        if not row_data.get('fecha'):
            # Fallback: search in full row text (first word often contains "31/12/2024 CONCEPT...")
            full_row_text = ' '.join([w.get('text', '').strip() for w in sorted_words])
            m = banorte_date_re.search(full_row_text)
            if m:
                row_data['fecha'] = m.group(0)
    
    # For Mercury, extract "Mon DD" (e.g. "Jul 01") from fecha column range (47-69)
    if bank_name == 'Mercury' and 'fecha' in columns and not row_data.get('fecha'):
        fecha_x0, fecha_x1 = columns['fecha']
        fecha_words = [w for w in sorted_words if fecha_x0 <= (w.get('x0', 0) + w.get('x1', 0)) / 2 <= fecha_x1]
        if fecha_words:
            fecha_text = ' '.join([w.get('text', '').strip() for w in fecha_words])
            mercury_date_re = re.compile(r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(0?[1-9]|[12][0-9]|3[01])\b', re.I)
            m = mercury_date_re.search(fecha_text)
            if m:
                row_data['fecha'] = m.group(0)
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
    
    # For HSBC, first try to reconstruct amounts that are split across multiple words
    # Example: "$", "30", ".40" or "$", "30.40" must be combined before detecting
    # Also: "30" and ".40" -> "30.40" (amounts without $ that are split)
    if bank_name == 'HSBC' and len(sorted_words) > 1:
        # Create a list of words with combined text for amounts
        combined_words = []
        i = 0
        while i < len(sorted_words):
            current_word = sorted_words[i].copy()
            current_text = current_word.get('text', '').strip()
            current_x0 = current_word.get('x0', 0)
            current_x1 = current_word.get('x1', 0)
            current_center = (current_x0 + current_x1) / 2
            
            # If the current word is "$" or starts with "$", try to combine with following words
            if current_text == '$' or (current_text.startswith('$') and len(current_text) == 1):
                # Search for following words that may be part of the amount
                combined_text = current_text
                combined_x0 = current_x0
                combined_x1 = current_x1
                j = i + 1
                
                # Combine up to 4 following words if they form a valid amount
                # Example: "$" + "30" + ".40" -> "$ 30.40"
                while j < len(sorted_words) and j < i + 5:
                    next_word = sorted_words[j]
                    next_text = next_word.get('text', '').strip()
                    next_x0 = next_word.get('x0', 0)
                    next_x1 = next_word.get('x1', 0)
                    next_center = (next_x0 + next_x1) / 2
                    
                    # Check if they are close horizontally (within 100 pixels)
                    if abs(next_center - current_center) < 100:
                        # Try to combine
                        test_combined = (combined_text + ' ' + next_text).strip()
                        # Check if the combined text forms a valid amount
                        # Improved pattern that accepts "$ 30.40", "$30.40", "$ 30 .40", etc.
                        hsbc_amount_test = re.compile(r'\$\s*\d{1,3}(?:[\s,]\d{3})*(?:\s*\.\s*\d{2}|\s+\d{2}|\s*,\s*\d{2})|\$\s*\d{1,3}\.\d{2}|\$\s*\d{1,3}\s*\.\s*\d{2}')
                        if hsbc_amount_test.search(test_combined):
                            combined_text = test_combined
                            combined_x1 = next_x1
                            combined_center = (combined_x0 + combined_x1) / 2
                            j += 1
                        else:
                            # If the next word is only ".XX" or "XX" (decimal part), continue combining
                            # This handles cases like "$" + "30" + ".40"
                            if re.match(r'^\.?\d{2}$', next_text):
                                combined_text = test_combined
                                combined_x1 = next_x1
                                j += 1
                            else:
                                break
                    else:
                        break
                
                # Validate and normalize the final combined text
                if combined_text != current_text:
                    # Normalize spaces: "$ 30 .40" -> "$ 30.40"
                    normalized = re.sub(r'\$\s*(\d+)\s+\.\s*(\d{2})', r'$ \1.\2', combined_text)
                    normalized = re.sub(r'\$\s*(\d+)\s+(\d{2})(?!\d)', r'$ \1.\2', normalized)
                    if normalized != combined_text:
                        combined_text = normalized
                
                # If something was combined, update the word
                if combined_text != current_text:
                    current_word['text'] = combined_text
                    current_word['x0'] = combined_x0
                    current_word['x1'] = combined_x1
                    # Skip the words that were combined
                    i = j
                else:
                    i += 1
            # Also combine numbers followed by ".XX" (e.g.: "30" + ".40" -> "30.40")
            elif re.match(r'^\d+$', current_text):  # Word is only digits
                # Search for next word that may be the decimal part
                if i + 1 < len(sorted_words):
                    next_word = sorted_words[i + 1]
                    next_text = next_word.get('text', '').strip()
                    next_x0 = next_word.get('x0', 0)
                    next_x1 = next_word.get('x1', 0)
                    next_center = (next_x0 + next_x1) / 2
                    
                    # Check if they are close horizontally (within 100 pixels)
                    if abs(next_center - current_center) < 100:
                        # Check if the next word is ".XX" or "XX" (decimal part)
                        if re.match(r'^\.?\d{2}$', next_text):  # ".40" or "40" (2 digits)
                            # Combine: "30" + ".40" -> "30.40" or "30" + "40" -> "30.40"
                            if next_text.startswith('.'):
                                combined_text = f"{current_text}{next_text}"
                            else:
                                combined_text = f"{current_text}.{next_text}"
                            combined_x1 = next_x1
                            current_word['text'] = combined_text
                            current_word['x0'] = current_x0
                            current_word['x1'] = combined_x1
                            # Skip the next word since it was combined
                            i += 2
                        else:
                            i += 1
                    else:
                        i += 1
                else:
                    i += 1
            else:
                i += 1
            
            combined_words.append(current_word)
        
        # Usar las palabras combinadas en lugar de las originales
        sorted_words = combined_words
    
    for word in sorted_words:
        text = word.get('text', '')
        x0 = word.get('x0', 0)
        x1 = word.get('x1', 0)
        center = (x0 + x1) / 2

        # Apply OCR error correction (only for HSBC)
        # DISABLED: No longer necessary after improving OCR resolution
        # This logic was incorrectly converting "30.40" to "$0.40"
        # if bank_name == 'HSBC' and columns:
        #     original_text = text
        #     text = fix_ocr_amount_errors(text, x0, columns, bank_name)
        #     # Update the text in the word dictionary so it's used in descriptions
        #     if text != original_text:
        #         word['text'] = text

        # Check if word is in description range BEFORE detecting amounts
        # For Banamex and HSBC: if word is in description range, don't detect amounts within it
        # This prevents values like "39.00 MXN" or "del 10 de" from being extracted as separate amounts
        word_in_description_range = False
        if 'descripcion' in columns:
            desc_x0, desc_x1 = columns['descripcion']
            # Validate and correct inverted ranges
            if desc_x0 > desc_x1:
                desc_x0, desc_x1 = desc_x1, desc_x0
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
            # For HSBC, search for all amounts in the word (there may be multiple)
            # IMPORTANT: Do not detect amounts if the word is in the description range
            if bank_name == 'HSBC':
                # Search for all amounts (with or without $) using an improved pattern that captures complete amounts
                # Skip if word is in description range (prevents "39.00 MXN" from being detected as amount)
                if not word_in_description_range:
                    # Improved pattern for HSBC that captures complete amounts even with spaces
                    # Includes amounts WITH thousand separators (1,666.37) and WITHOUT (12040506.2) so saldo is detected
                    hsbc_amount_pattern = re.compile(r'\$\s*(\d{1,3}(?:[\s,]\d{3})*\.\d{2}|\d{1,3}(?:[\s,]\d{3})*,\d{2}|\d{1,3}\.\d{2}|\d{1,3},\d{2}|\d+\.\d{2})')
                    # Also search for amounts without $ at the start
                    hsbc_amount_pattern_no_dollar = re.compile(r'(?<!\$)(?<![\.\d])(\d{1,3}(?:[\s,]\d{3})*\.\d{2}|\d{1,3}(?:[\s,]\d{3})*,\d{2}|\d{1,3}\.\d{2}|\d{1,3},\d{2}|\d+\.\d{2})(?![\.\d])')
                    
                    # First search for amounts with $
                    amount_matches = hsbc_amount_pattern.findall(text)
                    # If no amounts with $, search without $
                    if not amount_matches:
                        amount_matches = hsbc_amount_pattern_no_dollar.findall(text)
                    
                    if amount_matches:
                        for amt_match in amount_matches:
                            # Clean spaces from captured amount
                            amt_clean = re.sub(r'\s+', '', amt_match)
                            # Normalize: replace comma with dot if it's a decimal separator
                            if ',' in amt_clean and '.' not in amt_clean:
                                # Check if the comma is a decimal separator (last 2 digits after comma)
                                parts = amt_clean.split(',')
                                if len(parts) == 2 and len(parts[1]) == 2:
                                    amt_clean = parts[0] + '.' + parts[1]
                                else:
                                    # Es separador de miles, mantener pero normalizar
                                    amt_clean = amt_clean.replace(',', '')
                            elif ',' in amt_clean and '.' in amt_clean:
                                # Tiene ambos: la coma es separador de miles, el punto es decimal
                                amt_clean = amt_clean.replace(',', '')
                            
                            # Only add if the amount has at least one digit before the decimal point
                            # This avoids capturing only ".40" when it should be "30.40"
                            if amt_clean and re.match(r'^\d+', amt_clean):
                                amounts.append((amt_clean, center))
            else:
                # For Banamex: don't detect amounts if word is in description range
                if not (bank_name == 'Banamex' and word_in_description_range):
                    m = DEC_AMOUNT_RE.search(text)
                    if m:
                        amounts.append((m.group(), center))

        # For Banorte: split word that starts with DIA-MES-AÑO (DD-MMM-YY only) with no space after (e.g. "13-FEB-24COMPRA" or "27-FEB-2420..." -> fecha="13-FEB-24"/"27-FEB-24", rest to descripcion)
        if bank_name == 'Banorte' and 'fecha' in columns and 'descripcion' in columns:
            banorte_split = re.match(r'^(\d{1,2}-[A-Z]{3}-\d{2})(.*)$', text.strip(), re.I)
            if banorte_split:
                _date_text = banorte_split.group(1)
                _rest = banorte_split.group(2).strip()
                if not row_data.get('fecha'):
                    row_data['fecha'] = _date_text
                if _rest:
                    if row_data.get('descripcion'):
                        row_data['descripcion'] += ' ' + _rest
                    else:
                        row_data['descripcion'] = _rest
                continue  # skip normal column assignment for this word
        
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
                # Apply OCR error correction before searching for pattern
                text_corrected = fix_ocr_date_errors(text, bank_name)
                if text_corrected != text:
                    text = text_corrected
                    # Update the text in the word dictionary
                    word['text'] = text
                
                # The date should be exactly 2 digits (01-31) at the start
                # Extract the 2 digits and everything after as description
                # Pattern updated to handle cases like "06_1V.A." where underscore or other chars separate date from description
                # Also handle OCR errors like "2/" (should be "27") - the correction should have fixed it, but we check again
                hsbc_date_match = re.match(r'^(0[1-9]|[12][0-9]|3[01])([_\s\/]+.*)?', text)
                if hsbc_date_match:
                    date_text = hsbc_date_match.group(1)  # Just the 2 digits (01-31)
                    # Limpiar la fecha: quitar puntos, espacios y otros caracteres no numéricos
                    date_text = date_text.strip().rstrip('.,;:')
                    description_text = hsbc_date_match.group(2)  # Everything after (with leading underscore/space)
                    if description_text:
                        description_text = description_text.strip()  # Remove leading underscore/space
                        # Replace underscore with space for better readability
                        description_text = description_text.replace('_', ' ')
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
                elif debug_only_if_contains_iva:
                    print(f"      ❌ No se encontró patrón de fecha en: '{text}'")
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
            
            # For Banorte format "DIA-MES-AÑO", check if the date pattern captured the full date (DD-MMM-YY only)
            if bank_name == 'Banorte':
                # Pattern specifically for Banorte: DIA-MES-AÑO with 2-digit year only (e.g., "30-ENE-23"); \d{2} limits year to 2 digits so "27-FEB-2420..." yields "27-FEB-24"
                banorte_date_pattern = re.compile(r'(\d{1,2}-[A-Z]{3}-\d{2})', re.I)
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
            else:
                # Word is only the date (e.g. "31/12/2024") - assign to fecha and skip normal assignment
                if not row_data['fecha']:
                    row_data['fecha'] = date_text
                    continue
                elif bank_name != 'INTERCAM' and date_text not in row_data['fecha']:
                    row_data['fecha'] += ' ' + date_text
                    continue
                # INTERCAM: fecha already set (day only); do not append. Fall through to normal column assignment so
                # this word (e.g. "15" from "15 GRUPOS") is assigned to descripcion by X position.
        
        # Normal column assignment
        col_name = assign_word_to_column(x0, x1, columns)
        if col_name:
            # For Banamex and HSBC, prevent text from description range being assigned to cargos/abonos/saldo
            # Only assign to these columns if the word is actually an amount AND not in description range
            if bank_name in ('Banamex', 'HSBC') and col_name in ('cargos', 'abonos', 'saldo'):
                # If word is in description range, assign to description instead
                if word_in_description_range:
                    # This word is in description range, assign to description instead
                    if row_data.get('descripcion'):
                        row_data['descripcion'] += ' ' + text
                    else:
                        row_data['descripcion'] = text
                else:
                    # Word is not in description range, check if it's actually an amount
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
            # For Mercury: Cargos and Abonos share coordinates (430-480). Differentiate only by sign:
            # Negative (e.g. –$1,199.00, -$1,199.00) -> only Cargos. Positive (e.g. $6,830.00) -> only Abonos. No duplicates.
            elif bank_name == 'Mercury' and col_name in ('cargos', 'abonos'):
                is_amount = bool(DEC_AMOUNT_RE.search(text))
                if is_amount:
                    text_stripped = text.strip()
                    is_negative = (
                        text_stripped.startswith('-') or text_stripped.startswith('\u2013') or
                        re.match(r'^[-\u2013]\s*\$', text_stripped) or
                        re.search(r'\$\s*[-\u2013]', text_stripped) or
                        (text_stripped.startswith('(') and text_stripped.endswith(')'))
                    )
                    target_col = 'cargos' if is_negative else 'abonos'
                    other_col = 'abonos' if target_col == 'cargos' else 'cargos'
                    # Put amount only in the column for its sign; never in both (ignore col_name from shared range)
                    if row_data.get(target_col):
                        row_data[target_col] += ' ' + text
                    else:
                        row_data[target_col] = text
                    # Remove same amount from the other column so it never appears in both
                    other_val = (row_data.get(other_col) or '').strip()
                    if other_val:
                        try:
                            n_other = normalize_amount_str(other_val)
                            n_this = normalize_amount_str(text_stripped)
                            if n_other is not None and n_this is not None and abs(n_other) == abs(n_this):
                                row_data[other_col] = ''
                        except Exception:
                            pass
                else:
                    if row_data.get('descripcion'):
                        row_data['descripcion'] += ' ' + text
                    else:
                        row_data['descripcion'] = text
            # For BBVA, validate that amounts have 2 decimals before assigning to cargos/abonos/saldo
            elif bank_name == 'BBVA' and col_name in ('cargos', 'abonos', 'saldo'):
                # Validate that the text is a valid amount with 2 decimals
                # DEC_AMOUNT_RE requires: digits with optional thousands separators + 2 decimals
                is_valid_amount = bool(DEC_AMOUNT_RE.search(text))
                
                if not is_valid_amount:
                    # Not a valid amount, assign to description instead
                    if row_data.get('descripcion'):
                        row_data['descripcion'] += ' ' + text
                    else:
                        row_data['descripcion'] = text
                else:
                    # Valid amount, assign normally
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
    
    # Mercury: shared range (430-480) — amount must appear in only one of Cargos/Abonos (by sign). Remove duplicate.
    if bank_name == 'Mercury':
        c = (row_data.get('cargos') or '').strip()
        a = (row_data.get('abonos') or '').strip()
        if c and a:
            try:
                num_c = normalize_amount_str(c)
                num_a = normalize_amount_str(a)
                if num_c is not None and num_a is not None and abs(num_c) == abs(num_a):
                    is_neg = (c.startswith('-') or c.startswith('\u2013') or re.match(r'^[-\u2013]\s*\$', c))
                    if is_neg:
                        row_data['abonos'] = ''
                    else:
                        row_data['cargos'] = ''
            except Exception:
                pass
    
    # For HSBC: merge split saldo when it was parsed as "399" (description) + "344.88" (saldo) -> should be "399,344.88"
    if bank_name == 'HSBC' and row_data.get('saldo') and 'saldo' in columns:
        saldo_val = str(row_data.get('saldo', '')).strip()
        # Saldo is only decimal part (e.g. "344.88")?
        if re.match(r'^\d{1,3}\.\d{2}$', saldo_val):
            full_row_text = ' '.join([w.get('text', '') for w in words])
            # Look for "NNN DDD.DD" where DDD.DD is the current saldo (split amount)
            split_saldo_pattern = re.compile(r'(\d{1,3})\s+(\d{1,3}\.\d{2})\b')
            for m in split_saldo_pattern.finditer(full_row_text):
                if m.group(2) == saldo_val:
                    merged_saldo = m.group(1) + ',' + m.group(2)
                    row_data['saldo'] = merged_saldo
                    # Remove the trailing integer from description (e.g. "CGOF 31530 31531 399" -> "CGOF 31530 31531")
                    desc = str(row_data.get('descripcion') or '').strip()
                    parts = desc.rsplit(' ', 1)
                    if len(parts) == 2 and parts[1] == m.group(1):
                        row_data['descripcion'] = parts[0].strip()
                    break
    
    # For HSBC: merge space-separated amounts in cargos/abonos/saldo (e.g. "27 416.60" -> "27,416.60", "$ 27 416.60" -> "$ 27,416.60")
    # and remove extra space inside numbers (e.g. "4,697, 688.02" -> "4,697,688.02") so only space is between "$" and the number
    # and fix OCR "/" between digits (e.g. "6,2/1.85" -> "6,271.85" where "/" was misread "7")
    if bank_name == 'HSBC':
        _split_amount_re = re.compile(r'(\d{1,3})\s+(\d{1,3}\.\d{2})\b')
        _slash_between_digits_re = re.compile(r'(\d)/(\d)')
        def _merge_split_amount(txt):
            if not txt or not isinstance(txt, str):
                return txt
            return _split_amount_re.sub(r'\1,\2', txt)
        def _fix_slash_in_amount(txt):
            if not txt or not isinstance(txt, str):
                return txt
            # "/" between digits is often OCR misread of "7"; fix only in amount columns
            return _slash_between_digits_re.sub(r'\g<1>7\g<2>', txt)
        def _normalize_amount_spaces(txt):
            if not txt or not isinstance(txt, str):
                return txt
            s = _merge_split_amount(txt)
            # Remove comma-space inside number so "4,697, 688.02" -> "4,697,688.02"
            s = s.replace(', ', ',')
            s = _fix_slash_in_amount(s)
            return s
        for col in ('cargos', 'abonos', 'saldo'):
            if col in row_data and row_data[col]:
                row_data[col] = _normalize_amount_spaces(str(row_data[col]))
        # For HSBC: fecha is only day (01-31); if multiple numeric values ended up in fecha, keep only the first one
        if row_data.get('fecha'):
            fecha_val = str(row_data['fecha']).strip()
            day_only_re = re.compile(r'^(0[1-9]|[12][0-9]|3[01])$')
            parts = fecha_val.split()
            for part in parts:
                if day_only_re.match(part.strip()):
                    row_data['fecha'] = part.strip()
                    break
    
    # If fecha contains multiple date-like values (e.g. "02/ENE 01/ENE"): for BBVA (has liq column), first = Fecha Oper, second = Fecha Liq; else keep only the first
    if row_data.get('fecha'):
        fecha_val = str(row_data['fecha']).strip()
        first_date_re = re.compile(
            r'(?:0[1-9]|[12][0-9]|3[01])[/\-](?:0[1-9]|1[0-2]|[A-Za-z]{3})(?:[/\-]\d{2,4})?|'
            r'(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}(?:\s+\d{2,4})?',
            re.I
        )
        matches = first_date_re.findall(fecha_val)
        if len(matches) >= 2:
            row_data['fecha'] = matches[0].strip()
            if columns and 'liq' in columns:
                row_data['liq'] = (row_data.get('liq') or '').strip() or matches[1].strip()
            # else: other banks keep only first (fecha already set)
    
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
    
    # Banorte: fecha must be DD-MMM-YY only; if it has 4 digits after the month (e.g. 27-FEB-2420), keep DD-MMM-YY and move extra to descripcion
    if bank_name == 'Banorte' and row_data.get('fecha'):
        f = str(row_data['fecha']).strip()
        banorte_fecha_4digit = re.match(r'^(\d{1,2}-[A-Z]{3}-)(\d{2})(\d{2})$', f, re.I)
        if banorte_fecha_4digit:
            prefix, yy, rest = banorte_fecha_4digit.groups()
            row_data['fecha'] = prefix + yy
            extra = rest.lstrip('0') or rest
            if extra:
                desc = str(row_data.get('descripcion') or '').strip()
                row_data['descripcion'] = (extra + ' ' + desc).strip() if desc else extra

    # Banorte: cargos, abonos, saldo must be amount-like only; clear if they contain plain text (e.g. "nuestra")
    if bank_name == 'Banorte':
        for col in ('cargos', 'abonos', 'saldo'):
            if col not in row_data:
                continue
            val = str(row_data.get(col) or '').strip()
            if val and not DEC_AMOUNT_RE.fullmatch(val):
                row_data[col] = ''

    # HSBC: remove spaces within numeric values (e.g. "$ 5,121 ,926.11" -> "$ 5,121,926.11")
    if bank_name == 'HSBC':
        for col in ('cargos', 'abonos', 'saldo'):
            if col not in row_data:
                continue
            val = str(row_data.get(col) or '').strip()
            if val:
                row_data[col] = re.sub(r'\s*,\s*', ',', val)
        # HSBC: normalize fecha if it is 1-2 digits > 31 (OCR error e.g. 73 -> 23)
        if row_data.get('fecha'):
            row_data['fecha'] = fix_ocr_date_errors(str(row_data['fecha']).strip(), bank_name)

    # Banamex new format: single "monto" column; "-" prefix = Abono, "+" prefix = Cargo
    if bank_name == 'Banamex' and 'monto' in row_data:
        m = (row_data.get('monto') or '').strip()
        row_data['cargos'] = m if m.startswith('+') else ''
        row_data['abonos'] = m if m.startswith('-') else ''
        row_data['saldo'] = ''
        del row_data['monto']

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
    debug_mode = '--debug' in sys.argv
    debug_path = os.path.splitext(pdf_path)[0] + "_movements_debug.txt" if debug_mode else None
    
    # Normalize PDF path (handles UNC paths, spaces, etc.)
    pdf_path = os.path.normpath(pdf_path)
    if not os.path.isabs(pdf_path):
        pdf_path = os.path.abspath(pdf_path)
    
    # Check for --find mode
    if len(sys.argv) >= 3 and sys.argv[2] == '--find':
        page_num = int(sys.argv[3]) if len(sys.argv) > 3 else 1
        print(f"🔍 Buscando coordenadas en página {page_num}...")
        find_column_coordinates(pdf_path, page_num)
        sys.exit(0)

    if not os.path.isfile(pdf_path):
        print(f"❌ Error: File not found: {pdf_path}")
        sys.exit(1)

    if not pdf_path.lower().endswith(".pdf"):
        print(f"❌ Error: El archivo debe ser un PDF: {pdf_path}")
        sys.exit(1)
    
    # Verificar que el PDF no esté bloqueado/abierto por otro proceso
    try:
        with open(pdf_path, 'rb') as test_file:
            test_file.read(1)  # Intentar leer 1 byte
    except PermissionError:
        print(f"❌ Error: El archivo PDF está abierto o bloqueado por otro proceso: {pdf_path}")
        sys.exit(1)
    except IOError as e:
        print(f"❌ Error: No se puede acceder al archivo PDF: {pdf_path}")
        print(f"   Detalle: {e}")
        sys.exit(1)

    # Normalize output path
    output_excel = os.path.normpath(os.path.splitext(pdf_path)[0] + ".xlsx")
    
    # Validar permisos de escritura en el directorio de salida
    output_dir = os.path.dirname(output_excel) or os.getcwd()
    if not os.access(output_dir, os.W_OK):
        print(f"❌ Error: No write permissions in directory: {output_dir}")
        sys.exit(1)
    
    # Validate available disk space (optional but recommended)
    try:
        import shutil
        pdf_size = os.path.getsize(pdf_path)
        # Estimate Excel size (PDF size * 2 as safety margin)
        estimated_excel_size = pdf_size * 2
        disk_usage = shutil.disk_usage(output_dir)
        free_space = disk_usage.free
        
        if free_space < estimated_excel_size:
            print(f"❌ Error: Insufficient disk space. Available: {free_space:,} bytes, Required: {estimated_excel_size:,} bytes")
            sys.exit(1)
    except ImportError:
        # shutil not available, continue without validation
        pass
    except Exception as e:
        # If validation fails, continue (not critical)
        print(f"[WARNING] Could not validate disk space: {e}")

    print("Reading PDF...", flush=True)
    
    # Now extract full data
    extracted_data = extract_text_from_pdf(pdf_path)
    
    # Detectar si se usó OCR
    used_ocr = any(p.get('_used_ocr', False) for p in extracted_data)
    # When OCR was triggered for Banamex mixed, keep the pre-OCR bank (Banamex) instead of re-detecting from OCR text
    force_bank = (extracted_data[0].get('_force_bank') if extracted_data else None)
    
    # Detect bank: from extracted text if OCR was used, otherwise from PDF
    if force_bank:
        detected_bank = force_bank
    elif used_ocr:
        # If OCR was used, detect bank from extracted text.
        # Use only the first page for bank detection so the header (bank name) wins;
        # otherwise Phase 1 "first 20 lines" can match a different bank that appears
        # early (e.g. in a table) and HSBC later in the doc would be ignored.
        first_page_content = (extracted_data[0].get('content', '') if extracted_data else '')
        all_text = '\n'.join([p.get('content', '') for p in extracted_data])  # All pages (for extraction)
        detected_bank = detect_bank_from_text(first_page_content, from_ocr=True)
    else:
        # If OCR was not used, detect bank from PDF (normal method)
        detected_bank = detect_bank_from_pdf(pdf_path)
    
    print(f"🏦 Bank detected: {detected_bank}", flush=True)
    
    is_hsbc = (detected_bank == "HSBC")
    
    # split pages into lines (necessary for subsequent processing)
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
    elif bank_config['name'] == 'INTERCAM':
        # INTERCAM: fecha column is day only (1-31)
        date_pattern = re.compile(r"^(0?[1-9]|[12][0-9]|3[01])$")
    elif bank_config['name'] == 'Base':
        # For Base, date format is DD/MM/YYYY (e.g., "30/04/2024")
        date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(\d{4})\b')
    elif bank_config['name'] == 'Banorte':
        # Banorte: DD/MM/YYYY (e.g. 31/12/2024) and DIA-MES-AÑO (e.g. 12-ENE-23) — hyphen form is DD-MMM-YY only
        date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(\d{4})\b|\b(\d{1,2}-[A-Z]{3}-\d{2})\b', re.I)
    elif bank_config['name'] == 'Banbajío':
        # For BanBajío, date format is "DIA MES" (e.g., "3 ENE") without year
        date_pattern = re.compile(r'\b(0?[1-9]|[12][0-9]|3[01])\s+[A-Z]{3}\b', re.I)
    elif bank_config['name'] == 'Mercury':
        # Mercury: "Jul 01" - 3-char month + space + day (1-31)
        date_pattern = re.compile(r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(0?[1-9]|[12][0-9]|3[01])\b', re.I)
    else:
        date_pattern = day_re
    movement_start_found = False
    movement_start_page = None
    movement_start_index = None
    movements_lines = []
    
    # Generic movement_start_pattern from BANK_CONFIGS (for all banks)
    movement_start_pattern = None
    movement_start_string = bank_config.get('movements_start')
    if movement_start_string:
        # INTERCAM: header may appear as "DIA" or "DÍA" with variable spacing; normalize for match
        if bank_config['name'] == 'INTERCAM':
            movement_start_pattern = re.compile(
                r'D[IÍ]A\s+FOLIO\s+CONCEPTO\s+DEP[OÓ]SITOS\s+RETIROS\s+SALDO',
                re.I
            )
        else:
            movement_start_secondary = bank_config.get('movements_start_secondary')
            movement_start_tertiary = bank_config.get('movements_start_tertiary')  # list of regex strings (e.g. OCR-duplicated)
            if movement_start_secondary or movement_start_tertiary:
                parts = [re.escape(movement_start_string)]
                if movement_start_secondary:
                    # Secondary can be a string (literal) or list of strings (literals and/or regex)
                    sec_list = movement_start_secondary if isinstance(movement_start_secondary, list) else [movement_start_secondary]
                    for s in sec_list:
                        # Treat as regex if it contains '+' (one-or-more pattern) or other regex metacharacters we use
                        if '+' in s or '(\s' in s or '+)' in s:
                            parts.append(s)
                        else:
                            parts.append(re.escape(s))
                if movement_start_tertiary:
                    parts.extend(movement_start_tertiary)
                movement_start_pattern = re.compile(
                    r'(?:%s)' % '|'.join(parts),
                    re.I
                )
            else:
                # Create pattern from movements_start string (escape special chars)
                movement_start_pattern = re.compile(re.escape(movement_start_string), re.I)
    
    # Track if we've found the movements section start
    movement_section_found = False
    # Banamex new format: "DESGLOSE DE MOVIMIENTOS" (--find row Y≈802) or "CARGOS, ABONOS Y COMPRAS REGULARES"; single Monto column (- = Abono, + = Cargo)
    banamex_new_format = False
    
    for p in pages_lines:
        if not movement_start_found:
            for i, ln in enumerate(p['lines']):
                # Generic movement_start_pattern from BANK_CONFIGS (for all banks)
                if movement_start_pattern and not movement_section_found:
                    # Banamex new format: OCR may put "DESGLOSE", "DE", "MOVIMIENTOS" on separate lines (--find row Y≈802); try joined line
                    combined_banamex = None
                    if bank_config['name'] == 'Banamex' and not movement_section_found:
                        chunk = p['lines'][i:i+3]
                        combined_banamex = ' '.join((l or '').strip() for l in chunk).strip()
                        if len(combined_banamex) <= 100 and re.search(r"DESGLOSE\s+DE\s+MOVIMIENTOS", combined_banamex, re.I):
                            movement_section_found = True
                            banamex_new_format = True
                            print("[Banamex new format] movements_start line:", combined_banamex, flush=True)
                            print(flush=True)
                            continue
                        # OCR may split "CARGOS, ABONOS Y COMPRAS REGULARES" across lines; check combined chunk
                        if len(combined_banamex) <= 120 and re.search(r"CARGOS,?\s+ABONOS\s+Y\s+COMPRAS\s+REGULARES", combined_banamex, re.I):
                            movement_section_found = True
                            banamex_new_format = True
                            print("[Banamex new format] movements_start line (combined):", combined_banamex[:80], flush=True)
                            print(flush=True)
                            continue
                    if movement_start_pattern.search(ln):
                        ln_stripped = (ln or '').strip()
                        # Banamex new format: "DESGLOSE DE MOVIMIENTOS" (--find row Y≈802) or "CARGOS, ABONOS Y COMPRAS REGULARES"; only when line is short
                        is_banamex_short_header = (
                            bank_config['name'] == 'Banamex'
                            and len(ln_stripped) <= 80
                            and (
                                re.search(r"DESGLOSE\s+DE\s+MOVIMIENTOS", ln, re.I)
                                or re.search(r"CARGOS,?\s+ABONOS\s+Y\s+COMPRAS\s+REGULARES", ln, re.I)
                            )
                        )
                        is_other_bank_or_classic = (
                            bank_config['name'] != 'Banamex'
                            or re.search(r"DETALLE\s+DE\s+OPERACIONES", ln, re.I)
                        ) and len(ln_stripped) <= 120
                        if is_banamex_short_header:
                            movement_section_found = True
                            banamex_new_format = True
                            print("[Banamex new format] movements_start line:", ln_stripped, flush=True)
                            print(flush=True)
                            continue
                        if is_other_bank_or_classic:
                            movement_section_found = True
                            continue
                        # Long line that matched: only treat as section start if it does NOT look like a page header
                        # (e.g. "Banamex Pagina ... Número de tarjeta ... MENSAJES" would wrongly match regex)
                        if bank_config['name'] == 'Banamex' and (
                            re.search(r"DESGLOSE\s+DE\s+MOVIMIENTOS", ln, re.I)
                            or re.search(r"CARGOS,?\s+ABONOS\s+Y\s+COMPRAS\s+REGULARES", ln, re.I)
                        ):
                            if not re.search(r'Pagina|Número de tarjeta|MENSAJES ADICIONALES', (ln or ''), re.I):
                                movement_section_found = True
                                banamex_new_format = True
                                print("[Banamex new format] movements_start line (long):", (ln or '').strip()[:80], flush=True)
                                print(flush=True)
                                continue
                        # Long line that matched (e.g. paragraph or page header) - don't set section start, keep looking for real header
                        continue
                
                # After finding movements_start, look for first valid movement row (date)
                if movement_section_found:
                    # For Inbursa, use strict date pattern: "MES. DD" or "MES DD" at start of line
                    if bank_config['name'] == 'Inbursa':
                        inbursa_date_pattern = re.compile(r'^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\.?\s+(0[1-9]|[12][0-9]|3[01])\b', re.I)
                        if inbursa_date_pattern.search(ln.strip()):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            movements_lines.extend(p['lines'][i:])
                            break
                    # For Banbajío, accept either a date or "SALDO INICIAL"
                    elif bank_config['name'] == 'Banbajío':
                        if day_re.search(ln) or re.search(r'SALDO\s+INICIAL', ln, re.I):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            movements_lines.extend(p['lines'][i:])
                            break
                    # For Konfio, verify it's a valid date format (DIA MES AÑO, e.g., "14 mar 2023")
                    elif bank_config['name'] == 'Konfio':
                        konfio_date_in_line = date_pattern.search(ln)
                        if konfio_date_in_line:
                            konfio_full_date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}\s+\d{2,4}\b', re.I)
                            if konfio_full_date_pattern.search(ln):
                                movement_start_found = True
                                movement_start_page = p['page']
                                movement_start_index = i
                                movements_lines.extend(p['lines'][i:])
                                break
                    # For Clara, start when we find a date
                    elif bank_config['name'] == 'Clara':
                        if day_re.search(ln):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            movements_lines.extend(p['lines'][i:])
                            break
                    # For INTERCAM, fecha is day only (1-31); first data line starts with day number
                    elif bank_config['name'] == 'INTERCAM':
                        if re.match(r'^(0?[1-9]|[12][0-9]|3[01])\b', ln.strip()):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            movements_lines.extend(p['lines'][i:])
                            break
                    # For Mercury, first data line has "Mon DD" (e.g. "Jul 01")
                    elif bank_config['name'] == 'Mercury':
                        if date_pattern.search(ln):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            movements_lines.extend(p['lines'][i:])
                            break
                    # For other banks, look for date or header keywords
                    else:
                        if day_re.search(ln) or header_keywords_re.search(ln):
                            movement_start_found = True
                            movement_start_page = p['page']
                            movement_start_index = i
                            movements_lines.extend(p['lines'][i:])
                            break
                # For banks without movements_start, look for date or header directly
                elif not movement_start_pattern:
                    if day_re.search(ln) or header_keywords_re.search(ln):
                        movement_start_found = True
                        movement_start_page = p['page']
                        movement_start_index = i
                        movements_lines.extend(p['lines'][i:])
                        break
        else:
            # Already found movement start, collect all lines from this page
            # Filter out movements_start pattern if it appears again on subsequent pages
            # Headers will be automatically rejected by date validation during extraction
            if movement_start_pattern:
                filtered_lines = [ln for ln in p['lines'] if not movement_start_pattern.search(ln)]
                movements_lines.extend(filtered_lines)
            else:
                movements_lines.extend(p['lines'])

    # Banamex with OCR (mixed content): if new format not detected from lines, scan full content for new-format headers
    if (bank_config['name'] == 'Banamex' and used_ocr and not banamex_new_format and extracted_data):
        full_text = ' '.join((p.get('content') or '') for p in extracted_data)
        if re.search(r"DESGLOSE\s+DE\s+MOVIMIENTOS", full_text, re.I) or re.search(r"CARGOS,?\s+ABONOS\s+Y\s+COMPRAS\s+REGULARES", full_text, re.I):
            banamex_new_format = True
            print("[Banamex new format] detected from full content (OCR)", flush=True)

    # Banamex new format: use columns_new_format, date DD-mon-YYYY (e.g. 13-oct-2025), and movements_end_new_format
    if banamex_new_format and bank_config.get('columns_new_format'):
        columns_config = bank_config['columns_new_format'].copy()
        date_pattern = re.compile(r'\b(0?[1-9]|[12][0-9]|3[01])-[a-z]{3}-\d{4}\b', re.I)  # 13-oct-2025

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
    banamex_new_fmt_totals = {}  # Total cargos/abonos from "Cargos regulares (no a meses)" and "Pagos y abonos" for validation (Banamex new format only)
    # So trim block (later) can run for any path (e.g. HSBC OCR) without NameError
    movement_end_string = bank_config.get('movements_end')
    movement_end_pattern = None
    
    # Debug: collect (original_line, excel_cols_dict, disposition) for all banks (written to .txt at end)
    debug_movements_lines = [] if debug_mode else None

    def _banamex_new_fmt_log(orig, row_data, disp):
        """Print line info for Banamex new format only (normal operation, not gated by --debug)."""
        if not banamex_new_format:
            return
        orig_short = (orig or '')[:500]
        if disp and ('END' in disp or 'movements_end' in disp.lower()):
            print("[Banamex new format] movements_end line:", orig_short, flush=True)
            print(flush=True)
            return
        parts = []
        for k in ('fecha', 'descripcion', 'cargos', 'abonos', 'saldo'):
            v = (row_data.get(k) or '') if row_data else ''
            v = str(v).strip()[:80]
            parts.append(f"{k}={v}")
        print("[Banamex new format] ORIGINAL:", orig_short, flush=True)
        print("[Banamex new format] DIVIDED:", " | ".join(parts), flush=True)
        print("[Banamex new format] DISPOSITION:", disp, flush=True)
        print(flush=True)

    def _debug_mov_line(line_text, excel_row, disposition):
        row_for_banamex = excel_row if excel_row is not None else {}
        if debug_movements_lines is None:
            pass  # skip debug file
        else:
            excel_row = excel_row or {}
            parts = []
            for k in ('fecha', 'descripcion', 'cargos', 'abonos', 'saldo'):
                v = (excel_row.get(k) or '').strip()
                parts.append(f"{k}={v[:80] + '...' if len(v) > 80 else v}")
            excel_str = " | ".join(parts)
            orig = (line_text or '')[:500]
            debug_movements_lines.append({
                'original': orig,
                'excel': excel_str,
                'disposition': disposition,
            })
            # Also print to console (same format as debug file)
            print("ORIGINAL:", orig, flush=True)
            print("EXCEL:", excel_str, flush=True)
            print("DISPOSITION:", disposition, flush=True)
            print(flush=True)
        # Banamex new format: always log in normal operation (not gated by --debug)
        _banamex_new_fmt_log(line_text, row_for_banamex, disposition)
    
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
        end_strings_also = [bank_config['movements_end_secondary']] if bank_config.get('movements_end_secondary') else None
        
        # Filtrar palabras en la sección de movimientos (para por "Información CoDi" o "Información SPEI")
        filtered_words = filter_hsbc_movements_section(extracted_data, start_string, end_string, end_strings_also=end_strings_also)
        
        if not filtered_words:
            df_mov = pd.DataFrame(columns=['fecha', 'descripcion', 'cargos', 'abonos', 'saldo'])
        else:
            # Agrupar palabras por filas (igual que otros bancos)
            # Para HSBC, usar tolerancia Y más amplia para capturar montos que pueden estar ligeramente desalineados
            word_rows = group_words_by_row(filtered_words, y_tolerance=5)
            
            # Patrón de fecha para HSBC (solo día: 01-31)
            date_pattern = re.compile(r"^(0[1-9]|[12][0-9]|3[01])(?=\s|$)")
            
            # 🔍 HSBC DEBUG: Buffer para guardar últimas 2 filas válidas antes de movements_end
            last_two_valid_rows = []
            mov_section_line_num = 0
            
            # Extraer movimientos usando la misma lógica que otros bancos
            for row_idx, row_words in enumerate(word_rows):
                if not row_words:
                    continue
                mov_section_line_num += 1
                
                # Construir línea original desde palabras (sin ordenar)
                line_original = ' '.join([w.get('text', '') for w in row_words])
                # Solo imprimir debug si la línea contiene "I.V.A." o "IVA"
                contains_iva = 'I.V.A.' in line_original or 'IVA' in line_original or '1VA' in line_original
                
                # Extraer movimiento usando BANK_CONFIGS
                row_data = extract_movement_row(row_words, columns_config, 'HSBC', date_pattern, debug_only_if_contains_iva=contains_iva)
                
                # Obtener amounts_list
                amounts_list = row_data.get('_amounts', [])
                
                # Procesar montos detectados (_amounts) y asignarlos a columnas para HSBC
                # amounts_list ya fue obtenido arriba (en el bloque de debug o aquí)
                
                # Función auxiliar para validar que un valor es un monto válido
                def is_valid_amount(value):
                    """Valida que un valor sea un monto válido (no texto como '15 de mayo de').
                    Para HSBC acepta también montos sin separador de miles (ej. 12040506.2) para saldo."""
                    if not value or not isinstance(value, str):
                        return False
                    value_clean = value.strip()
                    if not value_clean:
                        return False
                    # HSBC: aceptar montos con solo decimales sin miles (ej. $12040506.2)
                    if re.match(r'^\$?\s*\d+[.,]\d{2}$', value_clean.replace(' ', '')):
                        return True
                    # Debe contener al menos un patrón de monto válido
                    if DEC_AMOUNT_RE.search(value_clean):
                        # Verificar que no sea solo texto sin números significativos
                        invalid_patterns = [
                            r'\b(de|del|al|la|el|en|por|para|con|sin)\b',
                            r'\b(mayo|enero|febrero|marzo|abril|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\b',
                            r'^\d+\s+(de|del|al)\s+',
                        ]
                        for pattern in invalid_patterns:
                            if re.search(pattern, value_clean, re.I):
                                return False
                        return True
                    return False
                
                # Función auxiliar para verificar si un monto está en el rango de descripción
                def is_in_description_range(center_x):
                    """Verifica si un monto está en el rango de la columna descripción"""
                    if 'descripcion' in columns_config:
                        desc_x0, desc_x1 = columns_config['descripcion']
                        # Validar y corregir rangos invertidos
                        if desc_x0 > desc_x1:
                            desc_x0, desc_x1 = desc_x1, desc_x0
                        return desc_x0 <= center_x <= desc_x1
                    return False
                
                if amounts_list and len(amounts_list) >= 1:
                    # Si hay montos pero no están asignados a columnas, asignarlos basándose en coordenadas X
                    if not row_data.get('cargos') and not row_data.get('abonos') and not row_data.get('saldo'):
                        # Asignar montos según coordenadas X (solo si son válidos Y no están en descripción)
                        for amt_text, center in amounts_list:
                            # Verificar que el monto sea válido y que NO esté en el rango de descripción
                            if is_valid_amount(amt_text) and not is_in_description_range(center):
                                col_name = assign_word_to_column(center - 50, center + 50, columns_config)
                                if col_name and col_name in ('cargos', 'abonos', 'saldo'):
                                    if not row_data.get(col_name):
                                        row_data[col_name] = amt_text
                    
                    # Cuando hay 2 montos, asignar por posición X: el más a la derecha = saldo (HSBC)
                    if len(amounts_list) == 2:
                        valid_two = [(a, c) for a, c in amounts_list if is_valid_amount(a)]
                        if len(valid_two) == 2:
                            by_x = sorted(valid_two, key=lambda x: x[1])
                            left_amt, left_center = by_x[0]
                            right_amt, right_center = by_x[1]
                            if not row_data.get('saldo'):
                                row_data['saldo'] = right_amt
                            if not row_data.get('cargos') and not row_data.get('abonos'):
                                col_name = assign_word_to_column(left_center - 50, left_center + 50, columns_config)
                                if col_name and col_name in ('cargos', 'abonos'):
                                    row_data[col_name] = left_amt
                                else:
                                    row_data['cargos'] = left_amt
                        else:
                            if not row_data.get('saldo'):
                                second_amt, second_center = amounts_list[1]
                                if is_valid_amount(second_amt) and not is_in_description_range(second_center):
                                    row_data['saldo'] = second_amt
                            if not row_data.get('cargos') and not row_data.get('abonos'):
                                first_amt, first_center = amounts_list[0]
                                if is_valid_amount(first_amt) and not is_in_description_range(first_center):
                                    col_name = assign_word_to_column(first_center - 50, first_center + 50, columns_config)
                                    if col_name and col_name in ('cargos', 'abonos'):
                                        row_data[col_name] = first_amt
                                    else:
                                        row_data['cargos'] = first_amt
                    elif len(amounts_list) == 1:
                        # Si solo hay un monto, asignarlo según coordenadas
                        first_amt, first_center = amounts_list[0]
                        if is_valid_amount(first_amt) and not is_in_description_range(first_center):
                            col_name = assign_word_to_column(first_center - 50, first_center + 50, columns_config)
                            if col_name and col_name in ('cargos', 'abonos', 'saldo'):
                                if not row_data.get(col_name):
                                    row_data[col_name] = first_amt
                
                # Limpiar valores inválidos de las columnas de montos
                # Si se asignó texto no numérico, eliminarlo
                for col in ['cargos', 'abonos', 'saldo']:
                    if col in row_data and row_data[col]:
                        if not is_valid_amount(row_data[col]):
                            # Si no es un monto válido, limpiar la columna
                            row_data[col] = ''
                
                # Limpiar la fecha antes de validar (por si acaso tiene puntos o caracteres extra)
                if 'fecha' in row_data and row_data['fecha']:
                    fecha_original = row_data['fecha']
                    fecha_clean = row_data['fecha'].strip().rstrip('.,;:')
                    if fecha_clean != fecha_original:
                        row_data['fecha'] = fecha_clean
                
                line_original = ' '.join([w.get('text', '') for w in row_words])
                # Si la descripción está vacía, rellenar con el texto residual de la línea original (HSBC OCR)
                if not row_data.get('descripcion') and line_original:
                    residual = line_original
                    # Quitar valores asignados: montos con espacios flexibles tras comas; fecha como palabra
                    for key in ('cargos', 'abonos', 'saldo'):
                        val = row_data.get(key)
                        if val and isinstance(val, str) and re.search(r'\d', val):
                            pat = re.escape(val).replace(',', r',\s*')
                            residual = re.sub(pat, '', residual, count=1)
                    # Fecha solo como token separado (no quitar "09" de "08045209")
                    fecha_val = row_data.get('fecha')
                    if fecha_val and isinstance(fecha_val, str):
                        residual = re.sub(r'(?<=\s)' + re.escape(fecha_val) + r'(?=\s|$)', ' ', residual)
                        residual = re.sub(r'^' + re.escape(fecha_val) + r'(?=\s|$)', '', residual)
                    row_data['descripcion'] = re.sub(r'\s+', ' ', residual).strip()
                # Si la descripción sigue vacía (sin residual o residual vacío), usar "."
                if not row_data.get('descripcion'):
                    row_data['descripcion'] = '.'
                
                # Validar si es una transacción válida
                is_valid = is_transaction_row(row_data, 'HSBC', debug_only_if_contains_iva=False)
                
                page_num_ocr = row_words[0].get('page', 0) if row_words else 0
                if is_valid:
                    # 🔍 HSBC DEBUG: Guardar últimas 2 filas válidas antes de movements_end
                    row_debug_info = {
                        'line_original': line_original,
                        'row_data': row_data.copy(),
                        'row_idx': row_idx,
                        'page_num': page_num_ocr
                    }
                    last_two_valid_rows.append(row_debug_info)
                    if len(last_two_valid_rows) > 2:
                        last_two_valid_rows.pop(0)  # Mantener solo las últimas 2
                    
                    _debug_mov_line(line_original, row_data, "ADDED")
                    movement_rows.append(row_data)
                else:
                    _debug_mov_line(line_original, row_data, "SKIPPED (HSBC OCR invalid)")
        
        df_mov = pd.DataFrame(movement_rows) if movement_rows else pd.DataFrame(columns=['fecha', 'descripcion', 'cargos', 'abonos', 'saldo'])
        
        # Debug: write movements debug file for HSBC OCR path (same format as coordinate path)
        if debug_path is not None and debug_movements_lines is not None and len(debug_movements_lines) > 0:
            with open(debug_path, 'w', encoding='utf-8') as f:
                for rec in debug_movements_lines:
                    f.write("ORIGINAL: " + (rec.get('original') or '') + "\n")
                    f.write("EXCEL: " + (rec.get('excel') or '') + "\n")
                    f.write("DISPOSITION: " + (rec.get('disposition') or '') + "\n")
                    f.write("\n")
            print(f"Debug: movements debug written to -> {debug_path}" + "\n", flush=True)
        
        # Extraer resumen desde texto OCR de la página 1
        pdf_summary = extract_hsbc_summary_from_ocr_text(extracted_data)
        
    else:
        # Flujo normal: procesar con coordenadas o texto plano
        # regex to detect decimal-like amounts (used to strip amounts from descriptions and detect amounts)
        dec_amount_re = re.compile(r"\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")
        # Pattern to detect end of movements table for specific banks
        # First, try to use movements_end from BANK_CONFIGS (for all banks)
        movement_end_pattern = None
        movement_end_string = bank_config.get('movements_end_new_format') if (bank_config['name'] == 'Banamex' and banamex_new_format) else bank_config.get('movements_end')
        if movement_end_string:
            # Create pattern from movements_end string (escape special chars)
            # For Santander, Banregio, Hey, and Clara, we need special handling for patterns with numbers
            if bank_config['name'] == 'Santander' or bank_config['name'] == 'Banregio' or bank_config['name'] == 'Hey' or bank_config['name'] == 'INTERCAM':
                # Santander/Banregio/Hey/INTERCAM: movements_end is "TOTAL"/"Total", match followed by 3 numeric amounts
                # Example: "TOTAL 821,646.20 820,238.73 1,417.18" or "Total 45,998.00 49,675.60 4,580.78"
                movement_end_pattern = re.compile(re.escape(movement_end_string) + r'\s+[\d,\.]+\s+[\d,\.]+\s+[\d,\.]+', re.I)
            elif bank_config['name'] == 'Clara':
                # Clara: movements_end is "Total MXN", followed by two amounts
                movement_end_pattern = re.compile(re.escape(movement_end_string) + r'\s+[\d,\.]+\s+[\d,\.\-]+', re.I)
            elif bank_config['name'] == 'Mercury':
                # Mercury: "Total" + 1 numeric value with 2 decimals (in Saldo column)
                movement_end_pattern = re.compile(
                    re.escape(movement_end_string) + r'\s+\$?[\d,\.\-]+\.\d{2}', re.I
                )
            elif bank_config['name'] == 'Banamex' and banamex_new_format:
                # Banamex new format: "Total +" or "Total cargos +" marks end
                movement_end_pattern = re.compile(r'Total\s+(?:cargos\s+)?\+', re.I)
            else:
                # For other banks, use simple string matching
                movement_end_pattern = re.compile(re.escape(movement_end_string), re.I)
        
        # For Banregio, also match "3 numeric amounts followed by Total" (variation seen in some PDFs)
        banregio_end_pattern_three_then_total = None
        if bank_config['name'] == 'Banregio' and movement_end_string:
            banregio_end_pattern_three_then_total = re.compile(
                r'[\d,\.]+\s+[\d,\.]+\s+[\d,\.]+\s+' + re.escape(movement_end_string), re.I
            )
        # For Banorte, create secondary pattern for alternative movements_end (from BANK_CONFIGS)
        banorte_secondary_end_pattern = None
        if bank_config['name'] == 'Banorte':
            banorte_secondary_end_string = bank_config.get('movements_end_secondary')
            if banorte_secondary_end_string:
                banorte_secondary_end_pattern = re.compile(re.escape(banorte_secondary_end_string), re.I)
        # For HSBC, create secondary pattern for "Información SPEI" (alternative to "Información CoDi")
        hsbc_secondary_end_pattern = None
        if bank_config['name'] == 'HSBC':
            hsbc_secondary_end_string = bank_config.get('movements_end_secondary')
            if hsbc_secondary_end_string:
                hsbc_secondary_end_pattern = re.compile(re.escape(hsbc_secondary_end_string), re.I)
        # For Banamex, create secondary pattern for "SALDO MINIMO REQUERIDO" (alternative to "SALDO PROMEDIO MINIMO REQUERIDO")
        banamex_secondary_end_pattern = None
        if bank_config['name'] == 'Banamex':
            banamex_secondary_end_string = bank_config.get('movements_end_secondary')
            if banamex_secondary_end_string:
                banamex_secondary_end_pattern = re.compile(re.escape(banamex_secondary_end_string), re.I)
        
        extraction_stopped = False
        # Debug: number each line in movements section (from movements_start until movements_end)
        mov_section_line_num = 0
        # For Banregio, initialize flag to track when we're in the commission zone
        in_comision_zone = False
        # For BanBajío, track detected rows for debugging
        banbajio_detected_rows = 0
        # For BBVA, track when we're in the movements section
        in_bbva_movements_section = False
        # 🔍 HSBC DEBUG: Buffer para guardar últimas 2 filas válidas antes de movements_end
        last_two_valid_rows = []
        total_pages = len(extracted_data)
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
            # Inbursa only: skip next row when it was consumed as fecha day (e.g. "01") for previous line
            inbursa_skip_next_row = False
            for row_idx, row_words in enumerate(word_rows):
                if not row_words or extraction_stopped:
                    continue
                if bank_config['name'] == 'Inbursa' and inbursa_skip_next_row:
                    inbursa_skip_next_row = False
                    continue
                # HSBC DEBUG: track if this row will be added to Excel
                hsbc_added = False
                row_was_added = False  # for 1,963.84 debug: set True before any movement_rows.append
                all_row_text = ' '.join([w.get('text', '') for w in row_words])
                all_row_text_orig = all_row_text  # for debug output
                _disp = "SKIPPED"  # for debug: ADDED | MERGED | SKIPPED | END (...)
                all_row_text_hsbc = all_row_text
                # Skip movements_start pattern if it appears during coordinate-based extraction
                # Headers will be automatically rejected by date validation
                if movement_start_pattern:
                    if movement_start_pattern.search(all_row_text):
                        # For BBVA, activate movements section when start pattern is found
                        if bank_config['name'] == 'BBVA' and not in_bbva_movements_section:
                            in_bbva_movements_section = True
                        mov_section_line_num += 1
                        _disp = "SKIPPED (movements_start header)"
                        _debug_mov_line(all_row_text_orig, None, _disp)
                        continue  # Skip the movements_start line
                
                mov_section_line_num += 1
                
                # Banamex new format only: capture Total Abonos/Cargos from summary lines for Valor en PDF validation
                # "Pagos y abonos" (-) $438.55 -> total_abonos; "Cargos regulares (no a meses)" (+) $1,127.00 -> total_cargos
                # OCR may output "(no meses)" without "a"; full multi-row scan is done later over all pages
                if bank_config['name'] == 'Banamex' and banamex_new_format and all_row_text_orig:
                    _amt_in_line = re.search(r'\$\s*([\d,]+\.\d{2})|(?<!\d)(\d{1,3}(?:,\d{3})*\.\d{2})(?=\s|$|[^\d])', all_row_text_orig)
                    if _amt_in_line:
                        _val = _amt_in_line.group(1) or _amt_in_line.group(2)
                        try:
                            _num = normalize_amount_str(_val)
                        except Exception:
                            _num = None
                        if _num is not None:
                            _norm = ' '.join(all_row_text_orig.split())
                            if re.search(r'Pagos\s+(?:\w+\s+)*y?\s+(?:\w+\s+)*abonos', _norm, re.I) and banamex_new_fmt_totals.get('total_abonos') is None:
                                banamex_new_fmt_totals['total_abonos'] = _num
                            if re.search(r'Cargos\s+regulares\s*\(?\s*no\s+a?\s*meses?', _norm, re.I) and banamex_new_fmt_totals.get('total_cargos') is None:
                                banamex_new_fmt_totals['total_cargos'] = _num
                
                # For Banregio, skip rows that start with "del 01 al" (irrelevant information)
                if bank_config['name'] == 'Banregio':
                    if re.search(r'^del\s+01\s+al', all_row_text, re.I):
                        _disp = "SKIPPED (Banregio del 01 al)"
                        _debug_mov_line(all_row_text_orig, None, _disp)
                        continue  # Skip irrelevant information rows

                # Check for end pattern (for Banamex, Santander, Banregio, Scotiabank, Konfio, Clara, etc.)
                if movement_end_pattern:
                    all_text = ' '.join([w.get('text', '') for w in row_words])
                    match_found = False
                    
                    # For Scotiabank, try flexible pattern matching (movements_end is "LAS TASAS DE INTERES")
                    if bank_config['name'] == 'Scotiabank':
                        # First check if movements_end string is present (flexible matching)
                        if movement_end_string and movement_end_string.lower() in all_text.lower():
                            match_found = True
                        else:
                            # Try patterns from BANK_CONFIGS if available (fallback)
                            end_patterns = bank_config.get('movements_end_patterns', [])
                            if end_patterns:
                                for pattern_str in end_patterns:
                                    pattern = re.compile(pattern_str, re.I)
                                    if pattern.search(all_text):
                                        match_found = True
                                        break
                    elif bank_config['name'] == 'Clara':
                        # For Clara, check for "Total MXN" (movements_end string) or "Total MXN X MXN Y" pattern
                        # First check if movements_end string is present
                        if movement_end_string and movement_end_string.lower() in all_text.lower():
                            match_found = True
                        else:
                            # Try patterns from BANK_CONFIGS if available (fallback)
                            end_patterns = bank_config.get('movements_end_patterns', [])
                            if end_patterns:
                                for pattern_str in end_patterns:
                                    pattern = re.compile(pattern_str, re.I)
                                    if pattern.search(all_text):
                                        match_found = True
                                        break
                    elif bank_config['name'] == 'Santander' or bank_config['name'] == 'Banregio' or bank_config['name'] == 'Hey':
                        # For Santander/Banregio/Hey, use the pattern with 3 numeric amounts (already created in movement_end_pattern)
                        if movement_end_pattern.search(all_text):
                            # Santander: valid movements_end is Total + 3 numeric values only; line must not contain % (resumen block has %)
                            if bank_config['name'] == 'Santander':
                                if '%' not in all_text:
                                    match_found = True
                            else:
                                match_found = True
                        # Banregio: also match "3 amounts + Total" (e.g. "11,973.34 12,973.34 1,000.00 Total")
                        if bank_config['name'] == 'Banregio' and banregio_end_pattern_three_then_total and banregio_end_pattern_three_then_total.search(all_text):
                            match_found = True
                    elif bank_config['name'] == 'Mercury':
                        # Mercury: "Total" + 1 numeric with 2 decimals (Saldo); or row with Total, no fecha, no descripcion, no Cargos/Abonos
                        if movement_end_pattern.search(all_text):
                            match_found = True
                        # Also check: row with "Total", no fecha, no descripcion, no Cargos/Abonos, but has Saldo value
                        if not match_found and movement_end_string and movement_end_string.upper() in all_text.upper() and columns_config:
                            has_fecha = False
                            has_descripcion = False
                            has_cargos = False
                            has_abonos = False
                            has_saldo = False
                            for w in row_words:
                                text = (w.get('text') or '').strip()
                                x0, x1 = w.get('x0', 0), w.get('x1', 0)
                                col = assign_word_to_column(x0, x1, columns_config)
                                if col == 'fecha' and date_pattern.search(text):
                                    has_fecha = True
                                elif col == 'descripcion' and text:
                                    has_descripcion = True
                                elif col == 'cargos' and DEC_AMOUNT_RE.search(text):
                                    has_cargos = True
                                elif col == 'abonos' and DEC_AMOUNT_RE.search(text):
                                    has_abonos = True
                                elif col == 'saldo' and DEC_AMOUNT_RE.search(text):
                                    has_saldo = True
                            # movements_end: Total + no fecha + no descripcion + no Cargos/Abonos + has Saldo
                            if not has_fecha and not has_descripcion and not has_cargos and not has_abonos and has_saldo:
                                match_found = True
                    elif bank_config['name'] == 'Banorte':
                        # For Banorte, try primary pattern first ("INVERSION ENLACE NEGOCIOS")
                        if movement_end_pattern and movement_end_pattern.search(all_text):
                            match_found = True
                        # If primary pattern not found, try secondary pattern ("CARGOS OBJETADOS EN EL PERÍODO")
                        elif banorte_secondary_end_pattern and banorte_secondary_end_pattern.search(all_text):
                            match_found = True
                    elif bank_config['name'] == 'HSBC':
                        # For HSBC, match "Información CoDi" (primary) or "Información SPEI" (secondary)
                        if movement_end_pattern and movement_end_pattern.search(all_text):
                            match_found = True
                        elif hsbc_secondary_end_pattern and hsbc_secondary_end_pattern.search(all_text):
                            match_found = True
                    elif bank_config['name'] == 'Inbursa':
                        # Inbursa: exact "Si desea recibir pagos" or flexible (phrase split/OCR) so we detect end on all PDFs
                        if movement_end_pattern and movement_end_pattern.search(all_text):
                            match_found = True
                        else:
                            # Flexible: row contains key phrase parts (handles line breaks / OCR)
                            all_text_lower = all_text.lower()
                            if 'si desea' in all_text_lower and 'recibir pagos' in all_text_lower:
                                match_found = True
                            elif 'si desea recibir' in all_text_lower or 'desea recibir pagos' in all_text_lower:
                                match_found = True
                    elif bank_config['name'] == 'Banamex':
                        # For Banamex, try primary ("SALDO PROMEDIO MINIMO REQUERIDO") then secondary ("SALDO MINIMO REQUERIDO")
                        if movement_end_pattern and movement_end_pattern.search(all_text):
                            match_found = True
                        elif banamex_secondary_end_pattern and banamex_secondary_end_pattern.search(all_text):
                            match_found = True
                    else:
                        # For other banks (Santander, Konfio, Banbajío, etc.), use the standard pattern
                        if movement_end_pattern.search(all_text):
                            match_found = True
                            #if bank_config['name'] == 'Banbajío':
                                #print(f"🛑 BanBajío: Fin de extracción detectado en página {page_num}, fila {row_idx+1}: {all_row_text[:100]}")
                    
                    if match_found:
                        # For BBVA, mark that we've left the movements section
                        if bank_config['name'] == 'BBVA' and in_bbva_movements_section:
                            in_bbva_movements_section = False
                        # Banamex new format: validation totals from "Pagos y abonos" / "Cargos regulares (no a meses)" only; treat movements_end like other banks
                        _disp = "END (movements_end matched, not added)"
                        _debug_mov_line(all_row_text_orig, None, _disp)
                        extraction_stopped = True
                        break

                # Extract structured row using coordinates
                # Pass bank_name and date_pattern to enable date/description separation
                # Santander: sanitize duplicated-character lines before extraction
                row_words_for_extract = _santander_sanitize_row_words_if_duplicated(row_words) if bank_config['name'] == 'Santander' else row_words
                row_data = extract_movement_row(row_words_for_extract, columns_config, bank_config['name'], date_pattern)
                
                # Banamex: override from line text for new format only (fecha, descripcion, cargo/abono from +/- in line or prev/next)
                # Run when header set banamex_new_format, OR when line has DD-mon-YYYY with lowercase month (e.g. 13-oct-2025); classic uses "30 ENE" (uppercase)
                if bank_config['name'] == 'Banamex' and all_row_text_orig:
                    # DD-mon-YYYY with optional space between day and month (OCR may output "13 oct-2025")
                    _bnf_date_re = re.compile(r'\b(0?[1-9]|[12][0-9]|3[01])[- ]([a-z]{3})[- ]?(\d{4})\b', re.I)
                    _bnf_date_re_strict = re.compile(r'\b(0?[1-9]|[12][0-9]|3[01])-[a-z]{3}-\d{4}\b', re.I)
                    # Amount: $ optional space digits/comma .XX (or standalone NNN.NN / N,NNN.NN not part of year)
                    _bnf_amt_re = re.compile(r'\$\s*[\d,]+\.\d{2}')
                    _bnf_amt_re_alt = re.compile(r'(?<!\d)(\d{1,3}(?:,\d{3})*\.\d{2})(?=\s|$|[^\d])')
                    dates = [f"{m.group(1)}-{m.group(2)}-{m.group(3)}" for m in _bnf_date_re.finditer(all_row_text_orig)]
                    if not dates:
                        dates = _bnf_date_re_strict.findall(all_row_text_orig)
                    amt_m = _bnf_amt_re.search(all_row_text_orig)
                    if not amt_m:
                        amt_m = _bnf_amt_re_alt.search(all_row_text_orig)
                    # Override only for new format: header detected OR line has DD-mon-YYYY with lowercase month (classic uses "30 ENE" uppercase)
                    _is_new_fmt_line = banamex_new_format or (bool(dates) and bool(re.search(r'-[a-z]{3}-', dates[0])))
                    if dates and amt_m and _is_new_fmt_line:
                        prev_line_text = ' '.join([w.get('text', '') for w in word_rows[row_idx - 1]]) if row_idx > 0 else ''
                        row_data['fecha'] = dates[-1]  # last date (e.g. 19-oct-2025); keep full DD-mon-YYYY
                        amount_val = amt_m.group().strip()
                        if not amount_val.startswith('$'):
                            amount_val = '$' + amount_val
                        desc_line = _bnf_date_re.sub(' ', all_row_text_orig)
                        desc_line = _bnf_amt_re.sub(' ', desc_line, count=1)
                        if _bnf_amt_re.search(desc_line) is None and _bnf_amt_re_alt.search(all_row_text_orig):
                            desc_line = _bnf_amt_re_alt.sub(' ', desc_line, count=1)
                        row_data['descripcion'] = ' '.join(desc_line.split()).strip().lstrip('|').strip()
                        # + on current or previous line = cargo; else abono (next line's + applies to next movement)
                        is_cargo = '+' in prev_line_text or '+' in all_row_text_orig
                        row_data['cargos'] = amount_val if is_cargo else ''
                        row_data['abonos'] = amount_val if not is_cargo else ''
                        row_data['saldo'] = ''
                        if 'monto' in row_data:
                            del row_data['monto']
                
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
                    elif bank_config['name'] == 'INTERCAM':
                        # INTERCAM: fecha column is day only (1-31)
                        has_date = bool(date_pattern.match(fecha_val.strip()))
                    elif bank_config['name'] == 'Clara':
                        # For Clara, check if fecha contains a valid date pattern (e.g., "01 ENE" or "01 E N E")
                        clara_date_pattern = re.compile(r'\d{1,2}\s+[A-Z]{3}|\d{1,2}\s+[A-Z]\s+[A-Z]\s+[A-Z]', re.I)
                        has_date = bool(clara_date_pattern.search(fecha_val))
                    elif bank_config['name'] == 'Banamex':
                        # Classic: day + 3-letter month (e.g. "30 ENE"). New format: DD-mon-YYYY (e.g. 13-oct-2025)
                        banamex_dd_mon_yyyy = re.compile(r'\b(0?[1-9]|[12][0-9]|3[01])-[a-z]{3}-\d{4}\b', re.I)
                        if banamex_new_format or banamex_dd_mon_yyyy.search(fecha_val):
                            has_date = bool(banamex_dd_mon_yyyy.search(fecha_val))
                        else:
                            # Classic Banamex only
                            banamex_date_pattern = re.compile(r'\b(0[1-9]|[12][0-9]|3[01])\s+[A-Z]{3}\b', re.I)
                            has_date = bool(banamex_date_pattern.search(fecha_val))
                    else:
                        has_date = bool(date_pattern.search(fecha_val))
                        # For Inbursa, use strict date pattern: "MES. DD" or "MES DD" at start of line
                        # Valid months: ENE, FEB, MAR, ABR, MAY, JUN, JUL, AGO, SEP, OCT, NOV, DIC
                        # This prevents false positives like "01 BAL" or "IVA 16"
                        if bank_config['name'] == 'Inbursa':
                            inbursa_date_pattern = re.compile(r'^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\.?\s+(0[1-9]|[12][0-9]|3[01])\b', re.I)
                            inbursa_month_only_re = re.compile(r'^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\.?$', re.I)
                            # Check if fecha_val matches full pattern (MES. DD) or month-only (date split across 2 lines)
                            has_date = bool(inbursa_date_pattern.match(fecha_val.strip()))
                            if not has_date:
                                all_row_text = ' '.join([w.get('text', '') for w in row_words])
                                if inbursa_date_pattern.search(all_row_text.strip()):
                                    has_date = True
                                    date_match = inbursa_date_pattern.search(all_row_text.strip())
                                    if date_match:
                                        row_data['fecha'] = date_match.group()
                                else:
                                    # Month-only: if next line has only numeric 01-31 (or starts with it), use as day to complete fecha
                                    month_match = inbursa_month_only_re.match(fecha_val.strip())
                                    if month_match:
                                        month_str = month_match.group(1)
                                        next_day_val = None
                                        next_row_rest = None  # rest of next line after day (if any)
                                        if row_idx + 1 < len(word_rows):
                                            next_row_text = ' '.join([w.get('text', '') for w in word_rows[row_idx + 1]]).strip()
                                            # Match exactly "01"-"31" or "1"-"9", or line starting with day then space and more
                                            day_only_match = re.match(r'^(0?[1-9]|[12][0-9]|3[01])$', next_row_text)
                                            if day_only_match:
                                                next_day_val = day_only_match.group(0).zfill(2)  # normalize to 01, 02, ...
                                            else:
                                                day_start_match = re.match(r'^(0?[1-9]|[12][0-9]|3[01])\s+(.*)$', next_row_text)
                                                if day_start_match:
                                                    next_day_val = day_start_match.group(1).zfill(2)
                                                    next_row_rest = day_start_match.group(2).strip()
                                                else:
                                                    # Day at end of next line (e.g. "something 30" -> use 30 as day)
                                                    day_end_match = re.search(r'\b(0?[1-9]|[12][0-9]|3[01])$', next_row_text)
                                                    if day_end_match:
                                                        next_day_val = day_end_match.group(1).zfill(2)
                                                        next_row_rest = next_row_text[:day_end_match.start()].strip()
                                        if next_day_val is not None:
                                            has_date = True
                                            row_data['fecha'] = month_str.capitalize() + '. ' + next_day_val
                                            inbursa_skip_next_row = True
                                            if next_row_rest:
                                                # Merge next line into current row: get structured data (descripcion, cargos, abonos, saldo) from next row's words (without leading day)
                                                next_row_words = word_rows[row_idx + 1]
                                                if len(next_row_words) > 1 and re.match(r'^(0?[1-9]|[12][0-9]|3[01])$', (next_row_words[0].get('text') or '').strip()):
                                                    words_without_day = next_row_words[1:]
                                                    if words_without_day:
                                                        next_row_data = extract_movement_row(words_without_day, columns_config, bank_config['name'], date_pattern)
                                                        for k in ('descripcion', 'cargos', 'abonos', 'saldo'):
                                                            if k in next_row_data and next_row_data.get(k) and not row_data.get(k):
                                                                row_data[k] = next_row_data[k]
                                                if not (row_data.get('descripcion') or row_data.get('cargos') or row_data.get('abonos') or row_data.get('saldo')):
                                                    existing_desc = str(row_data.get('descripcion') or '').strip()
                                                    row_data['descripcion'] = (existing_desc + ' ' + next_row_rest).strip() if existing_desc else next_row_rest
                                                    amt_in_rest = DEC_AMOUNT_RE.findall(next_row_rest)
                                                    if amt_in_rest:
                                                        row_data['saldo'] = amt_in_rest[-1]
                                        else:
                                            has_date = True
                                            row_data['fecha'] = month_str.capitalize() + '.'
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
                        # For Banorte, if fecha is still empty, extract date from full row text (DD/MM/YYYY or DIA-MES-AÑO)
                        if bank_config['name'] == 'Banorte' and not has_date:
                            all_row_text_banorte = ' '.join([w.get('text', '') for w in row_words])
                            date_match_banorte = date_pattern.search(all_row_text_banorte)
                            if date_match_banorte:
                                row_data['fecha'] = date_match_banorte.group(0)
                                has_date = True
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
                    # For BBVA, mark that we're in the movements section when we find the first date
                    # Only if movements_start was not already found (fallback mechanism)
                    if bank_config['name'] == 'BBVA' and not in_bbva_movements_section:
                        in_bbva_movements_section = True
                    
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
                            _disp = "SKIPPED (Konfio date format)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
                            continue
                        
                        # Check if this is a sub-row pattern (like "SBO 020902DX1 - 02 abr 2023 4316 FÍSICA")
                        # These should be rejected as they are continuation lines, not main movement rows
                        all_row_text = ' '.join([w.get('text', '') for w in row_words])
                        konfio_sub_row_pattern = re.compile(r'^[A-Z]{2,4}\s*[A-Z0-9]{8,15}\s*-\s*\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4}', re.I)
                        # If the row starts with this pattern, it's a sub-row and should be skipped
                        if konfio_sub_row_pattern.match(all_row_text.strip()):
                            # This is a sub-row, skip it (it will be handled as continuation if needed)
                            _disp = "SKIPPED (Konfio sub-row)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
                            continue
                        # Add row if it has date, even if description/amounts are empty (they'll be added in continuation rows)
                        row_was_added = True
                        _disp = "ADDED"
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
                            _disp = "SKIPPED (Banamex footer)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
                            continue
                        
                        #row_data['page'] = page_num
                        _disp = "ADDED"
                        movement_rows.append(row_data)
                        _debug_mov_line(all_row_text_orig, row_data, _disp)
                    elif has_description_or_amounts:
                        # For HSBC, validate row before adding
                        if bank_config['name'] == 'HSBC':
                            is_valid = is_transaction_row(row_data, 'HSBC', debug_only_if_contains_iva=False)
                            if is_valid:
                                # 🔍 HSBC DEBUG: Guardar últimas 2 filas válidas antes de movements_end (flujo normal, no OCR)
                                all_row_text = ' '.join([w.get('text', '') for w in row_words])
                                row_debug_info = {
                                    'line_original': all_row_text,
                                    'row_data': row_data.copy(),
                                    'row_idx': row_idx,
                                    'page_num': page_num
                                }
                                last_two_valid_rows.append(row_debug_info)
                                if len(last_two_valid_rows) > 2:
                                    last_two_valid_rows.pop(0)  # Mantener solo las últimas 2
                        
                        # For Banregio, only include rows where description starts with "TRA" or "DOC"
                        if bank_config['name'] == 'Banregio':
                            desc_val_check = str(row_data.get('descripcion') or '').strip()
                            if not (desc_val_check.startswith('TRA') or desc_val_check.startswith('DOC') or desc_val_check.startswith('INT') or desc_val_check.startswith('EFE')):
                                # Skip this row - doesn't start with TRA or DOC or INT or EFE
                                _disp = "SKIPPED (Banregio not TRA/DOC/INT/EFE)"
                                _debug_mov_line(all_row_text_orig, row_data, _disp)
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
                                _disp = "SKIPPED (Banamex footer)"
                                _debug_mov_line(all_row_text_orig, row_data, _disp)
                                continue
                        
                        # For INTERCAM, skip CFDI disclaimer and "Hoja de N" (page) metadata lines
                        if bank_config['name'] == 'INTERCAM':
                            all_row_text_intercam = ' '.join([w.get('text', '') for w in row_words])
                            if re.search(r'DOCUMENTO ES UNA REPRESENTACIÓN IMPRESA DE UN CFDI|REPRESENTACIÓN IMPRESA.*CFDI|Hoja de\s*\d+|Número Cliente\s*R\.F\.C\.\s*Sucursal', all_row_text_intercam, re.I):
                                _disp = "SKIPPED (INTERCAM CFDI/page)"
                                _debug_mov_line(all_row_text_orig, row_data, _disp)
                                continue
                        
                        #row_data['page'] = page_num
                        if bank_config['name'] == 'HSBC':
                            hsbc_added = True
                        row_was_added = True
                        _disp = "ADDED"
                        movement_rows.append(row_data)
                elif has_valid_data and bank_config['name'] == 'Banbajío':
                    # For BanBajío, also add rows without date if they contain "SALDO INICIAL"
                    desc_val = str(row_data.get('descripcion') or '').strip()
                    if re.search(r'SALDO\s+INICIAL', desc_val, re.I):
                        has_amounts = len(row_data.get('_amounts', [])) > 0
                        has_cargos_abonos = bool(row_data.get('cargos') or row_data.get('abonos') or row_data.get('saldo'))
                        if has_amounts or has_cargos_abonos:
                            #row_data['page'] = page_num
                            row_was_added = True
                            _disp = "ADDED"
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
                        _disp = "END (Banregio commission zone)"
                        _debug_mov_line(all_row_text_orig, row_data, _disp)
                        extraction_stopped = True
                        break  # Stop extraction - no more monthly commissions
                
                # If row has valid data but no date - treat as continuation or standalone row
                if not has_date and has_valid_data:
                    # Do not merge the movements_end line into the last movement: if this row contains
                    # the end marker (e.g. "Total", "SALDO TOTAL"), stop extraction without appending.
                    if movement_end_string and movement_end_string.strip():
                        end_upper = movement_end_string.upper().strip()
                        if end_upper in all_row_text.upper():
                            # Santander: TOTAL + numeric with % is resumen block, not valid movements_end
                            if bank_config['name'] == 'Santander' and '%' in all_row_text:
                                pass  # do not treat as end
                            else:
                                _disp = "END (movements_end string in line)"
                                _debug_mov_line(all_row_text_orig, row_data, _disp)
                                extraction_stopped = True
                                break
                    if movement_end_pattern and movement_end_pattern.search(all_row_text):
                        # Santander: Total + 3 numeric with % is resumen block, not valid movements_end
                        if bank_config['name'] == 'Santander' and '%' in all_row_text:
                            pass  # do not treat as end
                        else:
                            _disp = "END (movements_end pattern)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
                            extraction_stopped = True
                            break
                    # Row has valid data but no date - treat as continuation or standalone row
                    # For Santander: do not merge rows that contain "TOTAL" (no fecha) into previous movement
                    if bank_config['name'] == 'Santander':
                        all_row_text_sant = ' '.join([w.get('text', '') for w in row_words])
                        if 'TOTAL' in all_row_text_sant.upper():
                            _disp = "SKIPPED (Santander TOTAL)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
                            continue
                    
                    # For Mercury: rows without fecha share the last known fecha; add as new movement row (do not merge)
                    if bank_config['name'] == 'Mercury' and movement_rows:
                        prev = movement_rows[-1]
                        last_fecha = (prev.get('fecha') or '').strip()
                        if last_fecha:
                            new_row = {
                                'fecha': last_fecha,
                                'descripcion': str(row_data.get('descripcion') or '').strip(),
                                'cargos': str(row_data.get('cargos') or '').strip(),
                                'abonos': str(row_data.get('abonos') or '').strip(),
                                'saldo': str(row_data.get('saldo') or '').strip(),
                            }
                            if new_row.get('descripcion') or new_row.get('cargos') or new_row.get('abonos') or new_row.get('saldo'):
                                if '_amounts' in row_data:
                                    new_row['_amounts'] = row_data.get('_amounts', [])
                                row_was_added = True
                                _disp = "ADDED"
                                movement_rows.append(new_row)
                        _debug_mov_line(all_row_text_orig, row_data, _disp)
                        continue
                    
                    # Filter out footer information for Banamex (e.g., "000180.B61CHDA011.OD.0131.01")
                    if bank_config['name'] == 'Banamex':
                        desc_val_check = str(row_data.get('descripcion') or '').strip()
                        cargos_val_check = str(row_data.get('cargos') or '').strip()
                        abonos_val_check = str(row_data.get('abonos') or '').strip()
                        saldo_val_check = str(row_data.get('saldo') or '').strip()
                        
                        # For continuation rows, must have description AND it must not contain footer pattern
                        if not desc_val_check:
                            # No description - skip this continuation row
                            _disp = "SKIPPED (Banamex continuation no description)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
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
                            _disp = "SKIPPED (Banamex continuation footer)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
                            continue
                        
                        # Additional check: if saldo contains a value that looks like part of footer (e.g., "131.01")
                        # and there's no meaningful description, skip it
                        if saldo_val_check and re.match(r'^\d+\.\d{2}$', saldo_val_check) and len(desc_val_check) < 10:
                            # Check if this value appears near footer-related words
                            if any('OD' in w.get('text', '') or re.search(r'\d+\.\w+', w.get('text', ''), re.I) for w in row_words):
                                _disp = "SKIPPED (Banamex continuation footer 2)"
                                _debug_mov_line(all_row_text_orig, row_data, _disp)
                                continue
                    # For BanBajío, filter out informational rows like "1 DE ENERO AL 31 DE ENERO DE 2024 PERIODO:"
                    if bank_config['name'] == 'Banbajío':
                        all_row_text_check = ' '.join([w.get('text', '') for w in row_words])
                        all_row_text_check = all_row_text_check.strip()
                        # Check if row contains period information (e.g., "1 DE ENERO AL 31 DE ENERO DE 2024 PERIODO:")
                        if re.search(r'\bPERIODO\s*:?\s*$', all_row_text_check, re.I) or re.search(r'\d+\s+DE\s+[A-Z]+\s+AL\s+\d+\s+DE\s+[A-Z]+', all_row_text_check, re.I):
                            # Skip this row - it's period information, not a movement
                            _disp = "SKIPPED (Banbajío PERIODO)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
                            continue
                    
                    # For Banregio, if we're in commission zone, check for irrelevant rows and stop extraction
                    if bank_config['name'] == 'Banregio' and in_comision_zone:
                        # Collect all text from the row to check for irrelevant content
                        all_row_text_check = ' '.join([w.get('text', '') for w in row_words])
                        all_row_text_check = all_row_text_check.strip()
                        
                        # Check if row contains "Total" (summary line)
                        if re.search(r'\bTotal\b', all_row_text_check, re.I):
                            _disp = "END (Banregio commission Total)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
                            extraction_stopped = True
                            break  # Stop extraction - "Total" line detected
                        
                        # Check if text is too long (likely not part of a movement description)
                        if len(all_row_text_check) > 200:
                            _disp = "END (Banregio commission text too long)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
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
                            _disp = "END (Banregio commission irrelevant)"
                            _debug_mov_line(all_row_text_orig, row_data, _disp)
                            extraction_stopped = True
                            break  # Stop extraction - irrelevant information detected
                        
                        # If none of the above, stop extraction anyway (we're past commissions)
                        _disp = "END (Banregio commission past)"
                        _debug_mov_line(all_row_text_orig, row_data, _disp)
                        extraction_stopped = True
                        break  # Stop extraction - no more monthly commissions
                    
                    if movement_rows:
                        # For Banamex: additional validation before merging continuation row
                        if bank_config['name'] == 'Banamex':
                            desc_val_check = str(row_data.get('descripcion') or '').strip()
                            # Must have description to merge continuation row
                            if not desc_val_check:
                                _disp = "SKIPPED (Banamex merge no description)"
                                _debug_mov_line(all_row_text_orig, row_data, _disp)
                                continue
                            
                            # Check if description contains footer pattern
                            all_row_text = ' '.join([w.get('text', '') for w in row_words])
                            banamex_footer_pattern = re.compile(r'\d+\.\w+\.OD\.\d+\.\d+', re.I)
                            banamex_footer_partial = re.compile(r'\.OD\.\d+\.\d+', re.I)
                            has_od_pattern = any('OD' in w.get('text', '') and re.search(r'OD[\.\s]*\d+[\.\s]*\d+', w.get('text', ''), re.I) for w in row_words)
                            
                            if banamex_footer_pattern.search(desc_val_check) or banamex_footer_pattern.search(all_row_text) or banamex_footer_partial.search(desc_val_check) or has_od_pattern:
                                _disp = "SKIPPED (Banamex merge footer)"
                                _debug_mov_line(all_row_text_orig, row_data, _disp)
                                continue
                        
                        # Continuation row: append description-like text and amounts to previous movement
                        _disp = "MERGED"
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
                                # Mercury: cargos/abonos share range; assign only by sign (positive -> abonos, negative -> cargos)
                                if bank_config.get('name') == 'Mercury':
                                    amt_stripped = (amt_text or '').strip()
                                    is_neg = (
                                        amt_stripped.startswith('-') or amt_stripped.startswith('\u2013') or
                                        re.match(r'^[-\u2013]\s*\$', amt_stripped) or
                                        re.search(r'\$\s*[-\u2013]', amt_stripped) or
                                        (amt_stripped.startswith('(') and amt_stripped.endswith(')'))
                                    )
                                    target_col = 'cargos' if is_neg else 'abonos'
                                    if target_col in col_ranges:
                                        x0, x1 = col_ranges[target_col]
                                        if (x0 - tolerance) <= center <= (x1 + tolerance):
                                            existing = prev.get(target_col) or ''
                                            if not existing or amt_text not in existing:
                                                if existing:
                                                    prev[target_col] = (existing + ' ' + amt_text).strip()
                                                else:
                                                    prev[target_col] = amt_text
                                            assigned = True
                                if not assigned:
                                    for col in ('cargos', 'abonos', 'saldo'):
                                        if col in col_ranges:
                                            x0, x1 = col_ranges[col]
                                            if (x0 - tolerance) <= center <= (x1 + tolerance):
                                                # Only assign if the column is empty or if this is a better match
                                                existing = prev.get(col) or ''
                                                if not existing or amt_text not in existing:
                                                    # INTERCAM: do not concatenate amounts when prev already has a value
                                                    # (continuation row is often a separate movement that failed date validation)
                                                    if existing and bank_config.get('name') == 'INTERCAM':
                                                        pass
                                                    elif existing:
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
                                        # Mercury: if amount landed in cargos/abonos by proximity, assign only by sign
                                        if bank_config.get('name') == 'Mercury' and nearest in ('cargos', 'abonos'):
                                            amt_stripped = (amt_text or '').strip()
                                            is_neg = (
                                                amt_stripped.startswith('-') or amt_stripped.startswith('\u2013') or
                                                re.match(r'^[-\u2013]\s*\$', amt_stripped) or
                                                re.search(r'\$\s*[-\u2013]', amt_stripped) or
                                                (amt_stripped.startswith('(') and amt_stripped.endswith(')'))
                                            )
                                            nearest = 'cargos' if is_neg else 'abonos'
                                        if not descripcion_range or not (descripcion_range[0] <= center <= descripcion_range[1]):
                                            existing = prev.get(nearest) or ''
                                            if not existing or amt_text not in existing:
                                                if existing and bank_config.get('name') == 'INTERCAM':
                                                    pass
                                                elif existing:
                                                    prev[nearest] = (existing + ' ' + amt_text).strip()
                                                else:
                                                    prev[nearest] = amt_text
                        
                        # Mercury: after merging amounts into prev, ensure same value is not in both cargos and abonos
                        if bank_config.get('name') == 'Mercury':
                            c = (prev.get('cargos') or '').strip()
                            a = (prev.get('abonos') or '').strip()
                            if c and a:
                                try:
                                    num_c = normalize_amount_str(c)
                                    num_a = normalize_amount_str(a)
                                    if num_c is not None and num_a is not None and abs(num_c) == abs(num_a):
                                        is_neg = (c.startswith('-') or c.startswith('\u2013') or re.match(r'^[-\u2013]\s*\$', c))
                                        if is_neg:
                                            prev['abonos'] = ''
                                        else:
                                            prev['cargos'] = ''
                                except Exception:
                                    pass
                        
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
                        # Trim at movements_end so footer/summary line text is not appended to description
                        if cont_text and movement_end_string and movement_end_string.strip():
                            end_upper = movement_end_string.upper().strip()
                            idx = cont_text.upper().find(end_upper)
                            if idx != -1:
                                cont_text = cont_text[:idx].strip()
                        if movement_end_pattern and cont_text and movement_end_pattern.search(cont_text):
                            m = movement_end_pattern.search(cont_text)
                            cont_text = cont_text[:m.start()].strip()

                        if cont_text:
                            # append to previous 'descripcion' field
                            if prev.get('descripcion'):
                                prev['descripcion'] = (prev.get('descripcion') or '') + ' ' + cont_text
                            else:
                                prev['descripcion'] = cont_text
                            # Trim description at movements_end in case end marker leaked in
                            if movement_end_string and movement_end_string.strip():
                                d = prev['descripcion']
                                idx = d.upper().find(movement_end_string.upper().strip())
                                if idx != -1:
                                    prev['descripcion'] = d[:idx].strip()
                            if movement_end_pattern and prev.get('descripcion'):
                                m = movement_end_pattern.search(prev['descripcion'])
                                if m:
                                    prev['descripcion'] = prev['descripcion'][:m.start()].strip()
                        
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

                # Debug: log this row (generic for all banks; skipped/break paths already logged above)
                if debug_movements_lines is not None:
                    _debug_mov_line(all_row_text_orig, row_data, _disp)

        # Write debug file after coordinate-based extraction (all banks)
        if debug_path is not None and debug_movements_lines is not None and len(debug_movements_lines) > 0:
            with open(debug_path, 'w', encoding='utf-8') as f:
                for rec in debug_movements_lines:
                    f.write("ORIGINAL: " + (rec.get('original') or '') + "\n")
                    f.write("EXCEL: " + (rec.get('excel') or '') + "\n")
                    f.write("DISPOSITION: " + (rec.get('disposition') or '') + "\n")
                    f.write("\n")
            print(f"Debug: movements debug written to -> {debug_path}", flush=True)

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

        # Mercury: ensure no row has the same amount in both Cargos and Abonos (value must appear only in one by sign)
        # Mercury: remove negative sign from Cargos in Excel (show e.g. $1,300.00 instead of –$1,300.00)
        if bank_config.get('name') == 'Mercury' and movement_rows:
            for r in movement_rows:
                c = (r.get('cargos') or '').strip()
                a = (r.get('abonos') or '').strip()
                # Strip leading minus/en-dash from cargos so Excel shows positive amount
                if c:
                    c_clean = re.sub(r'^[-\u2013]\s*', '', c)
                    if c_clean != c:
                        r['cargos'] = c_clean
                    c = (r.get('cargos') or '').strip()
                if c and a:
                    try:
                        num_c = normalize_amount_str(c)
                        num_a = normalize_amount_str(a)
                        if num_c is not None and num_a is not None and abs(num_c) == abs(num_a):
                            is_neg = (c.startswith('-') or c.startswith('\u2013') or re.match(r'^[-\u2013]\s*\$', c))
                            if is_neg:
                                r['abonos'] = ''
                            else:
                                r['cargos'] = ''
                    except Exception:
                        pass
        # Trim any movement description that includes text after movements_end (footer/summary leakage)
        if movement_rows and (movement_end_string or movement_end_pattern):
            for r in movement_rows:
                desc = r.get('descripcion')
                if not desc:
                    continue
                if movement_end_string and movement_end_string.strip():
                    idx = desc.upper().find(movement_end_string.upper().strip())
                    if idx != -1:
                        r['descripcion'] = desc[:idx].strip()
                    desc = r.get('descripcion') or ''
                if movement_end_pattern and desc and movement_end_pattern.search(desc):
                    m = movement_end_pattern.search(desc)
                    r['descripcion'] = desc[:m.start()].strip()
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
    # Normalize 'fecha' column: keep only the first date in each cell (so "02/ENE 01/ENE" -> "02/ENE") before any later logic
    if 'fecha' in df_mov.columns:
        _first_date_re = re.compile(
            r'(?:0[1-9]|[12][0-9]|3[01])[/\-](?:0[1-9]|1[0-2]|[A-Za-z]{3})(?:[/\-]\d{2,4})?|'
            r'(?:0[1-9]|[12][0-9]|3[01])\s+[A-Za-z]{3}(?:\s+\d{2,4})?',
            re.I
        )
        def _first_date_only(txt):
            if not txt or not isinstance(txt, str) or str(txt).strip() in ('', 'nan', 'None'):
                return txt
            s = str(txt).strip()
            matches = _first_date_re.findall(s)
            if len(matches) >= 2:
                return matches[0].strip()
            return s
        df_mov['fecha'] = df_mov['fecha'].astype(str).apply(_first_date_only)
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
    elif bank_config['name'] == 'INTERCAM' and 'fecha' in df_mov.columns:
        # INTERCAM: fecha is day only (1-31), preserve as-is
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
    elif bank_config['name'] == 'Inbursa' and 'fecha' in df_mov.columns:
        # For Inbursa, preserve the date format "MES. DD" or "MES DD" (e.g., "ENE. 01" or "ABR 10") as-is
        df_mov['Fecha'] = df_mov['fecha'].astype(str)
        df_mov = df_mov.drop(columns=['fecha'])
    elif bank_config['name'] == 'Banorte' and 'fecha' in df_mov.columns:
        # For Banorte, preserve the date format DD/MM/YYYY or DIA-MES-AÑO as-is (don't use _extract_two_dates; it doesn't match DD/MM/YYYY)
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
        
        # For BBVA, create 'Fecha' from 'Fecha Liq'; when liq is empty (Fecha Liq is None), use Fecha Oper so we don't lose the date
        df_mov['Fecha'] = df_mov['Fecha Liq'].fillna(df_mov['Fecha Oper'])
        df_mov = df_mov.drop(columns=['Fecha Oper', 'Fecha Liq'])
        
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
            # Fallback: if _extract_two_dates returned (None, None), keep raw 'fecha' as Fecha Oper so we don't lose values like "02/ENE"
            if 'fecha' in df_mov.columns:
                def _valid_fecha(s):
                    if pd.isna(s):
                        return False
                    t = str(s).strip()
                    return t != '' and t.lower() != 'nan'
                mask = df_mov['Fecha Oper'].isna() & df_mov['fecha'].apply(_valid_fecha)
                if mask.any():
                    df_mov.loc[mask, 'Fecha Oper'] = df_mov.loc[mask, 'fecha'].astype(str).str.strip()

    # Remove original 'fecha' if present (only if not already removed)
    # Skip this for HSBC, Banregio, Base, Banbajío, Inbursa, Banorte, and INTERCAM - they already handled 'fecha' to 'Fecha' conversion
    if 'fecha' in df_mov.columns and bank_config['name'] != 'HSBC' and bank_config['name'] != 'Banregio' and bank_config['name'] != 'Base' and bank_config['name'] != 'Banbajío' and bank_config['name'] != 'Inbursa' and bank_config['name'] != 'Banorte' and bank_config['name'] != 'INTERCAM':
        df_mov = df_mov.drop(columns=['fecha'])
    
    # For non-BBVA banks, use only 'Fecha' column (based on Fecha Oper) and remove Fecha Liq
    # Skip this for Banregio, Base, BanBajío, Inbursa, Banorte, and HSBC (they already have 'Fecha' set above or from OCR)
    if bank_config['name'] != 'BBVA' and bank_config['name'] != 'Banregio' and bank_config['name'] != 'Base' and bank_config['name'] != 'Banbajío' and bank_config['name'] != 'Inbursa' and bank_config['name'] != 'Banorte' and bank_config['name'] != 'HSBC' and bank_config['name'] != 'INTERCAM':
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
        
        # For BBVA, create 'Saldo' from 'LIQUIDACIÓN' (first amount) and remove both saldo columns
        df_mov['Saldo'] = df_mov['LIQUIDACIÓN']
        df_mov = df_mov.drop(columns=['OPERACIÓN', 'LIQUIDACIÓN'])
        
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
        # For BBVA, OPERACIÓN and LIQUIDACIÓN are already converted to 'Saldo' above, so skip renaming
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
        # For BBVA: Fecha, Descripción, Abonos, Cargos, Saldo
        desired_order = ['Fecha', 'Descripción', 'Abonos', 'Cargos', 'Saldo']
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
        
        # Extract DIGITEM section using already-extracted data to avoid second PDF/OCR pass
        df_digitem = extract_digitem_section(pdf_path, columns_config, extracted_data=extracted_data)
        
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
    
    # Extract Santander METAS section (Mis Metas) into a separate tab (same structure as Movements)
    df_metas = None
    if bank_config['name'] == 'Santander':
        metas_start = bank_config.get('metas_start')
        metas_end = bank_config.get('metas_end')
        if metas_start and metas_end and columns_config:
            df_metas = extract_santander_metas_from_pdf(extracted_data, columns_config, metas_start, metas_end)
    
    # Extract summary from PDF and calculate totals for validation
    # IMPORTANT: Calculate totals AFTER removing DIGITEM rows and BEFORE adding the "Total" row
    #print("🔍 Extrayendo información de resumen del PDF para validación...")
    # Para HSBC con OCR, el resumen ya fue extraído desde el texto OCR arriba
    # Para otros casos, extraer desde PDF
    if not (is_hsbc and used_ocr):
        pdf_summary = extract_summary_from_pdf(pdf_path, movement_start_page=movement_start_page)
    # Banamex new format only: Valor en PDF from (1) "Total +" line → Total Cargos, (2) "Total" line (after "Total +") → Total Abonos.
    # Fallback: "Pagos y abonos" / "Cargos regulares (no a meses)" in RESUMEN block. Scan ALL PDF pages.
    if bank_config['name'] == 'Banamex':
        print(f"[Banamex new format Valor en PDF] banamex_new_format={banamex_new_format}, pages={len(extracted_data) if extracted_data else 0}", flush=True)
    if bank_config['name'] == 'Banamex' and banamex_new_format and extracted_data:
        need_abonos = banamex_new_fmt_totals.get('total_abonos') is None
        need_cargos = banamex_new_fmt_totals.get('total_cargos') is None
        # Always print so user sees Banamex new format Valor en PDF block ran (normal and --debug)
        print(f"[Banamex new format Valor en PDF] total_abonos={banamex_new_fmt_totals.get('total_abonos')}, total_cargos={banamex_new_fmt_totals.get('total_cargos')} | need_scan={need_abonos or need_cargos}, pages={len(extracted_data)}", flush=True)
        if need_abonos or need_cargos:
            re_amt = re.compile(r'\$\s*([\d,]+\.\d{2})|(?<!\d)(\d{1,3}(?:,\d{3})*\.\d{2})(?=\s|$|[^\d])')
            re_total_plus = re.compile(r'Total\s*\+', re.I)
            # Amount followed by "Total" (e.g. "$438.55 Total") for Total Abonos; avoids matching "Total de Movimientos" etc.
            re_amount_total = re.compile(r'[\d,]+\.\d{2}\s+Total(?:\s|$)', re.I)
            WINDOW_SIZE = 4
            debug_val = debug_mode
            # Always print in normal and --debug: confirm scan runs and what we need
            print(f"[Banamex new format Valor en PDF] Scanning {len(extracted_data)} pages for Total Cargos ('Total +') and Total Abonos ('Total')...", flush=True)
            if debug_val:
                print(f"[Banamex new format Valor en PDF] need_abonos={need_abonos}, need_cargos={need_cargos}", flush=True)
            for page_data in extracted_data:
                words = page_data.get('words', [])
                if not words:
                    if debug_val:
                        print(f"[Banamex new format Valor en PDF] Page {page_data.get('page')}: no words, skip", flush=True)
                    continue
                word_rows = group_words_by_row(words, y_tolerance=3)
                if debug_val:
                    print(f"[Banamex new format Valor en PDF] Page {page_data.get('page')}: {len(word_rows)} rows", flush=True)
                for r in range(len(word_rows)):
                    if not need_abonos and not need_cargos:
                        break
                    window_rows = word_rows[r:r + WINDOW_SIZE]
                    row_texts = []
                    for row_words in window_rows:
                        if not row_words:
                            continue
                        row_texts.append(' '.join([w.get('text', '') for w in row_words]))
                    block = ' '.join(row_texts)
                    block_norm = ' '.join(block.split())
                    all_amts = []
                    for m in re_amt.finditer(block):
                        _v = m.group(1) or m.group(2)
                        try:
                            all_amts.append(normalize_amount_str(_v))
                        except Exception:
                            pass
                    if not all_amts:
                        continue
                    has_total_plus = bool(re_total_plus.search(block_norm))
                    has_total_not_plus = bool(re_amount_total.search(block_norm)) and not re_total_plus.search(block_norm)
                    has_pagos_abonos = 'Pagos' in block_norm and 'abonos' in block_norm
                    has_cargos_regulares = 'Cargos' in block_norm and 'regulares' in block_norm and 'no' in block_norm and 'meses' in block_norm
                    if debug_val and (has_total_plus or has_total_not_plus or has_pagos_abonos or has_cargos_regulares):
                        snippet = (block_norm[:120] + '...') if len(block_norm) > 120 else block_norm
                        print(f"[Banamex new format Valor en PDF] Page {page_data.get('page')} window r={r}: amounts={all_amts} | Total+={has_total_plus} Total={has_total_not_plus} | snippet: {snippet!r}", flush=True)
                    # Primary: "Total +" line (e.g. "$1,127.00 Total +") → Total Cargos
                    if need_cargos and has_total_plus:
                        banamex_new_fmt_totals['total_cargos'] = all_amts[0]
                        need_cargos = False
                        print(f"[Banamex new format Valor en PDF] Total Cargos (Valor en PDF) set to {banamex_new_fmt_totals['total_cargos']} from line containing 'Total +'", flush=True)
                    # Primary: "Total" line after "Total +" (e.g. "$438.55 Total") → Total Abonos
                    if need_abonos and has_total_not_plus:
                        banamex_new_fmt_totals['total_abonos'] = all_amts[-1]
                        need_abonos = False
                        print(f"[Banamex new format Valor en PDF] Total Abonos (Valor en PDF) set to {banamex_new_fmt_totals['total_abonos']} from line containing 'Total'", flush=True)
                    # Fallback: RESUMEN block "Pagos y abonos" / "Cargos regulares (no a meses)" — last match wins (main RESUMEN usually last on page 1 or only occurrence)
                    # Skip zero for Total Abonos so we don't take $0.00 from another row; keep scanning so last non-zero wins
                    if need_abonos and has_pagos_abonos and all_amts[-1] != 0:
                        banamex_new_fmt_totals['total_abonos'] = all_amts[-1]
                        print(f"[Banamex new format Valor en PDF] Total Abonos (Valor en PDF) set to {banamex_new_fmt_totals['total_abonos']} from line(s) containing 'Pagos y abonos'", flush=True)
                    if need_abonos and has_pagos_abonos and all_amts[-1] == 0 and banamex_new_fmt_totals.get('total_abonos') is None:
                        banamex_new_fmt_totals['total_abonos'] = 0.0
                    if need_cargos and has_cargos_regulares:
                        banamex_new_fmt_totals['total_cargos'] = all_amts[0]
                        print(f"[Banamex new format Valor en PDF] Total Cargos (Valor en PDF) set to {banamex_new_fmt_totals['total_cargos']} from line(s) containing 'Cargos regulares (no a meses)'", flush=True)
                    # Do not break early: keep scanning all pages so last occurrence wins (main RESUMEN often appears once or last)
                # No early break for fallback; we scan all pages so last match wins
            # Always print result when something is missing (normal and --debug)
            if banamex_new_fmt_totals.get('total_abonos') is None or banamex_new_fmt_totals.get('total_cargos') is None:
                print(f"[Banamex new format Valor en PDF] After scan: total_abonos={banamex_new_fmt_totals.get('total_abonos')}, total_cargos={banamex_new_fmt_totals.get('total_cargos')} (missing = not found in any page)", flush=True)
            else:
                print(f"[Banamex new format Valor en PDF] After scan: total_abonos={banamex_new_fmt_totals.get('total_abonos')}, total_cargos={banamex_new_fmt_totals.get('total_cargos')}", flush=True)
    # Banamex new format: write summary-line totals into pdf_summary for Valor en PDF validation
    if bank_config['name'] == 'Banamex' and banamex_new_format and banamex_new_fmt_totals:
        if pdf_summary is None:
            pdf_summary = {}
        if banamex_new_fmt_totals.get('total_cargos') is not None:
            pdf_summary['total_cargos'] = banamex_new_fmt_totals['total_cargos']
            if debug_mode:
                print(f"[Banamex new format Valor en PDF] pdf_summary['total_cargos'] = {pdf_summary['total_cargos']}", flush=True)
        if banamex_new_fmt_totals.get('total_abonos') is not None:
            pdf_summary['total_abonos'] = banamex_new_fmt_totals['total_abonos']
            if debug_mode:
                print(f"[Banamex new format Valor en PDF] pdf_summary['total_abonos'] = {pdf_summary['total_abonos']}", flush=True)
    # Si es HSBC con OCR, pdf_summary ya fue extraído arriba en extract_hsbc_summary_from_ocr_text
    extracted_totals = calculate_extracted_totals(df_mov, bank_config['name'])
    
    # For INTERCAM and Mercury, use last Saldo from Bank Statement Report for validation "Valor en PDF" (Saldo Final)
    # When not in PDF we backfill; for Mercury always use last Saldo from report (like other banks) so Valor en PDF = last value in Saldo column
    if bank_config['name'] in ('INTERCAM', 'Mercury') and extracted_totals.get('saldo_final') is not None:
        if pdf_summary is not None:
            if bank_config['name'] == 'Mercury':
                pdf_summary['saldo_final'] = extracted_totals['saldo_final']
            elif pdf_summary.get('saldo_final') is None:
                pdf_summary['saldo_final'] = extracted_totals['saldo_final']
    
    # For Banamex: move rows containing "EMP" to separate sheet "Banca Electrónica Empresarial"
    df_banamex_emp = None
    if bank_config['name'] == 'Banamex' and not df_mov.empty:
        desc_col = 'Descripción' if 'Descripción' in df_mov.columns else 'descripcion'
        if desc_col in df_mov.columns:
            emp_mask = df_mov[desc_col].astype(str).str.contains('EMP', na=False)
            if emp_mask.any():
                df_banamex_emp = df_mov[emp_mask].copy()
                df_mov = df_mov[~emp_mask].reset_index(drop=True)

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
    
    # Append blank row, then RFC, Name, Period (all banks)
    _summary = pdf_summary or {}
    # Ensure at least 2 columns (e.g. HSBC OCR can yield single-column df when no valid movements)
    if len(df_mov.columns) < 2:
        default_mov_cols = ['Fecha', 'Descripción', 'Abonos', 'Cargos', 'Saldo']
        for c in default_mov_cols:
            if c not in df_mov.columns:
                df_mov[c] = ''
        df_mov = df_mov[[c for c in default_mov_cols if c in df_mov.columns]]
    col0 = df_mov.columns[0]
    col1 = df_mov.columns[1]
    row_blank = {c: '' for c in df_mov.columns}
    row_rfc = {c: '' for c in df_mov.columns}
    row_rfc[col0] = 'RFC'
    row_rfc[col1] = _summary.get('rfc') or ''
    row_name = {c: '' for c in df_mov.columns}
    row_name[col0] = 'Name'
    row_name[col1] = _summary.get('name') or ''
    row_period = {c: '' for c in df_mov.columns}
    row_period[col0] = 'Period:'
    row_period[col1] = _summary.get('period_text') or ''
    df_mov = pd.concat([
        df_mov,
        pd.DataFrame([row_blank, row_rfc, row_name, row_period]),
    ], ignore_index=True)
    
    # Create validation sheet
    #print("📋 Creando pestaña de validación...")
    # Check if bank has Saldo column in Movements
    has_saldo_column = 'Saldo' in df_mov.columns or 'saldo' in df_mov.columns
    df_validation = create_validation_sheet(pdf_summary, extracted_totals, has_saldo_column=has_saldo_column)
    #print(f"✅ DataFrame de validación creado con {len(df_validation)} filas")
    #print(f"   Columnas: {list(df_validation.columns)}")
    
    print("📊 Exporting to Excel...", flush=True)
    
    # Print validation summary to console
    print_validation_summary(pdf_summary, extracted_totals, df_validation, df_mov)
    
    # Determine number of sheets to write
    num_sheets = 3  # Summary, Movements, Data Validation
    if df_banamex_emp is not None and not df_banamex_emp.empty:
        num_sheets += 1  # Add Banca Electrónica Empresarial (Banamex EMP)
    if df_transferencias is not None and not df_transferencias.empty:
        num_sheets += 1  # Add Transferencias sheet
    if df_digitem is not None and not df_digitem.empty:
        num_sheets += 1  # Add DIGITEM sheet
    if df_metas is not None and not df_metas.empty:
        num_sheets += 1  # Add METAS sheet (Santander)
    
    # write sheets: summary, movements, validation, and optionally Transferencias, DIGITEM, METAS
    try:
        sheet_names = "Summary, Bank Statement Report, Data Validation"
        if df_banamex_emp is not None and not df_banamex_emp.empty:
            sheet_names += ", Banca Electrónica Empresarial"
        if df_transferencias is not None and not df_transferencias.empty:
            sheet_names += ", Transferencias"
        if df_digitem is not None and not df_digitem.empty:
            sheet_names += ", DIGITEM"
        if df_metas is not None and not df_metas.empty:
            sheet_names += ", METAS"
        #print(f"📝 Escribiendo Excel con {num_sheets} pestañas: {sheet_names}")
        # Clean amount columns for Banamex: extract only numeric amounts from mixed text
        def _banamex_extract_amount(value):
            """Extract only the numeric amount from a string that may contain text."""
            if pd.isna(value) or value == '':
                return ''
            value_str = str(value).strip()
            amounts = DEC_AMOUNT_RE.findall(value_str)
            if amounts:
                amount = amounts[-1]
                if '.' not in amount:
                    amount = amount + '.00'
                elif amount.count('.') == 1:
                    parts = amount.split('.')
                    if len(parts[1]) == 1:
                        amount = amount + '0'
                return amount
            return ''

        if bank_config['name'] == 'Banamex':
            for col in ['Cargos', 'Abonos', 'Saldo']:
                if col in df_mov.columns:
                    df_mov[col] = df_mov[col].apply(_banamex_extract_amount)
            # Apply same cleaning to Banca Electrónica Empresarial sheet if present
            if df_banamex_emp is not None and not df_banamex_emp.empty:
                for col in ['Cargos', 'Abonos', 'Saldo']:
                    if col in df_banamex_emp.columns:
                        df_banamex_emp[col] = df_banamex_emp[col].apply(_banamex_extract_amount)
        
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            # Set author in Excel properties
            writer.book.properties.creator = "CONTAAYUDA"
            
            #print("   - Escribiendo pestaña 'Summary'...")
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            df_mov.to_excel(writer, sheet_name='Bank Statement Report', index=False)

            # Write Banamex Banca Electrónica Empresarial (EMP) sheet if present
            if df_banamex_emp is not None and not df_banamex_emp.empty:
                df_banamex_emp.to_excel(writer, sheet_name='Banca Electrónica Empresarial', index=False)
            
            # Write Transferencias sheet if available
            if df_transferencias is not None and not df_transferencias.empty:
                #print("   - Escribiendo pestaña 'Transferencias'...")
                df_transferencias.to_excel(writer, sheet_name='Transferencias', index=False)
                #print(f"   ✅ 'Transferencias' sheet created successfully with {len(df_transferencias)} rows")
            
            # Write DIGITEM sheet if available
            if df_digitem is not None and not df_digitem.empty:
                # print("   - Escribiendo pestaña 'DIGITEM'...")
                df_digitem.to_excel(writer, sheet_name='DIGITEM', index=False)
                # print(f"   ✅ 'DIGITEM' sheet created successfully with {len(df_digitem)} rows")
            
            # Write METAS sheet if available (Santander Mis Metas section)
            if df_metas is not None and not df_metas.empty:
                df_metas.to_excel(writer, sheet_name='METAS', index=False)
            
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
        
        # Validar que el Excel se creó correctamente
        if not os.path.isfile(output_excel):
            print(f"❌ Error: El archivo Excel no se creó: {output_excel}")
            sys.exit(1)
        
        excel_size = os.path.getsize(output_excel)
        if excel_size == 0:
            print(f"❌ Error: El archivo Excel está vacío: {output_excel}")
            sys.exit(1)
        
        print(f"✅ Excel file created successfully -> {output_excel} ({excel_size:,} bytes)" + "\n", flush=True)
        sys.exit(0)
    except Exception as e:
        print(f"❌ Excel file not created -> {output_excel}")
        print(f"   Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
