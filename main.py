import sys
import os
import re
import pdfplumber
import pandas as pd


# Bank configurations with column coordinate ranges (X-axis)
# Use find_coordinates.py to get the exact ranges for your PDF
# python find_coordinates.py <pdf_path> <page_number>
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
            "cargos": (340, 388),          # Columna Cargos
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
        "columns": {
            "fecha": (17, 45),             # Columna Fecha de Operación
            "descripcion": (55, 260),      # Columna Descripción (ampliado para capturar mejor)
            "cargos": (275, 316),          # Columna Cargos (ampliado ligeramente)
            "abonos": (345, 395),          # Columna Abonos (ampliado ligeramente)
            "saldo": (425, 472),           # Columna Saldo (ampliado ligeramente)
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


def find_column_coordinates(pdf_path: str, page_number: int = 1):
    """Extract all words from a page and show their coordinates.
    Helps user find exact X ranges for columns.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_number > len(pdf.pages):
                # print(f"❌ El PDF solo tiene {len(pdf.pages)} páginas")
                return
            
            page = pdf.pages[page_number - 1]
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            
            if not words:
                # print("❌ No se encontraron palabras en la página")
                return
            
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


def detect_bank_from_pdf(pdf_path: str) -> str:
    """
    Detect the bank from PDF content by reading line by line.
    Returns the bank name if detected, otherwise returns DEFAULT_BANK.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Read first few pages (usually bank name appears early)
            max_pages_to_check = min(3, len(pdf.pages))
            
            for page_num in range(max_pages_to_check):
                page = pdf.pages[page_num]
                text = page.extract_text()
                
                if not text:
                    continue
                
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
    
    except Exception as e:
        pass
        # print(f"⚠️  Error al detectar banco: {e}")
    
    # If no bank detected, return default
    #print(f"⚠️  No se pudo detectar el banco, usando: {DEFAULT_BANK}")
    return DEFAULT_BANK


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
            
            # For Banregio, collect text from all pages to find "Total" line
            if bank_name == "Banregio":
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
                # Konfio:
                # "Saldo anterior $ 317,215.14"
                # "Pagos - $ 97,000.00"
                # "Compras y cargos $ 56,176.79"
                # "Saldo total al corte $ 312,227.05"
                #print(f"🔍 Buscando patrones Konfio en {len(all_lines)} líneas...")
                for i, line in enumerate(all_lines):
                    # Saldo anterior
                    if not summary_data['saldo_anterior']:
                        match = re.search(r'Saldo\s+anterior\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Konfio: Encontrado saldo anterior: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_anterior'] = amount
                    
                    # Compras y cargos
                    if not summary_data['total_cargos']:
                        match = re.search(r'Compras\s+y\s+cargos\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Konfio: Encontrado cargos: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['total_cargos'] = amount
                                summary_data['total_retiros'] = amount
                    
                    # Saldo total al corte
                    if not summary_data['saldo_final']:
                        match = re.search(r'Saldo\s+total\s+al\s+corte\s+\$\s*([\d,\.]+)', line, re.I)
                        if match:
                            amount = normalize_amount_str(match.group(1))
                            if amount > 0:
                                #print(f"✅ Konfio: Encontrado saldo total al corte: ${amount:,.2f} en línea {i+1}: {line[:80]}")
                                summary_data['saldo_final'] = amount
            
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
    """
    extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            # also extract words with positions for coordinate-based column detection
            try:
                words = page.extract_words(x_tolerance=3, y_tolerance=3)
            except Exception:
                words = []
            extracted_data.append({
                "page": page_number,
                "content": text if text else "",
                "words": words
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
    
    # Sort by top coordinate
    sorted_words = sorted(words, key=lambda w: w.get('top', 0))
    
    rows = []
    current_row = []
    current_y = None
    
    for word in sorted_words:
        word_y = word.get('top', 0)
        if current_y is None:
            current_y = word_y
        
        # If word is within y_tolerance of current row, add it
        if abs(word_y - current_y) <= y_tolerance:
            current_row.append(word)
        else:
            # Start a new row
            if current_row:
                rows.append(current_row)
            current_row = [word]
            current_y = word_y
    
    # Don't forget the last row
    if current_row:
        rows.append(current_row)
    
    return rows


def assign_word_to_column(word_x0, word_x1, columns):
    """Assign a word (with x0, x1 coordinates) to a column based on X-ranges.
    Returns column name or None if not in any range.
    Prioritizes numeric columns (cargos, abonos, saldo) over description when there's overlap.
    """
    word_center = (word_x0 + word_x1) / 2
    
    # First, check numeric columns (cargos, abonos, saldo) to prioritize them
    # This fixes the issue where cargos (360-398) overlaps with descripcion (160-400)
    numeric_cols = ['cargos', 'abonos', 'saldo']
    for col_name in numeric_cols:
        if col_name in columns:
            x_min, x_max = columns[col_name]
            if x_min <= word_center <= x_max:
                return col_name
    
    # Then check other columns (fecha, liq, descripcion, etc.)
    for col_name, (x_min, x_max) in columns.items():
        if col_name not in numeric_cols:  # Skip numeric cols, already checked
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
    # For Banregio: date is only 2 digits (01-31)
    # For Base: date format is DD/MM/YYYY (e.g., "30/04/2024")
    # For other banks: supports both "DIA MES" (01 ABR) and "MES DIA" (ABR 01) formats
    # Pattern for dates: supports multiple formats including "DIA MES AÑO" (06 mar 2023)
    has_date = False
    if bank_name == 'Banregio':
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
    
    # Pattern to detect dates (for separating date from description)
    if date_pattern is None:
        if bank_name == 'Banregio':
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

        # detect amount tokens inside the word
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
            # For Banregio, ensure we split the 2-digit date from the rest of the text
            if bank_name == 'Banregio':
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
            if row_data[col_name]:
                row_data[col_name] += ' ' + text
            else:
                row_data[col_name] = text

    # attach detected amounts for later disambiguation
    row_data['_amounts'] = amounts
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
    
    # Detect bank from PDF content (read PDF directly for detection)
    detected_bank = detect_bank_from_pdf(pdf_path)
    
    # Now extract full data
    extracted_data = extract_text_from_pdf(pdf_path)
    # split pages into lines
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
                        print(f"✅ BanBajío: Encontrado 'DETALLE DE LA CUENTA: CUENTA' en página {p['page']}, línea {i+1}: {ln[:100]}")
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
                            print(f"✅ BanBajío: Inicio de extracción detectado en página {p['page']}, línea {i+2}: {next_line[:100]}")
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
                        if day_re.search(ln):
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
    # regex to detect decimal-like amounts (used to strip amounts from descriptions and detect amounts)
    dec_amount_re = re.compile(r"\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")
    # Pattern to detect end of movements table for specific banks
    movement_end_pattern = None
    if bank_config['name'] == 'Banamex':
        movement_end_pattern = re.compile(r'SALDO\s+MINIMO\s+REQUERIDO', re.I)
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
        
        # Check if this page contains movements (page >= movement_start_page if found)
        if movement_start_found and page_num < movement_start_page:
            continue
        
        # Group words by row
        # Use y_tolerance=3 for all banks to avoid grouping multiple movements into one row
        # The split_row_if_multiple_movements function will handle cases where movements are still grouped
        word_rows = group_words_by_row(words, y_tolerance=3)
    
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
                        if bank_config['name'] == 'Banbajío':
                            print(f"🛑 BanBajío: Fin de extracción detectado en página {page_num}, fila {row_idx+1}: {all_row_text[:100]}")
                
                if match_found:
                    #print(f"🛑 Fin de tabla de movimientos detectado en página {page_num}")
                    extraction_stopped = True
                    break

            # Extract structured row using coordinates
            # Pass bank_name and date_pattern to enable date/description separation
            row_data = extract_movement_row(row_words, columns_config, bank_config['name'], date_pattern)

            # For BanBajío, debug: print all processed rows to see what's being filtered (only first 5 pages)
            if bank_config['name'] == 'Banbajío' and movement_start_found and page_num <= 4:
                all_row_text_debug = ' '.join([w.get('text', '') for w in row_words])
                fecha_val_debug = str(row_data.get('fecha') or '').strip()
                desc_val_debug = str(row_data.get('descripcion') or '').strip()
                cargos_val_debug = str(row_data.get('cargos') or '').strip()
                abonos_val_debug = str(row_data.get('abonos') or '').strip()
                saldo_val_debug = str(row_data.get('saldo') or '').strip()
                print(f"🔍 BanBajío: Procesando fila en página {page_num}, fila {row_idx+1} - Fecha: '{fecha_val_debug}', Desc: '{desc_val_debug[:60]}', Cargos: '{cargos_val_debug}', Abonos: '{abonos_val_debug}', Saldo: '{saldo_val_debug}'")
                print(f"   📝 Texto completo de la fila: {all_row_text_debug[:150]}")

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
                else:
                    has_date = bool(date_pattern.search(fecha_val))
                    # For BanBajío, debug: print date detection results (only first 5 pages)
                    if bank_config['name'] == 'Banbajío' and page_num <= 4:
                        date_match_result = date_pattern.search(fecha_val)
                        match_text = date_match_result.group() if date_match_result else 'None'
                        print(f"   🔍 Detección de fecha - Valor en columna fecha: '{fecha_val}', has_date: {has_date}, match: '{match_text}'")
                
                # Check if row has valid data (date, description, or amounts)
                has_valid_data = has_date
                if not has_valid_data:
                    # Check if row has description or amounts
                    desc_val = str(row_data.get('descripcion') or '').strip()
                    has_amounts = len(row_data.get('_amounts', [])) > 0
                    has_cargos_abonos = bool(row_data.get('cargos') or row_data.get('abonos') or row_data.get('saldo'))
                    has_valid_data = bool(desc_val or has_amounts or has_cargos_abonos)
                
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
                
                # For BanBajío, debug: print why rows are being skipped (only first 5 pages)
                if bank_config['name'] == 'Banbajío' and not has_description_or_amounts and page_num <= 4:
                    fecha_val = str(row_data.get('fecha') or '').strip()
                    print(f"⚠️ BanBajío: Fila rechazada (sin descripción ni montos) en página {page_num}, fila {row_idx+1} - Fecha: '{fecha_val}', Desc: '{desc_val[:60]}', Montos: {has_amounts}, Cargos/Abonos: {has_cargos_abonos}")
                
                if has_description_or_amounts:
                    # For Banregio, only include rows where description starts with "TRA" or "DOC"
                    if bank_config['name'] == 'Banregio':
                        desc_val_check = str(row_data.get('descripcion') or '').strip()
                        if not (desc_val_check.startswith('TRA') or desc_val_check.startswith('DOC') or desc_val_check.startswith('INT') or desc_val_check.startswith('EFE')):
                            # Skip this row - doesn't start with TRA or DOC or INT or EFE
                            continue
                    #row_data['page'] = page_num
                    movement_rows.append(row_data)
                    # For BanBajío, track detected rows and print first and second rows that will be exported to Excel
                    if bank_config['name'] == 'Banbajío':
                        banbajio_detected_rows += 1
                        if banbajio_detected_rows == 2:
                            all_row_text = ' '.join([w.get('text', '') for w in row_words])
                            print(f"✅ BanBajío: Segunda fila detectada (válida) en página {page_num}, fila {row_idx+1}: {all_row_text[:150]}")
                        if len(movement_rows) == 1:
                            fecha_val = str(row_data.get('fecha') or '').strip()
                            desc_val = str(row_data.get('descripcion') or '').strip()
                            cargos_val = str(row_data.get('cargos') or '').strip()
                            abonos_val = str(row_data.get('abonos') or '').strip()
                            saldo_val = str(row_data.get('saldo') or '').strip()
                            print(f"📊 BanBajío: Primera fila que se exportará al Excel - Fecha: '{fecha_val}', Desc: '{desc_val[:80]}', Cargos: '{cargos_val}', Abonos: '{abonos_val}', Saldo: '{saldo_val}'")
                        elif len(movement_rows) == 2:
                            fecha_val = str(row_data.get('fecha') or '').strip()
                            desc_val = str(row_data.get('descripcion') or '').strip()
                            cargos_val = str(row_data.get('cargos') or '').strip()
                            abonos_val = str(row_data.get('abonos') or '').strip()
                            saldo_val = str(row_data.get('saldo') or '').strip()
                            print(f"📊 BanBajío: Segunda fila que se exportará al Excel - Fecha: '{fecha_val}', Desc: '{desc_val[:80]}', Cargos: '{cargos_val}', Abonos: '{abonos_val}', Saldo: '{saldo_val}'")
            elif has_valid_data and bank_config['name'] == 'Banbajío':
                # For BanBajío, also add rows without date if they contain "SALDO INICIAL"
                desc_val = str(row_data.get('descripcion') or '').strip()
                if re.search(r'SALDO\s+INICIAL', desc_val, re.I):
                    has_amounts = len(row_data.get('_amounts', [])) > 0
                    has_cargos_abonos = bool(row_data.get('cargos') or row_data.get('abonos') or row_data.get('saldo'))
                    if has_amounts or has_cargos_abonos:
                        #row_data['page'] = page_num
                        movement_rows.append(row_data)
                        # For BanBajío, track detected rows and print first row that will be exported to Excel (SALDO INICIAL)
                        if bank_config['name'] == 'Banbajío':
                            banbajio_detected_rows += 1
                            if banbajio_detected_rows == 2:
                                all_row_text = ' '.join([w.get('text', '') for w in row_words])
                                print(f"✅ BanBajío: Segunda fila detectada (válida) en página {page_num}, fila {row_idx+1}: {all_row_text[:150]}")
                        if len(movement_rows) == 1:
                            fecha_val = str(row_data.get('fecha') or '').strip()
                            desc_val = str(row_data.get('descripcion') or '').strip()
                            cargos_val = str(row_data.get('cargos') or '').strip()
                            abonos_val = str(row_data.get('abonos') or '').strip()
                            saldo_val = str(row_data.get('saldo') or '').strip()
                            print(f"📊 BanBajío: Primera fila que se exportará al Excel - Fecha: '{fecha_val}', Desc: '{desc_val[:80]}', Cargos: '{cargos_val}', Abonos: '{abonos_val}', Saldo: '{saldo_val}'")
                    
                    # For Banregio, check if this is a monthly commission (last movement) - stop extraction
                    if bank_config['name'] == 'Banregio':
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
                
                # If row has date but no description/amounts, skip it (incomplete row)
                elif has_valid_data:
                    # Row has valid data but no date - treat as continuation or standalone row
                    # For BanBajío, filter out informational rows like "1 DE ENERO AL 31 DE ENERO DE 2024 PERIODO:"
                    if bank_config['name'] == 'Banbajío':
                        all_row_text_check = ' '.join([w.get('text', '') for w in row_words])
                        all_row_text_check = all_row_text_check.strip()
                        # Check if row contains period information (e.g., "1 DE ENERO AL 31 DE ENERO DE 2024 PERIODO:")
                        if re.search(r'\bPERIODO\s*:?\s*$', all_row_text_check, re.I) or re.search(r'\d+\s+DE\s+[A-Z]+\s+AL\s+\d+\s+DE\s+[A-Z]+', all_row_text_check, re.I):
                            # Skip this row - it's period information, not a movement
                            if page_num <= 4:
                                print(f"⏭️ BanBajío: Fila filtrada (información de período) en página {page_num}, fila {row_idx+1}: {all_row_text_check[:100]}")
                            continue
                        # For BanBajío, debug: print rows without date that are being processed as continuation (only first 5 pages)
                        if not movement_rows and page_num <= 4:
                            print(f"⚠️ BanBajío: Fila sin fecha y sin movimiento previo en página {page_num}, fila {row_idx+1}: {all_row_text_check[:100]}")
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
                        # Continuation row: append description-like text and amounts to previous movement
                        prev = movement_rows[-1]
                        
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
                    pass

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

        for r in movement_rows:
            amounts = r.get('_amounts', [])
            if not amounts:
                continue

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
                tolerance = 10
                assigned = False
                
                # First, check if amount is within any numeric column range
                # This takes priority over description range check
                for col in ('cargos', 'abonos', 'saldo'):
                    if col in col_ranges:
                        x0, x1 = col_ranges[col]
                        # Check with tolerance to handle amounts slightly outside the range
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
                    valid_cols = {}
                    for col in col_centers.keys():
                        if col in col_ranges:
                            x0, x1 = col_ranges[col]
                            # Only consider if center is reasonably close to the column
                            if center >= (x0 - 20) and center <= (x1 + 20):
                                valid_cols[col] = abs(center - col_centers[col])
                    
                    if valid_cols:
                        nearest = min(valid_cols.keys(), key=lambda c: valid_cols[c])
                        # Check if amount is in description range AND not in any numeric column
                        # If it's close to a numeric column, assign it even if it's also in description range
                        in_desc_range = descripcion_range and descripcion_range[0] <= center <= descripcion_range[1]
                        in_num_range = False
                        for col in col_ranges.keys():
                            x0, x1 = col_ranges[col]
                            if (x0 - 20) <= center <= (x1 + 20):
                                in_num_range = True
                                break
                        
                        # Only skip if in description range AND NOT in any numeric column range
                        if not in_desc_range or in_num_range:
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
                
                # Only skip if amount is in description range AND NOT assigned to any numeric column
                # This prevents amounts in cargos/abonos/saldo from being skipped
                if not assigned and descripcion_range:
                    if descripcion_range[0] <= center <= descripcion_range[1]:
                        # Amount is in description range and wasn't assigned to any numeric column
                        # Skip it to avoid assigning description amounts to numeric columns
                        continue

            # Remove amount tokens from descripcion if present
            if r.get('descripcion'):
                r['descripcion'] = DEC_AMOUNT_RE.sub('', r.get('descripcion'))

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
    if 'fecha' in df_mov.columns:
        df_mov = df_mov.drop(columns=['fecha'])
    
    # For non-BBVA banks, use only 'Fecha' column (based on Fecha Oper) and remove Fecha Liq
    # Skip this for Banregio and Base as they already have 'Fecha' column set
    if bank_config['name'] != 'BBVA' and bank_config['name'] != 'Banregio' and bank_config['name'] != 'Base':
        if 'Fecha Oper' in df_mov.columns:
            df_mov['Fecha'] = df_mov['Fecha Oper']
            df_mov = df_mov.drop(columns=['Fecha Oper', 'Fecha Liq'])

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
        if 'Descripcion' in df_mov.columns:
            column_rename['Descripcion'] = 'Descripción'
        if 'cargos' in df_mov.columns:
            column_rename['cargos'] = 'Cargos'
        if 'abonos' in df_mov.columns:
            column_rename['abonos'] = 'Abonos'
        if 'saldo' in df_mov.columns:
            column_rename['saldo'] = 'Saldo'
    
    if column_rename:
        df_mov = df_mov.rename(columns=column_rename)

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

    # Rename "Fecha Liq" to "Fecha Liq." for BBVA if needed
    if bank_config['name'] == 'BBVA' and 'Fecha Liq' in df_mov.columns:
        df_mov = df_mov.rename(columns={'Fecha Liq': 'Fecha Liq.'})

    # Reorder columns according to bank type
    if bank_config['name'] == 'BBVA':
        # For BBVA: Fecha Oper, Fecha Liq., Descripción, Cargos, Abonos, Operación, Liquidación
        desired_order = ['Fecha Oper', 'Fecha Liq.', 'Descripción', 'Cargos', 'Abonos', 'Operación', 'Liquidación']
        # Filter to only include columns that exist in the dataframe
        desired_order = [col for col in desired_order if col in df_mov.columns]
        # Add any remaining columns that are not in desired_order
        other_cols = [c for c in df_mov.columns if c not in desired_order]
        df_mov = df_mov[desired_order + other_cols]
    else:
        # For other banks: Fecha, Descripción, Cargos, Abonos, Saldo (if available)
        # Build desired_order based on what columns are configured for this bank
        desired_order = ['Fecha', 'Descripción', 'Cargos', 'Abonos']
        
        # Only include Saldo if it's in the bank's column configuration
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
    pdf_summary = extract_summary_from_pdf(pdf_path)
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
                numeric_values = df_mov[col].apply(lambda x: normalize_amount_str(x) if pd.notna(x) and str(x).strip() else 0.0)
                total = numeric_values.sum()
                # For Clara, always show total for Abonos (even if negative, zero, or positive)
                if bank_config['name'] == 'Clara' and col == 'Abonos':
                    # Format as currency with 2 decimals (always show, even if negative or zero)
                    total_row[col] = f"{total:,.2f}"
                elif total > 0:
                    # Format as currency with 2 decimals
                    total_row[col] = f"{total:,.2f}"
                else:
                    total_row[col] = ''
            except:
                # If conversion fails, leave empty
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
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            #print("   - Escribiendo pestaña 'Summary'...")
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
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
