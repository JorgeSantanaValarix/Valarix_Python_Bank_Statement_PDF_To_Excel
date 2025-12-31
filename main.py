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
            "fecha": (0, 80),              # Columna Fecha de Operación
            "liq": (80, 160),              # Columna LIQ (Liquidación)
            "descripcion": (160, 400),     # Columna Descripción
            "cargos": (362, 398),          # Columna Cargos
            "abonos": (422, 458),          # Columna Abonos
            "saldo": (539, 593),           # Columna Saldo
        }
    },
    # Add more banks here as needed
}

DEFAULT_BANK = "BBVA"

# Decimal / thousands amount regex (module-level so helpers can use it)
DEC_AMOUNT_RE = re.compile(r"\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")


def find_column_coordinates(pdf_path: str, page_number: int = 1):
    """Extract all words from a page and show their coordinates.
    Helps user find exact X ranges for columns.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_number > len(pdf.pages):
                print(f"❌ El PDF solo tiene {len(pdf.pages)} páginas")
                return
            
            page = pdf.pages[page_number - 1]
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            
            if not words:
                print("❌ No se encontraron palabras en la página")
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
            print("""
Ejemplo:
BANK_CONFIGS = {
    "BBVA": {
        "name": "BBVA",
        "columns": {
            "fecha": (x_min, x_max),           # Columna Fecha de Operación
            "liq": (x_min, x_max),              # Columna LIQ (Liquidación)
            "descripcion": (x_min, x_max),     # Columna Descripción
            "cargos": (x_min, x_max),          # Columna Cargos
            "abonos": (x_min, x_max),          # Columna Abonos
            "saldo": (x_min, x_max),           # Columna Saldo
        }
    },
}
            """)
    
    except Exception as e:
        print(f"❌ Error: {e}")


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
    print(f"✅ Excel file created: {output_path}")


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
    day_re = re.compile(r"\b(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}\b")
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
    """
    word_center = (word_x0 + word_x1) / 2
    for col_name, (x_min, x_max) in columns.items():
        if x_min <= word_center <= x_max:
            return col_name
    return None


def is_transaction_row(row_data):
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
    day_re = re.compile(r"\b(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}\b")
    has_date = bool(day_re.search(fecha))
    
    # Must have at least one numeric amount
    has_amount = bool(cargos or abonos or saldo)
    
    return has_date and has_amount


def extract_movement_row(words, columns):
    """Extract a structured movement row from grouped words using coordinate-based column assignment."""
    row_data = {col: '' for col in columns.keys()}
    amounts = []
    
    # Sort words by X coordinate within the row
    sorted_words = sorted(words, key=lambda w: w.get('x0', 0))
    
    for word in sorted_words:
        text = word.get('text', '')
        x0 = word.get('x0', 0)
        x1 = word.get('x1', 0)
        center = (x0 + x1) / 2

        # detect amount tokens inside the word
        m = DEC_AMOUNT_RE.search(text)
        if m:
            amounts.append((m.group(), center))

        col_name = assign_word_to_column(x0, x1, columns)
        if col_name:
            if row_data[col_name]:
                row_data[col_name] += ' ' + text
            else:
                row_data[col_name] = text

    # attach detected amounts for later disambiguation
    row_data['_amounts'] = amounts
    return row_data


def main():
    # Validate input
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python main2.py <input.pdf>              # Parse PDF and create Excel")
        print("  python main2.py <input.pdf> --find <page> # Find column coordinates on page N")
        print("\nExample:")
        print("  python main2.py BBVA.pdf")
        print("  python main2.py BBVA.pdf --find 2")
        sys.exit(1)

    pdf_path = sys.argv[1]
    
    # Check for --find mode
    if len(sys.argv) >= 3 and sys.argv[2] == '--find':
        page_num = int(sys.argv[3]) if len(sys.argv) > 3 else 1
        print(f"🔍 Buscando coordenadas en página {page_num}...")
        find_column_coordinates(pdf_path, page_num)
        return

    if not os.path.isfile(pdf_path):
        print("❌ File not found.")
        sys.exit(1)

    if not pdf_path.lower().endswith(".pdf"):
        print("❌ Only PDF files are supported.")
        sys.exit(1)

    output_excel = os.path.splitext(pdf_path)[0] + ".xlsx"

    print("📄 Reading PDF...")
    extracted_data = extract_text_from_pdf(pdf_path)
    # split pages into lines
    pages_lines = split_pages_into_lines(extracted_data)

    # Get bank config (default to BBVA)
    bank_config = BANK_CONFIGS.get(DEFAULT_BANK)
    if not bank_config:
        print(f"⚠️  Bank config not found for {DEFAULT_BANK}, using BBVA defaults")
        bank_config = BANK_CONFIGS["BBVA"]
    
    columns_config = bank_config["columns"]

    # find where movements start (first line anywhere that matches a date or contains header keywords)
    day_re = re.compile(r"\b(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}\b")
    # match lines that contain both 'fecha' AND 'descripcion', OR lines that contain 'concepto'
    # Implemented with lookahead for the AND case, and an alternation for 'concepto'
    header_keywords_re = re.compile(r"(?:(?=.*\bfecha\b)(?=.*\bdescripcion\b))|(?:\bconcepto\b)", re.I)
    # ensure a reusable date pattern is available for later checks
    date_pattern = day_re
    movement_start_found = False
    movement_start_page = None
    movement_start_index = None
    movements_lines = []
    for p in pages_lines:
        if not movement_start_found:
            for i, ln in enumerate(p['lines']):
                if day_re.search(ln) or header_keywords_re.search(ln):
                    movement_start_found = True
                    movement_start_page = p['page']
                    movement_start_index = i
                    # collect from this line onward in this page
                    movements_lines.extend(p['lines'][i:])
                    break
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
    movement_rows = []
    # regex to detect decimal-like amounts (used to strip amounts from descriptions and detect amounts)
    dec_amount_re = re.compile(r"\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")
    for page_data in extracted_data:
        page_num = page_data['page']
        words = page_data.get('words', [])
        
        if not words:
            continue
        
        # Check if this page contains movements (page >= movement_start_page if found)
        if movement_start_found and page_num < movement_start_page:
            continue
        
        # Group words by row
        word_rows = group_words_by_row(words)
        
        for row_words in word_rows:
            if not row_words:
                continue

            # Extract structured row using coordinates
            row_data = extract_movement_row(row_words, columns_config)

            # Determine if this row starts a new movement (contains a date)
            # A new movement begins when the 'fecha' column contains a date token.
            fecha_val = str(row_data.get('fecha') or '')
            has_date = bool(date_pattern.search(fecha_val))

            if has_date:
                row_data['page'] = page_num
                movement_rows.append(row_data)
            else:
                # Continuation row: append description-like text to previous movement if exists
                if movement_rows:
                    prev = movement_rows[-1]
                    # Collect possible text pieces from this row (prefer descripcion, then liq, then any other text)
                    cont_parts = []
                    for k in ('descripcion', 'fecha'):
                        v = row_data.get(k)
                        if v:
                            cont_parts.append(str(v))
                    # Also capture any stray text in other columns
                    for k, v in row_data.items():
                        if k in ('descripcion', 'fecha', 'cargos', 'abonos', 'saldo', 'page'):
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
                    # No previous movement to attach to; ignore or treat as header
                    continue

    # prepare DataFrames
    df_summary = pd.DataFrame({'informacion': summary_lines})
    
    # Reassign amounts to cargos/abonos/saldo by proximity when needed
    if movement_rows:
        # prepare column centers
        col_centers = {}
        for col in ('cargos', 'abonos', 'saldo'):
            if col in columns_config:
                x0, x1 = columns_config[col]
                col_centers[col] = (x0 + x1) / 2

        for r in movement_rows:
            amounts = r.get('_amounts', [])
            if not amounts:
                continue

            # If columns already have numbers, keep them but prefer reassignment
            # We'll assign each detected amount to the nearest numeric column
            for amt_text, center in amounts:
                # find nearest column among available numeric columns
                if not col_centers:
                    continue
                nearest = min(col_centers.keys(), key=lambda c: abs(center - col_centers[c]))
                # set or append
                existing = r.get(nearest) or ''
                if existing:
                    # avoid duplicating the same token
                    if amt_text not in existing:
                        r[nearest] = (existing + ' ' + amt_text).strip()
                else:
                    r[nearest] = amt_text

            # Remove amount tokens from descripcion if present
            if r.get('descripcion'):
                r['descripcion'] = DEC_AMOUNT_RE.sub('', r.get('descripcion'))

            # cleanup helper key
            if '_amounts' in r:
                del r['_amounts']

        df_mov = pd.DataFrame(movement_rows)
    else:
        movement_entries = group_entries_from_lines(movements_lines)
        df_mov = pd.DataFrame({'raw': movement_entries})

    # Split combined fecha values into two separate columns: Fecha Oper and Fecha Liq
    # Works for coordinate-based extraction (column 'fecha') and for fallback raw lines ('raw').
    date_pattern = re.compile(r"(?:0[1-9]|[12][0-9]|3[01])(?:[\/\-\s])[A-Za-z]{3}", re.I)

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

    if 'fecha' in df_mov.columns:
        dates = df_mov['fecha'].astype(str).apply(_extract_two_dates)
    elif 'raw' in df_mov.columns:
        dates = df_mov['raw'].astype(str).apply(_extract_two_dates)
    else:
        dates = pd.Series([(None, None)] * len(df_mov))

    df_mov['Fecha Oper'] = dates.apply(lambda t: t[0])
    df_mov['Fecha Liq'] = dates.apply(lambda t: t[1])

    # Remove original 'fecha' if present
    if 'fecha' in df_mov.columns:
        df_mov = df_mov.drop(columns=['fecha'])
    
    # For non-BBVA banks, use only 'Fecha' column (based on Fecha Oper) and remove Fecha Liq
    if bank_config['name'] != 'BBVA':
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
                # Return first two amounts
                return (amounts[0], amounts[1])
            elif len(amounts) == 1:
                # Only one amount found
                return (amounts[0], None)
            else:
                return (None, None)
        
        # Extract the two amounts from saldo column
        amounts = df_mov['saldo'].astype(str).apply(_extract_two_amounts)
        df_mov['OPERACIÓN'] = amounts.apply(lambda t: t[1])  # First amount is OPERACIÓN
        df_mov['LIQUIDACIÓN'] = amounts.apply(lambda t: t[0])  # Second amount is LIQUIDACIÓN
        
        # Remove the original 'saldo' column
        df_mov = df_mov.drop(columns=['saldo'])

    # Merge 'liq' and 'descripcion' into a single 'Descripcion' column.
    # Remove any date tokens and decimal amounts from the description text.
    dec_amount_re = re.compile(r"\d{1,3}(?:[\.,\s]\d{3})*(?:[\.,]\d{2})")
    # date_pattern already defined above; reuse it

    def _build_description(row):
        parts = []
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

    # Reorder columns according to bank type
    if bank_config['name'] == 'BBVA':
        # For BBVA: Fecha Oper, Fecha Liq, Descripción, Cargos, Abonos, Operación, Liquidación
        desired_order = ['Fecha Oper', 'Fecha Liq', 'Descripción', 'Cargos', 'Abonos', 'Operación', 'Liquidación']
        # Filter to only include columns that exist in the dataframe
        desired_order = [col for col in desired_order if col in df_mov.columns]
        # Add any remaining columns that are not in desired_order
        other_cols = [c for c in df_mov.columns if c not in desired_order]
        df_mov = df_mov[desired_order + other_cols]
    else:
        # For other banks: Fecha, Descripción, Cargos, Abonos, Saldo
        desired_order = ['Fecha', 'Descripción', 'Cargos', 'Abonos', 'Saldo']
        # Filter to only include columns that exist in the dataframe
        desired_order = [col for col in desired_order if col in df_mov.columns]
        # Add any remaining columns that are not in desired_order
        other_cols = [c for c in df_mov.columns if c not in desired_order]
        df_mov = df_mov[desired_order + other_cols]

    print("📊 Exporting to Excel...")
    # write two sheets: summary and movements
    try:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            df_mov.to_excel(writer, sheet_name='Movements', index=False)
        print(f"✅ Excel file created -> {output_excel}")
    except Exception as e:
        print('Error writing Excel:', e)


if __name__ == "__main__":
    main()
