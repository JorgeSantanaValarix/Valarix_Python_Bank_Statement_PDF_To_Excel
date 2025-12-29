import re
import os
import pdfplumber
import pandas as pd

# Regex: any alphabetic character (including Spanish accents) and a number with 2 decimals
# letters
alpha_re = re.compile(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]")
# numbers with optional thousands separators and 2 decimals (e.g. 1,123,215.95 or 161.32)
num_re = re.compile(r"\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})")
# Detect day with month like '01/JUN' (common in bank statements)
day_re = re.compile(r"\b(?:0[1-9]|[12][0-9]|3[01])/[A-Za-z]{3}\b")

PDF_PATH = r"C:\\Valarix\\pdf_to_excel\\Test\\BBVA.pdf"

def extract_lines(pdf_path):
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.splitlines():
                # normalize whitespace
                s = " ".join(line.split())
                if not s:
                    continue
                # filter boilerplate lines
                if is_boilerplate(s):
                    continue
                lines.append(s)
    return lines


def group_entries(lines):
    entries = []
    for line in lines:
        if day_re.search(line):
            # start a new entry
            entries.append(line)
        else:
            # continuation of previous entry (merge into description)
            if entries:
                entries[-1] = entries[-1] + ' ' + line
            else:
                # If we find continuation before any date, keep as standalone
                entries.append(line)
    return entries


def normalize_entry(entry: str) -> str:
    # find leading date tokens (e.g., '01/JUN 02/JUN')
    dates = day_re.findall(entry)
    date_part = ' '.join(dates) if dates else ''

    # find all numeric amounts with 2 decimals (keep order)
    amounts = num_re.findall(entry)

    # remove amounts from entry (only first occurrence each)
    s = entry
    for a in amounts:
        s = s.replace(a, '', 1)

    # remove leading date occurrences from the text
    if date_part:
        for d in dates:
            s = s.replace(d, '', 1)

    # normalize whitespace and trim
    desc = ' '.join(s.split()).strip()

    # build normalized line: dates + description + amounts at end
    parts = []
    if date_part:
        parts.append(date_part)
    if desc:
        parts.append(desc)
    if amounts:
        parts.append(' '.join(amounts))

    return '  '.join(parts)


# small integer regex
int_re = re.compile(r"\b\d+\b")


def parse_kv(line: str):
    """Try to parse a label + numeric value(s) from a line.
    Returns dict {'label': str, 'values': [str,...]} or None if no numeric content.
    """
    amounts = num_re.findall(line)
    ints = int_re.findall(line)

    # filter ints that are part of any amount (when removing separators)
    ints_clean = []
    for n in ints:
        in_amount = False
        for a in amounts:
            if n in re.sub(r'[.,]', '', a):
                in_amount = True
                break
        if not in_amount:
            ints_clean.append(n)

    values = amounts + ints_clean
    if not values:
        return None

    # find earliest numeric match to split label/value
    first_match = None
    m_amount = num_re.search(line)
    m_int = int_re.search(line)
    if m_amount and m_int:
        first_match = m_amount if m_amount.start() < m_int.start() else m_int
    else:
        first_match = m_amount or m_int

    if first_match:
        label = line[: first_match.start()].strip()
    else:
        label = ''

    return {'label': label, 'values': values, 'raw': line}


def process_lines(lines):
    transactions = []
    table_groups = []
    current_tx = None
    current_table = None

    i = 0
    while i < len(lines):
        line = lines[i]

        # If this line starts a transaction (date present), begin a new transaction
        if day_re.search(line):
            current_tx = line
            transactions.append(current_tx)
            # while building a transaction we shouldn't treat subsequent lines as table rows
            current_table = None
            i += 1
            continue

        # If currently building a transaction, treat any following line as continuation
        # until a new date is encountered. This prevents numeric reference lines
        # (account numbers, refs) from being mistaken as table rows.
        if current_tx is not None:
            transactions[-1] = transactions[-1] + ' ' + line
            i += 1
            continue

        # Not inside a transaction: detect table headings strictly and parse kv
        if is_heading(line):
            title = line.strip()
            current_table = {'title': title, 'items': []}
            table_groups.append(current_table)
            i += 1
            continue

        # try parse as key-value line
        kv = parse_kv(line)
        if kv:
            if current_table is None:
                current_table = {'title': 'Misc', 'items': []}
                table_groups.append(current_table)

            if not kv['label']:
                if current_table['items']:
                    current_table['items'][-1]['values'].extend(kv['values'])
                else:
                    current_table['items'].append(kv)
            else:
                current_table['items'].append(kv)

            # attach continuation text (non-numeric, non-date) to the label
            if i + 1 < len(lines) and not num_re.search(lines[i + 1]) and not day_re.search(lines[i + 1]):
                current_table['items'][-1]['label'] = (current_table['items'][-1].get('label', '') + ' ' + lines[i + 1]).strip()
                i += 1

            current_tx = None
            i += 1
            continue

        # otherwise skip / reset contexts
        current_table = None
        current_tx = None
        i += 1

    return transactions, table_groups


# Known heading keywords (common table titles)
HEADING_KEYWORDS = [
    "Rendimiento",
    "Comportamiento",
    "Detalle de Movimientos",
    "Información Financiera",
    "Comisiones de la cuenta",
    "Periodo",
    "Periodo DEL",
    "Saldo Promedio",
]


def is_heading(line: str) -> bool:
    # direct keyword match (case-insensitive)
    for k in HEADING_KEYWORDS:
        # match keyword as whole word or at line start to avoid combining multiple headings
        pattern = re.compile(r"(^|\b)" + re.escape(k) + r"(\b|$)", re.I)
        if pattern.search(line):
            return True

    # heuristic: short all-caps lines without numbers often act as headings
    if not num_re.search(line) and line.strip() and len(line) < 60:
        words = line.split()
        # require at least one alpha word and most words uppercase-ish
        alpha_words = [w for w in words if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", w)]
        if alpha_words and sum(1 for w in alpha_words if w.upper() == w) >= max(1, len(alpha_words) // 2):
            return True

    return False


def extract_tables_by_title(lines):
    tables = {}
    current_title = None
    last_kv = None

    for i, line in enumerate(lines):
        # stop grouping when transaction lines appear
        if day_re.search(line):
            current_title = None
            last_kv = None
            continue

        if is_heading(line):
            current_title = line.strip()
            tables.setdefault(current_title, [])
            last_kv = None
            continue

        kv = parse_kv(line)
        if kv:
            if current_title is None:
                current_title = 'Ungrouped'
                tables.setdefault(current_title, [])
            tables[current_title].append(kv)
            last_kv = kv
            continue

        # continuation lines: attach to last kv under current title
        if current_title and last_kv:
            # append as extra text to last kv label
            last_kv['label'] = (last_kv.get('label', '') + ' ' + line).strip()
            continue

    return tables


# Boilerplate patterns to drop from extracted text
BOILERPLATE_PATTERNS = [
    re.compile(r"Estimado Cliente", re.I),
    re.compile(r"Estado de Cuenta MAESTRA", re.I),
    re.compile(r"BBVA MEXICO", re.I),
    re.compile(r"Por Disposición Oficial", re.I),
    re.compile(r"Av\.? Paseo de la Reforma", re.I),
    re.compile(r"R\.F\.C\.|RFC", re.I),
    re.compile(r"Este documento es una representación impresa", re.I),
]


def is_boilerplate(s: str) -> bool:
    for p in BOILERPLATE_PATTERNS:
        if p.search(s):
            return True
    return False


if __name__ == '__main__':
    all_lines = extract_lines(PDF_PATH)
    # process sequentially to extract transactions and table key-values
    transactions, tables = process_lines(all_lines)
    # build transaction records with requested columns
    tx_records = []
    for entry in transactions:
        if is_boilerplate(entry):
            continue
        if alpha_re.search(entry) and num_re.search(entry):
            dates = day_re.findall(entry)
            fecha_op = dates[0] if len(dates) >= 1 else 'NULL'
            fecha_liq = dates[1] if len(dates) >= 2 else 'NULL'

            amounts = num_re.findall(entry)

            # remove amounts and dates to form description
            s = entry
            for a in amounts:
                s = s.replace(a, '', 1)
            for d in dates:
                s = s.replace(d, '', 1)
            desc = ' '.join(s.split()).strip()

            # try extract a 'Ref' token from description
            ref_match = re.search(r"(Ref\.?\s*[A-Za-z0-9\- ]{3,})", desc, re.I)
            if ref_match:
                referencia = ref_match.group(0).strip()
                desc = desc.replace(ref_match.group(0), '').strip()
            else:
                referencia = 'NULL'

            # map amounts into Cargos, Abonos, Operación, Liquidación using heuristic
            cargos = abonos = operacion = liquidacion = 'NULL'
            if len(amounts) >= 4:
                cargos, abonos, operacion, liquidacion = amounts[:4]
            elif len(amounts) == 3:
                cargos = amounts[0]
                abonos = 'NULL'
                operacion = amounts[1]
                liquidacion = amounts[2]
            elif len(amounts) == 2:
                operacion = amounts[0]
                liquidacion = amounts[1]
            elif len(amounts) == 1:
                cargos = amounts[0]

            tx_records.append({
                'Fecha de operación': fecha_op,
                'Fecha de liquidación': fecha_liq,
                'Descripción': desc if desc else 'NULL',
                'Referencia': referencia,
                'Cargos': cargos,
                'Abonos': abonos,
                'Operación': operacion,
                'Liquidación': liquidacion,
                'raw': entry,
            })

    # build table key-value records
    kv_records = []
    def extract_candidate_tokens(item):
        # Build a list of candidate tokens (percents, amounts, dates, or other trailing values)
        vals = item.get('values', []) or []
        tokens = []
        for v in vals:
            for p in v.split():
                if p.strip():
                    tokens.append(p.strip())

        # also check raw and label for date tokens (e.g., '01/JUN')
        raw = (item.get('raw') or '').strip()
        for d in day_re.findall(raw):
            if d not in tokens:
                tokens.append(d)

        label_text = (item.get('label') or '').strip()
        # if label contains a trailing ':' followed by a value, capture that value
        m = re.search(r":\s*(.+)$", label_text)
        clean_label = label_text
        if m:
            tail = m.group(1).strip()
            for p in tail.split():
                if p and p not in tokens:
                    tokens.append(p)
            clean_label = label_text[:m.start()].strip()

        return tokens, clean_label or label_text

    for tbl in tables:
        title = tbl.get('title', '')
        for item in tbl.get('items', []):
            orig_label = (item.get('label') or '').strip()
            tokens, label = extract_candidate_tokens(item)
            vals_joined = ''  # avoid joined pipe representation; use explicit Value/Percent columns

            # handle date ranges: if two date tokens present, emit two rows (Desde/Hasta)
            date_tokens = [t for t in tokens if day_re.search(t)]
            if len(date_tokens) >= 2:
                d1, d2 = date_tokens[0], date_tokens[1]
                kv_records.append({
                    'title': title,
                    'label': label + ' - Desde',
                    'Value': d1,
                    'Percent': '',
                    'values': '',
                    'raw': item.get('raw', ''),
                })
                kv_records.append({
                    'title': title,
                    'label': label + ' - Hasta',
                    'Value': d2,
                    'Percent': '',
                    'values': '',
                    'raw': item.get('raw', ''),
                })
                # remove these date tokens from further processing
                tokens = [t for t in tokens if t not in (d1, d2)]

            # if no tokens found, try to extract any numbers/dates from raw
            if not tokens:
                fallback = []
                for d in day_re.findall(item.get('raw', '')):
                    fallback.append(d)
                for a in num_re.findall(item.get('raw', '')):
                    fallback.append(a)
                tokens = fallback

            if not tokens:
                # still empty: produce single empty-value row
                kv_records.append({
                    'title': title,
                    'label': orig_label,
                    'Value': '',
                    'Percent': '',
                    'values': '',
                    'raw': item.get('raw', ''),
                })
                continue

            # Expand tokens into one record per logical pair.
            i = 0
            while i < len(tokens):
                t = tokens[i]
                # if token looks like a percent and there's a following token, pair them
                if '%' in t and i + 1 < len(tokens):
                    percent = t
                    value = tokens[i + 1]
                    kv_records.append({
                        'title': title,
                        'label': label,
                        'Value': value,
                        'Percent': percent,
                        'values': '',
                        'raw': item.get('raw', ''),
                    })
                    i += 2
                else:
                    # single value (could be date or amount)
                    kv_records.append({
                        'title': title,
                        'label': label,
                        'Value': t,
                        'Percent': '',
                        'values': '',
                        'raw': item.get('raw', ''),
                    })
                    i += 1

    # write to Excel
    OUTPUT_XLSX = os.path.join(os.getcwd(), 'parsed_output.xlsx')
    try:
        df_tx = pd.DataFrame(tx_records)
        df_kv = pd.DataFrame(kv_records)
        with pd.ExcelWriter(OUTPUT_XLSX, engine='openpyxl') as writer:
            if not df_tx.empty:
                df_tx.to_excel(writer, sheet_name='Transactions', index=False)
            if not df_kv.empty:
                df_kv.to_excel(writer, sheet_name='Table Key-Values', index=False)
        print(f"Wrote results to {OUTPUT_XLSX}")
    except ImportError:
        print('pandas or openpyxl not installed. Install requirements and retry.')
    except Exception as e:
        print('Error writing Excel:', e)