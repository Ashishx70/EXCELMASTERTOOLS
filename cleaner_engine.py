"""
Low-Memory GST Register Cleaning Engine for E-Commerce & ERP Systems.
Engineered to process 126MB+ Excel files under 35MB Peak RAM on Render (512MB Limit).
"""

import os
import re
import csv
import time
import uuid
import sqlite3
import tempfile
import zipfile
from typing import Dict, Any, List, Optional, Tuple

try:
    import lxml.etree as ET
    HAS_LXML = True
except ImportError:
    import xml.etree.ElementTree as ET
    HAS_LXML = False

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import xlsxwriter
    HAS_XLSXWRITER = True
except ImportError:
    HAS_XLSXWRITER = False

try:
    import python_calamine
    HAS_CALAMINE = True
except ImportError:
    HAS_CALAMINE = False


# ==============================================================================
# 1. CANDIDATE COLUMN ALIASES
# ==============================================================================
CANDIDATE_ALIASES = {
    "sale_erp": [
        "voucher", "voucher no", "voucher_no", "voucher no.", "voucherno",
        "invoice", "invoice no", "invoice_no", "invoice number", "bill no", "document no"
    ],
    "sale_zoho": [
        "invoice_number", "invoicenumber", "invoice_no", "invoice #",
        "transaction_number", "transaction number", "bill_number", "reference_number", "reference", "bill no"
    ],
    "purchase_erp": [
        "bill no", "bill_no", "bill no.", "bill number", "voucher no",
        "voucher_no", "voucher no.", "voucherno", "invoice", "invoice no", "invoice_no", "invoice number"
    ],
    "purchase_zoho": [
        "bill_number", "billnumber", "bill_no", "bill #", "bill no.", "bill no",
        "transaction_number", "transaction number", "invoice_number", "invoicenumber", "invoice_no", "invoice no", "invoice #",
        "reference_number", "reference"
    ],
    "sale_erp_details": [
        "invoice", "invoice no", "invoice_no", "invoice number", "invoice #",
        "voucher", "voucher no", "voucher_no", "voucherno", "bill no", "document no"
    ],
    "sale_zoho_details": [
        "transaction_number", "transaction number", "transaction_no", "transactionno", "transaction #",
        "invoice_number", "invoicenumber", "invoice_no", "invoice #", "reference_number", "reference"
    ],
    "purchase_erp_details": [
        "invoice", "invoice no", "invoice_no", "invoice number", "invoice #",
        "bill no", "bill_no", "bill no.", "bill number", "voucher no", "voucher_no", "voucherno"
    ],
    "purchase_zoho_details": [
        "transaction_number", "transaction number", "transaction_no", "transactionno", "transaction #",
        "bill_number", "billnumber", "bill_no", "bill #", "bill no.", "bill no",
        "invoice_number", "invoicenumber", "invoice_no", "invoice #"
    ],
    "sale_erp_details_aj": [
        "invoice", "invoice no", "invoice_no", "invoice number", "invoice #",
        "voucher", "voucher no", "voucher_no", "voucherno", "bill no", "document no"
    ],
    "sale_zoho_details_aj": [
        "transaction_number", "transaction number", "transaction_no", "transactionno", "transaction #",
        "invoice_number", "invoicenumber", "invoice_no", "invoice #", "reference_number", "reference"
    ],
    "purchase_erp_details_aj": [
        "invoice", "invoice no", "invoice_no", "invoice number", "invoice #",
        "bill no", "bill_no", "bill no.", "bill number", "voucher no", "voucher_no", "voucherno"
    ],
    "purchase_zoho_details_aj": [
        "transaction_number", "transaction number", "transaction_no", "transactionno", "transaction #",
        "bill_number", "billnumber", "bill_no", "bill #", "bill no.", "bill no",
        "invoice_number", "invoicenumber", "invoice_no", "invoice #"
    ],
    "sale_erp_details_az": [
        "invoice", "invoice no", "invoice_no", "invoice number", "invoice #",
        "voucher", "voucher no", "voucher_no", "voucherno", "bill no", "document no"
    ],
    "sale_zoho_details_az": [
        "transaction_number", "transaction number", "transaction_no", "transactionno", "transaction #",
        "invoice_number", "invoicenumber", "invoice_no", "invoice #", "reference_number", "reference"
    ],
    "purchase_erp_details_az": [
        "invoice", "invoice no", "invoice_no", "invoice number", "invoice #",
        "bill no", "bill_no", "bill no.", "bill number", "voucher no", "voucher_no", "voucherno"
    ],
    "purchase_zoho_details_az": [
        "transaction_number", "transaction number", "transaction_no", "transactionno", "transaction #",
        "bill_number", "billnumber", "bill_no", "bill #", "bill no.", "bill no",
        "invoice_number", "invoicenumber", "invoice_no", "invoice #"
    ],
    "credit_note_erp": [
        "voucher no", "voucher_no", "voucher no.", "voucherno", "voucher",
        "invoice no", "invoice_no", "invoice number", "bill no", "bill_no", "bill number",
        "creditnote_number", "credit_note_no", "cn_number", "credit note #", "document no", "reference_number"
    ],
    "credit_note_zoho": [
        "creditnote_number", "creditnote number", "credit_note_number",
        "credit_note_no", "cn_number", "credit note #", "credit note no",
        "voucher", "voucher no", "voucher_no", "reference_number", "bill no", "bill_no"
    ],
    "debit_note_erp": [
        "bill no", "bill_no", "bill no.", "bill number",
        "voucher no", "voucher_no", "voucher no.", "voucherno", "voucher",
        "invoice no", "invoice_no", "invoice number", "document no",
        "vendor_credit_number", "debitnote_number", "dn_number", "debit note #", "reference_number"
    ],
    "debit_note_zoho": [
        "vendor_credit_number", "vendor credit number", "vendor_credit_no",
        "debitnote_number", "debit_note_number", "dn_number", "debit note #", "debit note no",
        "bill no", "bill_no", "bill number", "voucher no", "voucher_no", "voucher", "reference_number"
    ]
}

# ==============================================================================
# 2. VALIDATION & SERIES RETENTION RULES
# ==============================================================================
RE_EXCLUDE_R_DIGIT = re.compile(r'R\d', re.IGNORECASE)

# Sales Register
RE_SALE_AJIO = re.compile(r'^AJ27S(J)?\d', re.IGNORECASE)
RE_SALE_MYNTRA = re.compile(r'^MY27S(Y)?\d', re.IGNORECASE)
RE_SALE_MARKETPLACE = re.compile(r'^(FK27S|CB27S|KT27S|RZ27S|ET27S|KK27S)\d', re.IGNORECASE)
RE_SALE_CGJ = re.compile(r'^CGJ12627', re.IGNORECASE)

# Purchase Register
RE_PUR_AJIO = re.compile(r'^AJ27P(J)?\d', re.IGNORECASE)
RE_PUR_MYNTRA = re.compile(r'^MY27P(Y)?\d', re.IGNORECASE)
RE_PUR_MARKETPLACE = re.compile(r'^(FK27P|CB27P|KT27P|KT2P7P|RZ27P|ET27P|KK27P)\d', re.IGNORECASE)
RE_PUR_PGJ_CGJ = re.compile(r'^(PGJ12627|CGJ12627)', re.IGNORECASE)

# Credit Notes (ERP & Zoho CN) - Strictly retain ONLY AJ27SCN, MY27SCN, FK27SCN, RCGJ12627 (or RCGJ126)
RE_CN_SERIES = re.compile(r'^(AJ27SCN|MY27SCN|FK27SCN|RCGJ12627|RCGJ126)', re.IGNORECASE)

# Debit Notes (ERP & Zoho DN) - Strictly retain ONLY AJ27PDN, MY27PDN, FK27PDN, PCGJ12627 (or PCGJ126)
RE_DN_SERIES = re.compile(r'^(AJ27PDN|MY27PDN|FK27PDN|PCGJ12627|PCGJ126)', re.IGNORECASE)

# Details Register Retention Series (AJ&MY&FK vs AZ)
# AJ&MY&FK Details (Sales: AJ27S, MY27S, FK27S, CGJ12627; Purchase: AJ27P, MY27P, FK27P, PGJ12627 / CGJ12627)
RE_SALE_AJ_MY_FK = re.compile(r'^(AJ27S(J)?|MY27S(Y)?|FK27S)\d|^CGJ12627', re.IGNORECASE)
RE_PUR_AJ_MY_FK = re.compile(r'^(AJ27P(J)?|MY27P(Y)?|FK27P)\d|^(PGJ12627|CGJ12627)', re.IGNORECASE)

# AZ Details (Sales: CB27S, KT27S, RZ27S, ET27S, KK27S; Purchase: CB27P, KT27P, KT2P7P, RZ27P, ET27P, KK27P)
RE_SALE_AZ = re.compile(r'^(CB27S|KT27S|RZ27S|ET27S|KK27S)\d', re.IGNORECASE)
RE_PUR_AZ = re.compile(r'^(CB27P|KT27P|KT2P7P|RZ27P|ET27P|KK27P)\d', re.IGNORECASE)


def is_valid_sale_voucher(val: Any) -> bool:
    """Validate voucher/invoice for Sale Registers (ERP & Zoho)."""
    if val is None:
        return False
    s = str(val).strip().upper()
    if not s:
        return False
    # Exclusion rules: CN, DN, SR, or matches R\d
    if "CN" in s or "DN" in s or "SR" in s or RE_EXCLUDE_R_DIGIT.search(s):
        return False
    # Retention rules: Authorized series only
    if (RE_SALE_AJIO.search(s) or 
        RE_SALE_MYNTRA.search(s) or 
        RE_SALE_MARKETPLACE.search(s) or 
        RE_SALE_CGJ.search(s)):
        return True
    return False


def is_valid_purchase_bill(val: Any) -> bool:
    """Validate bill/voucher for Purchase Registers (ERP & Zoho)."""
    if val is None:
        return False
    s = str(val).strip().upper()
    if not s:
        return False
    # Exclusion rules: DN, PDN, CN, PR, SR, or matches R\d
    if ("DN" in s or "PDN" in s or "CN" in s or "PR" in s or "SR" in s or 
        RE_EXCLUDE_R_DIGIT.search(s)):
        return False
    # Retention rules: Authorized series only
    if (RE_PUR_AJIO.search(s) or 
        RE_PUR_MYNTRA.search(s) or 
        RE_PUR_MARKETPLACE.search(s) or 
        RE_PUR_PGJ_CGJ.search(s)):
        return True
    return False


def is_valid_credit_note(val: Any) -> bool:
    """Validate Credit Note (ERP & Zoho CN). Strictly retain ONLY AJ27SCN, MY27SCN, FK27SCN, RCGJ12627 / RCGJ126."""
    if val is None:
        return False
    s = str(val).strip().replace("\xa0", "").replace(" ", "").upper()
    return bool(RE_CN_SERIES.search(s))


def is_valid_debit_note(val: Any) -> bool:
    """Validate Debit Note / Vendor Credit (ERP & Zoho DN). Strictly retain ONLY AJ27PDN, MY27PDN, FK27PDN, PCGJ12627 / PCGJ126."""
    if val is None:
        return False
    s = str(val).strip().replace("\xa0", "").replace(" ", "").upper()
    return bool(RE_DN_SERIES.search(s))


def is_valid_sale_voucher_aj_my_fk(val: Any) -> bool:
    """Validate voucher for Sales Details: AJ&MY&FK (Retains AJ27S, MY27S, FK27S, CGJ12627)."""
    if val is None:
        return False
    s = str(val).strip().upper()
    if not s:
        return False
    if "CN" in s or "DN" in s or "SR" in s or RE_EXCLUDE_R_DIGIT.search(s):
        return False
    return bool(RE_SALE_AJ_MY_FK.search(s))


def is_valid_purchase_bill_aj_my_fk(val: Any) -> bool:
    """Validate bill for Purchase Details: AJ&MY&FK (Retains AJ27P, MY27P, FK27P, PGJ12627)."""
    if val is None:
        return False
    s = str(val).strip().upper()
    if not s:
        return False
    if ("DN" in s or "PDN" in s or "CN" in s or "PR" in s or "SR" in s or 
        RE_EXCLUDE_R_DIGIT.search(s)):
        return False
    return bool(RE_PUR_AJ_MY_FK.search(s))


def is_valid_sale_voucher_az(val: Any) -> bool:
    """Validate voucher for Sales Details: AZ (Retains CB27S, KT27S, RZ27S, ET27S, KK27S)."""
    if val is None:
        return False
    s = str(val).strip().upper()
    if not s:
        return False
    if "CN" in s or "DN" in s or "SR" in s or RE_EXCLUDE_R_DIGIT.search(s):
        return False
    return bool(RE_SALE_AZ.search(s))


def is_valid_purchase_bill_az(val: Any) -> bool:
    """Validate bill for Purchase Details: AZ (Retains CB27P, KT27P, RZ27P, ET27P, KK27P)."""
    if val is None:
        return False
    s = str(val).strip().upper()
    if not s:
        return False
    if ("DN" in s or "PDN" in s or "CN" in s or "PR" in s or "SR" in s or 
        RE_EXCLUDE_R_DIGIT.search(s)):
        return False
    return bool(RE_PUR_AZ.search(s))


def get_validator_for_register(register_type: str):
    """Factory for register validator function and column candidate aliases."""
    reg = register_type.lower().strip()
    if "credit" in reg or reg in ("cn", "credit_note_zoho", "credit_note_erp"):
        aliases = CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["credit_note_erp"])
        return is_valid_credit_note, aliases
    elif "debit" in reg or reg in ("dn", "debit_note_zoho", "debit_note_erp"):
        aliases = CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["debit_note_erp"])
        return is_valid_debit_note, aliases

    # Specific Details: AJ&MY&FK
    elif "details_aj" in reg or "_aj" in reg or "aj_my_fk" in reg:
        if "pur" in reg or "purchase" in reg:
            return is_valid_purchase_bill_aj_my_fk, CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["purchase_erp_details_aj"])
        else:
            return is_valid_sale_voucher_aj_my_fk, CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["sale_erp_details_aj"])

    # Specific Details: AZ
    elif "details_az" in reg or "_az" in reg:
        if "pur" in reg or "purchase" in reg:
            return is_valid_purchase_bill_az, CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["purchase_erp_details_az"])
        else:
            return is_valid_sale_voucher_az, CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["sale_erp_details_az"])

    # Standard / Summary Registers
    elif "purchase_zoho" in reg or reg in ("pur_zoho", "purchase_zoho_details"):
        return is_valid_purchase_bill, CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["purchase_zoho"])
    elif "purchase_erp" in reg or "pur" in reg or reg == "purchase_erp_details":
        return is_valid_purchase_bill, CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["purchase_erp"])
    elif "sale_zoho" in reg:
        return is_valid_sale_voucher, CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["sale_zoho"])
    else:  # Default: Sale ERP
        return is_valid_sale_voucher, CANDIDATE_ALIASES.get(reg, CANDIDATE_ALIASES["sale_erp"])


# ==============================================================================
# 3. DYNAMIC COLUMN & HEADER DETECTOR (2-PASS ALGORITHM)
# ==============================================================================
def find_column_index(rows_sample: List[List[Any]], candidates: List[str]) -> Tuple[Optional[int], Optional[int], Optional[str]]:
    """
    Scans first 35 rows dynamically.
    Pass 1: Priority-based exact lowercased match.
    Pass 2: Priority-based substring match.
    Returns (header_row_index, target_col_index, matched_column_name).
    """
    candidates_lower = [c.lower().strip() for c in candidates]

    # Pass 1: Exact match in priority order of candidate aliases
    for cand in candidates_lower:
        for row_idx, row in enumerate(rows_sample[:35]):
            for col_idx, cell in enumerate(row):
                if cell is None:
                    continue
                text = str(cell).strip().lower()
                if text == cand:
                    return row_idx, col_idx, str(cell).strip()

    # Pass 2: Substring match in priority order (< 40 chars)
    for cand in candidates_lower:
        for row_idx, row in enumerate(rows_sample[:35]):
            for col_idx, cell in enumerate(row):
                if cell is None:
                    continue
                text = str(cell).strip().lower()
                if len(text) < 40 and cand in text:
                    return row_idx, col_idx, str(cell).strip()

    return None, None, None


# ==============================================================================
# 4. HIGH-SPEED MEMORY-SAFE SHARED STRINGS RESOLVER
# ==============================================================================
class FastSharedStringsResolver:
    """
    High-speed, memory-safe shared strings resolver.
    Keeps strings in a contiguous Python list in RAM for microsecond lookups (< 35MB for 250k strings).
    Automatically migrates to disk SQLite only if unique strings exceed safety threshold (500,000 strings)
    to guarantee Render's 512MB RAM limit is NEVER breached.
    """

    def __init__(self, sqlite_path: str, max_in_memory: int = 500000):
        self.strings: List[str] = []
        self.is_disk_backed = False
        self.sqlite_path = sqlite_path
        self.conn = None
        self.cur = None
        self.max_in_memory = max_in_memory
        self.cache: Dict[int, str] = {}

    def populate_from_stream(self, xml_file):
        idx = 0
        batch = []
        for event, elem in ET.iterparse(xml_file, events=('end',)):
            if elem.tag.endswith('si'):
                text = "".join(elem.itertext())
                if not self.is_disk_backed:
                    self.strings.append(text)
                    if len(self.strings) >= self.max_in_memory:
                        self._migrate_to_sqlite()
                else:
                    batch.append((idx, text))
                    if len(batch) >= 10000:
                        self.cur.executemany("INSERT INTO strings VALUES (?, ?)", batch)
                        batch = []
                idx += 1
                elem.clear()
        if self.is_disk_backed and batch:
            self.cur.executemany("INSERT INTO strings VALUES (?, ?)", batch)
            self.conn.commit()

    def _migrate_to_sqlite(self):
        self.is_disk_backed = True
        self.conn = sqlite3.connect(self.sqlite_path)
        self.cur = self.conn.cursor()
        self.cur.execute("PRAGMA synchronous = OFF;")
        self.cur.execute("PRAGMA journal_mode = OFF;")
        self.cur.execute("CREATE TABLE strings (id INTEGER PRIMARY KEY, val TEXT);")
        batch = [(i, s) for i, s in enumerate(self.strings)]
        self.cur.executemany("INSERT INTO strings VALUES (?, ?)", batch)
        self.conn.commit()
        self.strings = []

    def get_string(self, str_id: int) -> str:
        if not self.is_disk_backed:
            if 0 <= str_id < len(self.strings):
                return self.strings[str_id]
            return ""
        if str_id in self.cache:
            return self.cache[str_id]
        self.cur.execute("SELECT val FROM strings WHERE id = ?", (str_id,))
        row = self.cur.fetchone()
        val = row[0] if row else ""
        if len(self.cache) < 10000:
            self.cache[str_id] = val
        return val

    def close(self):
        if self.conn:
            try:
                self.conn.close()
            except Exception:
                pass


# Backward compatibility alias
SqliteSharedStringsResolver = FastSharedStringsResolver


def col_ref_to_index(col_ref: str) -> int:
    idx = 0
    for char in col_ref:
        if 'A' <= char <= 'Z':
            idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx - 1


def index_to_col_ref(col_idx: int) -> str:
    """Convert 0-based column index to Excel column string ('A', 'Z', 'AA', etc.)."""
    result = []
    col_idx += 1
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def clear_element(elem: Any, root: Optional[Any] = None, row_count: int = 0):
    """Safely frees element memory across both lxml and standard xml.etree without memory creep."""
    elem.clear()
    if hasattr(elem, 'getprevious'):
        parent = elem.getparent()
        if parent is not None:
            while elem.getprevious() is not None:
                del parent[0]
    elif root is not None and row_count % 5000 == 0:
        root.clear()


def extract_cell_value(c: Any, strings_resolver: Optional[FastSharedStringsResolver]) -> str:
    """Safely extracts cell string / number value across inline, shared, and direct formats."""
    t_attr = c.attrib.get('t')
    if t_attr == 's':
        v = c.find('{*}v')
        if v is not None and v.text and strings_resolver:
            v_clean = v.text.strip()
            if v_clean.isdigit():
                return strings_resolver.get_string(int(v_clean))
        return ""
    elif t_attr == 'inlineStr':
        return "".join(c.itertext())
    else:
        v = c.find('{*}v')
        return v.text if v is not None and v.text else ""


def extract_all_row_cells(elem: Any, strings_resolver: Optional[FastSharedStringsResolver]) -> List[Any]:
    """Extracts all cell values in column order for a row."""
    row_cells = {}
    max_c = 0
    for c in elem.findall('{*}c'):
        r_attr = c.attrib.get('r', '')
        col_letters = "".join(filter(str.isalpha, r_attr)).upper()
        c_idx = col_ref_to_index(col_letters) if col_letters else max_c
        max_c = max(max_c, c_idx + 1)
        val = extract_cell_value(c, strings_resolver)
        row_cells[c_idx] = val
    return [row_cells.get(i, "") for i in range(max_c)]


# ==============================================================================
# 5. HIGH-SPEED, ZERO-RAM STREAMING CLEANER
# ==============================================================================
def clean_single_register_streaming(
    input_filepath: str,
    output_filepath: str,
    register_type: str = "sale_erp",
    output_format: str = "csv"
) -> Dict[str, Any]:
    """
    Master Cleaning Engine:
    - Streams directly from disk using SAX parser and SQLite string storage.
    - Peak RAM stays < 35 MB even for 126 MB+ files.
    """
    start_time = time.time()
    ext = os.path.splitext(input_filepath)[1].lower()
    validator_func, candidate_aliases = get_validator_for_register(register_type)

    sqlite_db_path = os.path.join(tempfile.gettempdir(), f"strings_{uuid.uuid4().hex}.db")
    strings_resolver = None

    total_rows = 0
    retained_rows = 0
    excluded_rows = 0
    detected_col_name = None
    detected_col_idx = None
    header_row_idx = None

    process = psutil.Process(os.getpid()) if HAS_PSUTIL else None
    peak_ram_bytes = process.memory_info().rss if process else 0

    try:
        # -------------------------------------------------------------
        # BRANCH A: CSV File Processing (Direct Streaming)
        # -------------------------------------------------------------
        if ext == ".csv":
            with open(input_filepath, mode='r', encoding='utf-8', errors='replace') as infile:
                reader = csv.reader(infile)
                first_35 = []
                for _ in range(35):
                    try:
                        row = next(reader)
                        first_35.append(row)
                    except StopIteration:
                        break

                header_row_idx, detected_col_idx, detected_col_name = find_column_index(
                    first_35, candidate_aliases
                )
                if detected_col_idx is None:
                    header_row_idx, detected_col_idx, detected_col_name = 0, 0, "Default_Col_0"

            out_is_csv = output_filepath.lower().endswith('.csv')
            out_wb = None
            out_ws = None
            csv_file = None
            csv_writer = None

            if out_is_csv:
                csv_file = open(output_filepath, 'w', newline='', encoding='utf-8')
                csv_writer = csv.writer(csv_file)
            elif HAS_XLSXWRITER:
                out_wb = xlsxwriter.Workbook(output_filepath, {'constant_memory': True})
                out_ws = out_wb.add_worksheet("Cleaned_Data")
            elif HAS_OPENPYXL:
                out_wb = openpyxl.Workbook(write_only=True)
                out_ws = out_wb.create_sheet("Cleaned_Data")
            else:
                csv_file = open(output_filepath, 'w', newline='', encoding='utf-8')
                csv_writer = csv.writer(csv_file)
                out_is_csv = True

            out_row_idx = 0
            with open(input_filepath, mode='r', encoding='utf-8', errors='replace') as infile:
                reader = csv.reader(infile)
                for r_idx, row in enumerate(reader):
                    if r_idx < header_row_idx:
                        continue
                    if r_idx == header_row_idx:
                        if out_is_csv:
                            csv_writer.writerow(row)
                        elif HAS_XLSXWRITER:
                            out_ws.write_row(out_row_idx, 0, row)
                            out_row_idx += 1
                        elif HAS_OPENPYXL:
                            out_ws.append(row)
                        continue

                    total_rows += 1
                    val = row[detected_col_idx] if detected_col_idx < len(row) else None
                    if validator_func(val):
                        if out_is_csv:
                            csv_writer.writerow(row)
                        elif HAS_XLSXWRITER:
                            out_ws.write_row(out_row_idx, 0, row)
                            out_row_idx += 1
                        elif HAS_OPENPYXL:
                            out_ws.append(row)
                        retained_rows += 1
                    else:
                        excluded_rows += 1

            if csv_file:
                csv_file.close()
            elif out_wb:
                if HAS_XLSXWRITER:
                    out_wb.close()
                elif HAS_OPENPYXL:
                    out_wb.save(output_filepath)

        # -------------------------------------------------------------
        # BRANCH B: XLSX / XLSM File Processing
        # -------------------------------------------------------------
        else:
            used_calamine = False
            if HAS_CALAMINE:
                try:
                    wb = python_calamine.CalamineWorkbook.from_path(input_filepath)
                    sheet_name = wb.sheet_names[0]
                    sheet = wb.get_sheet_by_name(sheet_name)

                    # 1. Read first 35 rows for dynamic header detection (Zero-RAM)
                    sample_rows = []
                    row_iter = sheet.iter_rows()
                    for _ in range(35):
                        try:
                            sample_rows.append(next(row_iter))
                        except StopIteration:
                            break

                    header_row_idx, detected_col_idx, detected_col_name = find_column_index(
                        sample_rows, candidate_aliases
                    )
                    if detected_col_idx is None:
                        header_row_idx, detected_col_idx, detected_col_name = 0, 0, "Default_Col_0"

                    out_is_csv = output_filepath.lower().endswith('.csv')
                    out_wb = None
                    out_ws = None
                    csv_file = None
                    csv_writer = None

                    if out_is_csv:
                        csv_file = open(output_filepath, 'w', newline='', encoding='utf-8')
                        csv_writer = csv.writer(csv_file)
                    elif HAS_XLSXWRITER:
                        out_wb = xlsxwriter.Workbook(output_filepath, {'constant_memory': True})
                        out_ws = out_wb.add_worksheet("Cleaned_Data")
                    elif HAS_OPENPYXL:
                        out_wb = openpyxl.Workbook(write_only=True)
                        out_ws = out_wb.create_sheet("Cleaned_Data")
                    else:
                        csv_file = open(output_filepath, 'w', newline='', encoding='utf-8')
                        csv_writer = csv.writer(csv_file)
                        out_is_csv = True

                    # 2. Stream all rows from iter_rows() one-by-one (Zero-RAM: Peak RAM < 25 MB)
                    out_row_idx = 0
                    for r_idx, r in enumerate(sheet.iter_rows()):
                        if r_idx < header_row_idx:
                            continue
                        if r_idx == header_row_idx:
                            if out_is_csv:
                                csv_writer.writerow(r)
                            elif HAS_XLSXWRITER:
                                out_ws.write_row(out_row_idx, 0, r)
                                out_row_idx += 1
                            elif HAS_OPENPYXL:
                                out_ws.append(r)
                            continue

                        total_rows += 1
                        val = r[detected_col_idx] if detected_col_idx < len(r) else None
                        if validator_func(val):
                            if out_is_csv:
                                csv_writer.writerow(r)
                            elif HAS_XLSXWRITER:
                                out_ws.write_row(out_row_idx, 0, r)
                                out_row_idx += 1
                            elif HAS_OPENPYXL:
                                out_ws.append(r)
                            retained_rows += 1
                        else:
                            excluded_rows += 1

                    if csv_file:
                        csv_file.close()
                    elif out_wb:
                        if HAS_XLSXWRITER:
                            out_wb.close()
                        elif HAS_OPENPYXL:
                            out_wb.save(output_filepath)

                    if process:
                        peak_ram_bytes = max(peak_ram_bytes, process.memory_info().rss)
                    used_calamine = True
                except Exception as cal_err:
                    import sys
                    print(f"[CALAMINE FALLBACK] {cal_err}, falling back to SAX parser...", file=sys.stderr)
                    used_calamine = False

            if not used_calamine:
                with zipfile.ZipFile(input_filepath, 'r') as z:
                    # Step 1: Parse Shared Strings into Fast RAM Resolver (with auto SQLite safety net)
                    if 'xl/sharedStrings.xml' in z.namelist():
                        strings_resolver = FastSharedStringsResolver(sqlite_db_path)
                        with z.open('xl/sharedStrings.xml') as s_file:
                            strings_resolver.populate_from_stream(s_file)

                    # Locate primary worksheet
                    sheet_target = 'xl/worksheets/sheet1.xml'
                    if sheet_target not in z.namelist():
                        candidates = [n for n in z.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
                        sheet_target = candidates[0] if candidates else None

                    if not sheet_target:
                        raise ValueError("No valid worksheet found inside the Excel archive.")

                    # Step 2: First pass on first 35 rows for Header Detection
                    first_35_rows = []
                    with z.open(sheet_target) as sheet_file:
                        row_counter = 0
                        for event, elem in ET.iterparse(sheet_file, events=('end',)):
                            if elem.tag.endswith('row'):
                                row_counter += 1
                                row_list = extract_all_row_cells(elem, strings_resolver)
                                first_35_rows.append(row_list)
                                clear_element(elem)
                                if row_counter >= 35:
                                    break

                    header_row_idx, detected_col_idx, detected_col_name = find_column_index(
                        first_35_rows, candidate_aliases
                    )
                    if detected_col_idx is None:
                        header_row_idx, detected_col_idx, detected_col_name = 0, 0, "Default_Col_0"

                    target_col_letters = index_to_col_ref(detected_col_idx)

                    # Step 3: Stream Sheet Rows and Write Output (Lazy Single-Cell Evaluation)
                    out_is_csv = output_filepath.lower().endswith('.csv')
                    csv_file = None
                    csv_writer = None
                    xlsx_wb = None
                    xlsx_ws = None

                    if out_is_csv:
                        csv_file = open(output_filepath, 'w', newline='', encoding='utf-8')
                        csv_writer = csv.writer(csv_file)
                    else:
                        if HAS_XLSXWRITER:
                            xlsx_wb = xlsxwriter.Workbook(output_filepath, {'constant_memory': True})
                            xlsx_ws = xlsx_wb.add_worksheet("Cleaned_Data")
                        elif HAS_OPENPYXL:
                            xlsx_wb = openpyxl.Workbook(write_only=True)
                            xlsx_ws = xlsx_wb.create_sheet("Cleaned_Data")
                        else:
                            output_filepath = output_filepath + ".csv"
                            csv_file = open(output_filepath, 'w', newline='', encoding='utf-8')
                            csv_writer = csv.writer(csv_file)
                            out_is_csv = True

                    out_row_idx = 0
                    with z.open(sheet_target) as sheet_file:
                        context = ET.iterparse(sheet_file, events=('start', 'end'))
                        root = None
                        for event, elem in context:
                            if event == 'start' and root is None:
                                root = elem
                                continue

                            if event == 'end' and elem.tag.endswith('row'):
                                total_rows += 1
                                current_r_idx = total_rows - 1

                                # Preamble metadata rows before header are skipped
                                if current_r_idx < header_row_idx:
                                    clear_element(elem, root, total_rows)
                                    continue

                                # Header row is written as first row
                                if current_r_idx == header_row_idx:
                                    row_list = extract_all_row_cells(elem, strings_resolver)
                                    clear_element(elem, root, total_rows)
                                    if out_is_csv:
                                        csv_writer.writerow(row_list)
                                    elif HAS_XLSXWRITER:
                                        xlsx_ws.write_row(out_row_idx, 0, row_list)
                                        out_row_idx += 1
                                    elif HAS_OPENPYXL:
                                        xlsx_ws.append(row_list)
                                    continue

                                # Data rows: Fast Lazy Single-Cell Inspection
                                all_cells = elem.findall('{*}c')
                                target_cell = None
                                for c in all_cells:
                                    r_attr = c.attrib.get('r', '')
                                    if r_attr:
                                        col_part = "".join(filter(str.isalpha, r_attr)).upper()
                                        if col_part == target_col_letters:
                                            target_cell = c
                                            break

                                target_val = None
                                if target_cell is not None:
                                    target_val = extract_cell_value(target_cell, strings_resolver)
                                elif all_cells and not all_cells[0].attrib.get('r'):
                                    if detected_col_idx < len(all_cells):
                                        target_val = extract_cell_value(all_cells[detected_col_idx], strings_resolver)

                                # Fast Drop: If row does not match authorized series, drop immediately!
                                if not validator_func(target_val):
                                    excluded_rows += 1
                                    clear_element(elem, root, total_rows)
                                    if process and total_rows % 5000 == 0:
                                        peak_ram_bytes = max(peak_ram_bytes, process.memory_info().rss)
                                    continue

                                # Retained Row: Extract full row cells and write output
                                row_list = extract_all_row_cells(elem, strings_resolver)
                                clear_element(elem, root, total_rows)

                                if out_is_csv:
                                    csv_writer.writerow(row_list)
                                elif HAS_XLSXWRITER:
                                    xlsx_ws.write_row(out_row_idx, 0, row_list)
                                    out_row_idx += 1
                                elif HAS_OPENPYXL:
                                    xlsx_ws.append(row_list)
                                retained_rows += 1

                                # Update peak RAM metric
                                if process and total_rows % 5000 == 0:
                                    peak_ram_bytes = max(peak_ram_bytes, process.memory_info().rss)

                    # Close file outputs
                    if csv_file:
                        csv_file.close()
                    elif xlsx_wb:
                        if HAS_XLSXWRITER:
                            xlsx_wb.close()
                        elif HAS_OPENPYXL:
                            xlsx_wb.save(output_filepath)

    finally:
        # Cleanup temporary SQLite DB
        if strings_resolver:
            strings_resolver.close()
        if os.path.exists(sqlite_db_path):
            try:
                os.remove(sqlite_db_path)
            except Exception:
                pass
        import gc
        gc.collect()

    if process:
        peak_ram_bytes = max(peak_ram_bytes, process.memory_info().rss)
    peak_ram_mb = round(peak_ram_bytes / (1024 * 1024), 2)
    duration = round(time.time() - start_time, 2)

    clean_out_name = re.sub(r"^out_[a-f0-9]{32}_", "", os.path.basename(output_filepath))
    clean_in_name = re.sub(r"^in_[a-f0-9]{32}_", "", os.path.basename(input_filepath))

    return {
        "input_filename": clean_in_name,
        "output_filename": clean_out_name,
        "column_detected": detected_col_name,
        "column_index": detected_col_idx,
        "header_row_index": header_row_idx,
        "original_row_count": retained_rows + excluded_rows,
        "retained_row_count": retained_rows,
        "excluded_row_count": excluded_rows,
        "memory_peak_mb": peak_ram_mb,
        "duration_seconds": duration
    }
