import os
import re
import sys

sys.modules['numexpr'] = None
sys.modules['bottleneck'] = None

import pandas as pd
import os
from dotenv import load_dotenv
from sqlalchemy import create_engine, text


from sqlalchemy import create_engine


INPUT_FILE  = os.path.join("data", "facul.xlsx")
OUTPUT_FILE = os.path.join("data", "facul_clean_aca.xlsx")

CEDANT_COL   = "COMP_NAME"
CEDANT_VALUE = "PT. ASURANSI CENTRAL ASIA"

POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"
BROKER_COL  = "COMP_NAME2"

MAX_SPLIT_COLS = 5

POLIS_EXCEPTION_RE = re.compile(
    r"""
    MOP\s*MARINE
  | (?:LINE\s*SLIP|LINESLIP)
  | \b(?:P1|P2|P3|P73)\s*CANCEL
  | \b(?:P1|P2|P3|P73)\b
  | \bCANCEL\b
  | PENYELESAIAN
  | HUTANG\s*PIUTANG
    """,
    re.IGNORECASE | re.VERBOSE,
)

SLIP_EXCEPTION_RE = re.compile(
    r"""
    \bSUMMARY\b
  | \bBORDER[OA]\b
  | \bBORDRO\b
  | \bSINGGLESHIPMENT\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC", "PTE",
    "AND", "OR", "THE", "OF", "AS",
    "NON FOOD", "DIV",
    "BAPAK", "BPK", "IBU", "NYONYA", "NY", "MR", "MRS", "MS", "SDR", "SAUDARA", "SAUDARI",
})

INSURED_SPLIT_RE = re.compile(
    r"""
      \bAND\s*/\s*OR\b
    | \bC\s*/\s*Q\b
    | \bQ\s*/\s*Q\b
    | \bQ\.?Q\.?\b
    | (?<=[A-Z0-9])QQ(?=[A-Z0-9\s]|$)
    | (?<=\s)QQ(?=[A-Z0-9]|$)
    | ,(?![^(]*\))
    | /(?![^(]*\))
    """,
    re.IGNORECASE | re.VERBOSE,
)

INSURED_PREFIX_SUFFIX_RE = re.compile(
    r"""
    ^\s*(?:PT\.?|CV\.?|TBK\.?|\(PERSERO\)|\bPERSERO\b|LTD\.?|INC\.?|LLC\.?|UD\.?|PD\.?|NV\.?|BV\.?|GMBH\.?|SDN\s+BHD|BHD\.?)\s*
  | \s*(?:,?\s*\bTBK\b\s*(?:,?\s*PT\.?)?|,?\s*\bPT\.?\s*$|,?\s*\bCV\.?\s*$|\bPT\.?\b\s*$|\bCV\.?\b\s*$|,?\s*\bLTD\.?\s*$)
  | ^\s*\((?:PERSERO|FCI\.?\s*I|TBK)\)\s*
  | ^\s*(?:[A-Z]\.){1,}[A-Z]?\s*
  | \s*,?\s*\b(?:S\.?KOM|S\.?E|S\.?T|S\.?H|S\.?SI|S\.?SOS|S\.?IP|S\.?TP|S\.?PSI|S\.?KED|M\.?M|M\.?B\.?A|M\.?SI|M\.?T|M\.?KN|M\.?H|DRS?|DRA?|IR|PROF|PH\.?D)\.?\b\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

INSURED_ENTITY_ANYWHERE_RE = re.compile(
    r"(?i)(?:\b|\.|\,)\s*(?:PT|CV|TBK|PERSERO|\(PERSERO\)|LTD|PTE(?:\s+LTD)?|INC|LLC|UD|PD|NV|BV|GMBH|SDN\s+BHD|BHD)\b\.?\s*"
)

HONORIFICS_TITLES_RE = re.compile(
    r"""
    (?i)\b(?:
        BAPAK|BPK|IBU|NYONYA|NY|MR|MRS|MS|SDR|SAUDARA|SAUDARI
      | S\.?KOM|S\.?E|S\.?T|S\.?H|S\.?SI|S\.?SOS|S\.?IP|S\.?TP|S\.?PSI|S\.?KED|M\.?M|M\.?B\.?A|M\.?SI|M\.?T|M\.?KN|M\.?H|DRS?|DRA?|IR|PROF|PH\.?D|HJ?
    )\b\.?\s*
    """,
    re.VERBOSE,
)

_SLIP_NOISE_RE = re.compile(
    r"""
    \b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS
        |SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b
  | \b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST
        |SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b
  | \b(?:IDR|USD|ENG)\b
  | \b(?:P1|P2|P3|P73)\b
  | \bNEW\b
  | \bVARIOUS\b
  | \b20[0-9]{2}\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

_SLIP_TOKEN_RE = re.compile(r"[A-Z0-9][A-Z0-9\-]{6,}", re.IGNORECASE)

# Regex untuk mengekstrak nomor sertifikat dari polis.
_CERT_FROM_POLIS_RE = re.compile(r'[-\s]+\s*(\d{1,6})(?:[^0-9]|$)')

# Regex untuk range S/D pada sertifikat: '{1-6digit}S/D{1-6digit}'
_CERT_SD_RE = re.compile(r'(\d{1,6})\s*S/D\s*(\d{1,6})', re.IGNORECASE)

# Sertifikat valid: 1-6 digit pure numeric
_SERTIF_DIGIT_RE = re.compile(r'^\d{1,6}$')

# Max kolom sertifikat
_MAX_SERTIF_COLS = 5

# =============================================================================
# HELPERS
# =============================================================================

def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


def extract_cert_from_polis(val) -> list:
    """Ekstrak nomor sertifikat dari polis_ori di FACUL.

    Format 1 (suffix): '{base_polis} - {1-6digit_sertif}'
    Format 2 (S/D range): '{base_polis} - {1-6digit}S/D{1-6digit}'

    Aturan:
    - Cari suffix digit (1-6) setelah tanda '-' atau '- '
    - Jika format S/D: breakdown range, maksimal 5 nilai
    - Zero-pad setiap nilai ke 6 digit

    Contoh:
        '100030825120000155-001007' -> ['001007']
        '7283738299277344 - 000211' -> ['000211']
        '100030825120000100S/D000104' -> ['000100','000101','000102','000103','000104']
        '100030825120000155'        -> []  (tidak ada sertif)
    """
    if pd.isna(val):
        return []
    s = str(val).strip()
    if not s:
        return []

    # Cari pola S/D terlebih dahulu
    m_sd = _CERT_SD_RE.search(s)
    if m_sd:
        start = int(m_sd.group(1))
        end   = int(m_sd.group(2))
        if start > end:
            start, end = end, start
        certs = [str(i).zfill(6) for i in range(start, end + 1)]
        return certs[:_MAX_SERTIF_COLS]

    # Cari suffix digit (1-6) setelah tanda '-'
    m = _CERT_FROM_POLIS_RE.search(s)
    if m:
        raw = m.group(1)
        if _SERTIF_DIGIT_RE.match(raw):
            return [raw.zfill(6)]

    return []


def _is_valid_polis_token(tok: str) -> bool:
    """True if token is ≥5 chars, contains a digit, and does not end with TBA."""
    tok = tok.strip()
    return (
        len(tok) >= 5
        and not re.search(r"TBA$", tok, re.IGNORECASE)
        and bool(re.search(r"\d", tok))
    )


def _is_valid_slip_token(tok: str) -> bool:
    """True if token is ≥7 chars, contains a digit, and is not a 4-digit year."""
    tok = tok.strip()
    return (
        len(tok) >= 7
        and bool(re.search(r"\d", tok))
        and not re.match(r"^\d{4}$", tok)
    )


def _strip_polis_base(tok: str) -> str:
    """Strip short numeric suffixes after dash to return the base policy number.

    Examples:
      '131030817120000016 - 000149'      → '131030817120000016'
      '21001032011000063-066-517-552-418' → '21001032011000063'
      '210010421100000031-1/1'            → '210010421100000031'
    """
    tok = re.sub(r"\s*-\s*", "-", tok.strip())
    parts = tok.split("-")
    base = parts[0].strip()
    if len(parts) >= 2 and len(re.sub(r"\D", "", base)) >= 10:
        if all(len(re.sub(r"\D", "", s)) <= 6 for s in parts[1:]):
            return base
    m = re.match(r"^(.+?)-(\d{1,6})$", tok)
    if m and len(m.group(1)) > len(m.group(2)):
        return m.group(1)
    return tok


def _strip_slip_suffix(tok: str) -> str:
    """Strip short numeric suffix after dash (e.g. -03, -000059)."""
    m = re.match(r"^(.+?)-(\d{1,6})$", tok)
    if m and len(m.group(1)) > len(m.group(2)):
        return m.group(1)
    return tok


def _extract_polis_tokens(text: str) -> list:
    """Extract valid policy number tokens from a text string."""
    tokens = []
    for block in re.split(r"\s{2,}", text.strip()):
        for tok in re.split(r"\+|\s+", block.strip()):
            tok = tok.strip().strip("-/")
            if _is_valid_polis_token(tok):
                if len(re.sub(r"\D", "", tok)) >= 10:
                    tok = _strip_polis_base(tok)
                tokens.append(tok)
    return tokens


def _expand_plus_suffix(text: str) -> list | None:
    """Expand BASE + s1 + s2 + ... format using trailing-digit suffix-replace.

    Rules:
    - ``text`` must be a single block (no double-space gaps).
    - BASE (first token) must be pure digits, length >= 10.
    - Each segment s_i must be pure digits and shorter than BASE.
    - s_i replaces the **last len(s_i) digits** of BASE.
    - Returns [BASE, expanded_s1, ...] in original order.
    - Returns None if format is invalid → caller should fallback to existing logic.

    Examples::

        _expand_plus_suffix("100010324110001373 + 362 + 384+64+75+166+166")
        # → ["100010324110001373", "100010324110001362", "100010324110001384",
        #     "100010324110001364", "100010324110001375", "100010324110001166",
        #     "100010324110001166"]

        _expand_plus_suffix("100090325110000155+188+202")
        # → ["100090325110000155", "100090325110000188", "100090325110000202"]
    """
    # Strip surrounding whitespace and split by '+'
    parts = [p.strip() for p in text.split("+")]
    if len(parts) < 2:
        return None

    base = parts[0].strip()
    # BASE must be purely numeric and at least 10 digits long
    if not base.isdigit() or len(base) < 10:
        return None

    results = [base]
    for seg in parts[1:]:
        seg = seg.strip()
        if not seg.isdigit():
            return None  # non-numeric segment → abort
        if len(seg) >= len(base):
            return None  # segment not shorter than base → abort
        # Replace the last len(seg) digits of BASE with seg
        expanded = base[: len(base) - len(seg)] + seg
        results.append(expanded)

    return results


def _extract_slip_tokens(text: str) -> list:
    """Extract valid slip number tokens from a text string."""
    results = []
    for block in re.split(r"\s{2,}", text.strip()):
        clean = _SLIP_NOISE_RE.sub(" ", block)
        clean = re.sub(r"^[\s\-/+,]+|[\s\-/+,]+$", "", clean).strip()

        for cand in _SLIP_TOKEN_RE.findall(clean):
            cand = cand.strip("-")
            parts = cand.split("-")
            if len(parts) == 2:
                a, b = parts
                if _is_valid_slip_token(a) and _is_valid_slip_token(b) and abs(len(a) - len(b)) <= 2:
                    results.extend([a, b])
                    continue
            if _is_valid_slip_token(cand):
                results.append(_strip_slip_suffix(cand))

    return results


def _clean_insured_name(name: str) -> str:
    """Strip legal entity indicators, honorifics (Bapak, Ibu, Ny, Mr, Mrs, Ms, Sdr), titles, slashes, dashes, and quotes."""
    name = name.strip(' "\'“”«»')
    name = re.sub(r"[\/\-]", " ", name)
    name = _normalize_spaces(name)
    for _ in range(4):
        name = INSURED_ENTITY_ANYWHERE_RE.sub(" ", name)
        name = HONORIFICS_TITLES_RE.sub(" ", name)
        name = INSURED_PREFIX_SUFFIX_RE.sub("", name)
        name = re.sub(r"^\s*(?:[A-Z]\.){1,}[A-Z]?\s*", " ", name)
        name = re.sub(r"\(\s*\)", " ", name)
        name = name.strip(' "\'“”«»-.,/;:')
        cleaned = _normalize_spaces(name)
        if cleaned == name:
            break
        name = cleaned
    return name


def _cap_or_join(items: list) -> list:
    """If item count exceeds MAX_SPLIT_COLS, join into a single comma-separated string."""
    return [",".join(items)] if len(items) > MAX_SPLIT_COLS else items


# =============================================================================
# CLEAN POLIS
# =============================================================================

def clean_polis(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    if POLIS_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]
    if re.search(r"\d+TBA\d+", val, re.IGNORECASE):
        return [_normalize_spaces(val)]

    val_upper = val.upper().strip()
    if re.match(r"^VARIOUS\s*$", val_upper):
        return [val.strip()]
    if re.match(r"^VARIOUS\s*-\s*SEE\s+ATTACH", val_upper):
        return [val.strip()]
    if re.match(r"^TBA\s*$", val_upper):
        return [val.strip()]

    # Dot-suffix pattern: "010114003555.3431" → ["010114003555"]
    m = re.match(r"^(\d[\d\-]+)\.\d{1,6}$", val.strip())
    if m:
        base = m.group(1).strip()
        if _is_valid_polis_token(base):
            return [base]

    if re.search(r"S/D", val, re.IGNORECASE):
        sd_parts = re.split(r"\s{2,}", val.strip())
        left = sd_parts[0] if sd_parts else val
        base_m = (re.match(r"^(\d[\d\-]+?)-\d+S/D\d+", left, re.IGNORECASE) or
                  re.match(r"^(\d[\d\-]+)S/D\d+",       left, re.IGNORECASE))
        results = [base_m.group(1).strip()] if base_m else []
        for part in sd_parts[1:]:
            part = re.sub(r"\+?\s*\bTBA\b\s*", "", part, flags=re.IGNORECASE).strip().strip("+")
            results += [t for t in part.split() if _is_valid_polis_token(t.strip())]
        if results:
            return _cap_or_join(results)

    if re.match(r"^\s*TBA\s{2,}", val, re.IGNORECASE):
        rest = re.sub(r"^\s*TBA\s+", "", val, flags=re.IGNORECASE).strip()
        tokens = _extract_polis_tokens(rest)
        if tokens:
            return _cap_or_join(tokens)

    if "&" in val:
        tokens = _extract_polis_tokens(val)
        if tokens:
            return _cap_or_join(tokens)

    if "+" in val:
        # --- Skenario A: ada entry terpisah (double-space gap + blok non-'+') ---
        if re.search(r"\s{2,}", val):
            blocks = re.split(r"\s{2,}", val.strip())
            plus_blocks  = [b for b in blocks if "+" in b]
            entry_blocks = [b.strip() for b in blocks if "+" not in b and b.strip()]
            if plus_blocks and entry_blocks:
                # Blok '+' diabaikan; hanya entry terpisah yang diambil sebagai token
                entry_tokens = []
                for eb in entry_blocks:
                    for tok in eb.split():
                        tok = tok.strip().strip("/")
                        if _is_valid_polis_token(tok):
                            if len(re.sub(r"\D", "", tok)) >= 10:
                                tok = _strip_polis_base(tok)
                            entry_tokens.append(tok)
                if entry_tokens:
                    return _cap_or_join(entry_tokens)

        # --- Skenario B: satu blok '+' tanpa entry terpisah → expand suffix-replace ---
        expanded = _expand_plus_suffix(val.strip())
        if expanded:
            return _cap_or_join(expanded)

        # --- Fallback: logic lama ---
        tokens = _extract_polis_tokens(val)
        if tokens:
            return _cap_or_join(tokens)

    if re.search(r"\s{2,}", val):
        tokens = _extract_polis_tokens(val)
        if tokens:
            return _cap_or_join(tokens)

    val = re.sub(r"\bVARIOUS\b\s*", "", val, flags=re.IGNORECASE).strip()
    val = re.sub(r"\bVAR\b\s*",     "", val, flags=re.IGNORECASE).strip()
    val = val.replace("¿", "")

    parts   = re.split(r"\s*/\s*", val)
    cleaned = [_normalize_spaces(p).strip(" -/") for p in parts if _normalize_spaces(p).strip(" -/")]

    if len(cleaned) > 1 and all(len(c) < 7 for c in cleaned):
        return [val.strip()] if val.strip() else []

    return _cap_or_join(cleaned) if cleaned else []


# =============================================================================
# CLEAN SLIP
# =============================================================================

def clean_slip(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    # --- Skenario A: ada entry terpisah (double-space gap + blok non-'+') ---
    if "+" in val and re.search(r"\s{2,}", val):
        blocks = re.split(r"\s{2,}", val.strip())
        plus_blocks  = [b for b in blocks if "+" in b]
        entry_blocks = [b.strip() for b in blocks if "+" not in b and b.strip()]
        if plus_blocks and entry_blocks:
            entry_tokens = []
            for eb in entry_blocks:
                tokens = _extract_slip_tokens(eb)
                entry_tokens.extend(tokens)
            if entry_tokens:
                return _cap_or_join(entry_tokens)

    blocks = re.split(r"\s{2,}", val.strip())
    if len(blocks) >= 2:
        left_check = re.sub(
            r"\b(?:SUMMARY|BORDERO?|BORDERA|BORDRO|SINGGLESHIPMENT|VARIOUS)\b",
            "", blocks[0], flags=re.IGNORECASE,
        )
        left_check = _SLIP_NOISE_RE.sub("", left_check)
        left_check = re.sub(r"[\s\+\-/]+", "", left_check).strip()
        if not _is_valid_slip_token(left_check):
            tokens = _extract_slip_tokens("  ".join(blocks[1:]))
            if tokens:
                return _cap_or_join(tokens)

    tokens = _extract_slip_tokens(val)
    if tokens:
        return _cap_or_join(tokens)

    cleaned = re.sub(r"\bVARIOUS\b\s*", "", val, flags=re.IGNORECASE)
    cleaned = re.sub(r"^\s*-\s*", "", cleaned).strip()
    cleaned = _normalize_spaces(cleaned)
    return [cleaned] if cleaned else []


# =============================================================================
# CLEAN INSURED
# =============================================================================

def clean_insured(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip(' "\'“”«»')
    if not val:
        return []

    cleaned = []
    for p in INSURED_SPLIT_RE.split(val):
        p = _normalize_spaces(p.strip())
        if len(p) <= 2 or p.upper().strip() in INSURED_JUNK_WORDS:
            continue
        if re.match(r"^[^a-zA-Z0-9]+$", p):
            continue
        p_clean = _clean_insured_name(p)
        if not p_clean or len(p_clean) <= 2 or p_clean.upper().strip() in INSURED_JUNK_WORDS:
            continue
        cleaned.append(p_clean)

    if not cleaned:
        fallback = _clean_insured_name(_normalize_spaces(val))
        return [fallback] if fallback else []

    return _cap_or_join(cleaned)


# =============================================================================
# BUSINESS PARTNER
# =============================================================================

def get_business_partner(comp_name2, comp_name) -> str:
    """Return comp_name2 (broker) unless empty or 'DIRECT', in which case return comp_name (cedant)."""
    broker = "" if pd.isna(comp_name2) else str(comp_name2).strip()
    if not broker or broker.upper() == "DIRECT":
        return "" if pd.isna(comp_name) else str(comp_name).strip()
    return broker


# =============================================================================
# PROCESS DATA
# =============================================================================

def _insert_clean_columns(df: pd.DataFrame, all_lists: list, prefix: str, max_cols: int = MAX_SPLIT_COLS) -> list:
    """Insert clean columns (capped at MAX_SPLIT_COLS). Returns list of added names."""
    n_cols = min(max_cols, MAX_SPLIT_COLS)
    prefix_map = {
        "polis": "POLICY_CLEAN",
        "sertif": "SERTIF_CLEAN",
        "slip": "SLIP_CLEAN",
        "insured": "INSURED_CLEAN",
    }
    col_prefix = prefix_map.get(prefix.lower(), f"{prefix.upper()}_CLEAN")
    added = []
    for i in range(1, n_cols + 1):
        col_name = f"{col_prefix}_{i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else "" for lst in all_lists]
        added.append(col_name)
    return added


def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Reading: {input_file} (membaca file Excel besar, mohon tunggu)...", flush=True)
    df = pd.read_excel(input_file, header=0)
    print(f"      Total rows: {len(df):,}", flush=True)

    # Remove illegal XML control characters (e.g. \x1f) that corrupt Excel workbooks
    for c in df.select_dtypes(include=['object']).columns:
        df[c] = (df[c]
                 .fillna('')
                 .astype(str)
                 .str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True))

    if CEDANT_COL not in df.columns:
        print(f"\n[ERROR] Column '{CEDANT_COL}' not found. Available: {list(df.columns)}")
        return

    df[CEDANT_COL] = df[CEDANT_COL].astype(str).str.strip()
    df = df[df[CEDANT_COL] == CEDANT_VALUE].copy()
    print(f"[2/5] Filter '{CEDANT_VALUE}': {len(df):,} rows.")

    if df.empty:
        print("\n[WARN] No data after filter. Stopping.")
        return

    for col in [POLIS_COL, SLIP_COL, INSURED_COL]:
        if col not in df.columns:
            print(f"\n[ERROR] Column '{col}' not found. Available: {list(df.columns)}")
            return

    business_partner_values = [get_business_partner(row[BROKER_COL], row[CEDANT_COL]) for _, row in df.iterrows()]

    if "COMP_NAME.1" in df.columns:
        df.rename(columns={"COMP_NAME.1": "COMP_NAME2"}, inplace=True)

    df.rename(columns={POLIS_COL: "FAC_POLICY_NO", SLIP_COL: "FAC_SLIP", INSURED_COL: "FAC_INSURED"}, inplace=True)

    insert_pos = list(df.columns).index(BROKER_COL) + 1 if BROKER_COL in df.columns else len(df.columns)
    df.insert(insert_pos, "BUSSINESS_PATNER", business_partner_values)

    print("[3/5] Cleaning ...")

    # Vectorized map() — jauh lebih cepat dari iterrows() untuk dataset besar
    print(f"      Cleaning polis ({len(df):,} baris) ...", flush=True)
    all_polis = df["FAC_POLICY_NO"].map(clean_polis).tolist()

    print(f"      Cleaning slip ...", flush=True)
    all_slip  = df["FAC_SLIP"].map(clean_slip).tolist()

    print(f"      Cleaning insured ...", flush=True)
    all_ins   = df["FAC_INSURED"].map(clean_insured).tolist()

    # Ekstrak sertifikat 6-digit dari polis_ori
    print(f"      Extracting sertifikat dari polis ...", flush=True)
    all_sertif = df["FAC_POLICY_NO"].map(extract_cert_from_polis).tolist()
    max_sertif = min(3, max(1, max((len(x) for x in all_sertif), default=1)))

    max_polis = max(1, max((len(x) for x in all_polis), default=1))
    max_slip  = max(1, max((len(x) for x in all_slip),  default=1))
    max_ins   = max(1, max((len(x) for x in all_ins),   default=1))

    print(f"      -> polis cols: {max_polis}, slip cols: {max_slip}, insured cols: {max_ins}")

    print("[4/5] Building output columns ...")
    _insert_clean_columns(df, all_ins, "insured", 5)
    _insert_clean_columns(df, all_polis, "polis", 5)
    _insert_clean_columns(df, all_sertif, "sertif", 3)
    _insert_clean_columns(df, all_slip, "slip", 5)

    # Standarisasi urutan kolom FACUL CLEAN (Rule 1: sertif selalu sehabis polis)
    FACUL_TARGET_COLS = [
        "FAC_CODE", "FAC_CEDANT", "COMP_NAME", "FAC_BROKER", "COMP_NAME2", "BUSSINESS_PATNER",
        "FAC_INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2", "INSURED_CLEAN_3", "INSURED_CLEAN_4", "INSURED_CLEAN_5",
        "FAC_POLICY_NO", "POLICY_CLEAN_1", "SERTIF_CLEAN_1", "POLICY_CLEAN_2", "SERTIF_CLEAN_2", "POLICY_CLEAN_3", "SERTIF_CLEAN_3",
        "POLICY_CLEAN_4", "POLICY_CLEAN_5",
        "FAC_SLIP", "SLIP_CLEAN_1", "SLIP_CLEAN_2", "SLIP_CLEAN_3", "SLIP_CLEAN_4", "SLIP_CLEAN_5",
        "FAC_CURRENCY", "FAC_INP_DATE", "FAC_COM_DATE", "FAC_EXP_DATE", "FAC_SUB_CLASS", "FAC_RISK", "FAC_DESC", "FAC_ACC_STS", "FAC_STS_SLIP"
    ]
    for c in FACUL_TARGET_COLS:
        if c not in df.columns:
            df[c] = ""
    extra_cols = [c for c in df.columns if c not in FACUL_TARGET_COLS]
    df = df[FACUL_TARGET_COLS + extra_cols]

    print(f"[5/5] Saving to: {output_file} ...")
    df.to_excel(output_file, index=False)

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(output_file, "facul_clean")
    except Exception as e:
        print(f"  [WARN] Styling failed: {e}")


    # ---------------------------------------------------------
    # EXPORT TO POSTGRESQL
    # ---------------------------------------------------------
    
    load_dotenv()
    print("  Exporting to PostgreSQL ...")
    db_user = os.environ.get("DB_USER", "postgres")
    db_pass = os.environ.get("DB_PASS", "postgres")
    db_host = os.environ.get("DB_HOST", "localhost")
    db_port = os.environ.get("DB_PORT", "5432")
    db_name = os.environ.get("DB_NAME", "postgres")
    
    conn_str = f"postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"
    try:
        engine = create_engine(conn_str)
        table_name = "SUSPENSE_DATA_FACUL_CLEAN"
        
        with engine.connect() as conn:
            result = conn.execute(text(f'SELECT * FROM "{table_name}" LIMIT 0'))
            db_cols = list(result.keys())
            
        export_cols = [c for c in df.columns if c in db_cols]
        df_export = df[export_cols]
        
        df_export = df_export.replace({"": None})
        df_export.to_sql(table_name, con=engine, if_exists="append", index=False)
        print(f"      -> Successfully exported {len(df):,} rows to table '{table_name}'")
    except Exception as e:
        print(f"  [WARN] PostgreSQL export failed: {e}")

    print(f"\n{'=' * 50}")
    print(f"  Done. {len(df):,} rows, {len(df.columns)} cols -> {output_file}")
    print(f"{'=' * 50}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)
