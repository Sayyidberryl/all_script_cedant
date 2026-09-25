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


INPUT_FILE  = os.path.join("data", "osbal.xlsx")
OUTPUT_FILE = os.path.join("data", "osbal_clean_aca.xlsx")

CEDANT_COL   = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT. ASURANSI CENTRAL ASIA"

POLIS_COL        = "FAC_POLICY_NO"
SLIP_COL         = "FAC_SLIP"
INSURED_COL      = "FAC_INSURED"
CLSDT_POLIS_COL  = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL   = "CLSDT_SLIP_NO"
CLSDT_SERTF_COL  = "CLSDT_SERTF_NO"

MAX_SPLIT_COLS = 5

POLIS_EXCEPTION_RE = re.compile(
    r"""
    MOP\s*MARINE
  | (?:LINE\s*SLIP|LINESLIP)
  | \b(?:P1|P2|P3|P4|P73)\s*CANCEL
  | \b(?:P1|P2|P3|P4|P73)\b
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


# =============================================================================
# HELPERS
# =============================================================================

def clean_sertif(val) -> list:
    """Bersihkan nilai kolom sertifikat OSBAL (CLSDT_SERTF_NO)."""
    if pd.isna(val):
        return []
        
    if isinstance(val, float) and val.is_integer():
        val = int(val)
        
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
        
    if not s or s == '#':
        return []
    # _SERTIF_DIGIT_RE = re.compile(r'^\d{1,6}$') - using regex from above
    if re.match(r'^\d{1,6}$', s):
        return [s.zfill(6)]
    return []


def extract_cert_from_polis_osbal(val) -> list:
    """Extract certificate range SD from OSBAL policy."""
    if pd.isna(val): return []
    val = str(val).strip()
    if not val: return []

    m = re.search(r"^(\d{10,})\s*-\s*(\d+)\s*(?:S/D|SD)\s*(\d+)$", val, re.IGNORECASE)
    if not m:
        m = re.search(r"^(\d{10,})\s+(\d+)\s*(?:S/D|SD)\s*(\d+)$", val, re.IGNORECASE)
    if m:
        start_str = m.group(2)
        end_str = m.group(3)
        pad_len = max(len(start_str), len(end_str), 6)
        return [f"{start_str.zfill(pad_len)} SD {end_str.zfill(pad_len)}"]

    val_slash = val.replace('¿', '').strip()
    if "/" in val_slash or re.search(r"\d{10,}\s+\d{1,4}(?:/|\s)", val_slash):
        parts = [p.strip() for p in re.split(r'[\s/]+', val_slash) if p.strip()]
        if len(parts) >= 2:
            base = parts[0]
            if base.isdigit() and len(base) >= 10:
                is_valid = True
                for seg in parts[1:]:
                    if not seg.isdigit() or len(seg) >= len(base):
                        is_valid = False
                        break
                if is_valid:
                    start_str = parts[1]
                    end_str = parts[-1]
                    return [f"{start_str.zfill(6)} SD {end_str.zfill(6)}"]
    
    return []


def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


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


def _extract_polis_tokens(text: str) -> list:
    """Extract valid policy number tokens from a text string."""
    tokens = []
    for block in re.split(r"\s{2,}", text.strip()):
        for tok in re.split(r"\+|\s+", block.strip()):
            tok = tok.strip().strip("/")
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
    if not base.isdigit() or len(base) < 10:
        return None

    results = [base]
    for seg in parts[1:]:
        seg = seg.strip()
        if not seg.isdigit():
            return None
        if len(seg) >= len(base):
            return None
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

            if len(parts) >= 3:
                results.append(_strip_polis_base(cand))
                continue

            if len(parts) == 2:
                a, b = parts
                if _is_valid_slip_token(a) and _is_valid_slip_token(b) and abs(len(a) - len(b)) <= 2:
                    results.extend([a, b])
                    continue

            if _is_valid_slip_token(cand):
                results.append(_strip_polis_base(cand))

    return results


def _clean_insured_name(name: str) -> str:
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
    return [",".join(items)] if len(items) > MAX_SPLIT_COLS else items


def _clsdt_result_is_valid(tokens: list) -> bool:
    if not tokens:
        return False
    for tok in tokens:
        if not tok:
            continue
        s = str(tok)
        if re.search(r"\d{10,}", s):
            return True
        if re.search(r"\d{7,}", s) and not re.search(r"\bTBA\b", s, re.IGNORECASE):
            return True
    return False


# =============================================================================
# CLEAN POLIS
# =============================================================================

def clean_polis(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    # Standarisasi variasi S/D -> SD terlebih dahulu sebelum pemrosesan
    val_norm = re.sub(r"\s*S\s*/\s*D\s*", " SD ", val, flags=re.IGNORECASE)
    val_norm = re.sub(r"(?<=\w)SD(?=\w)", " SD ", val_norm)
    val_norm = _normalize_spaces(val_norm)

    # Handle sequence base polis + cert range (contoh: 100170221120000071 - 000001 SD 001127)
    m = re.search(r"^(\d{10,})\s*[-:]\s*(\d+)\s*SD\s*(\d+)$", val_norm, re.IGNORECASE)
    if not m:
        m = re.search(r"^(\d{10,})\s+(\d+)\s*SD\s*(\d+)$", val_norm, re.IGNORECASE)
    if m:
        if len(m.group(2)) <= 6 and len(m.group(3)) <= 6:
            return [m.group(1)]
        else:
            return [f"{m.group(1)} SD {m.group(3)}"]

    # Handle range 2 nomor polis utuh (contoh: 022100109150 SD 022100109210)
    m_range = re.search(r"^(\d{10,})\s*SD\s*(\d{10,})$", val_norm, re.IGNORECASE)
    if m_range:
        return [f"{m_range.group(1)} SD {m_range.group(2)}"]

    # Handle suffix sequence like: 101030819040000012 80/81/83/¿
    val_slash = val_norm.replace('¿', '').strip()
    if "/" in val_slash or re.search(r"\d{10,}\s+\d{1,4}(?:/|\s)", val_slash):
        parts = [p.strip() for p in re.split(r'[\s/]+', val_slash) if p.strip()]
        if len(parts) >= 2:
            base = parts[0]
            if base.isdigit() and len(base) >= 10:
                is_valid = True
                for seg in parts[1:]:
                    if not seg.isdigit() or len(seg) >= len(base):
                        is_valid = False
                        break
                if is_valid:
                    return [base]

    if POLIS_EXCEPTION_RE.search(val_norm):
        tokens = _extract_polis_tokens(val_norm)
        long_tokens = [t for t in tokens if len(re.sub(r"\D", "", t)) >= 10]
        long_tokens = [re.sub(r"S/D", "SD", t, flags=re.IGNORECASE) for t in long_tokens]
        return _cap_or_join(long_tokens) if long_tokens else [_normalize_spaces(re.sub(r"S/D", "SD", val_norm, flags=re.IGNORECASE))]

    if re.search(r"\d+TBA\d+", val_norm, re.IGNORECASE):
        return [_normalize_spaces(re.sub(r"S/D", "SD", val_norm, flags=re.IGNORECASE))]

    val_upper = val_norm.upper().strip()
    if re.match(r"^VARIOUS\s*$", val_upper):
        return [val_norm.strip()]
    if re.match(r"^VARIOUS\s*-\s*SEE\s+ATTACH", val_upper):
        return [val_norm.strip()]
    if re.match(r"^TBA\s*$", val_upper):
        return [val_norm.strip()]

    if re.search(r"\bSD\b", val_norm, re.IGNORECASE):
        sd_parts = re.split(r"\s{2,}", val_norm.strip())
        left = sd_parts[0] if sd_parts else val_norm
        m_dash = re.match(r"^(\d[\d\-]+?)\s*[-:]\s*\d{1,6}\s*SD\s*\d{1,6}$", left, re.IGNORECASE)
        if m_dash:
            results = [m_dash.group(1).strip()]
        elif re.match(r"^(\d{10,})\s*SD\s*(\d{10,})$", left, re.IGNORECASE):
            results = [left]
        else:
            tokens = [t.strip() for t in re.split(r"\s*SD\s*|\s+", left) if t.strip()]
            results = [t for t in tokens if _is_valid_polis_token(t)]
        for part in sd_parts[1:]:
            part = re.sub(r"\+?\s*\bTBA\b\s*", "", part, flags=re.IGNORECASE).strip().strip("+")
            results += [t for t in part.split() if _is_valid_polis_token(t.strip())]
        if results:
            results = [re.sub(r"S/D", "SD", t, flags=re.IGNORECASE) for t in results]
            return _cap_or_join(results)

    if re.match(r"^\s*TBA\s{2,}", val_norm, re.IGNORECASE):
        rest = re.sub(r"^\s*TBA\s+", "", val_norm, flags=re.IGNORECASE).strip()
        tokens = _extract_polis_tokens(rest)
        if tokens:
            tokens = [re.sub(r"S/D", "SD", t, flags=re.IGNORECASE) for t in tokens]
            return _cap_or_join(tokens)

    if "+" in val_norm:
        if re.search(r"\s{2,}", val_norm):
            blocks = re.split(r"\s{2,}", val_norm.strip())
            plus_blocks  = [b for b in blocks if "+" in b]
            entry_blocks = [b.strip() for b in blocks if "+" not in b and b.strip()]
            if plus_blocks and entry_blocks:
                entry_tokens = []
                for eb in entry_blocks:
                    for tok in eb.split():
                        tok = tok.strip().strip("/")
                        if _is_valid_polis_token(tok):
                            if len(re.sub(r"\D", "", tok)) >= 10:
                                tok = _strip_polis_base(tok)
                            entry_tokens.append(re.sub(r"S/D", "SD", tok, flags=re.IGNORECASE))
                if entry_tokens:
                    return _cap_or_join(entry_tokens)

        expanded = _expand_plus_suffix(val_norm.strip())
        if expanded:
            expanded = [re.sub(r"S/D", "SD", t, flags=re.IGNORECASE) for t in expanded]
            return _cap_or_join(expanded)

        tokens = _extract_polis_tokens(val_norm)
        if tokens:
            tokens = [re.sub(r"S/D", "SD", t, flags=re.IGNORECASE) for t in tokens]
            return _cap_or_join(tokens)

    if re.search(r"\s{2,}", val_norm):
        tokens = _extract_polis_tokens(val_norm)
        if tokens:
            tokens = [re.sub(r"S/D", "SD", t, flags=re.IGNORECASE) for t in tokens]
            return _cap_or_join(tokens)

    val_clean = re.sub(r"\bVARIOUS\b\s*", "", val_norm, flags=re.IGNORECASE).strip()
    val_clean = re.sub(r"\bVAR\b\s*",     "", val_clean, flags=re.IGNORECASE).strip()
    val_clean = val_clean.replace("¿", "")

    if re.search(r"\d{10,}.*-\d", val_clean) and "/" in val_clean:
        base_only = _strip_polis_base(val_clean.split("/")[0])
        if len(re.sub(r"\D", "", base_only)) >= 10:
            return [re.sub(r"S/D", "SD", base_only, flags=re.IGNORECASE)]

    parts   = re.split(r"\s*/\s*", val_clean)
    cleaned = [_normalize_spaces(p).strip(" -/") for p in parts if _normalize_spaces(p).strip(" -/")]

    final = []
    for c in cleaned:
        final.append(_strip_polis_base(c) if len(re.sub(r"\D", "", c)) >= 10 else c)
    cleaned = [c for c in final if len(c) >= 3]

    if len(cleaned) > 1 and all(len(c) < 7 for c in cleaned):
        ret = re.sub(r"S/D", "SD", val.strip(), flags=re.IGNORECASE)
        return [ret] if ret else []

    cleaned = [re.sub(r"S/D", "SD", c, flags=re.IGNORECASE) for c in cleaned]
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

    # Ubah S/D menjadi SD
    val = re.sub(r"\s*S\s*/\s*D\s*", " SD ", val, flags=re.IGNORECASE)
    val = re.sub(r"(?<=\w)SD(?=\w)", " SD ", val)
    val = _normalize_spaces(val)

    if "+" in val and re.search(r"\s{2,}", val):
        blocks = re.split(r"\s{2,}", val.strip())
        plus_blocks  = [b for b in blocks if "+" in b]
        entry_blocks = [b.strip() for b in blocks if "+" not in b and b.strip()]
        if plus_blocks and entry_blocks:
            entry_tokens = []
            for eb in entry_blocks:
                tokens = _extract_slip_tokens(eb)
                entry_tokens.extend([re.sub(r"S/D", "SD", t, flags=re.IGNORECASE) for t in tokens])
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
                tokens = [re.sub(r"S/D", "SD", t, flags=re.IGNORECASE) for t in tokens]
                return _cap_or_join(tokens)

    tokens = _extract_slip_tokens(val)
    if tokens:
        tokens = [re.sub(r"S/D", "SD", t, flags=re.IGNORECASE) for t in tokens]
        return _cap_or_join(tokens)

    cleaned = re.sub(r"\bVARIOUS\b\s*", "", val, flags=re.IGNORECASE)
    cleaned = re.sub(r"^\s*-\s*", "", cleaned).strip()
    cleaned = _normalize_spaces(cleaned)
    cleaned = re.sub(r"S/D", "SD", cleaned, flags=re.IGNORECASE)
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
# PROCESS DATA
# =============================================================================

def _insert_clean_columns(df: pd.DataFrame, all_lists: list, prefix: str, max_cols: int) -> list:
    """Insert clean columns into df. Returns list of added column names."""
    prefix_map = {
        "polis": "POLICY_CLEAN",
        "sertif": "SERTIF_CLEAN",
        "slip": "SLIP_CLEAN",
        "insured": "INSURED_CLEAN",
    }
    col_prefix = prefix_map.get(prefix.lower(), f"{prefix.upper()}_CLEAN")
    added = []
    for i in range(1, max_cols + 1):
        col_name = f"{col_prefix}_{i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added


def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Reading: {input_file} (membaca file Excel besar, mohon tunggu)...", flush=True)
    df = pd.read_excel(input_file, header=0)
    print(f"      Total rows: {len(df):,}", flush=True)

    for c in df.select_dtypes(include=['object']).columns:
        df[c] = df[c].astype(str).str.replace(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', regex=True)

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

    has_clsdt_polis = CLSDT_POLIS_COL in df.columns
    has_clsdt_slip  = CLSDT_SLIP_COL  in df.columns
    has_clsdt_sertf = CLSDT_SERTF_COL  in df.columns
    if not has_clsdt_polis:
        print(f"  [INFO] Column '{CLSDT_POLIS_COL}' not found, will use FAC polis only.")
    if not has_clsdt_slip:
        print(f"  [INFO] Column '{CLSDT_SLIP_COL}' not found, will use FAC slip only.")

    df.rename(columns={POLIS_COL: "FAC_POLICY_NO", SLIP_COL: "FAC_SLIP", INSURED_COL: "FAC_INSURED"}, inplace=True)

    print("[3/5] Cleaning ...")

    clsdt_polis_col = CLSDT_POLIS_COL if has_clsdt_polis else None
    clsdt_slip_col  = CLSDT_SLIP_COL  if has_clsdt_slip  else None

    # --- POLIS ---
    print(f"      Cleaning polis ({len(df):,} baris) ...", flush=True)
    if clsdt_polis_col:
        _clsdt_polis = df[clsdt_polis_col].map(clean_polis).tolist()
        _fac_polis   = df["FAC_POLICY_NO"].map(clean_polis).tolist()
        all_polis = [
            cp if _clsdt_result_is_valid(cp) else cf
            for cp, cf in zip(_clsdt_polis, _fac_polis)
        ]
    else:
        all_polis = df["FAC_POLICY_NO"].map(clean_polis).tolist()

    # --- SERTIFIKAT ---
    print(f"      Cleaning sertifikat ...", flush=True)
    _polis_sertif = df["FAC_POLICY_NO"].map(extract_cert_from_polis_osbal).tolist()
    if has_clsdt_sertf:
        _clsdt_sertif = df[CLSDT_SERTF_COL].map(clean_sertif).tolist()
        all_sertif = [
            ps if ps else cs 
            for ps, cs in zip(_polis_sertif, _clsdt_sertif)
        ]
    else:
        all_sertif = _polis_sertif

    # --- SLIP ---
    print(f"      Cleaning slip ...", flush=True)
    if clsdt_slip_col:
        _clsdt_slip = df[clsdt_slip_col].map(clean_slip).tolist()
        _fac_slip   = df["FAC_SLIP"].map(clean_slip).tolist()
        all_slip = [
            cs if _clsdt_result_is_valid(cs) else cf
            for cs, cf in zip(_clsdt_slip, _fac_slip)
        ]
    else:
        all_slip = df["FAC_SLIP"].map(clean_slip).tolist()

    # --- INSURED (selalu dari FAC, tidak terpengaruh CLSDT) ---
    print(f"      Cleaning insured ...", flush=True)
    all_ins = df["FAC_INSURED"].map(clean_insured).tolist()

    max_polis = max((len(x) for x in all_polis), default=1)
    max_sertif = max((len(x) for x in all_sertif), default=1)
    max_slip  = max((len(x) for x in all_slip),  default=1)
    max_ins   = max((len(x) for x in all_ins),   default=1)

    print(f"      -> polis cols: {max_polis}, sertif cols: {max_sertif}, slip cols: {max_slip}, insured cols: {max_ins}")

    print("[4/5] Building output columns ...")
    _insert_clean_columns(df, all_ins, "insured", 5)
    _insert_clean_columns(df, all_polis, "polis", 5)
    _insert_clean_columns(df, all_sertif, "sertif", 1)
    _insert_clean_columns(df, all_slip, "slip", 5)

    # Pastikan di semua kolom clean polis, slip, dan sertif tidak ada S/D (diubah menjadi SD)
    clean_cols_to_check = [
        c for c in df.columns 
        if any(k in c.lower() for k in ["polis", "policy", "slip", "sertif"]) and "clean" in c.lower()
    ]
    for c in clean_cols_to_check:
        df[c] = df[c].map(
            lambda x: re.sub(r"\s*S\s*/\s*D\s*", " SD ", str(x), flags=re.IGNORECASE).strip()
            if pd.notna(x) and re.search(r"S\s*/\s*D", str(x), re.IGNORECASE)
            else x
        )

    # Standarisasi urutan kolom OSBAL CLEAN (Rule 1: SERTIF berurutan setelah POLIS)
    OSBAL_TARGET_COLS = [
        "CCOS_DOC_NO", "CCOS_DATE", "CCOS_REF_CODE", "CCOS_COMP", "CCOS_COMP_NAME", "CCOS_REF_COMP", "CCOS_REF_COMP_NAME",
        "FAC_INSURED", "INSURED_CLEAN_1", "INSURED_CLEAN_2", "INSURED_CLEAN_3", "INSURED_CLEAN_4", "INSURED_CLEAN_5",
        "CCOS_CURR", "CCOS_OR_BAL", "CCOS_BAL_DUE", "CCOS_OR_BAL_IN_IDR", "CCOS_BAL_DUE_IN_IDR",
        "FAC_COM_DATE", "FAC_EXP_DATE", "FAC_DUE_DATES", "FAC_SUB_CLASS",
        "FAC_POLICY_NO", "POLICY_CLEAN_1", "SERTIF_CLEAN_1", "POLICY_CLEAN_2", "SERTIF_CLEAN_2", "POLICY_CLEAN_3", "SERTIF_CLEAN_3", "POLICY_CLEAN_4", "POLICY_CLEAN_5",
        "FAC_SLIP", "SLIP_CLEAN_1", "SLIP_CLEAN_2", "SLIP_CLEAN_3", "SLIP_CLEAN_4", "SLIP_CLEAN_5",
        "CLASS_CODE", "CLASS_NAME", "CLSDT_POLICY_NO", "CLSDT_SLIP_NO", "CLSDT_SERTF_NO"
    ]
    for c in OSBAL_TARGET_COLS:
        if c not in df.columns:
            df[c] = None
    extra_cols = [c for c in df.columns if c not in OSBAL_TARGET_COLS]
    df = df[OSBAL_TARGET_COLS + extra_cols]

    print(f"[5/5] Saving to: {output_file} ...")
    df.to_excel(output_file, index=False)

    try:
        from excel_styler import apply_purple_column_style
        apply_purple_column_style(output_file, "osbal_clean")
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
        table_name = "SUSPENSE_DATA_OSBAL_CLEAN"
        df.to_sql(table_name, con=engine, if_exists="append", index=False)
        print(f"      -> Successfully exported {len(df):,} rows to table '{table_name}'")
    except Exception as e:
        print(f"  [WARN] PostgreSQL export failed: {e}")

    print(f"\n{'=' * 50}")
    print(f"  Done. {len(df):,} rows, {len(df.columns)} cols -> {output_file}")
    print(f"{'=' * 50}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)
