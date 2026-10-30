import os
import re
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_FILE = BASE_DIR / "data_raw" / "2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx"
OUTPUT_FILE = BASE_DIR / "data_cleaned" / "tripa_output_osbal.xlsx"

CEDANT_COL   = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT ASURANSI TRIPAKARTA"

POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"

MAX_SPLIT_COLS = 5


# ============================================================
# EXCEPTION POLIS
# ============================================================

POLIS_EXCEPTION_RE = re.compile(
    r"""
    MOP\s*MARINE
  | (?:LINE\s*SLIP|LINESLIP)
  | \b(?:P1|P2|P3|P73)\s*CANCEL
  | \b(?:P1|P2|P3|P73)\b
  | \bCANCEL\b
  | PENYELESAIAN(?:\s+SUSPENSE)?
  | HUTANG
  | UTANG
    """,
    re.IGNORECASE | re.VERBOSE,
)


# ============================================================
# EXCEPTION SLIP
# ============================================================

SLIP_EXCEPTION_RE = re.compile(
    r"""
    \bSUMMARY\b
  | \bBORDER[OA]\b
  | \bBORDRO\b
  | \bSINGGLESHIPMENT\b
  | \bP1\b
  | \bP2\b
  | \bP3\b
  | PENYELESAIAN(?:\s+SUSPENSE)?
  | HUTANG
  | UTANG
    """,
    re.IGNORECASE | re.VERBOSE,
)


MONTH_NAMES = frozenset({
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER",
    "JANUARY", "FEBRUARY", "MARCH", "MAY", "JUNE", "JULY", "AUGUST",
    "OCTOBER",
})


# ============================================================
# INSURED
# ============================================================

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS",
    "NON FOOD", "DIV",
    "MR", "MR.", "MRS", "MRS.", "MS", "MS.",
})


INSURED_SUFFIX_RE = re.compile(
    r"""
    ,?\s*\bTBK\b\s*(?:,?\s*PT\.?)?
  | ,?\s*\bPT\.?\s*$
  | ,?\s*\bCV\.?\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)


INSURED_PT_QQ_RE = re.compile(
    r"""
    \bPT\.?\b
  | \bQQ\b
    """,
    re.IGNORECASE | re.VERBOSE,
)


# ============================================================
# SLIP NOISE
# ============================================================

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


_SLIP_TOKEN_RE = re.compile(
    r"[A-Z0-9][A-Z0-9\-]{6,}",
    re.IGNORECASE,
)


# ============================================================
# HELPERS
# ============================================================

def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


def _is_valid_slip_token(tok: str) -> bool:
    tok = tok.strip()

    return (
        len(tok) >= 7
        and bool(re.search(r"\d", tok))
        and not re.match(r"^\d{4}$", tok)
    )


def _strip_slip_suffix(tok: str) -> str:
    m = re.match(r"^(.+?)-(\d{1,6})$", tok)

    if m and len(m.group(1)) > len(m.group(2)):
        return m.group(1)

    return tok


def _extract_slip_tokens(text: str) -> list:
    results = []

    for block in re.split(r"\s{2,}", text.strip()):

        clean = _SLIP_NOISE_RE.sub(" ", block)

        clean = re.sub(
            r"^[\s\-/+,]+|[\s\-/+,]+$",
            "",
            clean
        ).strip()

        for cand in _SLIP_TOKEN_RE.findall(clean):

            cand = cand.strip("-")

            parts = cand.split("-")

            if len(parts) == 2:

                a, b = parts

                if (
                    _is_valid_slip_token(a)
                    and _is_valid_slip_token(b)
                    and abs(len(a) - len(b)) <= 2
                ):
                    results.extend([a, b])
                    continue

            if _is_valid_slip_token(cand):
                results.append(
                    _strip_slip_suffix(cand)
                )

    return results


# ============================================================
# CLEAN INSURED NAME
# ============================================================

def _clean_insured_name(name: str) -> str:

    name = _normalize_spaces(name)

    if re.fullmatch(
        r"AEON\s+MALL\s*\(TENANTS\)",
        name,
        re.IGNORECASE
    ):
        return name.upper()

    for _ in range(3):

        cleaned = _normalize_spaces(
            INSURED_SUFFIX_RE.sub(
                "",
                name
            ).strip().strip(",")
        )

        if cleaned == name:
            break

        name = cleaned

    name = INSURED_PT_QQ_RE.sub(
        " ",
        name
    )

    if re.fullmatch(
        r"\s*AEON\s+MALL\s+\(TENANTS\)\s*",
        name,
        flags=re.IGNORECASE
    ):
        return "AEON MALL (TENANTS)"

    name = re.sub(
        r"\(\s*PERSERO\s*\)",
        "",
        name,
        flags=re.IGNORECASE,
    )

    name = re.sub(
        r"[()]",
        " ",
        name
    )

    name = name.replace(
        "(",
        " "
    )

    name = name.replace(
        ")",
        " "
    )

    name = _normalize_spaces(name)

    name = re.sub(
        r"^[\s.,\-]+|[\s.,\-]+$",
        "",
        name
    )

    return name


def _cap_or_join(items: list) -> list:

    if len(items) > MAX_SPLIT_COLS:
        return [",".join(items)]

    return items


# ============================================================
# CLEAN POLIS TRIPA
# ============================================================

def _build_repeated_polis(base: str, suffix: str) -> str:
    base = re.sub(
        r"[^A-Z0-9]",
        "",
        str(base).upper()
    )

    suffix = re.sub(
        r"[^A-Z0-9]",
        "",
        str(suffix).upper()
    )

    if not base or not suffix:
        return ""

    if len(suffix) >= len(base):
        return suffix

    return base[:-len(suffix)] + suffix


def _is_short_suffix(value: str) -> bool:
    value = value.strip().upper()

    return bool(
        re.fullmatch(
            r"\d{1,6}",
            value
        )
    )


def clean_polis_raw(val) -> list:
    if pd.isna(val):
        return []

    original = str(val).strip().upper()

    if not original:
        return []

    repeat_match = re.fullmatch(
        r"(\d+)\s*-\s*(\d{1,7})",
        original,
    )

    if repeat_match:

        base = repeat_match.group(1)
        suffix = repeat_match.group(2)

        if len(suffix) < len(base):
            results = [base]

            # Jika suffix adalah nomor sertifikat (biasanya 6 digit diawali banyak 0)
            # Jangan dibuat perulangan polis
            if len(suffix) == 6 and suffix.startswith("000"):
                return results

            candidate = _build_repeated_polis(
                base,
                suffix,
            )

            if candidate and candidate not in results:
                results.append(candidate)

            return results

    pr_pattern = (
        re.search(
            r"^\s*P1\s*/",
            original,
            flags=re.IGNORECASE,
        )
        or
        re.search(
            r"^\s*\d+\s*-\s*\d{4,7}(?:\s*/|$)",
            original,
            flags=re.IGNORECASE,
        )
        or
        re.search(
            r"^\s*\d+\s*-\s*\d{4,7}\s*/\s*\d+",
            original,
            flags=re.IGNORECASE,
        )
    )

    if pr_pattern:

        pr_val = original

        pr_val = re.sub(
            r"^\s*P1\s*/\s*",
            "",
            pr_val,
            flags=re.IGNORECASE,
        ).strip()

        pr_val = re.sub(
            r"\s*/\s*(?:VAR|VARIOUS)\s*$",
            "",
            pr_val,
            flags=re.IGNORECASE,
        ).strip()

        parts = re.split(
            r"\s*/\s*",
            pr_val,
        )

        results = []

        for part in parts:

            part = part.strip()

            if not part:
                continue

            if part.upper() in {
                "P1",
                "VAR",
                "VARIOUS",
            }:
                continue

            if re.fullmatch(r"\d{1,7}", part):
                continue

            part = re.sub(
                r"^(\d+)\s*-\s*\d{4,7}$",
                r"\1",
                part,
            ).strip()

            part = re.sub(
                r"\s*/\s*\d{1,7}$",
                "",
                part,
            ).strip()

            cleaned = re.sub(
                r"[^A-Z0-9]",
                "",
                part.upper(),
            )

            if (
                cleaned
                and cleaned not in {
                    "P1",
                    "VAR",
                    "VARIOUS",
                }
            ):
                results.append(cleaned)

        if results:
            return results

    if '-' in original and '+' in original:
        parts = re.split(r'\s*\+\s*', original)
        results = []
        current_base = None
        valid = True
        for part in parts:
            part = part.strip()
            if not part: continue
            if '-' in part:
                subparts = part.split('-')
                if len(subparts) == 2:
                    base = re.sub(r'[^A-Z0-9]', '', subparts[0])
                    suffix = re.sub(r'[^A-Z0-9]', '', subparts[1])
                    current_base = base
                    results.append(base + suffix)
                else:
                    valid = False
                    break
            else:
                if current_base:
                    suffix = re.sub(r'[^A-Z0-9]', '', part)
                    results.append(current_base + suffix)
                else:
                    valid = False
                    break
        if valid and results:
            return results

    if original == "SUSPENSE VARIOUS CEDANT 2018-2019":
        return [original]

    if re.search(r"\bSUSPENSE\b", original, flags=re.IGNORECASE):
        return [original]

    if original in {
        "SWADHARMA QQ SECTOOR",
        "VARIOUS",
        "TBA",
        "ENDORSEMENT 2024",
    }:
        return [original]

    if re.search(
        r"\b(PENYELESAIAN|HUTANG|UTANG|PIUTANG|ENDORSEMENT)\b",
        original,
        flags=re.IGNORECASE,
    ):
        return [original]

    sd_match = re.fullmatch(
        r"(.+?)\s+S\s*/?\s*D\s+(.+)",
        str(val),
        flags=re.IGNORECASE,
    )

    if sd_match:

        start_raw = sd_match.group(1).strip()
        end_raw = sd_match.group(2).strip()

        start_clean = re.sub(
            r"[^A-Z0-9]",
            "",
            start_raw.upper()
        )

        end_clean = re.sub(
            r"[^A-Z0-9]",
            "",
            end_raw.upper()
        )

        if (
            re.fullmatch(r"\d+", start_clean)
            and re.fullmatch(r"\d+", end_clean)
            and len(start_clean) == len(end_clean)
        ):

            start_num = int(start_clean)
            end_num = int(end_clean)

            step = 1 if end_num >= start_num else -1
            jumlah = abs(end_num - start_num) + 1

            if jumlah > MAX_SPLIT_COLS:
                return [original]

            hasil_sd = [
                str(num).zfill(len(start_clean))
                for num in range(
                    start_num,
                    end_num + step,
                    step
                )
            ]

            return hasil_sd

        return [original]
    
    val_str = re.sub(
        r"\s*/\s*(?:VARIOUS|VAR)\s*$",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    if re.match(r"^TBA\s*/\s*", val_str):
        val_str = re.sub(
            r"^TBA\s*/\s*",
            "",
            val_str,
            flags=re.IGNORECASE,
        ).strip()

    val_str = re.sub(
        r"-0$",
        "",
        val_str,
    ).strip()

    if re.search(r"\bS\s*/\s*D\b", val_str, flags=re.IGNORECASE):

        parts_sd = re.split(
            r"\s*S\s*/\s*D\s*",
            val_str,
            flags=re.IGNORECASE,
        )

        if len(parts_sd) != 2:
            return [original]

        start_raw = re.sub(
            r"[^0-9]",
            "",
            parts_sd[0],
        )

        end_raw = re.sub(
            r"[^0-9]",
            "",
            parts_sd[1],
        )

        if not start_raw or not end_raw:
            return [original]

        if len(start_raw) != len(end_raw):
            return [original]

        start_num = int(start_raw)
        end_num = int(end_raw)

        step = 1 if end_num >= start_num else -1
        jumlah = abs(end_num - start_num) + 1

        if jumlah > MAX_SPLIT_COLS:
            return [original]

        hasil_sd = [
            str(num).zfill(len(start_raw))
            for num in range(
                start_num,
                end_num + step,
                step
            )
        ]

        return hasil_sd

    if val_str in {"TBA", "VAR", "VARIOUS"}:
        return []

    m = re.match(
        r"^(\d+)-(\d{1,7})$",
        val_str,
    )

    if m:
        base = m.group(1)
        suffix = m.group(2)

        if len(suffix) < len(base):
            candidate = _build_repeated_polis(base, suffix)

            results = [base]

            if candidate and candidate not in results:
                results.append(candidate)

            return results

    m = re.match(
        r"^(\d+)-(\d{3}(?:-\d{3})+)$",
        val_str,
    )

    if m:
        base = m.group(1)
        suffixes = re.findall(r"\d{3}", m.group(2))

        results = [base]
        prefix = base[:-3]

        for suffix in suffixes:
            candidate = prefix + suffix

            if candidate not in results:
                results.append(candidate)

        if len(results) > MAX_SPLIT_COLS:
            return [original]

        return results

    val_str = re.sub(
        r"^P1\s*/\s*",
        "",
        val_str,
        flags=re.IGNORECASE,
    ).strip()

    if "/" in val_str:

        slash_parts = [
            p.strip()
            for p in re.split(r"\s*/\s*", val_str)
            if p.strip()
        ]

        results_slash = []

        for p in slash_parts:
            if p in {"VAR", "VARIOUS"}:
                continue

            if p:
                results_slash.append(p)

        if len(results_slash) > 1:
            return results_slash

    parts = re.split(
        r"\s*(?:\+|&|,)\s*",
        val_str,
    )

    parts = [
        p.strip()
        for p in parts
        if p.strip()
    ]

    results = []
    current_base = None

    for p in parts:

        m = re.match(
            r"^([A-Z]*\d+)-(\d{1,6})$",
            p,
        )

        if m:

            base = m.group(1)
            suffix = m.group(2)
            current_base = base

            if len(suffix) < len(base):

                candidate = (
                    base[:-len(suffix)]
                    + suffix
                )

                if base not in results:
                    results.append(base)

                if candidate not in results:
                    results.append(candidate)

                continue

            cleaned = re.sub(
                r"[^A-Z0-9]",
                "",
                p,
            )

            if cleaned:
                results.append(cleaned)

            continue

        if (
            current_base
            and re.fullmatch(r"\d{1,6}", p)
            and len(p) < len(current_base)
        ):

            candidate = (
                current_base[:-len(p)]
                + p
            )

            if candidate not in results:
                results.append(candidate)

            continue

        cleaned = re.sub(
            r"[^A-Z0-9]",
            "",
            p,
        )

        if (
            cleaned
            and cleaned not in {"TBA", "VAR", "VARIOUS"}
        ):
            results.append(cleaned)

            if re.search(r"\d", cleaned):
                current_base = cleaned

    if len(parts) == 1:

        space_parts = re.split(
            r"\s+",
            val_str.strip(),
        )

        if len(space_parts) > 1:

            temp = []

            for p in space_parts:

                cleaned = re.sub(
                    r"[^A-Z0-9]",
                    "",
                    p,
                )

                if (
                    cleaned
                    and cleaned not in {"TBA", "VAR", "VARIOUS"}
                ):
                    temp.append(cleaned)

            if len(temp) > 1:
                results = temp

    results = [
        re.sub(
            r"^VARIOUS",
            "",
            x,
            flags=re.IGNORECASE,
        ).strip()
        for x in results
    ]

    results = [
        x for x in results
        if x
    ]

    unique_results = []

    for item in results:
        if item not in unique_results:
            unique_results.append(item)

    if len(unique_results) > MAX_SPLIT_COLS:
        return [original]

    return unique_results


def clean_polis(val) -> list:
    if pd.isna(val):
        return []

    original = str(val).strip().upper()

    if not original:
        return []

    if original == "SUSPENSE VARIOUS CEDANT 2018-2019":
        return [original]

    original = re.sub(
        r"\s*/\s*0\s*/\s*(?:VAR|VARIOUS)\s*$",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    original = re.sub(
        r"\s*-\s*\d+\s*/\s*(?:VAR|VARIOUS)\s*$",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    original = re.sub(
        r"\s*-\s*(?:VAR|VARIOUS)\s*$",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    # Hapus daftar sertifikat setelah dash yang dipisahkan koma
    # Contoh: "10703082400027 - 00009,000010,000012" -> "10703082400027"
    original = re.sub(
        r"\s*-\s*\d+(?:\s*,\s*\d+)+.*$",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    # Hapus suffix certificate yang berupa range SD (misal: - 000003 SD 000015)
    # Agar output clean_polis hanya mengembalikan base polisnya saja
    original = re.sub(
        r"\s*-\s*\d{1,6}\s*S\s*/?\s*D\s*\d{1,6}.*",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    hasil = clean_polis_raw(original)
    hasil_clean = []

    for x in hasil:

        x = str(x).strip().upper()

        if (
            x in {
                "SWADHARMA QQ SECTOOR",
                "VARIOUS",
                "TBA",
                "ENDORSEMENT 2024",
            }
            or re.search(
                r"\b(PENYELESAIAN|HUTANG|UTANG|PIUTANG|ENDORSEMENT)\b",
                x,
                flags=re.IGNORECASE,
            )
            or re.search(r"\bS\s*/\s*D\b", x, flags=re.IGNORECASE)
        ):
            cleaned = x
        else:
            if re.fullmatch(
                r"\d+\s*-\s*\d+",
                x,
            ):
                cleaned = x

            else:
                cleaned = re.sub(
                    r"[^A-Z0-9]",
                    "",
                    x,
                )

        if cleaned and cleaned not in hasil_clean:
            hasil_clean.append(cleaned)

    return hasil_clean


# ============================================================
# CLEAN SLIP
# ============================================================

def clean_slip(val):
    if pd.isna(val):
        return []

    original = str(val).strip()

    if not original:
        return []

    s = original.upper().strip()

    if ";" in s:

        semicolon_parts = [
            p.strip()
            for p in re.split(r"\s*;\s*", s)
            if p.strip()
        ]

        if len(semicolon_parts) <= 5:
            return semicolon_parts

        return [
            _normalize_spaces(s)
        ]

    special_slip_match = re.match(
        r"^\s*(\d{10,})\s*/\s*0\s*/\s*VAR\b",
        s,
        flags=re.IGNORECASE,
    )

    if special_slip_match:
        return [
            special_slip_match.group(1)
        ]

    def fallback_original():
        return [
            _normalize_spaces(original.upper())
        ]

    def build_repeated_number(base, suffix):

        base = re.sub(
            r"[^A-Z0-9]",
            "",
            str(base).upper()
        )

        suffix = re.sub(
            r"[^0-9]",
            "",
            str(suffix).upper()
        )

        if not base or not suffix:
            return ""

        if len(suffix) >= len(base):
            return suffix

        return (
            base[:-len(suffix)]
            + suffix.zfill(len(suffix))
        )

    cn_slip_match = re.match(
        r"^\s*(\d{7,})\s*-\s*\d+\s*/\s*CN\b",
        s,
        flags=re.IGNORECASE,
    )

    if cn_slip_match:
        return [
            cn_slip_match.group(1)
        ]

    if s == "TBA":
        return ["TBA"]

    if s == "VARIOUS":
        return ["VARIOUS"]

    if s == "P1":
        return ["P1"]

    if re.search(
        r"\bPENYELESAIAN\b|\bSUSPENSE\b|\bHUTANG PIUTANG\b|\bUTANG PIUTANG\b",
        s,
        flags=re.IGNORECASE,
    ):
        return [
            _normalize_spaces(s)
        ]

    if "SUSPENSE VARIOUS CEDANT" in s:
        return [
            _normalize_spaces(s)
        ]

    if "PENYELESAIANSUSPEN" in s:
        return [
            _normalize_spaces(s)
        ]

    if re.search(
        r"\bENDORSEMENT\b",
        s,
        flags=re.IGNORECASE,
    ):
        return [
            _normalize_spaces(s)
        ]

    s = re.sub(
        r"^\s*TBA\s*/\s*",
        "",
        s,
        flags=re.IGNORECASE,
    ).strip()

    month_pattern = (
        r"JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|"
        r"AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|"
        r"JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|"
        r"AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER"
    )

    month_sd_match = re.search(
        rf"(\d{{7,}})\s+S\s*/?\s*D\s+(\d{{7,}})",
        s,
        flags=re.IGNORECASE,
    )

    if month_sd_match:

        slip_start = month_sd_match.group(1)
        slip_end = month_sd_match.group(2)

        start_num = int(slip_start)
        end_num = int(slip_end)

        jumlah = abs(end_num - start_num) + 1

        if jumlah > MAX_SPLIT_COLS:

            return [
                f"{slip_start} S/D {slip_end}"
            ]

        step = (
            1
            if end_num >= start_num
            else -1
        )

        return [
            str(num).zfill(len(slip_start))
            for num in range(
                start_num,
                end_num + step,
                step,
            )
        ]

    if re.search(
        rf"\b(?:{month_pattern})\b\s+\d{{4}}",
        s,
        flags=re.IGNORECASE,
    ):

        month_numbers = re.findall(
            r"(?<!\d)\d{7,}(?!\d)",
            s,
        )

        month_numbers = list(
            dict.fromkeys(month_numbers)
        )

        if month_numbers:

            if len(month_numbers) > MAX_SPLIT_COLS:

                return [
                    _normalize_spaces(s)
                ]

            return month_numbers

    if "BORDERO" in s:

        return [
            _normalize_spaces(s)
        ]

    sd_match = re.search(
        r"(\d{7,})\s+S\s*/?\s*D\s+(\d{7,})",
        s,
        flags=re.IGNORECASE,
    )

    if sd_match:

        slip_start = sd_match.group(1)
        slip_end = sd_match.group(2)

        start_num = int(slip_start)
        end_num = int(slip_end)

        step = 1 if end_num >= start_num else -1
        jumlah = abs(end_num - start_num) + 1

        if jumlah > MAX_SPLIT_COLS:
            return [
                _normalize_spaces(s)
            ]

        return [
            str(num).zfill(len(slip_start))
            for num in range(
                start_num,
                end_num + step,
                step,
            )
        ]

    if re.fullmatch(
        r"(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|"
        r"AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)"
        r"\s+\d{4}",
        s,
        re.IGNORECASE,
    ):
        return [
            _normalize_spaces(s)
        ]

    sd_match = re.search(
        r"^"
        r"(\d{10,})"
        r"\s*-\s*"
        r"(\d{1,6})"
        r"\s*S\s*/?\s*D\s*"
        r"(\d{1,6})"
        r"$",
        s,
        flags=re.IGNORECASE,
    )

    if sd_match:

        base = sd_match.group(1)
        suffix_start = sd_match.group(2)
        suffix_end = sd_match.group(3)

        result = [
            base,
            build_repeated_number(
                base,
                suffix_start
            ),
            build_repeated_number(
                base,
                suffix_end
            ),
        ]

        result = [
            x for x in result
            if x
        ]

        result = list(
            dict.fromkeys(result)
        )

        if len(result) > 5:
            return fallback_original()

        return result

    s = re.sub(
        r"^\s*P1\s*/\s*",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = re.sub(
        r"\s*/\s*P1\s*$",
        "",
        s,
        flags=re.IGNORECASE,
    )

    cn_numbers = re.findall(
        r"(?<!\d)(\d{4,})\s*/\s*CN\b",
        s,
        flags=re.IGNORECASE,
    )

    if cn_numbers:

        result = list(
            dict.fromkeys(cn_numbers)
        )

        if len(result) > 5:
            return fallback_original()

        return result

    s = re.sub(
        r"(?<!\d)(\d{10,})\s*-\s*(\d)(?!\d)",
        r"\1\2",
        s,
    )

    plus_repeat_match = re.fullmatch(
        r"(\d{10,})-(\d{1,6})(?:\+(\d{1,6}))+",
        s,
    )

    if plus_repeat_match:

        base = plus_repeat_match.group(1)
        suffix_part = s.split("-", 1)[1]

        suffixes = re.findall(
            r"\d{1,6}",
            suffix_part,
        )

        result = [base]

        for suffix in suffixes:

            candidate = build_repeated_number(
                base,
                suffix,
            )

            if candidate and candidate not in result:
                result.append(candidate)

        if len(result) > 5:
            return fallback_original()

        return result

    repeated_parts = re.split(
        r"\s*[;/]\s*",
        s,
    )

    if len(repeated_parts) > 1:

        repeated_result = []
        current_base = None

        for part in repeated_parts:

            part = part.strip()

            if not part:
                continue

            if part in {
                "VARIOUS",
                "VAR",
                "TBA",
            }:
                continue

            m = re.fullmatch(
                r"(\d{10,})\s*-\s*(\d{1,6})",
                part,
            )

            if m:

                base = m.group(1)
                suffix = m.group(2)
                current_base = base

                if base not in repeated_result:
                    repeated_result.append(base)

                if len(suffix) == 6 and suffix.startswith("000"):
                    continue

                candidate = build_repeated_number(
                    base,
                    suffix
                )

                if candidate and candidate not in repeated_result:
                    repeated_result.append(candidate)

                continue

            if re.fullmatch(
                r"\d{10,}",
                part,
            ):

                if part not in repeated_result:
                    repeated_result.append(part)

                current_base = part
                continue

            if (
                current_base
                and re.fullmatch(
                    r"\d{1,6}",
                    part,
                )
            ):

                candidate = build_repeated_number(
                    current_base,
                    part
                )

                if candidate and candidate not in repeated_result:
                    repeated_result.append(candidate)

                continue

        if repeated_result:

            if len(repeated_result) > 5:
                return fallback_original()

            return repeated_result

    if "+" in s:

        parts = [
            x.strip()
            for x in re.split(
                r"\s*\+\s*",
                s
            )
            if x.strip()
        ]

        if (
            parts
            and re.fullmatch(
                r"\d{10,}",
                parts[0]
            )
        ):

            first = parts[0]
            prefix = first[:-4]

            result = [first]

            for part in parts[1:]:

                if part in {
                    "VAR",
                    "VARIOUS",
                    "TBA",
                    "P1",
                }:
                    continue

                if re.fullmatch(
                    r"\d{1,4}",
                    part
                ):

                    result.append(
                        prefix
                        + part.zfill(4)
                    )
                    continue

                if re.fullmatch(
                    r"\d{10,}",
                    part
                ):

                    result.append(part)
                    continue

            result = list(
                dict.fromkeys(result)
            )

            if len(result) > 5:
                return fallback_original()

            return result

    s = re.sub(
        r"\s*/\s*307\s*$",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = re.sub(
        r"\bVARIOUS\b",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = re.sub(
        r"\bVAR\b",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = re.sub(
        r"\bTBA\b",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = s.strip(" /+-;")

    if re.search(
        r"SUB\s*CLASS\s+OF\s+BUSINESS",
        original,
        flags=re.IGNORECASE,
    ):

        return [
            _normalize_spaces(original.upper())
        ]

    long_numbers = re.findall(
        r"(?<!\d)\d{10,}(?!\d)",
        s,
    )

    if long_numbers:

        long_numbers = list(
            dict.fromkeys(long_numbers)
        )

        if len(long_numbers) > 5:
            return fallback_original()

        return long_numbers

    if re.fullmatch(
        r"\d{4}\.[A-Z]+\.\d+\.\d{4}",
        original.upper(),
    ):

        return [
            _normalize_spaces(
                original.upper()
            )
        ]

    sd_match = re.search(
        r"(?<!\d)"
        r"(\d{7,})"
        r"\s+S\s*/?\s*D\s+"
        r"(\d{1,})"
        r"(?!\d)",
        s,
        flags=re.IGNORECASE,
    )

    if sd_match:

        start_raw = sd_match.group(1)
        end_raw = sd_match.group(2)

        if len(start_raw) == len(end_raw):

            start_num = int(start_raw)
            end_num = int(end_raw)

            step = 1 if end_num >= start_num else -1
            jumlah = abs(end_num - start_num) + 1

            if jumlah > 5:
                return fallback_original()

            result = [
                str(num).zfill(len(start_raw))
                for num in range(
                    start_num,
                    end_num + step,
                    step
                )
            ]

            return result

        return fallback_original()

    parts = re.split(
        r"\s*[;/]\s*",
        s,
    )

    result = []

    for part in parts:

        part = part.strip()

        if not part:
            continue

        if part in {
            "VAR",
            "VARIOUS",
            "TBA",
            "P1",
        }:
            continue

        cn_match = re.search(
            r"(?<!\d)(\d{4,})\s*CN\b",
            part,
            flags=re.IGNORECASE,
        )

        if cn_match:

            number = cn_match.group(1)

            if number not in result:
                result.append(number)

            continue

        if re.fullmatch(
            r"\d+",
            part
        ):

            if part not in result:
                result.append(part)

            continue

        part = re.sub(
            r"\b(?:VAR|VARIOUS|TBA)\b",
            "",
            part,
            flags=re.IGNORECASE,
        ).strip()

        if not part:
            continue

        cleaned = _normalize_spaces(part)

        if not cleaned:
            continue

        if cleaned.upper() in {
            "VAR",
            "VARIOUS",
            "TBA",
        }:
            continue

        if cleaned not in result:
            result.append(cleaned)

    if len(result) > 5:
        return fallback_original()

    if not result:

        return [
            _normalize_spaces(
                original.upper()
            )
        ]

    return result


# ============================================================
# CLEAN INSURED
# ============================================================

def clean_insured(val) -> list:
    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val:
        return []
    
    if re.fullmatch(
        r"BIO\s+FARMA\s*,\s*PT\s+CLUSTER\s+E2",
        val,
        flags=re.IGNORECASE
    ):
        return ["BIO FARMA CLUSTER E2"]

    val = re.sub(
        r"\s*/\s*BORD(?:ERO|O)?\b.*$",
        "",
        val,
        flags=re.IGNORECASE
    ).strip()

    bracket_parts = re.findall(
        r"\(([^)]*)\)",
        val
    )

    val = re.sub(
        r"\([^)]*\)",
        "",
        val
    )

    val = re.sub(
        r"/\s*VARIOUS\s*$",
        "",
        val,
        flags=re.IGNORECASE,
    ).strip()

    val = re.sub(
        r"\s*&\s*FIXED\s+ASSET\s*$",
        "",
        val,
        flags=re.IGNORECASE,
    ).strip()

    parts = re.split(
        r"/|,|QQ",
        val,
        flags=re.IGNORECASE
    )

    parts.extend(bracket_parts)
    cleaned = []

    for p in parts:

        p = _normalize_spaces(p.strip())

        p = re.sub(
            r"^(MR|MRS|MS|DR|IR)\.?\s+",
            "",
            p,
            flags=re.IGNORECASE,
        )

        if len(p) <= 2:
            continue

        if p.upper() in INSURED_JUNK_WORDS:
            continue

        if re.match(
            r"^[^A-Za-z0-9]+$",
            p
        ):
            continue

        p_clean = _clean_insured_name(p)

        if not p_clean:
            continue

        if p_clean.upper() in INSURED_JUNK_WORDS:
            continue

        if p_clean not in cleaned:
            cleaned.append(p_clean.upper())

    if not cleaned:

        fallback = _clean_insured_name(
            _normalize_spaces(val)
        )

        return (
            [fallback.upper()]
            if fallback
            else []
        )

    return _cap_or_join(cleaned)


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN CERTIFICATE
# ─────────────────────────────────────────────────────────────────────────────
def clean_certificate(polis_ori):
    """
    Mengambil CERTIFICATE dari POLIS ORI dengan validasi ketat.
    """
    if pd.isna(polis_ori):
        return ""

    polis = str(polis_ori).strip().upper()

    if not polis:
        return ""

    # Hapus data kotor / suspense
    if re.search(r"\bSUSPENSE\b|\bENDORSEMENT\b", polis):
        return ""

    # 1. Hapus suffix fraksi di paling belakang (contoh: -1/1, 4/1)
    polis_clean = re.sub(r"[\s-]+\d+/\d+$", "", polis).strip()

    # 2. Hapus suffix / VAR / VARIOUS / TBA
    polis_clean = re.sub(
        r"\s*/\s*(?:VAR|VARIOUS|TBA).*$",
        "",
        polis_clean,
        flags=re.IGNORECASE,
    ).strip()

    # 3. ABAIKAN jika suffix bernilai "-0" atau "-0/VAR"
    if re.search(r"-0$", polis_clean):
        return ""

    # 4. Handling Pola RANGE S/D / SD
    if re.search(r"\bS\s*/?\s*D\b", polis_clean):
        parts_sd = re.split(r"\s*S\s*/?\s*D\s*", polis_clean)
        if len(parts_sd) == 2:
            left, right = parts_sd[0].strip(), parts_sd[1].strip()

            m_left = re.search(r"(?:^|[\s-])(\d{4,6})$", left)
            cert_left = m_left.group(1).zfill(6) if m_left else ""

            m_right = re.search(r"(?:^|[\s-])(\d{4,6})$", right)
            cert_right = m_right.group(1).zfill(6) if m_right else ""

            if cert_left and cert_right:
                return f"{cert_left} SD {cert_right}"

    # 5. Handling Multi Certificate (Pemisah Koma)
    # Wajib angka yang dipisahkan oleh koma (contoh: - 00009,000010,000012)
    m_multi = re.search(r"-\s*(\d+(?:\s*,\s*\d+)+)$", polis_clean)
    if m_multi:
        parts = m_multi.group(1).split(',')
        return ", ".join(p.strip().zfill(6) for p in parts)

    # 6. Handling Single Certificate (Pemisah Dash)
    # Wajib 4 - 6 digit angka untuk membedakan dari pecahan polis (seperti -020, -022)
    m_single_dash = re.search(r"-(\d{4,6})$", polis_clean)
    if m_single_dash:
        return m_single_dash.group(1).zfill(6)

    # 7. Handling Single Certificate (Pemisah Spasi)
    m_single_space = re.search(r"(?<=\s)(\d{4,6})$", polis_clean)
    if m_single_space:
        return m_single_space.group(1).zfill(6)

    return ""
  
# ============================================================
# PROCESS DATA
# ============================================================

def _insert_clean_columns(
    df: pd.DataFrame,
    all_lists: list,
    prefix: str,
    max_cols: int,
) -> list:

    added = []

    for i in range(1, max_cols + 1):
        col_name = f"{prefix}_{i}"

        df[col_name] = [
            lst[i - 1]
            if i - 1 < len(lst)
            else None
            for lst in all_lists
        ]

        added.append(col_name)

    return added


def _is_valid_clstd(val) -> bool:
    if pd.isna(val):
        return False

    val = str(val).strip().upper()

    if not val:
        return False

    if val in {"TBA", "VAR", "VARIOUS"}:
        return False

    return True


def _get_clstd_source(clstd_value, fac_value):
    if _is_valid_clstd(clstd_value):
        return (clstd_value, "CLSDT")

    return (fac_value, "FAC")


def _normalize_cedant(val) -> str:
    if pd.isna(val):
        return ""

    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(val).upper()
    )


def process_data(input_file: Path, output_file: Path) -> None:

    print(f"[1/5] Membaca data dari: {input_file} ...")

    from io import StringIO
    from xlsx2csv import Xlsx2csv

    print(f"[1/5] Membaca data dari: {input_file} (mode kencang) ...")

    # Stream xlsx langsung ke memori sebagai string CSV
    buffer = StringIO()
    Xlsx2csv(str(input_file), skip_empty_lines=True).convert(buffer)
    buffer.seek(0)

    # BACA DENGAN PANDAS READ_CSV (Hanya butuh 3-5 detik!)
    df = pd.read_csv(buffer, dtype=str, low_memory=False)

    print(f"      Total baris keseluruhan: {len(df):,}")

    required_cols = [
        CEDANT_COL,
        POLIS_COL,
        SLIP_COL,
        INSURED_COL,
        "CLSDT_POLICY_NO",
        "CLSDT_SLIP_NO",
    ]

    for col in required_cols:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    cedant_target = _normalize_cedant(CEDANT_VALUE)
    cedant_normalized = df[CEDANT_COL].map(_normalize_cedant)
    is_tripa = cedant_normalized.eq(cedant_target)

    print(
        f"[2/5] Filter cedant '{CEDANT_VALUE}': "
        f"{is_tripa.sum():,} baris TRIPA "
        f"dari total {len(df):,} baris."
    )

    if not is_tripa.any():
        print("\n[ERROR] Cedant TRIPA tidak ditemukan.")
        print("Contoh nama cedant:")
        print(
            df[CEDANT_COL]
            .dropna()
            .astype(str)
            .str.strip()
            .value_counts()
            .head(20)
            .to_string()
        )
        return

    # Buang semua baris selain cedant TRIPA -> output hanya TRIPA
    total_lain = len(df) - int(is_tripa.sum())
    df = df[is_tripa].reset_index(drop=True)
    is_tripa = pd.Series(True, index=df.index)
    print(f"      Baris cedant lain dibuang: {total_lain:,} | sisa: {len(df):,}")

    print("[3/5] Menjalankan proses cleaning hanya untuk TRIPA ...")

    # Ambil data kolom yang dibutuhkan dalam bentuk list untuk pemrosesan cepat
    polis_ori_list = df[POLIS_COL].tolist()
    slip_ori_list = df[SLIP_COL].tolist()
    insured_ori_list = df[INSURED_COL].tolist()
    clsdt_polis_list = df["CLSDT_POLICY_NO"].tolist()
    clsdt_slip_list = df["CLSDT_SLIP_NO"].tolist()
    is_tripa_list = is_tripa.tolist()

    all_clean_polis = []
    all_clean_slip = []
    all_clean_ins = []
    all_certificates = []

    max_polis = 1
    max_slip = 1
    max_ins = 1

    total_clstd_used = 0
    total_fallback_fac = 0
    total_rows = len(df)

    for idx in range(total_rows):

        if (idx + 1) % 50_000 == 0:
            print(f"      Progress: {idx + 1:,} / {total_rows:,} baris diproses...")

        if is_tripa_list[idx]:

            p_ori = polis_ori_list[idx]
            c_p_no = clsdt_polis_list[idx]

            polis_source, polis_source_name = _get_clstd_source(c_p_no, p_ori)

            if polis_source_name == "CLSDT":
                total_clstd_used += 1
            else:
                total_fallback_fac += 1

            c_polis = clean_polis(polis_source)

            s_ori = slip_ori_list[idx]
            c_s_no = clsdt_slip_list[idx]

            slip_source, _ = _get_clstd_source(c_s_no, s_ori)
            c_slip = clean_slip(slip_source)

            c_ins = clean_insured(insured_ori_list[idx])
            c_certificate = clean_certificate(p_ori)

        else:
            c_polis = []
            c_slip = []
            c_ins = []
            c_certificate = ""

        max_polis = max(max_polis, len(c_polis))
        max_slip = max(max_slip, len(c_slip))
        max_ins = max(max_ins, len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)
        all_certificates.append(c_certificate)

    print("      Selesai diproses!")
    print(f"      -> Jumlah kolom clean polis  : {max_polis}")
    print(f"      -> Jumlah kolom clean slip   : {max_slip}")
    print(f"      -> Jumlah kolom clean insured: {max_ins}\n")

    print("      SUMBER POLIS CLEAN:")
    print(f"      -> Menggunakan CLSDT : {total_clstd_used:,} baris")
    print(f"      -> Fallback ke FAC   : {total_fallback_fac:,} baris\n")

    print("[4/5] Menyusun kolom output ...")

    # TAMBAHKAN CERTIFICATE KE DATAFRAME TERLEBIH DAHULU (Solusi KeyError)
    df["CERTIFICATE_1"] = all_certificates

    new_columns = []
    for col in df.columns:
        # Hindari memasukkan CERTIFICATE pada posisi aslinya
        if col == "CERTIFICATE_1":
            continue

        if col == POLIS_COL:
            new_columns.append(col)

            clean_polis_columns = _insert_clean_columns(
                df,
                all_clean_polis,
                "FAC_POLICY_NO_CLEAN",
                max_polis,
            )

            if clean_polis_columns:
                new_columns.append(clean_polis_columns[0])
                new_columns.append("CERTIFICATE_1")
                new_columns += clean_polis_columns[1:]
            else:
                new_columns.append("CERTIFICATE_1")

        elif col == SLIP_COL:
            new_columns.append(col)
            new_columns += _insert_clean_columns(
                df,
                all_clean_slip,
                "FAC_SLIP_CLEAN",
                max_slip,
            )

        elif col == INSURED_COL:
            new_columns.append(col)
            new_columns += _insert_clean_columns(
                df,
                all_clean_ins,
                "FAC_INSURED_CLN",
                max_ins,
            )

        else:
            new_columns.append(col)

    df = df[new_columns]

    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")

    Path(output_file).parent.mkdir(parents=True, exist_ok=True)

    # Tambahkan engine="xlsxwriter"
    df.to_excel(output_file, index=False, engine="xlsxwriter")

    print("\n" + "=" * 55)
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print("=" * 55)


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("MASUK MAIN")
    process_data(INPUT_FILE, OUTPUT_FILE)