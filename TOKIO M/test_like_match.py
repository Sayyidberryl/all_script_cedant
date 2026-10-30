"""Unit tests untuk _like_match() yang sudah di-refactor.

Kasus uji:
  1. TRIMITRA MEBELLINDO  vs TRIMITRA MEBELINDO GROUP  -> dokumentasi limit typo
  2. PANARUB INDUSTRY (x2 baris)  -> MATCH ke kedua baris, tidak di-narrow jadi 1
  3. Negatif: ABC COMPANY vs XYZ GROUP                -> NO MATCH
  4. Backward compat: exact substring                 -> MATCH via SUBSTR
  5. Edge case: query hanya stopword                  -> NO MATCH
  6. Threshold polis ketat (0.80): partial overlap    -> NO MATCH

Cara jalankan:
  python test_like_match.py
"""
import sys
import os

# Tambahkan root IBS ke path agar bisa import matching_ibs
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from matching_tokio import (
    _like_match,
    _build_token_index,
    _normalize,
    _INSURED_STOPWORDS,
    _INSURED_LIKE_THRESHOLD,
    _POLIS_LIKE_THRESHOLD,
)


def _make_rows_and_index(names: list, col: str = "insured_cln_1") -> tuple:
    """Buat list rows dan token index dari nama-nama referensi."""
    rows = [{col: _normalize(n)} for n in names]
    tok_idx, tok_cache = _build_token_index(rows, [col], "insured_ori")
    return rows, tok_idx, [col]


# ═══════════════════════════════════════════════════════════════════════════
# TEST CASES
# ═══════════════════════════════════════════════════════════════════════════

def test_1_typo_overlap_documented_limit():
    """Kasus 1: TRIMITRA MEBELLINDO (typo satu huruf) vs TRIMITRA MEBELINDO GROUP.

    Token analysis:
      query tokens  = {TRIMITRA, MEBELLINDO}
      sig_query     = {TRIMITRA, MEBELLINDO}   (tidak ada stopword)
      ref tokens    = {TRIMITRA, MEBELINDO, GROUP}
      sig_ref       = {TRIMITRA, MEBELINDO}    (GROUP di-exclude)
      intersection  = {TRIMITRA} = 1
      overlap_coeff = 1/min(2,2) = 0.50

    Threshold standar 0.65 → TIDAK match (limit implementasi token-integer).
    Threshold longgar 0.45 → MATCH (0.50 >= 0.45).

    Dokumentasi: untuk kasus typo lebih dari 1 token perlu turunkan threshold
    atau tambahkan fuzzy per-token (edit distance) di versi berikutnya.
    """
    ref_names = ["TRIMITRA MEBELINDO GROUP"]
    rows, tok_idx, ref_cols = _make_rows_and_index(ref_names)
    query = [_normalize("TRIMITRA MEBELLINDO")]

    result_strict = _like_match(
        query, tok_idx, rows, ref_cols,
        stopwords=_INSURED_STOPWORDS,
        overlap_threshold=_INSURED_LIKE_THRESHOLD,   # 0.65
    )
    result_loose = _like_match(
        query, tok_idx, rows, ref_cols,
        stopwords=_INSURED_STOPWORDS,
        overlap_threshold=0.45,
    )

    assert result_strict == [], (
        f"[TEST 1a] Threshold 0.65: diharapkan NO MATCH (overlap=0.50), dapat {result_strict}"
    )
    assert result_loose == [0], (
        f"[TEST 1b] Threshold 0.45: diharapkan MATCH (overlap=0.50>=0.45), dapat {result_loose}"
    )
    print("PASS TEST 1 -- TRIMITRA MEBELLINDO vs TRIMITRA MEBELINDO GROUP: "
          "overlap=0.50 | strict(0.65)=NoMatch | loose(0.45)=Match [limit terdokumentasi]")


def test_2_multi_candidate_preserved():
    """Kasus 2: PANARUB INDUSTRY vs 2 baris PANARUB GROUP.

    Token analysis:
      sig_query = {PANARUB}   (INDUSTRY di-exclude sebagai stopword)
      sig_ref   = {PANARUB}   (GROUP di-exclude)
      overlap   = 1/1 = 1.0  -> MATCH untuk KEDUANYA

    Ekspektasi: result berisi 2 indeks (0 dan 1), TIDAK di-narrow.
    Ini memvalidasi bahwa _like_match tidak memilih salah satu -- keputusan
    narrowing ada di FASE 4 (currency/periode), bukan di sini.
    """
    ref_names = ["PANARUB GROUP", "PANARUB GROUP"]
    rows, tok_idx, ref_cols = _make_rows_and_index(ref_names)
    query = [_normalize("PANARUB INDUSTRY")]

    log = []
    result = _like_match(
        query, tok_idx, rows, ref_cols,
        stopwords=_INSURED_STOPWORDS,
        overlap_threshold=_INSURED_LIKE_THRESHOLD,
        match_log=log,
    )

    assert len(result) == 2, (
        f"[TEST 2] Diharapkan 2 kandidat (multi-fac tidak di-narrow), "
        f"dapat {len(result)}: {result}"
    )
    assert set(result) == {0, 1}, f"[TEST 2] Kandidat salah: {result}"
    methods = {m for _, m in log}
    assert "SUBSET" in methods or "OVERLAP" in methods, (
        f"[TEST 2] Method harus SUBSET atau OVERLAP, dapat: {methods}"
    )
    print(f"PASS TEST 2 -- PANARUB INDUSTRY vs 2x PANARUB GROUP: "
          f"{len(result)} kandidat, methods={methods}")


def test_3_negative_no_match():
    """Negatif: ABC COMPANY vs XYZ GROUP -- tidak ada token overlap signifikan."""
    ref_names = ["XYZ GROUP"]
    rows, tok_idx, ref_cols = _make_rows_and_index(ref_names)
    query = [_normalize("ABC COMPANY")]

    result = _like_match(
        query, tok_idx, rows, ref_cols,
        stopwords=_INSURED_STOPWORDS,
        overlap_threshold=_INSURED_LIKE_THRESHOLD,
    )

    assert result == [], (
        f"[TEST 3] Diharapkan NO MATCH, dapat: {result}"
    )
    print("PASS TEST 3 -- ABC COMPANY vs XYZ GROUP: no match (benar)")


def test_4_backward_compat_substr():
    """Backward compat: 'TRIMITRA MEBELINDO' in 'TRIMITRA MEBELINDO GROUP' -> SUBSTR."""
    ref_names = ["TRIMITRA MEBELINDO GROUP"]
    rows, tok_idx, ref_cols = _make_rows_and_index(ref_names)
    query = [_normalize("TRIMITRA MEBELINDO")]   # exact substring dari ref

    log = []
    result = _like_match(
        query, tok_idx, rows, ref_cols,
        stopwords=_INSURED_STOPWORDS,
        overlap_threshold=_INSURED_LIKE_THRESHOLD,
        match_log=log,
    )

    assert result == [0], f"[TEST 4] Diharapkan match, dapat: {result}"
    methods = {m for _, m in log}
    # "TRIMITRA MEBELINDO" is a substring of "TRIMITRA MEBELINDO GROUP"
    assert "SUBSTR" in methods or "SUBSET" in methods, (
        f"[TEST 4] Method diharapkan SUBSTR atau SUBSET, dapat: {methods}"
    )
    print(f"PASS TEST 4 -- backward compat substring: methods={methods}")


def test_5_stopword_only_query():
    """Edge case: query hanya stopword -> sig_query kosong -> NO MATCH via OVERLAP/SUBSET."""
    ref_names = ["PANARUB GROUP"]
    rows, tok_idx, ref_cols = _make_rows_and_index(ref_names)
    # "PT GROUP INDONESIA" -> semua token adalah stopword
    query = [_normalize("PT GROUP INDONESIA")]

    result = _like_match(
        query, tok_idx, rows, ref_cols,
        stopwords=_INSURED_STOPWORDS,
        overlap_threshold=_INSURED_LIKE_THRESHOLD,
    )

    # Kandidat ada (GROUP ada di token_index), tapi:
    # SUBSTR: "PT GROUP INDONESIA" in "PANARUB GROUP" -> False
    # SUBSET: sig_query={}, kosong -> False
    # OVERLAP: sig_query={} -> False
    assert result == [], (
        f"[TEST 5] Diharapkan NO MATCH (query stopword-only), dapat: {result}"
    )
    print("PASS TEST 5 -- stopword-only query: no match (benar)")


def test_6_polis_threshold_strict():
    """Threshold polis 0.80: satu token overlap dari tiga -> overlap=0.33 -> NO MATCH."""
    ref_names = ["POLICY ABC 12345"]
    rows, tok_idx, ref_cols = _make_rows_and_index(ref_names, col="polis_cln_1")
    query = [_normalize("POLICY XYZ 99999")]
    # tokens query = {POLICY, XYZ, 99999}
    # tokens ref   = {POLICY, ABC, 12345}
    # intersection = {POLICY} -> 1/3 = 0.33 < 0.80

    result = _like_match(
        query, tok_idx, rows, ref_cols,
        stopwords=frozenset(),
        overlap_threshold=_POLIS_LIKE_THRESHOLD,   # 0.80
    )

    assert result == [], (
        f"[TEST 6] Threshold polis 0.80: diharapkan NO MATCH (overlap=0.33), dapat: {result}"
    )
    print("PASS TEST 6 -- polis threshold 0.80: partial overlap tidak cukup")


# ═══════════════════════════════════════════════════════════════════════════
# RUNNER
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=" * 60)
    print("  Unit Test: _like_match() Token-Overlap Refactor")
    print("=" * 60)

    tests = [
        test_1_typo_overlap_documented_limit,
        test_2_multi_candidate_preserved,
        test_3_negative_no_match,
        test_4_backward_compat_substr,
        test_5_stopword_only_query,
        test_6_polis_threshold_strict,
    ]

    passed = 0
    failed = 0
    for t in tests:
        try:
            t()
            passed += 1
        except AssertionError as e:
            print(f"FAIL: {e}")
            failed += 1
        except Exception as e:
            import traceback
            print(f"ERROR in {t.__name__}: {e}")
            traceback.print_exc()
            failed += 1

    print("=" * 60)
    print(f"  Hasil: {passed} passed, {failed} failed / {len(tests)} total")
    print("=" * 60)
    sys.exit(0 if failed == 0 else 1)
