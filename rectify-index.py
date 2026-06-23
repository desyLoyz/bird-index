#!/usr/bin/env python3
"""
Laubmann General Index Parser  –  v3
======================================
Produces a clean, one-row-per-species table:

  Bird Name (German, edited) | Original Name (if different) | Family (Latin) | References (rectified)


No AI technologies used. Auto-installs missing dependencies.
"""

import subprocess, sys
def _ensure(pkg, imp=None):
    try: __import__(imp or pkg)
    except ImportError:
        print(f"Installing {pkg} …")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

_ensure("pandas")
_ensure("openpyxl")

# ─────────────────────────────────────────────────────────────────────────────
import json, re
from pathlib import Path
import pandas as pd

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
JSON_PATH    = "Laubmann_35_gemini_edits-170.json"
CSV_PATH     = "laubmann_index.csv"          # existing CSV to compare against
OUTPUT_EXCEL = "laubmann_index_v3.xlsx"
OUTPUT_CSV   = "laubmann_index_v3.csv"
INDEX_START  = "2d6d52f0-b55f-431d-81ac-e2b8ce255fb2_0004_R"

# Roman numerals longest-first so regex is greedy
ROMAN_LIST = [
    "XXXV","XXXIV","XXXIII","XXXII","XXXI","XXX",
    "XXIX","XXVIII","XXVII","XXVI","XXV","XXIV",
    "XXIII","XXII","XXI","XX","XIX","XVIII","XVII",
    "XVI","XV","XIV","XIII","XII","XI","X",
    "IX","VIII","VII","VI","V","IV","III","II","I"
]
ROMAN_PAT = "|".join(ROMAN_LIST)

ROMAN_TO_INT = {r: i+1 for i, r in enumerate(reversed(ROMAN_LIST))}

# ── TEXT CLEANING ─────────────────────────────────────────────────────────────

def strip_tags(t):
    return re.sub(r'<[^>]+>', '', t)

def strip_strikethrough(t):
    """Remove ~~struck-through~~ content."""
    return re.sub(r'~~[^~]*~~', '', t)

def clean(t):
    t = strip_strikethrough(t)
    t = strip_tags(t)
    t = re.sub(r'^#[^\n]*\n?', '', t, flags=re.MULTILINE)   # markdown headings
    t = re.sub(r'^\s*-\s', '',   t, flags=re.MULTILINE)      # list dashes
    return t

# ── FAMILY HEADER DETECTION ───────────────────────────────────────────────────

def extract_family(text):
    """
    A family header is a line containing ONLY a Latin name (1–2 capitalised
    words, no digits, no colon, length > 4 chars).
    """
    for line in text.splitlines():
        s = line.strip().rstrip('.')
        if re.fullmatch(r'[A-Z][a-z]{3,}(?:\s+[A-Z][a-z]{3,})?', s):
            return s
    return None

# ── REFERENCE RECTIFICATION ───────────────────────────────────────────────────

def rectify(raw):
    """
    Normalise any messy reference string to canonical 'VOL, PAGE; VOL, PAGE; …'
    Handles: dot-as-separator, missing comma, comma-as-separator,
             stray OCR letters, line-breaks, trailing punctuation.
    """
    s = re.sub(r'\s*\n\s*', ' ', raw).strip()

    # Remove stray single lowercase letters between refs (OCR noise)
    s = re.sub(r'(?<=\d)\s+[a-km-wyz]\s+(?=' + ROMAN_PAT + r')', ' ', s)

    # dot-as-separator: digit . ROMAN  →  digit ; ROMAN
    s = re.sub(
        r'(\d)\s*\.\s*(?=(?:' + ROMAN_PAT + r')(?:\s*[,;]|\s+\d))',
        r'\1; ', s)

    # comma-as-separator: digit , ROMAN  →  digit ; ROMAN
    s = re.sub(
        r'(\d)\s*,\s*(?=(?:' + ROMAN_PAT + r')(?:\s*[,;]|\s+\d))',
        r'\1; ', s)

    # missing comma: ROMAN<space>digit  →  ROMAN, digit
    s = re.sub(r'(' + ROMAN_PAT + r')\s+(\d)', r'\1, \2', s)

    # strip leading junk
    s = re.sub(r'^[\s;:,.\-]+', '', s)

    # normalise semicolons and trailing
    s = re.sub(r'\s*;\s*', '; ', s)
    s = re.sub(r'[;,.\s]+$', '', s).strip()
    return s

# ── REASSEMBLE WRAPPED LINES ──────────────────────────────────────────────────

def reassemble(text):
    """
    Join lines that are continuations of a reference list.
    A continuation line starts with a Roman numeral or a bare digit.
    """
    lines  = text.splitlines()
    result = []
    for line in lines:
        s = line.strip()
        if not s:
            result.append('')
            continue
        is_cont = bool(
            re.match(r'^(?:' + ROMAN_PAT + r')\s*[,\s]', s) or
            re.match(r'^\d', s)
        )
        if is_cont and result and result[-1]:
            result[-1] = result[-1].rstrip() + ' ' + s
        else:
            result.append(s)
    return '\n'.join(result)

# ── SPECIES ENTRY EXTRACTION ──────────────────────────────────────────────────

SPECIES_RE = re.compile(
    # Name: 1–70 chars starting with uppercase (German, may contain spaces,
    #        brackets, dots as in "Sib. Tannenhäher")
    r'^((?:[A-ZÄÖÜ\[][^:\n]{1,70}?))\s*:+\s*'
    # Refs: one or more Roman+page tokens
    r'((?:(?:' + ROMAN_PAT + r')\s*,?\s*\d[\d\s,;.\[\]?]*)+)',
    re.MULTILINE
)

def is_family_header(name):
    return bool(re.fullmatch(r'[A-Z][a-z]{3,}(?:\s+[A-Z][a-z]{3,})?', name))

def parse_entries(text):
    """
    Return list of (name, refs_raw) from a cleaned, reassembled page text.
    """
    entries = []
    for m in SPECIES_RE.finditer(text):
        name = m.group(1).strip().rstrip(':. -')
        refs = m.group(2).strip()
        if is_family_header(name) or len(name) < 3:
            continue
        entries.append((name, refs))
    return entries

# ── ORIGINAL-NAME LOOKUP ──────────────────────────────────────────────────────

def original_name_for(orig_text, edited_name):
    """
    Find the species name in the *original* text that sits at the same
    structural position as edited_name.
    Strategy: collect all species names from original; if exactly one differs
    from edited_name and the refs overlap, that is the original name.
    Returns '' if original name == edited name (no change needed).
    """
    orig_entries = parse_entries(reassemble(orig_text))
    orig_names = [n for n, _ in orig_entries
                  if not is_family_header(n) and len(n) >= 3]
    # Direct lookup first
    if edited_name in orig_names:
        return ''   # same name in original
    # Return first name that differs (single-entry pages)
    if len(orig_names) == 1 and orig_names[0] != edited_name:
        return orig_names[0]
    # Multi-entry page: try to match by refs similarity
    edited_vols = set(re.findall(ROMAN_PAT,
                                 next((r for n,r in
                                       parse_entries(reassemble(orig_text))
                                       if n == edited_name), '')))
    best, best_score = '', 0
    for n, r in orig_entries:
        if n == edited_name:
            continue
        vols = set(re.findall(ROMAN_PAT, r))
        score = len(edited_vols & vols)
        if score > best_score:
            best, best_score = n, score
    return best if best_score > 0 else ''

# ── MAIN PIPELINE ─────────────────────────────────────────────────────────────

def main():
    json_file = Path(JSON_PATH)
    if not json_file.exists():
        raise FileNotFoundError(
            f"'{JSON_PATH}' not found. Place it in the same folder as this script.")

    with open(json_file, encoding='utf-8') as fh:
        data = json.load(fh)

    all_keys   = [k for k in data if not k.startswith('_')]
    try:
        start_idx  = all_keys.index(INDEX_START)
    except ValueError:
        print(f"Warning: start key not found, processing all pages.")
        start_idx = 0
    index_keys = all_keys[start_idx:]
    print(f"Processing {len(index_keys)} index pages …")

    # ── Pass 1: extract all raw records ──────────────────────────────────────
    raw_records = []
    current_family = None

    for pk in index_keys:
        entry = data[pk]
        edited_raw   = entry.get('edited',   '')
        original_raw = entry.get('original', '')

        edited_clean   = clean(edited_raw)
        original_clean = clean(original_raw)

        # Update family
        fam = extract_family(edited_clean)
        if fam:
            current_family = fam

        # Reassemble wrapped lines, then extract entries
        edited_asm   = reassemble(edited_clean)
        original_asm = reassemble(original_clean)

        for name_e, refs_e in parse_entries(edited_asm):
            name_o = original_name_for(original_asm, name_e)

            raw_records.append({
                'Bird Name (German)' : name_e,
                'Original Name'      : name_o,
                'Family (Latin)'     : current_family or '',
                'References'         : rectify(refs_e),
                '_page'              : pk,
            })

    print(f"Raw entries before dedup: {len(raw_records)}")

    # ── Pass 2: merge duplicate entries (same name + family, split across pages)
    df_raw = pd.DataFrame(raw_records)

    def merge_refs(series):
        """Merge and re-rectify reference strings, removing duplicates."""
        combined = '; '.join(s for s in series if s)
        # Parse all vol,page pairs, deduplicate, sort by volume number
        pairs = re.findall(r'(' + ROMAN_PAT + r'),\s*(\d+)', combined)
        seen  = {}
        for vol, page in pairs:
            if vol not in seen:
                seen[vol] = page
        sorted_pairs = sorted(seen.items(), key=lambda x: ROMAN_TO_INT.get(x[0], 99))
        return '; '.join(f'{v}, {p}' for v, p in sorted_pairs)

    def first_nonempty(series):
        for v in series:
            if v: return v
        return ''

    df = (
        df_raw
        .groupby(['Bird Name (German)', 'Family (Latin)'], sort=False, as_index=False)
        .agg({
            'Original Name' : first_nonempty,
            'References'    : merge_refs,
            '_page'         : lambda s: ', '.join(s.unique()),
        })
        .rename(columns={'_page': 'Source Page(s)'})
    )

    # Sort by family then name
    df = df.sort_values(['Family (Latin)', 'Bird Name (German)']).reset_index(drop=True)
    df.index += 1  # 1-based index

    print(f"Unique species entries: {len(df)}\n")
    pd.set_option('display.max_colwidth', 90)
    print(df[['Bird Name (German)', 'Original Name',
              'Family (Latin)', 'References']].to_string())


    # ── Export ────────────────────────────────────────────────────────────────
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        df[['Bird Name (German)', 'Original Name',
            'Family (Latin)', 'References',
            'Source Page(s)']].to_excel(writer, sheet_name='Species Index', index=True)


        df_raw.rename(columns={'_page':'Source Page'}).to_excel(
            writer, sheet_name='Raw Entries', index=False)

    df[['Bird Name (German)', 'Original Name',
        'Family (Latin)', 'References']].to_csv(
        OUTPUT_CSV, index=True, encoding='utf-8-sig')

    print(f"\n✓  Excel  →  {OUTPUT_EXCEL}")
    print(f"✓  CSV    →  {OUTPUT_CSV}")

if __name__ == '__main__':
    main()