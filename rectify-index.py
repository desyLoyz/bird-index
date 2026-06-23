#!/usr/bin/env python3
"""
Laubmann General Index Parser  –  v7
=====================================
Kernstrategie: Arbeite DIREKT auf dem <u>-Token-Stream.
Jeder bedeutsame Begriff (Name, Bandnummer, Familienname) ist in <u>...</u>
eingeschlossen. Die Seitenzahlen stehen AUSSERHALB der Tags (als plain text).

Fixes gegenüber v6:
- is_family() prüft NUR gegen die explizite Whitelist (kein Regex-Overmatch)
- Seitenzahlen werden aus dem Text ZWISCHEN den <u>-Tags extrahiert
- Malformed tags (<u>Text,/u>) werden korrekt behandelt
- Bullet-Zeichen vor Tags werden ignoriert
- <u>Garten u. Waldbaumläufer</u> wird als ein Token erkannt
- Original Name wird IMMER befüllt (auch wenn identisch mit edited)
"""

import subprocess, sys
def _ensure(pkg):
    try: __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
_ensure("pandas"); _ensure("openpyxl")

import json, re
from pathlib import Path
import pandas as pd

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
JSON_PATH    = "Laubmann_35_gemini_edits-170.json"
OUTPUT_EXCEL = "laubmann_index_v7.xlsx"
OUTPUT_CSV   = "laubmann_index_v7.csv"
INDEX_START  = "2d6d52f0-b55f-431d-81ac-e2b8ce255fb2_0004_R"

ROMAN_LIST = [
    "XXXV","XXXIV","XXXIII","XXXII","XXXI","XXX",
    "XXIX","XXVIII","XXVII","XXVI","XXV","XXIV",
    "XXIII","XXII","XXI","XX","XIX","XVIII","XVII",
    "XVI","XV","XIV","XIII","XII","XI","X",
    "IX","VIII","VII","VI","V","IV","III","II","I"
]
ROMAN_PAT    = "|".join(ROMAN_LIST)
ROMAN_TO_INT = {r: i+1 for i, r in enumerate(reversed(ROMAN_LIST))}

# EXPLICIT whitelist only — no regex overmatch!
LATIN_FAMILIES = {
    'Alaudidae','Anatidae','Ardeidae','Bombycillidae','Certhiidae',
    'Charadriidae','Columbidae','Colymbidae','Corvidae','Falconidae',
    'Fringillidae','Hirundinidae','Laridae','Lanidae','Laniidae',
    'Motacillidae','Muscicapidae','Paridae','Pelecanidae',
    'Phalacrocoracidae','Phasianidae','Phoenicopteridae','Picidae',
    'Plegadidae','Procellariidae','Rallidae','Scolopacidae',
    'Strigidae','Sylviidae','Troglodytidae','Vulturidae',
    'Accipitridae','Apodidae','Caprimulgidae','Ciconiidae',
    'Cuculidae','Emberizidae','Gaviidae','Gruidae','Haematopodidae',
    'Otididae','Passeridae','Podicipedidae','Recurvirostridae',
    'Sturnidae','Threskiornithidae','Turdidae','Upupidae',
    'Alcidae','Hierapotidae','Regulidae',
}

# ── TOKEN EXTRACTION ──────────────────────────────────────────────────────────
# Key insight: page numbers are OUTSIDE <u> tags, between them.
# Structure: <u>NAME</u>: <u>VOL</u>, PAGE; <u>VOL</u>, PAGE; ...
# So we extract (tag_content, following_plain_text) pairs.

def extract_tag_pairs(raw):
    """
    Returns list of (tag_text, text_after_tag) tuples.
    text_after_tag = plain text until the next < character.
    Also handles malformed <u>TEXT,/u> tags.
    """
    pairs = []
    # Match both well-formed and malformed closing tags
    pattern = re.compile(r'<u>(.*?)(?:</u>|,?/u>)', re.DOTALL)
    last_end = 0
    for m in pattern.finditer(raw):
        tag_text = m.group(1).strip()
        # Text between this tag's end and the next tag's start
        after_start = m.end()
        next_tag = raw.find('<', after_start)
        if next_tag == -1:
            after_text = raw[after_start:]
        else:
            after_text = raw[after_start:next_tag]
        pairs.append((tag_text, after_text))
    return pairs

def is_roman(t):
    return bool(re.fullmatch(ROMAN_PAT, t.strip()))

def is_family(t):
    return t.strip().rstrip('.') in LATIN_FAMILIES

def is_name_token(t):
    t = t.strip()
    if not t: return False
    if is_roman(t): return False
    if is_family(t): return False
    if re.fullmatch(r'\d+', t): return False
    # Must start with uppercase (incl. umlauts and \ufffd OCR replacement)
    if not re.match(r'^[A-ZÄÖÜ\ufffd\[]', t): return False
    if len(t) < 2: return False
    return True

# ── ENTRY PARSING FROM TAG PAIRS ─────────────────────────────────────────────
def parse_entries_from_pairs(pairs):
    """
    Parse (tag_text, after_text) pairs into (name, refs) entries.
    
    Page numbers are in after_text: e.g. tag='I', after=', 153; '
    So for each Roman numeral tag, the page is the first number in after_text.
    """
    entries = []
    i = 0
    n = len(pairs)
    current_family = None

    while i < n:
        tag, after = pairs[i]

        if is_family(tag):
            current_family = tag.rstrip('.')
            i += 1
            continue

        if is_name_token(tag):
            # Look ahead: is there a Roman numeral within next 3 tags?
            lookahead_tags = [pairs[j][0] for j in range(i+1, min(i+4, n))]
            if not any(is_roman(t) for t in lookahead_tags):
                i += 1
                continue

            name = tag
            i += 1
            ref_pairs = []

            while i < n:
                t, a = pairs[i]

                if is_roman(t):
                    # Page number is first digit sequence in `a`
                    page_m = re.search(r'\d+', a)
                    if page_m:
                        ref_pairs.append((t, page_m.group()))
                    i += 1

                elif is_family(t):
                    current_family = t.rstrip('.')
                    i += 1
                    break

                elif is_name_token(t):
                    # New entry starts if followed by Roman numeral
                    la2 = [pairs[j][0] for j in range(i+1, min(i+4, n))]
                    if any(is_roman(x) for x in la2):
                        break
                    else:
                        i += 1  # noise token, skip

                else:
                    i += 1

            if ref_pairs:
                refs_str = '; '.join(f'{v}, {p}' for v, p in ref_pairs)
                entries.append((name, refs_str, current_family or ''))

        else:
            i += 1

    return entries

# ── ORIGINAL NAME LOOKUP ──────────────────────────────────────────────────────
def get_original_name(original_raw, position):
    """
    Extract names from original token stream at ordinal position.
    Always returns the name (even if same as edited).
    """
    pairs = extract_tag_pairs(original_raw)
    orig_entries = parse_entries_from_pairs(pairs)
    if position < len(orig_entries):
        return orig_entries[position][0]
    return ''

# ── REFERENCE MERGING ─────────────────────────────────────────────────────────
def merge_refs(series):
    combined = '; '.join(s for s in series if s)
    pairs = re.findall(r'(' + ROMAN_PAT + r'),\s*(\d+)', combined)
    seen = {}
    for vol, page in pairs:
        if vol not in seen:
            seen[vol] = page
    sorted_pairs = sorted(seen.items(), key=lambda x: ROMAN_TO_INT.get(x[0], 99))
    return '; '.join(f'{v}, {p}' for v, p in sorted_pairs)

def first_nonempty(series):
    for v in series:
        if v: return v
    return ''

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    json_file = Path(JSON_PATH)
    if not json_file.exists():
        raise FileNotFoundError(f"'{JSON_PATH}' nicht gefunden.")

    with open(json_file, encoding='utf-8') as fh:
        data = json.load(fh)

    all_keys  = [k for k in data if not k.startswith('_')]
    try:
        start_idx = all_keys.index(INDEX_START)
    except ValueError:
        print("Warnung: Start-Key nicht gefunden.")
        start_idx = 0
    index_keys = all_keys[start_idx:]
    print(f"Verarbeite {len(index_keys)} Index-Seiten …")

    raw_records = []

    for pk in index_keys:
        entry        = data[pk]
        edited_raw   = entry.get('edited',   '')
        original_raw = entry.get('original', '')

        edited_pairs = extract_tag_pairs(edited_raw)
        page_entries = parse_entries_from_pairs(edited_pairs)

        for pos, (name_e, refs_e, family_e) in enumerate(page_entries):
            name_o = get_original_name(original_raw, pos)
            raw_records.append({
                'Bird Name (German)' : name_e,
                'Original Name'      : name_o,
                'Family (Latin)'     : family_e,
                'References'         : refs_e,
                '_page'              : pk,
            })

    print(f"Rohdaten vor Dedup: {len(raw_records)}")

    df_raw = pd.DataFrame(raw_records)

    # Merge split entries across pages
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

    df = df.sort_values(['Family (Latin)', 'Bird Name (German)']).reset_index(drop=True)
    df.index += 1

    print(f"Eindeutige Arten: {len(df)}\n")
    pd.set_option('display.max_colwidth', 120)
    print(df[['Bird Name (German)', 'Original Name',
              'Family (Latin)', 'References']].to_string())

    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        df[['Bird Name (German)', 'Original Name',
            'Family (Latin)', 'References',
            'Source Page(s)']].to_excel(
            writer, sheet_name='Species Index', index=True)
        df_raw.rename(columns={'_page': 'Source Page'}).to_excel(
            writer, sheet_name='Raw Entries', index=False)

    df[['Bird Name (German)', 'Original Name',
        'Family (Latin)', 'References']].to_csv(
        OUTPUT_CSV, index=True, encoding='utf-8-sig')

    print(f"\n✓  Excel  →  {OUTPUT_EXCEL}")
    print(f"✓  CSV    →  {OUTPUT_CSV}")

if __name__ == '__main__':
    main()