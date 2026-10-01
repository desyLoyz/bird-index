#!/usr/bin/env python3
"""
Laubmann Index Enricher  –  v8
================================
Nimmt ein Laubmann-Index-Workbook (vorzugsweise Species Index) und reichert
es mit Daten aus dem Nomenclator (fts.txt) an.

Aktuelle Merkmale:
- Verarbeitung von .xls und .xlsx
- Fokus auf das korrigierte `Species Index`-Blatt
- Familienprüfung als zusätzliches Matching-Kriterium
- Eigene Spalte `Family Match (Source/Nomenclator)` zur Qualitätskontrolle
- Originaleinträge bleiben unverändert; zusätzlich werden die Rohdaten als
  `Raw Entries` mit ausgegeben
"""

import subprocess, sys
def _ensure(pkg):
    try: __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
_ensure("pandas"); _ensure("openpyxl"); _ensure("rapidfuzz")

import re
from pathlib import Path
import pandas as pd
from rapidfuzz import process, fuzz

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
# Standardpfad für die korrigierte XLS-Datei. Das Skript akzeptiert aber auch
# einen Dateinamen per CLI und verarbeitet automatisch .xls sowie .xlsx.
DEFAULT_EXCEL_IN = "laubmann_index_v10-corr.xls"
FTS_PATH         = "fts.txt"
MATCH_THRESHOLD  = 70

# `pd.read_excel()` benötigt für .xls zusätzlich xlrd. Für .xlsx reicht
# openpyxl, das bereits im Projekt eingebunden ist.
_ensure("xlrd")

EXCEL_IN = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL_IN
if not Path(EXCEL_IN).exists():
    for candidate in [
        Path(EXCEL_IN),
        Path(EXCEL_IN).with_suffix('.xls'),
        Path(EXCEL_IN).with_suffix('.xlsx'),
        Path(DEFAULT_EXCEL_IN),
    ]:
        if candidate.exists():
            EXCEL_IN = str(candidate)
            break

EXCEL_PATH = Path(EXCEL_IN)
OUTPUT_EXCEL = str(EXCEL_PATH.with_name(EXCEL_PATH.stem + "_nomenclator.xlsx"))
OUTPUT_CSV   = str(EXCEL_PATH.with_name(EXCEL_PATH.stem + "_nomenclator.csv"))

# ── STEP 1: PARSE NOMENCLATOR (fts.txt) ───────────────────────────────────────
def normalize_spaces(t):
    return re.sub(r'  +', ' ', t).strip()

def parse_nomenclator(path):
    text = Path(path).read_text(encoding='utf-8')

    sec_start = text.find("I.  Verzeichnis  der  mit  Sicherheit")
    sec_end   = text.find("II.  Verzeichnis  der  Vogelarten")
    if sec_end == -1:
        sec_end = text.find("II. Verzeichnis")
    if sec_end == -1:
        sec_end = len(text)
    section = text[sec_start:sec_end]

    family_re = re.compile(
        r'^([A-Z][a-z]{4,}(?:idae|inae|oidae))\s*\.\s*$',
        re.MULTILINE
    )

    results = []
    current_family = ''
    lines = section.split('\n')
    i = 0
    while i < len(lines):
        line = normalize_spaces(lines[i])

        fm = family_re.match(line)
        if fm:
            current_family = fm.group(1)
            i += 1
            continue

        nm = re.match(r'^(\d+)\.\s+(.+)', line)
        if nm:
            num  = int(nm.group(1))
            rest = nm.group(2)

            j = i + 1
            citation_lines = [rest]
            while j < len(lines):
                next_line = normalize_spaces(lines[j])
                if re.match(r'^\d+\.', next_line) or next_line == '':
                    break
                if family_re.match(next_line):
                    break
                citation_lines.append(next_line)
                j += 1

            full_text = normalize_spaces(' '.join(citation_lines))

            dash_m  = re.search(r'[—–]\s*([^—–\n]+?)(?:\s*\.|$)', full_text)
            german  = dash_m.group(1).strip().rstrip('.') if dash_m else ''

            sci_m = re.match(
                r'^([A-Z][a-zA-Zäöü]+(?:\s+[a-zA-Zäöü]+){1,3}?)'
                r'\s+(?=\(|[A-Z][a-z]+,|\[)',
                full_text
            )
            if not sci_m:
                sci_m = re.match(
                    r'^([A-Z][a-zA-Zäöü]+(?:\s+[a-zA-Zäöü]+){1,3})',
                    full_text
                )
            sci_name = sci_m.group(1).strip() if sci_m else ''

            author_m = re.search(r'\(([^)]+\d{4}[^)]*)\)', full_text)
            if not author_m:
                author_m = re.search(
                    r'([A-Z][a-z]+(?:\s+[A-Z][a-z.]+)?,\s*\d{4})',
                    full_text
                )
            author_year = author_m.group(1).strip() if author_m else ''

            terra_m = re.search(
                r'terra\s+typica[:\s]+([^;)\n.]+)',
                full_text, re.IGNORECASE
            )
            if not terra_m:
                terra_m = re.search(r'—\s*([A-Z][^;)\n]{3,30})\)', full_text)
            terra = terra_m.group(1).strip().rstrip('.)') if terra_m else ''

            citation = full_text[:dash_m.start()].strip() if dash_m else full_text

            if num and german:
                results.append({
                    'num'         : num,
                    'sci_name'    : sci_name,
                    'german_name' : german,
                    'family_nom'  : current_family,
                    'author_year' : author_year,
                    'terra_typica': terra,
                    'citation'    : citation,
                })
            i = j
            continue

        i += 1

    return results

# ── STEP 2: FUZZY MATCHING ────────────────────────────────────────────────────
def clean_name(n):
    if not isinstance(n, str): return ''
    n = n.strip().rstrip('.:')
    n = n.replace('\ufffd', '')
    n = re.sub(r'^Sib\.\s+', 'Sibirischer ', n)
    return n.lower().strip()

def build_nom_lookup(nom_list):
    return {clean_name(e['german_name']): e for e in nom_list}

def normalize_family(value):
    # Familie wird robust normalisiert: Leerwerte bleiben leer, Zeilenumbrüche
    # werden zu Leerzeichen, Punkt am Ende entfernt und alles auf Kleinbuchstaben
    # reduziert, damit Vergleiche bei Groß-/Kleinschreibung und Formatierungen
    # keine Unterschiede mehr machen.
    if pd.isna(value):
        return ''
    return re.sub(r'\s+', ' ', str(value).strip()).rstrip('.').lower()


def family_matches(source_family, nom_family):
    # Vergleich der Familienangaben aus Quelle und Nomenclator. `None` bedeutet,
    # dass beide Seiten keine Familie hatten; `False` bedeutet, dass nur eine
    # Seite eine Familie enthält oder die Familien verschieden sind.
    src = normalize_family(source_family)
    nom = normalize_family(nom_family)
    if not src and not nom:
        return None
    if not src or not nom:
        return False
    return src == nom


def match_species(bird_name, nom_lookup, source_family=None, threshold=MATCH_THRESHOLD):
    # Die Familienangabe dient als zusätzlicher Filter: Wenn im Index eine
    # passende Family vorhanden ist, werden nur noch Nomenclator-Einträge derselben
    # Familie für den eigentlichen Fuzzy-Match geprüft. Dadurch werden
    # Verwechslungen wie ähnliche deutsche Namen aus anderen Familien reduziert.
    query = clean_name(bird_name)
    if not query:
        return None

    source_family_norm = normalize_family(source_family)
    candidate_entries = list(nom_lookup.values())
    if source_family_norm:
        candidate_entries = [
            e for e in candidate_entries
            if normalize_family(e.get('family_nom')) == source_family_norm
        ]

    if not candidate_entries:
        candidate_entries = list(nom_lookup.values())

    candidate_lookup = {clean_name(e['german_name']): e for e in candidate_entries}

    if query in candidate_lookup:
        return candidate_lookup[query]

    result = process.extractOne(
        query, list(candidate_lookup.keys()),
        scorer=fuzz.token_sort_ratio,
        score_cutoff=threshold
    )
    if result:
        return candidate_lookup[result[0]]

    # Fallback to the full nomenclator if family-based filtering finds nothing.
    if candidate_lookup != nom_lookup:
        result = process.extractOne(
            query, list(nom_lookup.keys()),
            scorer=fuzz.token_sort_ratio,
            score_cutoff=threshold
        )
        if result:
            return nom_lookup[result[0]]
    return None

# ── STEP 3: PAGE-SORT KEY ─────────────────────────────────────────────────────
def page_sort_key(page_str):
    if not isinstance(page_str, str):
        return (9999, 'Z')
    first = page_str.split(',')[0].strip()
    m = re.search(r'_(\d{4})_(L|R)$', first)
    if m:
        return (int(m.group(1)), m.group(2))
    return (9999, 'Z')

# ── MAIN ──────────────────────────────────────────────────────────────────────
def select_sheet_names(path):
    # Die Layouts sind nicht immer identisch; daher bevorzugen wir explizit die
    # relevanten Blätter und benutzen nur die übrigen als Fallback.
    workbook = pd.ExcelFile(path)
    sheet_names = workbook.sheet_names
    preferred = ['Species Index', 'Raw Entries']
    for name in preferred:
        if name in sheet_names:
            yield name
    for name in sheet_names:
        if name not in preferred:
            yield name


def read_excel_sheet(path, sheet_names):
    # Hilfsfunktion für den Fall, dass ein Blatt im Workbook zwar vorhanden ist,
    # aber nicht immer exakt mit dem erwarteten Namen übereinstimmt.
    for name in sheet_names:
        try:
            return pd.read_excel(path, sheet_name=name)
        except ValueError:
            continue
    raise ValueError(f"Keine passende Tabelle in {path!r} gefunden: {sheet_names!r}")


def main():
    print(f"Lade Excel: {EXCEL_IN} …")
    # `Species Index` ist die korrigierte Basisdatei; `Raw Entries` dient nur als
    # Zusatzblatt und wird nicht für das Matching verwendet.
    sheet_names = list(select_sheet_names(EXCEL_IN))
    species_sheet = 'Species Index' if 'Species Index' in sheet_names else sheet_names[0]
    raw_sheet = 'Raw Entries' if 'Raw Entries' in sheet_names else None

    df_species = pd.read_excel(EXCEL_IN, sheet_name=species_sheet)
    df_raw = pd.read_excel(EXCEL_IN, sheet_name=raw_sheet) if raw_sheet else pd.DataFrame()

    if 'Unnamed: 0' in df_species.columns:
        df_species = df_species.rename(columns={'Unnamed: 0': 'Orig_Idx'})

    if 'Family (Latin)' in df_species.columns:
        # Verhindert Leerstellen bzw. führende Nachlaufzeichen in der Familienstamm-
        # Spalte, damit die spätere Familienprüfung zuverlässig funktioniert.
        df_species['Family (Latin)'] = df_species['Family (Latin)'].apply(
            lambda v: '' if pd.isna(v) else str(v).strip()
        )

    print("Parse Nomenclator …")
    nom_list   = parse_nomenclator(FTS_PATH)
    nom_lookup = build_nom_lookup(nom_list)
    print(f"  → {len(nom_list)} Einträge im Nomenclator")

    print("Sortiere nach Seiten-ID …")
    df_species['_sort_key'] = df_species['Source Page(s)'].apply(page_sort_key)
    df_species = df_species.sort_values('_sort_key').reset_index(drop=True)
    df_species.index += 1
    df_species = df_species.drop(columns=['_sort_key', 'Orig_Idx'], errors='ignore')

    print("Matche Arten …")
    # Für jede Zeile in `Species Index` wird ein Nomenclator-Eintrag gesucht. Neben
    # dem deutschen Namen wird zusätzlich die Familie berücksichtigt, damit nur
    # passende Einträge derselben Familie in die Fuzzy-Übereinstimmung einfließen.
    nums, sci_names, dt_names_nom, author_years = [], [], [], []
    terra_typicas, citations, match_scores, family_match_flags = [], [], [], []
    unmatched = []

    for _, row in df_species.iterrows():
        bird = row['Bird Name (German)']
        source_family = row.get('Family (Latin)', '')
        match = match_species(bird, nom_lookup, source_family=source_family)
        if match:
            nums.append(match['num'])
            sci_names.append(match['sci_name'])
            dt_names_nom.append(match['german_name'])
            author_years.append(match['author_year'])
            terra_typicas.append(match['terra_typica'])
            citations.append(match['citation'])
            score = fuzz.token_sort_ratio(
                clean_name(bird), clean_name(match['german_name']))
            match_scores.append(score)
            family_match_flags.append(family_matches(source_family, match['family_nom']))
        else:
            nums.append(None)
            sci_names.append('')
            dt_names_nom.append('')
            author_years.append('')
            terra_typicas.append('')
            citations.append('')
            match_scores.append(0)
            family_match_flags.append(False if source_family else None)
            unmatched.append(bird)

    df_species.insert(0, 'Nr. (Nomenclator)', nums)
    df_species['Wissenschaftlicher Name'] = sci_names
    df_species['Dt. Name (Nomenclator)']  = dt_names_nom
    df_species['Autor & Jahr']            = author_years
    df_species['Terra typica']            = terra_typicas
    df_species['Zitation (Nomenclator)']  = citations
    df_species['Family Match (Source/Nomenclator)'] = family_match_flags
    df_species['Match-Score']             = match_scores

    print(f"\n✓ Gematcht:       {sum(1 for s in match_scores if s > 0)} / {len(df_species)}")
    if unmatched:
        print(f"✗ Nicht gematcht ({len(unmatched)}):")
        for u in unmatched:
            print(f"    - {u!r}")

    col_order = [
        'Nr. (Nomenclator)',
        'Bird Name (German)',
        'Original Name',
        'Family (Latin)',
        'Family Match (Source/Nomenclator)',
        'Wissenschaftlicher Name',
        'Dt. Name (Nomenclator)',
        'Autor & Jahr',
        'Terra typica',
        'References',
        'Source Page(s)',
        'Match-Score',
        'Zitation (Nomenclator)',
    ]
    col_order = [c for c in col_order if c in df_species.columns]

    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        df_species[col_order].to_excel(
            writer, sheet_name='Species Index', index=True, index_label='#')
        df_raw.to_excel(
            writer, sheet_name='Raw Entries', index=False)
        df_nom = pd.DataFrame(nom_list).rename(columns={
            'num'         : 'Nr.',
            'sci_name'    : 'Wissenschaftlicher Name',
            'german_name' : 'Deutscher Name',
            'family_nom'  : 'Familie',
            'author_year' : 'Autor & Jahr',
            'terra_typica': 'Terra typica',
            'citation'    : 'Zitation',
        })
        df_nom.to_excel(writer, sheet_name='Nomenclator', index=False)

    df_species[col_order].to_csv(OUTPUT_CSV, index=True, encoding='utf-8-sig')
    print(f"\n✓  Excel  →  {OUTPUT_EXCEL}")
    print(f"✓  CSV    →  {OUTPUT_CSV}")

if __name__ == '__main__':
    main()