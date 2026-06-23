#!/usr/bin/env python3
"""
Laubmann Index Enricher  –  v8
================================
Nimmt laubmann_index_v7.xlsx und reichert sie mit Daten aus dem
Nomenclator (fts.txt) an:
- Sortierung nach Seiten-ID (Tagebuch-Reihenfolge)
- Neue Spalten: Lfd. Nr. (Nomenclator), Wissenschaftlicher Name, Autor & Jahr, Terra typica
- Originaleinträge bleiben UNVERÄNDERT
"""

import subprocess, sys
def _ensure(pkg):
    try: __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
_ensure("pandas"); _ensure("openpyxl"); _ensure("rapidfuzz")

import re, json
from pathlib import Path
import pandas as pd
from rapidfuzz import process, fuzz

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
EXCEL_IN     = "laubmann_index_v8.xlsx"
FTS_PATH     = "fts.txt"
OUTPUT_EXCEL = EXCEL_IN.replace('.xlsx', '_nomenclator.xlsx')
OUTPUT_CSV   = EXCEL_IN.replace('.xlsx', '_nomenclator.csv')   
MATCH_THRESHOLD = 70  # Fuzzy-Score-Schwelle (0–100)

# ── STEP 1: PARSE NOMENCLATOR (fts.txt) ───────────────────────────────────────
def normalize_spaces(t):
    """OCR-Text hat oft doppelte Leerzeichen — normalisieren."""
    return re.sub(r'  +', ' ', t).strip()

def parse_nomenclator(path):
    """
    Parst den Abschnitt I. des Nomenclators.
    Gibt Liste von Dicts zurück:
      {
        'num':          int,           # Laufende Nummer
        'sci_name':     str,           # Wissenschaftlicher Name (trinomial)
        'german_name':  str,           # Deutscher Name
        'family':       str,           # Familienname (Latin)
        'citation':     str,           # Volle Zitationszeile (Author, Jahr, ...)
        'author_year':  str,           # Kurzform: "Linnaeus, 1758"
        'terra_typica': str,           # Terra typica wenn vorhanden
      }
    """
    text = Path(path).read_text(encoding='utf-8')

    # Relevanten Abschnitt extrahieren
    sec_start = text.find("I.  Verzeichnis  der  mit  Sicherheit")
    sec_end   = text.find("II.  Verzeichnis  der  Vogelarten")
    if sec_end == -1:
        sec_end = text.find("II. Verzeichnis")
    if sec_end == -1:
        sec_end = len(text)
    section = text[sec_start:sec_end]

    # Familiennamen erkennen: Zeile endet mit " ." oder ist nur ein Wort + Punkt
    # z.B. "Corvidae ." oder "Fringillidae."
    family_re = re.compile(
        r'^([A-Z][a-z]{4,}(?:idae|inae|oidae))\s*\.\s*$',
        re.MULTILINE
    )

    # Eintragszeile: "N.  Genus species subspecies Author  —  Deutscher Name."
    # OCR hat doppelte Leerzeichen, daher \s+ überall
    entry_re = re.compile(
        r'^(\d+)\.\s+'                          # Nummer
        r'([A-Z][a-z]+(?:\s+[a-z]+){1,3})'     # Wissenschaftlicher Name (2–4 Teile)
        r'\s+(?:\([^)]+\)|[A-Z][a-z.]+(?:\s+[A-Z][a-z.]+)?)\s*'  # Autor (geklammert oder nicht)
        r'[.—–-]+\s*'                           # Trennzeichen
        r'(.+?)\s*\.',                          # Deutscher Name
        re.MULTILINE
    )

    # Einfachere, robustere Variante:
    # Zeile beginnt mit Zahl + Punkt, enthält "—" als Trennzeichen
    entry_re2 = re.compile(
        r'^(\d+)\.\s+'                          # Nummer
        r'([A-Z][a-zA-Z]+(?:\s+[a-zA-Z]+){1,3}?)'  # Wissenschaftlicher Name
        r'\s+[A-Z(].*?'                         # Autor (beginnt mit Großbuchstabe oder Klammer)
        r'[—–]\s*'                              # Gedankenstrich
        r'([^\n.]+)',                           # Deutscher Name (bis Zeilenende oder Punkt)
        re.MULTILINE
    )

    results = []
    current_family = ''

    # Zeilenweise verarbeiten für bessere Kontrolle
    lines = section.split('\n')
    i = 0
    while i < len(lines):
        line = normalize_spaces(lines[i])

        # Familienname?
        fm = family_re.match(line)
        if fm:
            current_family = fm.group(1)
            i += 1
            continue

        # Eintragsnummer am Zeilenanfang?
        nm = re.match(r'^(\d+)\.\s+(.+)', line)
        if nm:
            num = int(nm.group(1))
            rest = nm.group(2)

            # Nächste Zeile(n) sammeln bis zur nächsten Nummer oder Leerzeile
            j = i + 1
            citation_lines = [rest]
            while j < len(lines):
                next_line = normalize_spaces(lines[j])
                if re.match(r'^\d+\.', next_line) or next_line == '':
                    break
                # Familienname → auch abbrechen
                if family_re.match(next_line):
                    break
                citation_lines.append(next_line)
                j += 1

            full_text = ' '.join(citation_lines)
            full_text = normalize_spaces(full_text)

            # Deutschen Namen extrahieren: nach "—" oder "–"
            dash_m = re.search(r'[—–]\s*([^—–\n]+?)(?:\s*\.|$)', full_text)
            german = dash_m.group(1).strip().rstrip('.') if dash_m else ''

            # Wissenschaftlichen Namen extrahieren: alles vor dem ersten Autor-Token
            # Autor beginnt mit Großbuchstabe nach dem wiss. Namen, oder in Klammern
            sci_m = re.match(
                r'^([A-Z][a-zA-Zäöü]+(?:\s+[a-zA-Zäöü]+){1,3}?)'
                r'\s+(?=\(|[A-Z][a-z]+,|\[)',
                full_text
            )
            if not sci_m:
                # Fallback: alles bis zum ersten Großbuchstaben-Wort nach Leerzeichen
                sci_m = re.match(
                    r'^([A-Z][a-zA-Zäöü]+(?:\s+[a-zA-Zäöü]+){1,3})',
                    full_text
                )
            sci_name = sci_m.group(1).strip() if sci_m else ''

            # Autor & Jahr: erste Klammer oder "Autor, Jahr" Muster
            author_m = re.search(
                r'\(([^)]+\d{4}[^)]*)\)',
                full_text
            )
            if not author_m:
                author_m = re.search(
                    r'([A-Z][a-z]+(?:\s+[A-Z][a-z.]+)?,\s*\d{4})',
                    full_text
                )
            author_year = author_m.group(1).strip() if author_m else ''

            # Terra typica
            terra_m = re.search(
                r'terra\s+typica[:\s]+([^;)\n.]+)',
                full_text, re.IGNORECASE
            )
            if not terra_m:
                # Manchmal: "— Ort)." am Ende der Klammer
                terra_m = re.search(
                    r'—\s*([A-Z][^;)\n]{3,30})\)',
                    full_text
                )
            terra = terra_m.group(1).strip().rstrip('.)') if terra_m else ''

            # Zitation: alles nach dem wiss. Namen bis zum deutschen Namen
            citation = full_text
            if dash_m:
                citation = full_text[:dash_m.start()].strip()

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
    """Normalisiere deutschen Namen für Matching."""
    if not isinstance(n, str): return ''
    n = n.strip().rstrip('.:')
    # OCR-Ersatzzeichen \ufffd → ignorieren beim Matching
    n = n.replace('\ufffd', '')
    # Sib. Tannenhäher → Sibirischer Tannenhäher
    n = re.sub(r'^Sib\.\s+', 'Sibirischer ', n)
    return n.lower().strip()

def build_nom_lookup(nom_list):
    """Erstelle Dict: clean_name → nom_entry"""
    lookup = {}
    for e in nom_list:
        key = clean_name(e['german_name'])
        lookup[key] = e
    return lookup

def match_species(bird_name, nom_lookup, threshold=MATCH_THRESHOLD):
    """
    Versucht, einen deutschen Vogelnamen im Nomenclator zu finden.
    Gibt das beste Match zurück oder None.
    """
    query = clean_name(bird_name)
    if not query:
        return None

    # Direkter Treffer
    if query in nom_lookup:
        return nom_lookup[query]

    # Fuzzy-Suche
    keys = list(nom_lookup.keys())
    result = process.extractOne(
        query, keys,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=threshold
    )
    if result:
        matched_key, score, _ = result
        return nom_lookup[matched_key]

    return None

# ── STEP 3: PAGE-SORT KEY ─────────────────────────────────────────────────────
def page_sort_key(page_str):
    """
    Extrahiert Seitennummer aus UUID_NNNN_L/R für Sortierung.
    Mehrere Seiten: nimmt die erste.
    """
    if not isinstance(page_str, str):
        return (9999, 'Z')
    first = page_str.split(',')[0].strip()
    m = re.search(r'_(\d{4})_(L|R)$', first)
    if m:
        return (int(m.group(1)), m.group(2))
    return (9999, 'Z')

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    # 1. Excel laden
    print("Lade Excel …")
    df_species = pd.read_excel(EXCEL_IN, sheet_name='Species Index')
    df_raw     = pd.read_excel(EXCEL_IN, sheet_name='Raw Entries')

    # Spalte umbenennen falls nötig
    if 'Unnamed: 0' in df_species.columns:
        df_species = df_species.rename(columns={'Unnamed: 0': 'Orig_Idx'})

    # 2. Nomenclator parsen
    print("Parse Nomenclator …")
    nom_list   = parse_nomenclator(FTS_PATH)
    nom_lookup = build_nom_lookup(nom_list)
    print(f"  → {len(nom_list)} Einträge im Nomenclator gefunden")

    # Debug: erste 10 Einträge zeigen
    for e in nom_list[:10]:
        print(f"  [{e['num']:3d}] {e['german_name']:30s} | {e['sci_name']:40s} | {e['author_year']}")

    # 3. Species Index nach Seite sortieren
    print("\nSortiere nach Seiten-ID …")
    df_species['_sort_key'] = df_species['Source Page(s)'].apply(page_sort_key)
    df_species = df_species.sort_values('_sort_key').reset_index(drop=True)
    df_species.index += 1
    df_species = df_species.drop(columns=['_sort_key', 'Orig_Idx'], errors='ignore')

    # 4. Nomenclator-Spalten hinzufügen
    print("Matche Arten mit Nomenclator …")
    nums, sci_names, author_years, terra_typicas, citations, match_scores = \
        [], [], [], [], [], []

    unmatched = []
    for _, row in df_species.iterrows():
        bird = row['Bird Name (German)']
        match = match_species(bird, nom_lookup)
        if match:
            nums.append(match['num'])
            sci_names.append(match['sci_name'])
            author_years.append(match['author_year'])
            terra_typicas.append(match['terra_typica'])
            citations.append(match['citation'])
            # Score berechnen
            score = fuzz.token_sort_ratio(clean_name(bird), clean_name(match['german_name']))
            match_scores.append(score)
        else:
            nums.append(None)
            sci_names.append('')
            author_years.append('')
            terra_typicas.append('')
            citations.append('')
            match_scores.append(0)
            unmatched.append(bird)

    df_species.insert(0, 'Nr. (Nomenclator)', nums)
    df_species['Wissenschaftlicher Name'] = sci_names
    df_species['Autor & Jahr']            = author_years
    df_species['Terra typica']            = terra_typicas
    df_species['Zitation (Nomenclator)']  = citations
    df_species['Match-Score']             = match_scores

    print(f"\n✓ Gematcht: {sum(1 for s in match_scores if s > 0)} / {len(df_species)}")
    if unmatched:
        print(f"✗ Nicht gematcht ({len(unmatched)}):")
        for u in unmatched:
            print(f"    - {u!r}")

    # 5. Export
    print(f"\nExportiere nach {OUTPUT_EXCEL} …")
    col_order = [
        'Nr. (Nomenclator)',
        'Bird Name (German)',
        'Original Name',
        'Family (Latin)',
        'Wissenschaftlicher Name',
        'Autor & Jahr',
        'Terra typica',
        'References',
        'Source Page(s)',
        'Match-Score',
        'Zitation (Nomenclator)',
    ]
    # Nur vorhandene Spalten
    col_order = [c for c in col_order if c in df_species.columns]

    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        df_species[col_order].to_excel(
            writer, sheet_name='Species Index', index=True, index_label='#')
        df_raw.to_excel(
            writer, sheet_name='Raw Entries', index=False)

        # Zusatz-Sheet: Nomenclator-Referenz
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

    df_species[col_order].to_csv(
        OUTPUT_CSV, index=True, encoding='utf-8-sig')

    print(f"✓  Excel  →  {OUTPUT_EXCEL}")
    print(f"✓  CSV    →  {OUTPUT_CSV}")

if __name__ == '__main__':
    main()