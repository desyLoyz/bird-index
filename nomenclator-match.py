#!/usr/bin/env python3
"""
Laubmann Index Enricher  –  v9
================================
Nimmt ein Laubmann-Index-Workbook (vorzugsweise Species Index) und reichert
es mit Daten aus dem Nomenclator (Hellmayr & Laubmann 1916, fts.txt aus BHL) an.

Änderungen gegenüber v8 (Ursachen der Fehlzuordnungen behoben):
- Parser: Zeilenumbruch-Trennungen ("Kirsch¬ / kernbeißer", "Waldlaub- / sänger")
  werden zusammengefügt; Fußnotenzeichen ("2)", "*)") entfernt.
- Parser: Familienüberschriften werden OCR-tolerant gegen eine feste Liste der
  50 Nomenclator-Familien erkannt (v8 erkannte nur 28 und vererbte sonst die
  vorherige Familie, z. B. alle Fringillidae als "Oriolidae").
- Parser: Abschnitt I und II werden getrennt erkannt (v8 fand das Ende von
  Abschnitt I nicht und vergab für Abschnitt II die Nummern 1–14 doppelt).
  Jeder Eintrag erhält eine eindeutige ID ("I-14", "II-7").
- Parser: Nummernfolge wird geprüft; OCR-Fehler (z. B. "58." statt 53) werden
  korrigiert, Einträge ohne Punkt nach der Nummer ("94 Acrocephalus") erkannt,
  Einträge ohne deutschen Namen nicht mehr verworfen. Alles im Parser-Report.
- Matching: Reihenfolge exakt (normalisiert) → Synonymtabelle → Fuzzy.
  Fuzzy-Treffer werden NICHT mehr in Nr./Wissenschaftlicher Name übernommen,
  sondern nur als Kandidat zur Prüfung ausgegeben. Kein Familien-Hartfilter.
- Hybride/Sammeleinträge ("Grünfink [x] Zeisig", "Tafelente + Reiherente")
  werden erkannt und nicht zugeordnet.
- Neue Spalten: Nomenclator-ID, Abschnitt, Match-Methode, Kandidaten,
  Mehrfachvergabe. Neues Blatt "Parser-Report".
"""

import subprocess, sys
def _ensure(pkg):
    try: __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
_ensure("pandas"); _ensure("openpyxl"); _ensure("rapidfuzz")

import re
from collections import Counter
from pathlib import Path
import pandas as pd
from rapidfuzz import process, fuzz

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
DEFAULT_EXCEL_IN = "laubmann_index_v14.xlsx"
FTS_PATH         = "fts.txt"
SYNONYMS_CSV     = "nomenclator_synonyms.csv"   # optional, ergänzt SYNONYMS
FUZZY_MIN        = 88   # Mindestscore (fuzz.ratio) für einen Fuzzy-Kandidaten
FUZZY_MARGIN     = 3    # Abstand zum zweitbesten Kandidaten

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

# Familien des Nomenclators in der Reihenfolge des Buches (Inhaltsübersicht).
NOMENCLATOR_FAMILIES = [
    "Corvidae", "Sturnidae", "Oriolidae", "Fringillidae", "Alaudidae",
    "Motacillidae", "Certhiidae", "Sittidae", "Paridae", "Regulidae",
    "Laniidae", "Bombycillidae", "Muscicapidae", "Sylviidae", "Prunellidae",
    "Troglodytidae", "Hirundinidae", "Micropodidae", "Caprimulgidae",
    "Meropidae", "Upupidae", "Coraciidae", "Alcedinidae", "Cuculidae",
    "Picidae", "Strigidae", "Falconidae", "Vulturidae", "Ciconiidae",
    "Plegadidae", "Plataleidae", "Ardeidae", "Phoenicopteridae", "Anatidae",
    "Mergidae", "Phalacrocoracidae", "Pelecanidae", "Procellariidae",
    "Colymbidae", "Charadriidae", "Scolopacidae", "Laridae", "Alcidae",
    "Otididae", "Megalornithidae", "Rallidae", "Columbidae", "Pteroclidae",
    "Phasianidae", "Tetraonidae",
]
_FAMILY_KEYS = [f.lower() for f in NOMENCLATOR_FAMILIES]

# Historische / abweichende deutsche Namen im Index → Nomenclator-ID(s).
# Mehrere IDs = mehrdeutig (wird nicht automatisch übernommen).
SYNONYMS = {
    "Wasseramsel":            (["I-124"], "Nomenclator: Wasserschmätzer; I-124 = mitteldeutsche Form medius, I-125/126 weitere Unterarten"),
    "Ohrentaucher":           (["I-240"], "Nomenclator: Ohrensteißfuß"),
    "Zwergtaucher":           (["I-243"], "Nomenclator: Zwergsteißfuß"),
    "Mönchsgrasmücke":        (["I-98"],  "Nomenclator: Schwarzplättchen"),
    "Große Bekassine":        (["I-280"], "Nomenclator: Große Sumpfschnepfe"),
    "Kleine Bekassine":       (["I-282"], "Nomenclator: Kleine Sumpfschnepfe"),
    "Brachvogel":             (["I-278"], "Nomenclator: Großer Brachvogel"),
    "Rostrote Uferschnepfe":  (["I-276"], "Nomenclator: Pfuhlschnepfe"),
    "Ohrenlerche":            (["I-48"],  "Nomenclator: Alpenlerche"),
    "Weidenlaubvogel":        (["I-84"],  "Nomenclator: Weidenlaubsänger"),
    "Waldlaubvogel":          (["I-87"],  "Nomenclator: Waldlaubsänger"),
    "Fitislaubvogel":         (["I-85"],  "Nomenclator: Fitis-Laubsänger"),
    "Berglaubvogel":          (["I-86"],  "Nomenclator: Berglaubsänger"),
    "Binsenrohrsänger":       (["I-94"],  "Nomenclator: Wasserrohrsänger (Acrocephalus aquaticus)"),
    "Kormoran":               (["I-235"], "Nomenclator: Kormoranscharbe"),
    "Storch":                 (["I-189"], "Nomenclator: Weißer Storch"),
    "Schwarzstorch":          (["I-190"], "Nomenclator: Schwarzer Storch"),
    "Brandgans":              (["I-211"], "Nomenclator: Brandente"),
    "Teichhuhn":              (["I-313"], "Nomenclator: Grünfüßiges Teichhuhn"),
    "Weißwangengans":         (["I-210"], "Nomenclator: Nonnengans"),
    "Silberreiher":           (["I-195"], "Nomenclator: Edelreiher"),
    "Ringdrossel":            (["I-105"], "Nomenclator: Nordische Ringdrossel"),
    "Weißstern-Blaukehlchen": (["I-119"], "Nomenclator führt nur L. svecica cyanecula"),
    "Gartenbaumläufer":       (["I-60"],  "OCR im Nomenclator: 'Garten- banmlänfer'"),
    "Rotkehlpieper":          (["I-53"],  "Nomenclator: Rotkehliger Pieper (OCR 'Kotkehliger')"),
    "Wüstenrennvogel":        (["I-249"], "Nomenclator: Rennvogel"),
    "Rotrückenwürger":        (["I-78"],  "Nomenclator: Rotrückiger Würger"),
    "Brachwasserläufer":      (["I-268"], "Nomenclator: Bruchwasserläufer"),
    "Frauenfliegenschnäpper": (["I-81"],  "vermutlich Lesefehler für Trauerfliegenschnäpper"),
    "Nebelkrähe":             (["I-2"],   "OCR im Nomenclator: 'Nebelkrätie'"),
    "Zwergsumpfhuhn":         (["I-311"], "OCR im Nomenclator: 'Zwergsumpfhuhu'"),
    "Sturmschwalbe":          (["I-238"], "Nomenclator: Kleine Sturmschwalbe"),
    "Meerstrandläufer":       (["I-264"], "Nomenclator: Seestrandläufer"),
    "Gimpel":                 (["I-25", "I-26"], "Großer oder Kleiner Gimpel – Unterart nicht bestimmbar"),
    "Weidenmeise":            (["I-69", "I-70", "II-6"], "drei Weidenmeisen-Formen im Nomenclator"),
}

# Abkürzungen im Index
ABBREVIATIONS = {
    r'^sib\.\s*':          'sibirischer ',
    r'^nord\.\s*':         'nordische ',
    r'^südl\.\s*':         'südliche ',
    r'^isländ\.\s*':       'isländischer ',
    r'weißbärt\.\s*-?':    'weißbärtige ',
    r'schnäbl\.\s*':       'schnabel ',
}

HYBRID_RE = re.compile(r'\[x\]|\s[x×]\s|\+|&|\su\.\s', re.IGNORECASE)
KNOWN_HYBRIDS = {'rackelhuhn', 'bastardkrähe'}   # Tetrao × Lyrurus, Corvus corone × cornix

# ── STEP 1: PARSE NOMENCLATOR (fts.txt) ───────────────────────────────────────
def normalize_spaces(t):
    return re.sub(r'\s+', ' ', t).strip()

# Fußnotenzeichen: "Brehm1)", "Tunst.* 2)", "Billb. *)", "ente 1)"
FOOTNOTE_RE = re.compile(r"(?:(?<=[A-Za-zäöüß.'’)])|\s+)\d\)|\s*\*\s*\)?")

def strip_footnotes(t):
    return FOOTNOTE_RE.sub('', FOOTNOTE_RE.sub('', t))

def detect_family(line):
    """Gibt den kanonischen Familiennamen zurück, wenn die Zeile eine
    (ggf. OCR-beschädigte) Familienüberschrift ist, sonst None."""
    s = strip_footnotes(line).strip().rstrip('.,:; ')
    if not s or len(s) > 30 or re.match(r'^\d', s):
        return None
    compact = re.sub(r'[^A-Za-zÄÖÜäöü]', '', s).lower()
    if not 6 <= len(compact) <= 20:
        return None
    hit = process.extractOne(compact, _FAMILY_KEYS, scorer=fuzz.ratio, score_cutoff=80)
    return NOMENCLATOR_FAMILIES[hit[2]] if hit else None

ENTRY_RE = re.compile(r'^(\d{1,3})(\.?)\s+([A-ZÄÖÜßa-zäöü][^\s]*)\s+\S')

def join_wrapped(a, b):
    """Fügt eine umbrochene Kopfzeile zusammen und löst Trennungen auf."""
    a = a.rstrip()
    if a.endswith(('¬', '-')):
        a = a[:-1]
        return a + b if b[:1].islower() else a + '-' + b
    return a + ' ' + b

def header_complete(h):
    h = strip_footnotes(h).rstrip(" '’\"")
    return h.endswith('.')

def clean_german(g):
    g = strip_footnotes(g)
    g = re.sub(r'\s+[a-z]\)\W*$', '', g)                    # "Sturmvogel s)" (OCR für 3))
    g = re.sub(r'(\w)¬\s*(?=[a-zäöüß])', r'\1', g)        # "Rheini¬ scher"
    g = re.sub(r'(\w)-\s+(?=[a-zäöüß])', r'\1', g)        # "Garten- banm…"
    g = re.sub(r'(\w)-\s+(?=[A-ZÄÖÜ])', r'\1-', g)        # "Breitschnabel- Wassertreter"
    g = re.sub(r'\b([A-ZÄÖÜ]) (?=[a-zäöüß]{3,})', r'\1', g)  # "W asserschmätzer"
    g = g.strip().rstrip(".'’,;: ")
    return normalize_spaces(g)

def parse_sci(before_dash):
    """Trennt wissenschaftlichen Namen und Autor aus dem Teil vor dem Gedankenstrich."""
    s = strip_footnotes(before_dash)
    tokens = s.split()
    if not tokens:
        return '', ''
    name = [tokens[0]]
    i = 1
    while i < len(tokens) and len(name) < 3 and re.fullmatch(r'[A-Za-zäöüß]+', tokens[i]):
        name.append(tokens[i].lower())
        i += 1
    author = ' '.join(tokens[i:]).strip().rstrip('.') if i < len(tokens) else ''
    return ' '.join(name), author

def find_section(text, patterns, required=True):
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.start()
    if required:
        raise ValueError(f"Abschnittsmarke nicht gefunden: {patterns}")
    return None

def parse_nomenclator(path):
    text = Path(path).read_text(encoding='utf-8')
    s1 = find_section(text, [r'I\.\s+Verzeichnis\s+der\s+mit\s+Sicherheit'])
    s2 = find_section(text, [r'II\.\s+Verzeichnis\s+jener', r'II\.\s+Verzeichnis'])
    s3 = find_section(text, [r'III\.\s+Nomenclator\s+der\s+Gattungsnamen',
                             r'III\.\s+Nomenclator'], required=False) or len(text)
    sections = [('I', text[s1:s2]), ('II', text[s2:s3])]

    results, report = [], []
    for sec, body in sections:
        lines = [normalize_spaces(l) for l in body.split('\n')]
        raw = []                       # (zeile, gedruckte nr, familie)
        family = ''
        last_printed = 0
        for idx, line in enumerate(lines[1:], start=1):
            fam = detect_family(line)
            if fam and not ENTRY_RE.match(line):
                family = fam
                continue
            m = ENTRY_RE.match(line)
            if not m:
                continue
            num, dot = int(m.group(1)), m.group(2)
            if not dot and num != last_printed + 1:
                continue               # Fußnotenzeile wie "9  Wie schon …"
            if not dot:
                report.append((sec, num, 'Eintrag ohne Punkt nach der Nummer erkannt', line[:60]))
            raw.append((idx, num, family))
            last_printed = num

        # Nummernfolge prüfen und OCR-Fehler korrigieren
        expected = 1
        for k, (idx, printed, fam) in enumerate(raw):
            num = printed
            if printed != expected:
                nxt = raw[k + 1][1] if k + 1 < len(raw) else None
                if nxt == expected + 1:
                    report.append((sec, expected, f'Nummer OCR-korrigiert ({printed} → {expected})', lines[idx][:60]))
                    num = expected
                else:
                    report.append((sec, printed, f'Lücke in der Nummernfolge (erwartet {expected})', lines[idx][:60]))
            expected = num + 1

            # Kopfzeile (ggf. umbrochen) und Zitatzeilen sammeln
            end = raw[k + 1][0] if k + 1 < len(raw) else len(lines)
            header = re.sub(r'^\d{1,3}\.?\s+', '', lines[idx])
            j = idx + 1
            while not header_complete(header) and j < end and j - idx <= 2:
                if not lines[j] or detect_family(lines[j]) or lines[j].startswith('['):
                    break
                header = join_wrapped(header, lines[j])
                j += 1
            citation_lines = [l for l in lines[j:end] if l and not detect_family(l)]

            dash = re.search(r'\s[—–]\s*', header)
            if dash:
                german = clean_german(header[dash.end():])
                sci_name, author = parse_sci(header[:dash.start()])
            else:
                german = ''
                sci_name, author = parse_sci(header)
                report.append((sec, num, 'kein deutscher Name im Eintrag', header[:60]))

            full_text = normalize_spaces(' '.join([header] + citation_lines))
            author_m = re.search(r'\(([^)]+\d{4}[^)]*)\)', full_text) or \
                re.search(r'([A-Z][a-z]+(?:\s+[A-Z][a-z.]+)?,\s*\d{4})', full_text)
            terra_m = re.search(r'terra\s+typica[:\s]+([^;)\n.]+)', full_text, re.IGNORECASE) or \
                re.search(r'—\s*([A-Z][^;)\n]{3,30})\)', full_text)

            results.append({
                'nom_id'      : f'{sec}-{num}',
                'section'     : sec,
                'num'         : num,
                'sci_name'    : sci_name,
                'author'      : author,
                'german_name' : german,
                'family_nom'  : fam if sec == 'I' else '',
                'author_year' : author_m.group(1).strip() if author_m else '',
                'terra_typica': terra_m.group(1).strip().rstrip('.)') if terra_m else '',
                'citation'    : normalize_spaces(' '.join(citation_lines)),
            })

    # Fehlende Familienüberschriften melden
    seen = {e['family_nom'] for e in results if e['family_nom']}
    for f in NOMENCLATOR_FAMILIES:
        if f not in seen:
            report.append(('I', None, f'Familienüberschrift nicht gefunden: {f}', ''))
    return results, report

# ── STEP 2: MATCHING ──────────────────────────────────────────────────────────
def norm_name(n):
    if not isinstance(n, str):
        return ''
    n = n.replace('�', '').replace('¬', '').strip().lower()
    for pat, rep in ABBREVIATIONS.items():
        n = re.sub(pat, rep, n)
    n = re.sub(r"[.\-'’,:;]", ' ', n)
    return normalize_spaces(n)

def compact(n):
    return norm_name(n).replace(' ', '')

def normalize_family(value):
    if pd.isna(value):
        return ''
    return re.sub(r'\s+', ' ', str(value).strip()).rstrip('.').lower()

def family_matches(source_family, nom_family):
    src, nom = normalize_family(source_family), normalize_family(nom_family)
    if not src and not nom:
        return None
    if not src or not nom:
        return False
    return src == nom

def load_synonyms():
    syn = {compact(k): v for k, v in SYNONYMS.items()}
    p = Path(SYNONYMS_CSV)
    if p.exists():
        df = pd.read_csv(p, dtype=str).fillna('')
        for _, r in df.iterrows():
            ids = [i.strip() for i in r['nomenclator_id'].split('|') if i.strip()]
            syn[compact(r['index_name'])] = (ids, r.get('note', ''))
    return syn

def match_species(bird_name, nom_by_id, nom_by_compact, synonyms):
    """Gibt dict mit method, entries (Liste), note, candidate, score zurück."""
    res = {'method': 'none', 'entries': [], 'note': '', 'candidate': None, 'score': 0}
    if not isinstance(bird_name, str) or not bird_name.strip():
        return res
    if HYBRID_RE.search(bird_name) or compact(bird_name) in KNOWN_HYBRIDS:
        res.update(method='hybrid/aggregate', note='Hybrid oder Sammeleintrag – nicht zugeordnet')
        return res
    key = compact(bird_name)

    if key in synonyms:
        ids, note = synonyms[key]
        entries = [nom_by_id[i] for i in ids if i in nom_by_id]
        res.update(entries=entries, note=note,
                   method='synonym' if len(entries) == 1 else 'synonym-ambiguous')
        return res

    if key in nom_by_compact:
        entries = nom_by_compact[key]
        res.update(entries=entries, method='exact' if len(entries) == 1 else 'exact-ambiguous',
                   score=100)
        return res

    choices = list(nom_by_compact.keys())
    top = process.extract(key, choices, scorer=fuzz.ratio, limit=2)
    if top and top[0][1] >= FUZZY_MIN:
        second = top[1][1] if len(top) > 1 else 0
        if top[0][1] - second >= FUZZY_MARGIN:
            res.update(method='fuzzy-review', candidate=nom_by_compact[top[0][0]][0],
                       score=round(top[0][1]))
            return res
    if top:
        res.update(candidate=nom_by_compact[top[0][0]][0], score=round(top[0][1]))
    return res

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
    workbook = pd.ExcelFile(path)
    sheet_names = workbook.sheet_names
    preferred = ['Species Index', 'Raw Entries']
    for name in preferred:
        if name in sheet_names:
            yield name
    for name in sheet_names:
        if name not in preferred:
            yield name

def main():
    print(f"Lade Excel: {EXCEL_IN} …")
    sheet_names = list(select_sheet_names(EXCEL_IN))
    species_sheet = 'Species Index' if 'Species Index' in sheet_names else sheet_names[0]
    raw_sheet = 'Raw Entries' if 'Raw Entries' in sheet_names else None

    df_species = pd.read_excel(EXCEL_IN, sheet_name=species_sheet)
    df_raw = pd.read_excel(EXCEL_IN, sheet_name=raw_sheet) if raw_sheet else pd.DataFrame()

    if 'Unnamed: 0' in df_species.columns:
        df_species = df_species.rename(columns={'Unnamed: 0': 'Orig_Idx'})
    if 'Family (Latin)' in df_species.columns:
        df_species['Family (Latin)'] = df_species['Family (Latin)'].apply(
            lambda v: '' if pd.isna(v) else str(v).strip())

    print("Parse Nomenclator …")
    nom_list, report = parse_nomenclator(FTS_PATH)
    nom_by_id = {e['nom_id']: e for e in nom_list}
    nom_by_compact = {}
    for e in nom_list:
        if e['german_name']:
            nom_by_compact.setdefault(compact(e['german_name']), []).append(e)
    synonyms = load_synonyms()
    for k, (ids, _) in synonyms.items():
        for i in ids:
            if i not in nom_by_id:
                report.append(('-', None, f'Synonym verweist auf unbekannte ID {i}', k))
    n1 = sum(e['section'] == 'I' for e in nom_list)
    n2 = sum(e['section'] == 'II' for e in nom_list)
    print(f"  → {n1} Einträge in Abschnitt I, {n2} in Abschnitt II")
    for sec, num, msg, ctx in report:
        print(f"  ! [{sec}-{num}] {msg}  {ctx}")

    if 'Source Page(s)' in df_species.columns:
        df_species['_sort_key'] = df_species['Source Page(s)'].apply(page_sort_key)
        df_species = df_species.sort_values('_sort_key', kind='stable').reset_index(drop=True)
    df_species.index += 1
    df_species = df_species.drop(columns=['_sort_key', 'Orig_Idx'], errors='ignore')

    print("Matche Arten …")
    out = {k: [] for k in ['num', 'nom_id', 'section', 'sci', 'dt', 'auth', 'terra', 'cit',
                           'fam', 'score', 'method', 'note', 'cand']}
    for _, row in df_species.iterrows():
        bird = row['Bird Name (German)']
        source_family = row.get('Family (Latin)', '')
        r = match_species(bird, nom_by_id, nom_by_compact, synonyms)
        single = r['entries'][0] if len(r['entries']) == 1 else None
        out['method'].append(r['method'])
        out['note'].append(r['note'])
        if single:
            out['num'].append(single['num'])
            out['nom_id'].append(single['nom_id'])
            out['section'].append(single['section'])
            out['sci'].append(single['sci_name'])
            out['dt'].append(single['german_name'])
            out['auth'].append(single['author_year'])
            out['terra'].append(single['terra_typica'])
            out['cit'].append(single['citation'])
            out['fam'].append(family_matches(source_family, single['family_nom']))
            out['score'].append(r['score'] if r['method'] == 'exact' else None)
            out['cand'].append('')
        else:
            for k in ['num', 'nom_id', 'section', 'score', 'fam']:
                out[k].append(None)
            for k in ['sci', 'dt', 'auth', 'terra', 'cit']:
                out[k].append('')
            cands = r['entries'] or ([r['candidate']] if r['candidate'] else [])
            txt = ' | '.join(f"{c['nom_id']} {c['german_name']} ({c['sci_name']})" for c in cands)
            if r['candidate'] and not r['entries']:
                txt += f" [Score {r['score']}]"
            out['cand'].append(txt)

    df_species.insert(0, 'Nr. (Nomenclator)', pd.array(out['num'], dtype='Int64'))
    df_species.insert(1, 'Nomenclator-ID', out['nom_id'])
    df_species = df_species.drop(columns=['Original Name'], errors='ignore')
    df_species['Abschnitt']               = out['section']
    df_species['Wissenschaftlicher Name'] = out['sci']
    df_species['Dt. Name (Nomenclator)']  = out['dt']
    df_species['Autor & Jahr']            = out['auth']
    df_species['Terra typica']            = out['terra']
    df_species['Zitation (Nomenclator)']  = out['cit']
    df_species['Family Match (Source/Nomenclator)'] = out['fam']
    df_species['Match-Methode']           = out['method']
    df_species['Match-Score']             = out['score']
    df_species['Kandidaten (prüfen)']     = out['cand']
    df_species['Hinweis']                 = out['note']
    counts = Counter(i for i in out['nom_id'] if i)
    df_species['Nr. mehrfach vergeben'] = [bool(i) and counts[i] > 1 for i in out['nom_id']]

    mc = Counter(out['method'])
    print(f"\n✓ Eindeutig zugeordnet: {mc['exact'] + mc['synonym']} / {len(df_species)}"
          f"  (exakt {mc['exact']}, Synonym {mc['synonym']})")
    for m in ['fuzzy-review', 'synonym-ambiguous', 'exact-ambiguous', 'hybrid/aggregate', 'none']:
        if mc[m]:
            print(f"  {m}: {mc[m]}")
    multi = sorted(i for i, c in counts.items() if c > 1)
    if multi:
        print(f"  Mehrfach vergebene Nomenclator-IDs: {', '.join(multi)}")

    col_order = [
        'Nr. (Nomenclator)', 'Nomenclator-ID', 'Abschnitt',
        'Bird Name (German)', 'Family (Latin)', 'Family Match (Source/Nomenclator)',
        'Wissenschaftlicher Name', 'Dt. Name (Nomenclator)', 'Autor & Jahr',
        'Terra typica', 'Match-Methode', 'Match-Score', 'Kandidaten (prüfen)',
        'Hinweis', 'Nr. mehrfach vergeben', 'References', 'Source Page(s)',
        'Zitation (Nomenclator)',
    ]
    col_order = [c for c in col_order if c in df_species.columns]

    df_report = pd.DataFrame(report, columns=['Abschnitt', 'Nr.', 'Meldung', 'Kontext'])
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        df_species[col_order].to_excel(writer, sheet_name='Species Index', index=True, index_label='#')
        df_raw.to_excel(writer, sheet_name='Raw Entries', index=False)
        pd.DataFrame(nom_list).rename(columns={
            'nom_id': 'ID', 'section': 'Abschnitt', 'num': 'Nr.',
            'sci_name': 'Wissenschaftlicher Name', 'author': 'Autor (Kopfzeile)',
            'german_name': 'Deutscher Name', 'family_nom': 'Familie',
            'author_year': 'Autor & Jahr', 'terra_typica': 'Terra typica', 'citation': 'Zitation',
        }).to_excel(writer, sheet_name='Nomenclator', index=False)
        df_report.to_excel(writer, sheet_name='Parser-Report', index=False)

    df_species[col_order].to_csv(OUTPUT_CSV, index=True, index_label='#', encoding='utf-8-sig')
    print(f"\n✓  Excel  →  {OUTPUT_EXCEL}")
    print(f"✓  CSV    →  {OUTPUT_CSV}")

if __name__ == '__main__':
    main()
