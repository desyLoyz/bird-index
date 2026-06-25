#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import pandas as pd
from pathlib import Path
from difflib import SequenceMatcher

# =========================
# KONFIG
# =========================
EXCEL_IN = "laubmann_index_v9_nomenclator.xlsx"
SHEET_NAME = "Species Index"
CORPUS_MD = "corpus.md"

EXCEL_OUT = "laubmann_index_v9_with_all_corpus_hits.xlsx"
CSV_OUT = "laubmann_index_v9_with_all_corpus_hits.csv"

FUZZY_THRESHOLD = 0.78
INITIAL_HIT_COLS = 20

ROMAN_TO_INT = {
    "I":1, "II":2, "III":3, "IV":4, "V":5, "VI":6, "VII":7, "VIII":8, "IX":9, "X":10,
    "XI":11, "XII":12, "XIII":13, "XIV":14, "XV":15, "XVI":16, "XVII":17, "XVIII":18, "XIX":19, "XX":20,
    "XXI":21, "XXII":22, "XXIII":23, "XXIV":24, "XXV":25, "XXVI":26, "XXVII":27, "XXVIII":28, "XXIX":29, "XXX":30,
    "XXXI":31, "XXXII":32, "XXXIII":33, "XXXIV":34, "XXXV":35, 
}
# =========================
# HELPERS
# =========================
def normalize_name(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("ä","ae").replace("ö","oe").replace("ü","ue").replace("ß","ss")
    s = re.sub(r"\[.*?\]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def sim(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize_name(a), normalize_name(b)).ratio()

def parse_all_references(refs: str):
    """
    "I, 162; II, 170; XVI, 188" -> [("I",162),("II",170),("XVI",188)]
    """
    if not isinstance(refs, str):
        return []
    return [(m.group(1), int(m.group(2)))
            for m in re.finditer(r"\b([IVXLCDM]+)\s*,\s*(\d+)\b", refs)]

def split_pages(corpus_text: str):
    pattern = re.compile(r"(^##\s+Vol\.\s+\d+\s+·\s+scan\s+\d+\s+·\s+p\.\s+\d+\.?\s*$)", re.MULTILINE)
    matches = list(pattern.finditer(corpus_text))
    pages = []

    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i+1].start() if i+1 < len(matches) else len(corpus_text)
        block = corpus_text[start:end]

        hm = re.search(r"Vol\.\s+(\d+)\s+·\s+scan\s+(\d+)\s+·\s+p\.\s+(\d+)", m.group(1))
        if not hm:
            continue
        pages.append({
            "vol": int(hm.group(1)),
            "scan": int(hm.group(2)),
            "page": int(hm.group(3)),
            "header": m.group(1).strip(),
            "text": block
        })
    return pages

def index_pages_by_vol_page(pages):
    return {(p["vol"], p["page"]): p for p in pages}

def extract_u_entries(page_text: str):
    """
    <u>Name</u>: 33; 99; 100;
    """
    pat = re.compile(r"<u>(.*?)</u>\s*:\s*([^\n\r]+)", re.IGNORECASE)
    out = []
    for m in pat.finditer(page_text):
        out.append((m.group(1).strip(), m.group(2).strip()))
    return out

def extract_positions(rhs: str):
    return [int(x) for x in re.findall(r"\b\d+\b", rhs)]

def find_best_entry(entries, target_name):
    best = None
    best_score = -1.0
    for name, rhs in entries:
        s = sim(name, target_name)
        if s > best_score:
            best_score = s
            best = (name, rhs)
    return best, best_score

# =========================
# MAIN
# =========================
def main():
    df = pd.read_excel(EXCEL_IN, sheet_name=SHEET_NAME)
    corpus_text = Path(CORPUS_MD).read_text(encoding="utf-8")

    pages = split_pages(corpus_text)
    page_idx = index_pages_by_vol_page(pages)

    # Name-Spalte erkennen
    if "Bird Name (German)" in df.columns:
        bird_col = "Bird Name (German)"
    elif "Bird Name" in df.columns:
        bird_col = "Bird Name"
    else:
        raise ValueError("Keine Namensspalte gefunden ('Bird Name (German)' oder 'Bird Name').")

    if "References" not in df.columns:
        raise ValueError("Spalte 'References' fehlt.")

    # Zusatzspalten
    if "Corpus_Debug" not in df.columns:
        df["Corpus_Debug"] = ""

    for i in range(1, INITIAL_HIT_COLS + 1):
        c = f"CorpusHit_{i}"
        if c not in df.columns:
            df[c] = ""

    max_hits_used = 0

    for ridx, row in df.iterrows():
        bird = str(row.get(bird_col, "")).strip()
        refs = row.get("References", "")
        ref_pairs = parse_all_references(refs)

        row_hits = []
        debug_parts = []

        if not ref_pairs:
            row_hits.append("<not-found!>")
            debug_parts.append("no-refs")
        else:
            for roman, p in ref_pairs:
                vol = ROMAN_TO_INT.get(roman)
                if not vol:
                    row_hits.append(f"{roman}, <not-found!>")
                    debug_parts.append(f"{roman},{p}:bad-roman")
                    continue

                page_obj = page_idx.get((vol, p))
                if not page_obj:
                    row_hits.append(f"{roman}, <not-found!>")
                    debug_parts.append(f"{roman},{p}:page-missing")
                    continue

                entries = extract_u_entries(page_obj["text"])
                if not entries:
                    row_hits.append(f"{roman}, <not-found!>")
                    debug_parts.append(f"{roman},{p}:no-u-entries")
                    continue

                best, score = find_best_entry(entries, bird)
                if not best or score < FUZZY_THRESHOLD:
                    row_hits.append(f"{roman}, <not-found!>")
                    debug_parts.append(f"{roman},{p}:no-match({score:.2f})")
                    continue

                matched_name, rhs = best
                nums = extract_positions(rhs)
                if not nums:
                    row_hits.append(f"{roman}, <not-found!>")
                    debug_parts.append(f"{roman},{p}:no-positions")
                    continue

                # Für jede lokale Position eigener Hit
                for n in nums:
                    row_hits.append(f"{roman}, {n}")

                debug_parts.append(
                    f"{roman},{p}:ok->{matched_name}({score:.2f})[{len(nums)}]"
                )

        # dynamisch Spalten erweitern
        if len(row_hits) > INITIAL_HIT_COLS:
            for j in range(INITIAL_HIT_COLS + 1, len(row_hits) + 1):
                c = f"CorpusHit_{j}"
                if c not in df.columns:
                    df[c] = ""

        for j, val in enumerate(row_hits, start=1):
            df.at[ridx, f"CorpusHit_{j}"] = val

        df.at[ridx, "Corpus_Debug"] = " | ".join(debug_parts)
        max_hits_used = max(max_hits_used, len(row_hits))

    # Ungenutzte Hit-Spalten entfernen
    hit_cols = [c for c in df.columns if c.startswith("CorpusHit_")]
    keep = {f"CorpusHit_{i}" for i in range(1, max_hits_used + 1)}
    drop = [c for c in hit_cols if c not in keep]
    if drop:
        df = df.drop(columns=drop)

    with pd.ExcelWriter(EXCEL_OUT, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=SHEET_NAME, index=False)

    df.to_csv(CSV_OUT, index=False, encoding="utf-8-sig")

    print(f"Fertig: {EXCEL_OUT}")
    print(f"Fertig: {CSV_OUT}")

if __name__ == "__main__":
    main()
