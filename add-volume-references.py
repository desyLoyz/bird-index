#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# This script adds the diary pages of the small volume registers (vols 1-34) to the index (Volume 35).

import argparse
import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

# =========================
# KONFIG
# =========================
EXCEL_IN = "laubmann_index_v13_ergaenzt_nomenclator.xlsx"
SHEET_NAME = "Species Index"
CORPUS_SOURCE = "json"  # "md" or "json"
CORPUS_FILE = {
    "md": "corpus.md",
    "json": "corpus.json",
}.get(CORPUS_SOURCE, "corpus.md")
SYNONYMS_FILE = Path(__file__).with_name("register-synonyms.csv")
READINGS_FILE = Path(__file__).with_name("register-readings.csv")

def build_output_paths(excel_in: str | Path, corpus_source: str, corpus_file: str | Path):
    excel_path = Path(excel_in)
    corpus_path = Path(corpus_file)
    corpus_label = corpus_path.stem
    if corpus_label.endswith("-edits"):
        corpus_label = corpus_label[:-len("-edits")]
    out_stem = f"{excel_path.stem}_{corpus_source.upper()}-{corpus_label}_with_all_corpus_hits"
    return (
        excel_path.with_name(out_stem + excel_path.suffix),
        excel_path.with_name(out_stem + ".csv"),
    )

FUZZY_THRESHOLD = 0.78
INITIAL_HIT_COLS = 20
MIN_REGISTER_ENTRIES = 3

ROMAN_TO_INT = {
    "I":1, "II":2, "III":3, "IV":4, "V":5, "VI":6, "VII":7, "VIII":8, "IX":9, "X":10,
    "XI":11, "XII":12, "XIII":13, "XIV":14, "XV":15, "XVI":16, "XVII":17, "XVIII":18, "XIX":19, "XX":20,
    "XXI":21, "XXII":22, "XXIII":23, "XXIV":24, "XXV":25, "XXVI":26, "XXVII":27, "XXVIII":28, "XXIX":29, "XXX":30,
    "XXXI":31, "XXXII":32, "XXXIII":33, "XXXIV":34, "XXXV":35,
}
OLD_SPELLINGS = (("th", "t"), ("moewe", "moeve"), ("rauh", "rau"), ("ck", "k"))
NAME_NOISE = {"x", "u", "und", "bastard"}
LETTER = "A-Za-zÄÖÜäöüß"
ENTRY_UNDERLINED = re.compile(r"^\s*(?:[-–]\s*)?<u>(?P<name>[^<]*?)</u>(?P<rhs>.*)$", re.IGNORECASE)
ENTRY_PLAIN = re.compile(rf"^\s*(?:[-–]\s*)?(?P<name>[A-ZÄÖÜ][^\d:;<>]*?[{LETTER}.)\]])\s*:(?P<rhs>.*)$")
NUMBER_LIKE = re.compile(rf"(?<![{LETTER}0-9])[0-9ilOo]{{1,3}}[a-d]?(?![{LETTER}0-9])")
LOOKALIKES = str.maketrans({"i": "1", "l": "1", "O": "0", "o": "0"})
ROMAN_MONTH = re.compile(rf"(?<![{LETTER}])[IVX]{{1,4}}(?![{LETTER}])")

# =========================
# HELPERS
# =========================
def normalize_name(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("ä","ae").replace("ö","oe").replace("ü","ue").replace("ß","ss")
    s = re.sub(r"[^a-z]+", " ", s)
    return s.strip()

def compact_name(s: str) -> str:
    s = normalize_name(s).replace(" ", "")
    for old, new in OLD_SPELLINGS:
        s = s.replace(old, new)
    return s

def sorted_words(s: str) -> str:
    return " ".join(sorted(w for w in normalize_name(s).split() if w not in NAME_NOISE))

def sim(a: str, b: str) -> float:
    return max(
        SequenceMatcher(None, normalize_name(a), normalize_name(b)).ratio(),
        SequenceMatcher(None, compact_name(a), compact_name(b)).ratio(),
        SequenceMatcher(None, sorted_words(a), sorted_words(b)).ratio(),
    )

def parse_all_references(refs: str):
    """
    "I, 162; II, 170; XVI, 188" -> [("I",162),("II",170),("XVI",188)]
    """
    if not isinstance(refs, str):
        return []
    return [(m.group(1), int(m.group(2)))
            for m in re.finditer(r"\b([IVXLCDM]+)\s*,\s*(\d+)\b", refs)]

def extract_positions(rhs: str):
    """ "12; 15; ioi; 5a" -> ["12", "15", "101", "5a"]; years and other 4-digit numbers are no pages """
    rhs = re.sub(r"</?sup>", " ", rhs)
    positions = []
    for m in NUMBER_LIKE.finditer(rhs):
        token = m.group(0)
        if any(ch.isdigit() for ch in token) or len(token.rstrip("abcd")) >= 2:
            positions.append(token.translate(LOOKALIKES))
    return positions

def is_number_list(text: str) -> bool:
    text = re.sub(r"</?sup>|</?u>|\bp\.", " ", text)
    numbers = extract_positions(text)
    letters = len(re.findall(f"[{LETTER}]", NUMBER_LIKE.sub(" ", text)))
    return bool(numbers) and letters <= 2 + len(numbers) // 4 and not ROMAN_MONTH.search(text)

def is_empty_rhs(rhs: str) -> bool:
    return not rhs.strip(" :;.-–")

def split_entry_line(line: str):
    """ "<u>Amsel</u>: 12; 15 <u>Drossel</u>: 7; 9" -> two entry lines """
    parts = [part for part in re.split(r"(?=<u>)", line) if part.strip()]
    matches = [ENTRY_UNDERLINED.match(part) for part in parts]
    if len(parts) > 1 and all(m and (is_number_list(m.group("rhs")) or is_empty_rhs(m.group("rhs"))) for m in matches):
        return parts
    return [line]

def extract_u_entries(page_text: str):
    """
    <u>Name</u>: 33; 99; 100;     <u>Name:</u> 33;     Name: 33; 99;
    followed by lines with numbers only, which continue the entry above.
    Returns the entries and whether the page looks like a register page.
    """
    entries = []
    lines = [part for line in (page_text or "").splitlines() if line.strip() for part in split_entry_line(line)]
    register_lines = 0
    in_entry = False
    for line in lines:
        m = ENTRY_UNDERLINED.match(line)
        if not m:
            plain = ENTRY_PLAIN.match(line)
            m = plain if plain and is_number_list(plain.group("rhs")) else None
        if m:
            rhs = m.group("rhs").strip().lstrip(":;").strip()
            name = m.group("name").strip().rstrip(":;").strip()
            if name:
                entries.append({"name": name, "rhs": rhs, "listed": is_number_list(rhs) or is_empty_rhs(rhs)})
                register_lines += 1
                in_entry = True
                continue
        if in_entry and "<u>" not in line and is_number_list(line) and not re.search(f"[{LETTER}]{{3,}}", line):
            entries[-1]["rhs"] += " " + line.strip()
            register_lines += 1
            continue
        in_entry = False
    for entry in entries:
        entry["positions"] = extract_positions(entry["rhs"])
    listed = sum(1 for e in entries if e["listed"] and e["positions"])
    is_register = listed >= MIN_REGISTER_ENTRIES and register_lines >= 0.5 * len(lines)
    return entries, is_register

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


def load_corpus_pages(corpus_path: Path, readings=None):
    corpus_path = Path(corpus_path)
    if corpus_path.suffix.lower() == ".json":
        with corpus_path.open("r", encoding="utf-8") as fh:
            payload = json.load(fh)

        if not isinstance(payload, list):
            raise ValueError(f"{corpus_path} does not contain a list of pages.")

        pages = []
        for item in payload:
            volume = item.get("volume")
            if volume is None:
                continue
            text = "\n".join(
                region.get("text", "")
                for region in item.get("regions", [])
                if isinstance(region, dict)
            )
            try:
                page = int(str(item.get("page_number", "")).strip().rstrip("."))
            except ValueError:
                page = None
            pages.append({
                "page_id": item.get("page_id", ""),
                "vol": int(volume),
                "scan": int(item.get("scan", 0) or 0),
                "page": page,
                "header": f"Vol. {volume} · scan {item.get('scan', 0) or 0} · p. {page}",
                "text": text,
            })
    else:
        pages = split_pages(corpus_path.read_text(encoding="utf-8"))

    readings = load_readings(READINGS_FILE) if readings is None else readings
    for page_no, page in enumerate(pages):
        page["entries"], page["is_register"] = extract_u_entries(page["text"])
        for entry in page["entries"]:
            fix = readings.get((page.get("page_id", ""), entry["name"], " ".join(extract_positions(entry["rhs"]))))
            if fix:
                entry["read_as"], entry["name"] = entry["name"], fix
        for entry_no, entry in enumerate(page["entries"]):
            entry["id"] = (page_no, entry_no)
            entry["page"] = page["page"]
            entry["on_register_page"] = page["is_register"]
            entry["scan"] = page["scan"]
    return pages


def load_readings(path):
    """register-readings.csv: headwords of the small registers checked on the scan
    (page_id, read_name, positions, scan_name); the entry is identified by its page, its reading and its numbers"""
    path = Path(path) if path else None
    if not path or not path.exists():
        return {}
    table = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    return {(r.page_id, r.read_name, " ".join(r.positions.replace(";", " ").split())): r.scan_name
            for r in table.itertuples(index=False)}


def load_synonyms(path):
    path = Path(path) if path else None
    if not path or not path.exists() or not path.read_text(encoding="utf-8-sig").strip():
        return {}
    table = pd.read_csv(path, dtype=str, keep_default_na=False, sep=None, engine="python", encoding="utf-8-sig")
    if table.shape[1] < 2:
        raise ValueError(f"{path}: expected the columns 'Bird Name (German)' and 'Register Name'.")
    synonyms = defaultdict(list)
    for row in table.itertuples(index=False):
        bird, name = str(row[0]).strip(), str(row[1]).strip()
        if bird and compact_name(name):
            synonyms[bird].append(name)
    return dict(synonyms)


@dataclass
class Reference:
    row: int
    roman: str
    page: int
    vol: int | None
    bird: str
    status: str = ""
    entry: dict | None = None
    score: float = 0.0


class RegisterMatcher:
    """
    Gives every (species, volume, register page) reference of the index one entry of the small register:
    1. an entry with the species' name (or a synonym) on the cited page; if a register page and a diary page
       share the cited number, only the register page counts
    2. an entry with that name on another register page of the volume, entries with page numbers first, then
       the nearest page (the index cites misread page numbers, e.g. 187 for 181, and some register page numbers
       are misread); only the run of register pages around the pages that gave at least three entries in step 1
    3. the most similar entry on the cited page, if no other index species is as similar to it and the entry
       does not carry the name (or a synonym) of another index species; a misread headword ("Gänsesäger"
       for "Gänsegeier") is corrected in register-readings.csv after a look at the scan, never guessed here
    Each entry goes to one species only.
    """

    def __init__(self, pages, birds, synonyms):
        self.by_number = defaultdict(list)
        self.registers = defaultdict(list)
        for page in pages:
            if page["page"] is not None:
                self.by_number[(page["vol"], page["page"])].append(page)
            if page["is_register"]:
                self.registers[page["vol"]].append(page)
        self.names = {bird: [name for name in [bird] + synonyms.get(bird, []) if compact_name(name)] or [bird] for bird in birds}
        self.owners = defaultdict(set)
        for bird, names in self.names.items():
            for name in names:
                self.owners[compact_name(name)].add(bird)
        self.score_cache = {}

    def is_name_of(self, bird, entry):
        return bird in self.owners.get(compact_name(entry["name"]), ())

    def score(self, bird, entry):
        return max(sim(name, entry["name"]) for name in self.names[bird])

    def is_closest(self, bird, entry):
        key = entry["name"]
        if key not in self.score_cache:
            self.score_cache[key] = {other: self.score(other, entry) for other in self.names}
        scores = self.score_cache[key]
        return all(scores[bird] > s for other, s in scores.items() if other != bird)

    def cited_entries(self, ref):
        pages = self.by_number.get((ref.vol, ref.page), [])
        registers = [page for page in pages if page["is_register"]]
        return [entry for page in (registers or pages) for entry in page["entries"]]

    def register_scans(self, refs, gap=2):
        claims = defaultdict(Counter)
        for ref in refs:
            if ref.status and ref.entry["on_register_page"]:
                claims[ref.vol][ref.entry["scan"]] += 1
        ranges = {}
        for vol, counts in claims.items():
            scans = [scan for scan, n in counts.items() if n >= 3]
            if not scans:
                continue
            low, high = min(scans), max(scans)
            register_scans = sorted({page["scan"] for page in self.registers[vol]})
            for scan in reversed(register_scans):
                if low - gap <= scan < low:
                    low = scan
            for scan in register_scans:
                if high < scan <= high + gap:
                    high = scan
            ranges[vol] = (low, high)
        return ranges

    def match(self, refs):
        claimed = {}
        open_refs = [ref for ref in refs if ref.vol]

        def take(ref, entry, status, score):
            ref.entry, ref.status, ref.score = entry, status, score
            claimed[entry["id"]] = ref.bird

        for ref in open_refs:
            for entry in sorted(self.cited_entries(ref), key=lambda e: not e["positions"]):
                if not ref.status and self.is_name_of(ref.bird, entry) and claimed.get(entry["id"], ref.bird) == ref.bird:
                    take(ref, entry, "ok", 1.0)

        scan_ranges = self.register_scans(open_refs)
        candidates = []
        for order, ref in enumerate(open_refs):
            if ref.status:
                continue
            low, high = scan_ranges.get(ref.vol, (float("-inf"), float("inf")))
            for page in self.registers[ref.vol]:
                if page["page"] == ref.page or not low <= page["scan"] <= high:
                    continue
                for entry in page["entries"]:
                    if self.is_name_of(ref.bird, entry):
                        distance = abs(page["page"] - ref.page) if page["page"] is not None else 1000
                        candidates.append((not entry["positions"], distance, order, ref, entry))
        for *_, ref, entry in sorted(candidates, key=lambda c: c[:3]):
            if not ref.status and entry["id"] not in claimed:
                take(ref, entry, "ok-other-page", 1.0)
        moved = defaultdict(list)
        for ref in open_refs:
            if ref.status == "ok-other-page":
                moved[(ref.vol, ref.bird)].append(ref)
        for group in moved.values():
            entries = sorted((ref.entry for ref in group), key=lambda e: (e["page"] is None, e["page"] or 0))
            for ref, entry in zip(sorted(group, key=lambda r: r.page), entries):
                ref.entry = entry

        candidates = []
        for order, ref in enumerate(open_refs):
            if ref.status:
                continue
            for entry in self.cited_entries(ref):
                score = self.score(ref.bird, entry)
                ref.score = max(ref.score, score)
                owners = self.owners.get(compact_name(entry["name"]), set())
                if score >= FUZZY_THRESHOLD and entry["id"] not in claimed and not owners - {ref.bird}:
                    candidates.append((-score, order, ref, entry))
        for negative_score, _, ref, entry in sorted(candidates, key=lambda c: c[:2]):
            if not ref.status and entry["id"] not in claimed and self.is_closest(ref.bird, entry):
                take(ref, entry, "ok", -negative_score)

        for ref in open_refs:
            if ref.status:
                if not ref.entry["positions"]:
                    ref.status = "no-positions"
                continue
            pages = self.by_number.get((ref.vol, ref.page), [])
            if not pages:
                ref.status = "page-missing"
            elif not any(page["entries"] for page in pages):
                ref.status = "no-u-entries"
            else:
                ref.status = "no-match"
        return claimed

    def synonym_candidates(self, refs, claimed):
        pairs = defaultdict(set)
        for ref in refs:
            if ref.status != "no-match":
                continue
            for page in self.by_number.get((ref.vol, ref.page), []):
                for entry in page["entries"]:
                    if entry["id"] not in claimed and compact_name(entry["name"]) not in self.owners and entry["positions"]:
                        pairs[(ref.bird, entry["name"])].add(ref.roman)
        rows = [
            {"Bird Name (German)": bird, "Register Name": name, "volumes": len(vols), "volume_list": ", ".join(sorted(vols, key=ROMAN_TO_INT.get))}
            for (bird, name), vols in pairs.items() if len(vols) >= 2
        ]
        return pd.DataFrame(rows, columns=["Bird Name (German)", "Register Name", "volumes", "volume_list"]).sort_values(
            ["volumes", "Bird Name (German)"], ascending=[False, True])

# =========================
# MAIN
# =========================
def parse_args():
    parser = argparse.ArgumentParser(description="Add single-volume corpus references to the bird index workbook.")
    parser.add_argument("--excel-in", default=EXCEL_IN, help="Input Excel workbook path.")
    parser.add_argument("--sheet-name", default=SHEET_NAME, help="Worksheet name with the species index.")
    parser.add_argument("--corpus-file", default=CORPUS_FILE, help="Corpus source file: corpus.md or corpus.json.")
    parser.add_argument("--corpus-source", choices=("md", "json"), help="Corpus source type. Overrides default if set.")
    parser.add_argument("--synonyms", default=str(SYNONYMS_FILE), help="CSV with register names of index species (empty string: none).")
    parser.add_argument("--readings", default=str(READINGS_FILE), help="CSV with scan-checked headword readings of the small registers (empty string: none).")
    return parser.parse_args()


def main():
    args = parse_args()
    excel_in = args.excel_in
    sheet_name = args.sheet_name
    corpus_source = args.corpus_source or CORPUS_SOURCE
    corpus_file = args.corpus_file if args.corpus_file != CORPUS_FILE else {
        "md": "corpus.md",
        "json": "corpus.json",
    }.get(corpus_source, CORPUS_FILE)

    excel_out, csv_out = build_output_paths(excel_in, corpus_source, corpus_file)

    df = pd.read_excel(excel_in, sheet_name=sheet_name)
    pages = load_corpus_pages(Path(corpus_file), load_readings(args.readings))

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

    birds = {str(row.get(bird_col, "")).strip() for _, row in df.iterrows()}
    matcher = RegisterMatcher(pages, birds, load_synonyms(args.synonyms))
    refs = [
        Reference(ridx, roman, p, ROMAN_TO_INT.get(roman), str(row.get(bird_col, "")).strip())
        for ridx, row in df.iterrows()
        for roman, p in parse_all_references(row.get("References", ""))
    ]
    claimed = matcher.match(refs)
    refs_by_row = defaultdict(list)
    for ref in refs:
        refs_by_row[ref.row].append(ref)

    max_hits_used = 0
    stats = Counter()
    per_volume_stats = {}

    def get_volume_stats(vol):
        key = str(vol)
        if key not in per_volume_stats:
            per_volume_stats[key] = Counter(volume=vol)
        return per_volume_stats[key]

    for ridx, row in df.iterrows():
        row_hits = []
        debug_parts = []
        row_refs = refs_by_row.get(ridx, [])

        if not row_refs:
            row_hits.append("<not-found!>")
            debug_parts.append("no-refs")
            stats["missed_matches"] += 1
            stats["total_reference_pairs"] += 1
            stats["rows_without_match"] += 1
        else:
            for ref in row_refs:
                stats["total_reference_pairs"] += 1
                vol_stats = get_volume_stats(ref.vol) if ref.vol else Counter()
                vol_stats["total_reference_pairs"] += 1

                if not ref.vol:
                    row_hits.append(f"{ref.roman}, <not-found!>")
                    debug_parts.append(f"{ref.roman},{ref.page}:bad-roman")
                    stats["missed_matches"] += 1
                    continue

                if not ref.status.startswith("ok"):
                    row_hits.append(f"{ref.roman}, <not-found!>")
                    score = f"({ref.score:.2f})" if ref.status == "no-match" else ""
                    debug_parts.append(f"{ref.roman},{ref.page}:{ref.status}{score}")
                    stats["missed_matches"] += 1
                    vol_stats["missed_matches"] += 1
                    continue

                kind = "exact_matches" if ref.score >= 1.0 else "fuzzy_matches"
                for counter in (stats, vol_stats):
                    counter["successful_matches"] += 1
                    counter[kind] += 1
                    if ref.status == "ok-other-page":
                        counter["other_page_matches"] += 1

                for n in ref.entry["positions"]:
                    row_hits.append(f"{ref.roman}, {n}")

                found = "" if ref.status == "ok" else f"@{ref.entry['page'] if ref.entry['page'] is not None else '?'}"
                debug_parts.append(
                    f"{ref.roman},{ref.page}:ok{found}->{ref.entry['name']}({ref.score:.2f})[{len(ref.entry['positions'])}]"
                )

            matched_volumes = {ref.vol for ref in row_refs if ref.status.startswith("ok")}
            if matched_volumes:
                stats["rows_with_any_match"] += 1
                for vol in matched_volumes:
                    get_volume_stats(vol)["rows_with_any_match"] += 1
            else:
                stats["rows_without_match"] += 1
                for ref in row_refs:
                    if ref.vol:
                        get_volume_stats(ref.vol)["rows_without_match"] += 1

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

    metrics = ["total_reference_pairs", "successful_matches", "missed_matches", "exact_matches", "fuzzy_matches",
               "other_page_matches", "rows_with_any_match", "rows_without_match"]

    def success_rate(counter):
        return round(counter["successful_matches"] / counter["total_reference_pairs"] * 100, 2) if counter["total_reference_pairs"] else 0.0

    stats_df = pd.DataFrame(
        [{"metric": m, "value": stats[m]} for m in metrics]
        + [{"metric": "success_rate_percent", "value": success_rate(stats)}]
    )

    volume_stats_df = pd.DataFrame([
        {"volume": v["volume"], **{m: v[m] for m in metrics}, "success_rate_percent": success_rate(v)}
        for v in sorted(per_volume_stats.values(), key=lambda x: x["volume"])
    ])

    with pd.ExcelWriter(excel_out, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=SHEET_NAME, index=False)
        stats_df.to_excel(w, sheet_name="Corpus_Stats", index=False)
        volume_stats_df.to_excel(w, sheet_name="Volume_Stats", index=False)
        matcher.synonym_candidates(refs, claimed).to_excel(w, sheet_name="Synonym_Candidates", index=False)

    df.to_csv(csv_out, index=False, encoding="utf-8-sig")

    print(f"Fertig: {excel_out}")
    print(f"Fertig: {csv_out}")
    print("Statistik:")
    for _, row in stats_df.iterrows():
        print(f"- {row['metric']}: {row['value']}")
    print("\nVolume_Stats:")
    print(volume_stats_df.to_string(index=False))

if __name__ == "__main__":
    main()
