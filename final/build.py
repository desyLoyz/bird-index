"""Final register dataset: Laubmann's general index (vol 35) resolved through the small registers of vols 1-34.

    python final/build.py --corpus <register corpus.json> --out <folder>

Steps
 1. index v14 (rectify-index.py on "Laubmann_35_gemini_edits (18).json" in the repo root)
    + index-fixes.csv                                          -> laubmann_index_v15.xlsx
 2. nomenclator-match.py (v9) + nomenclator_synonyms.csv       -> Nomenclator 1916 entry per species
 3. add-volume-references.py (v3) + register-synonyms.csv + register-readings.csv
                                                               -> one small-register entry per index reference
 4. final/taxa.py                                              -> GBIF backbone taxon per species
 5. tables: species, references, diary_pages, index_correction_candidates
 6. KG check (laubmann-htr eval/validate_kg.py): every KG occurrence whose species a register lists on its
    diary page; DwC-A copy with identificationVerificationStatus and a subset of the listed records.

The KG is used in step 4 (taxon names) and step 6 (flags) only; no register reading or reference depends on it.
"""
import argparse
import csv
import importlib.util
import json
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
HTR = Path(r"C:\Users\totom\Projects\laubmann-htr")
DWCA = Path(r"C:\Users\totom\Projects\laubmann-kg_final\data\exports\kg_exports_2026-10-06_regions\dwca")
ROMAN = "I II III IV V VI VII VIII IX X XI XII XIII XIV XV XVI XVII XVIII XIX XX XXI XXII XXIII XXIV XXV XXVI " \
        "XXVII XXVIII XXIX XXX XXXI XXXII XXXIII XXXIV".split()
VERIFICATION = {
    "listed on page": "species listed for this diary page in Laubmann's own registers",
    "listed on neighbour page": "species listed for the neighbouring diary page in Laubmann's own registers",
    "not listed": "species not listed for this diary page in Laubmann's registers",
    "no diary page number": "diary page number unknown; not checked against Laubmann's registers",
}


def load_module(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run(cmd, cwd):
    print("$", " ".join(str(c) for c in cmd))
    subprocess.run([str(c) for c in cmd], cwd=cwd, check=True, env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"})


def match_kind(ref, matcher, M):
    name = M.compact_name(ref.entry["name"])
    if name == M.compact_name(ref.bird):
        kind = "exact"
    elif matcher.is_name_of(ref.bird, ref.entry):
        kind = "synonym"
    else:
        kind = "fuzzy"
    return kind


def build(corpus, out):
    work = out / "work"
    work.mkdir(parents=True, exist_ok=True)
    py = sys.executable

    # 1 index
    if not (ROOT / "laubmann_index_v14.xlsx").exists():
        run([py, ROOT / "rectify-index.py"], ROOT)
    index_xlsx = work / "laubmann_index_v15.xlsx"
    run([py, ROOT / "final/fix_index.py", ROOT / "laubmann_index_v14.xlsx", ROOT / "index-fixes.csv", index_xlsx], ROOT)

    # 2 Nomenclator
    for name in ("nomenclator_synonyms.csv", "inputs/fts.txt"):
        shutil.copy(ROOT / name, work / Path(name).name)
    run([py, ROOT / "nomenclator-match.py", index_xlsx.name], work)
    nomenclator_xlsx = work / "laubmann_index_v15_nomenclator.xlsx"

    # 3 matching
    M = load_module(ROOT / "add-volume-references.py", "avr")
    index = pd.read_excel(index_xlsx, sheet_name="Species Index")
    index["Bird Name (German)"] = index["Bird Name (German)"].astype(str).str.strip()
    pages = M.load_corpus_pages(Path(corpus), M.load_readings(ROOT / "register-readings.csv"))
    readings = pd.read_csv(ROOT / "register-readings.csv", dtype=str, keep_default_na=False)
    applied = sum("read_as" in e for p in pages for e in p["entries"])
    assert applied == len(readings), f"{applied} of {len(readings)} register readings applied"
    for p in pages:
        for e in p["entries"]:
            e["page_id"] = p.get("page_id", "")
    birds = set(index["Bird Name (German)"])
    matcher = M.RegisterMatcher(pages, birds, M.load_synonyms(ROOT / "register-synonyms.csv"))
    refs = [M.Reference(i, roman, page, M.ROMAN_TO_INT.get(roman), row["Bird Name (German)"])
            for i, row in index.iterrows() for roman, page in M.parse_all_references(row["References"])]
    matcher.match(refs)

    fixes = pd.read_csv(ROOT / "index-fixes.csv", dtype=str, keep_default_na=False)
    fixed = {(f.species, f.roman, int(f.fixed_page)): int(f.cited_page) for f in fixes.itertuples() if f.action == "page"}
    v14 = pd.read_excel(ROOT / "laubmann_index_v14.xlsx", sheet_name="Species Index")
    v14_name = v14["Bird Name (German)"].astype(str).str.strip()
    merged = {(f.fixed_page, roman, page): f.species for f in fixes.itertuples() if f.action == "merge"
              for roman, page in M.parse_all_references(v14[v14_name == f.species].iloc[0]["References"])}

    rows = []
    for n, ref in enumerate(refs, start=1):
        ok = ref.status.startswith("ok")
        e = ref.entry if ok else None
        rows.append({
            "reference_id": f"R{n:05d}",
            "species": ref.bird,
            "volume": ref.vol or "",
            "volume_roman": ref.roman,
            "register_page_cited": fixed.get((ref.bird, ref.roman, ref.page), ref.page),
            "register_page_index_v15": ref.page,
            "status": "resolved" if ok else ref.status,
            "match": (match_kind(ref, matcher, M) + ("" if ref.status == "ok" else ", other register page")) if ok else "",
            "register_page": (e["page"] if e["page"] is not None else "") if ok else "",
            "register_scan": e["page_id"] if ok else "",
            "register_entry": e["name"] if ok else "",
            "register_entry_htr": e.get("read_as", e["name"]) if ok else "",
            "score": round(ref.score, 3),
            "diary_pages": " ".join(e["positions"]) if ok else "",
            "index_fix": ("page restored" if (ref.bird, ref.roman, ref.page) in fixed else "")
                         or (f"from the '{merged[(ref.bird, ref.roman, ref.page)]}' row"
                             if (ref.bird, ref.roman, ref.page) in merged else ""),
        })
    references = pd.DataFrame(rows)

    # eval-compatible copy for taxa.py and validate_kg.py
    compat = references.assign(
        status=references.status.map(lambda s: "ok" if s == "resolved" else s),
        matched_name=references.register_entry,
        found_page=[str(r.register_page) if "other" in r.match else "" for r in references.itertuples()],
        register_page=references.register_page_index_v15,
    )[["species", "volume", "register_page", "status", "score", "matched_name", "found_page", "diary_pages"]]
    compat_csv = work / "references_compat.csv"
    compat.to_csv(compat_csv, index=False, encoding="utf-8")

    # 4 taxa
    taxa_csv = work / "taxa.csv"
    run([py, ROOT / "final/taxa.py", nomenclator_xlsx, compat_csv, taxa_csv], ROOT)
    taxa = pd.read_csv(taxa_csv, keep_default_na=False, dtype=str)

    # 5 tables
    data = out / "data"
    data.mkdir(exist_ok=True)
    tax_cols = ["nomenclator_id", "scientificName", "acceptedSpecies", "gbif_usageKey"]
    references = references.merge(taxa[["index_name"] + tax_cols].rename(columns={"index_name": "species"}), on="species", how="left")
    references.to_csv(data / "references.csv", index=False, encoding="utf-8")

    diary = []
    for r in references[references.status == "resolved"].itertuples():
        for token in dict.fromkeys(r.diary_pages.split()):
            diary.append({"species": r.species, "nomenclator_id": r.nomenclator_id, "scientificName": r.scientificName,
                          "acceptedSpecies": r.acceptedSpecies, "gbif_usageKey": r.gbif_usageKey,
                          "volume": r.volume, "volume_roman": r.volume_roman, "diary_page": token,
                          "reference_id": r.reference_id, "register_page": r.register_page, "match": r.match})
    diary = pd.DataFrame(diary)
    diary.to_csv(data / "diary_pages.csv", index=False, encoding="utf-8")

    counts = references.groupby("species").agg(
        references=("reference_id", "size"),
        resolved=("status", lambda s: (s == "resolved").sum()),
        volumes=("volume", "nunique"),
    )
    pages_per_species = diary.groupby("species").size().rename("diary_pages")
    species = taxa.merge(counts, left_on="index_name", right_index=True, how="left") \
                  .merge(pages_per_species, left_on="index_name", right_index=True, how="left")
    species["diary_pages"] = species["diary_pages"].fillna(0).astype(int)
    species.to_csv(data / "species.csv", index=False, encoding="utf-8")

    def correction_kind(cited, found):
        a, b = str(cited), str(found)
        if len(a) == len(b) and sum(x != y for x, y in zip(a, b)) == 1:
            return "one digit differs"
        if sorted(a) == sorted(b):
            return "digits swapped"
        return "other"
    other = references[references.match.str.contains("other register page")].copy()
    other["kind"] = [correction_kind(a, b) for a, b in zip(other.register_page_index_v15, other.register_page)]
    other[["reference_id", "species", "volume_roman", "register_page_index_v15", "register_page", "kind",
           "register_entry", "register_scan", "diary_pages"]].to_csv(data / "index_correction_candidates.csv", index=False, encoding="utf-8")
    references[references.status != "resolved"][["reference_id", "species", "volume_roman", "register_page_index_v15", "status"]] \
        .to_csv(data / "unresolved_references.csv", index=False, encoding="utf-8")
    readings.to_csv(data / "register_readings_checked.csv", index=False, encoding="utf-8")
    fixes.to_csv(data / "index_fixes.csv", index=False, encoding="utf-8")

    # 6 KG check
    sys.path.insert(0, str(HTR / "eval"))
    V = load_module(HTR / "eval/validate_kg.py", "validate_kg")
    V.INDEX_XLSX = index_xlsx
    V.OUT_ROOT = work
    summary = V.run(compat_csv, corpus, "kg_check", ROOT / "add-volume-references.py")
    records = pd.read_csv(work / "kg_check/records.csv", keep_default_na=False, dtype=str)
    records = records.rename(columns={"either": "register_check"})[
        ["occurrence", "date", "species", "taxon", "volume", "diary_page", "register_check", "index", "registers"]]
    records.to_csv(data / "kg_register_check.csv", index=False, encoding="utf-8")
    summary.to_csv(data / "kg_register_check_summary.csv", index=False, encoding="utf-8")
    write_dwca(records, out / "dwca")

    stats = {
        "species_in_index": int(len(species)),
        "references": int(len(references)),
        "resolved": int((references.status == "resolved").sum()),
        "resolved_percent": round(100 * (references.status == "resolved").mean(), 2),
        "match": references[references.status == "resolved"].match.value_counts().to_dict(),
        "unresolved": references[references.status != "resolved"].status.value_counts().to_dict(),
        "species_volume_diary_pages": int(len(diary)),
        "distinct_volume_diary_pages": int(diary[["volume", "diary_page"]].drop_duplicates().shape[0]),
        "register_readings_checked": int(len(readings)),
        "index_fixes": int(len(fixes)),
        "taxon_check": species.taxon_check.str.split(";").str[0].value_counts().to_dict(),
        "kg_records": int(len(records)),
        "kg_register_check": records.register_check.value_counts().to_dict(),
    }
    (data / "summary.json").write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(stats, ensure_ascii=False, indent=2))


def write_dwca(records, out):
    """KG DwC-A with identificationVerificationStatus per occurrence (absence records: not checked) and a subset
    with the records listed on their own diary page."""
    out.mkdir(parents=True, exist_ok=True)
    status_of = dict(zip(records.occurrence, records.register_check))
    term = "http://rs.tdwg.org/dwc/terms/identificationVerificationStatus"
    meta = (DWCA / "meta.xml").read_text(encoding="utf-8")
    with open(DWCA / "occurrence.txt", encoding="utf-8", newline="") as f:
        header = f.readline().rstrip("\n").split("\t")
        lines = [line.rstrip("\n").split("\t") for line in f if line.strip()]
    position = len(header)
    ext_start = meta.index('rowType="http://rs.tdwg.org/dwc/terms/Occurrence"')
    ext_end = meta.index("</extension>", ext_start)
    meta = meta[:ext_end] + f'  <field index="{position}" term="{term}"/>\n  ' + meta[ext_end:]
    for subset, keep in (("all", lambda s: True), ("listed_on_page", lambda s: s == "listed on page")):
        folder = out / subset
        folder.mkdir(exist_ok=True)
        for name in ("eml.xml", "event.txt", "measurementorfact.txt"):
            shutil.copy(DWCA / name, folder / name)
        (folder / "meta.xml").write_text(meta, encoding="utf-8")
        kept, events = set(), set()
        with open(folder / "occurrence.txt", "w", encoding="utf-8", newline="") as f:
            f.write("\t".join(header + ["identificationVerificationStatus"]) + "\n")
            for row in lines:
                occurrence = row[1].rsplit("/", 1)[-1]
                status = status_of.get(occurrence, "absence record")
                if keep(status):
                    kept.add(row[1])
                    events.add(row[0])
                    f.write("\t".join(row + [VERIFICATION.get(status, "absence record; not checked against the registers")]) + "\n")
        if subset != "all":
            with open(DWCA / "measurementorfact.txt", encoding="utf-8") as src, \
                    open(folder / "measurementorfact.txt", "w", encoding="utf-8", newline="") as dst:
                dst.write(src.readline())
                for line in src:
                    event, occurrence = line.split("\t")[:2]
                    if occurrence in kept or (not occurrence and event in events):
                        dst.write(line)
            with open(DWCA / "event.txt", encoding="utf-8") as src, \
                    open(folder / "event.txt", "w", encoding="utf-8", newline="") as dst:
                dst.write(src.readline())
                for line in src:
                    if line.split("\t")[0] in events:
                        dst.write(line)
        shutil.make_archive(str(out / f"laubmann_dwca_{subset}"), "zip", folder)
        shutil.rmtree(folder)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--corpus", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()
    build(Path(args.corpus), Path(args.out))
