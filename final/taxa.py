"""Species table: index name, names in the small registers, Nomenclator 1916 entry, GBIF backbone taxon.

The GBIF taxon is the KG's taxon for the index name, its register synonyms and the names the small registers
use (GBIF backbone, linked through GBIF/Wikidata vernaculars and reviewed), unless taxa-overrides.csv says
otherwise. The Nomenclator name (OCR of the BHL scan, repaired by a few letter swaps until GBIF matches it) is
the independent check: GBIF resolves many 1916 names to the wrong species, so it is never the source.

    python final/taxa.py <index_v15_nomenclator.xlsx> <references.csv> <out.csv>
"""
import json
import re
import sys
import time
import urllib.parse
import urllib.request
from collections import Counter
from pathlib import Path

import pandas as pd

HTR = Path(r"C:\Users\totom\Projects\laubmann-htr")
sys.path.insert(0, str(HTR / "eval"))
import confirm as C  # noqa: E402

CACHE = Path(__file__).resolve().parents[1] / "runs" / "gbif_cache.json"
OCR_SWAPS = [("tli", "th"), ("li", "h"), ("rn", "m"), ("in", "m"), ("iu", "in"), ("nn", "un"), ("ii", "u"),
             ("u", "n"), ("n", "u"), ("c", "e"), ("e", "c"), ("y", "v"), ("ß", "B"), ("ö", "O"), ("ü", "O"),
             ("ca", "ea"), ("ae", "ac"), ("ac", "ae"), ("sra", "ga"), ("ar", "ra")]
cache = json.loads(CACHE.read_text(encoding="utf-8")) if CACHE.exists() else {}


def gbif_match(name):
    if name not in cache:
        url = "https://api.gbif.org/v1/species/match?" + urllib.parse.urlencode({"name": name, "class": "Aves"})
        req = urllib.request.Request(url, headers={"User-Agent": "bird-index-taxa/1.0"})
        cache[name] = json.load(urllib.request.urlopen(req, timeout=30))
        time.sleep(0.1)
    return cache[name]


def species_level(m):
    return m.get("rank") in ("SPECIES", "SUBSPECIES") and m.get("matchType") in ("EXACT", "FUZZY")


def variants(name):
    words = name.split()
    yield " ".join(words)
    yield " ".join(words[:2])
    seen = set()
    for old, new in OCR_SWAPS:
        for n in (" ".join(words[:2]), " ".join(words)):
            for i in [m.start() for m in re.finditer(re.escape(old), n)]:
                v = n[:i] + new + n[i + len(old):]
                if v not in seen:
                    seen.add(v)
                    yield v
    if len(words) >= 2 and words[0].lower() == words[1].lower():
        yield words[0]


def resolve(sci_name):
    """first variant GBIF resolves at species level: (matched variant, match)"""
    if not isinstance(sci_name, str) or not sci_name.strip():
        return None, None
    clean = re.sub(r"[^A-Za-zäöüßé\- ]", "", sci_name).strip()
    for v in variants(clean):
        m = gbif_match(v)
        if species_level(m):
            return v, m
    return None, gbif_match(" ".join(clean.split()[:2]))


def accepted_species(m):
    return (m or {}).get("species"), (m or {}).get("speciesKey")


def main(nomenclator_xlsx, references_csv, out_csv):
    index = pd.read_excel(nomenclator_xlsx, sheet_name="Species Index")
    refs = pd.read_csv(references_csv, keep_default_na=False)
    reg_names = refs[refs.status.str.startswith("ok")].groupby("species").matched_name.agg(lambda s: Counter(s).most_common(3))
    kg = C.SpeciesOnPage(C.kg_records())
    root = Path(__file__).resolve().parents[1]
    syn = pd.read_csv(root / "register-synonyms.csv", dtype=str, keep_default_na=False)
    overrides = {r.index_name: (r.scientific_name, r.note) for r in
                 pd.read_csv(root / "taxa-overrides.csv", dtype=str, keep_default_na=False).itertuples(index=False)}
    rows = []
    for r in index.itertuples(index=False):
        bird = str(r[index.columns.get_loc("Bird Name (German)")]).strip()
        get = lambda col: r[index.columns.get_loc(col)]  # noqa: E731
        sci = get("Wissenschaftlicher Name") if get("Match-Methode") in ("exact", "synonym") else None
        variant, match = resolve(sci) if sci else (None, None)
        names = [bird] + list(syn[syn.iloc[:, 0] == bird].iloc[:, 1]) + [n for n, _ in reg_names.get(bird, [])]
        kg_taxa = Counter(kg.kg_taxon_of_name.get(C.fold(n)) for n in names) - Counter({None: 99})
        kg_taxon = kg_taxa.most_common(1)[0][0] if kg_taxa else None
        kg_match = gbif_match(kg_taxon) if kg_taxon else None
        n_sp, n_key = accepted_species(match if variant else None)
        k_sp, k_key = accepted_species(kg_match if kg_match and kg_match.get("rank") in ("SPECIES", "SUBSPECIES") else None)
        if n_key and k_key:
            agree = "nomenclator agrees" if n_key == k_key else "nomenclator differs"
        elif sci:
            agree = "nomenclator name not resolved by GBIF"
        elif get("Match-Methode") in ("synonym-ambiguous", "fuzzy-review"):
            agree = "nomenclator entry ambiguous"
        else:
            agree = "no nomenclator entry"
        if bird in overrides:
            name, note = overrides[bird]
            source, final = "override", (gbif_match(name) if name else None)
            agree = f"{agree}; {note}"
        elif kg_match and kg_match.get("matchType") != "NONE":
            source, final = "kg", kg_match
        else:
            source, final = "", None
        rows.append({
            "index_name": bird,
            "family": get("Family (Latin)"),
            "register_names": "; ".join(f"{n} ({c})" for n, c in reg_names.get(bird, [])),
            "nomenclator_id": get("Nomenclator-ID") if sci else "",
            "nomenclator_name_ocr": sci or "",
            "nomenclator_match": get("Match-Methode"),
            "nomenclator_note": get("Hinweis") if isinstance(get("Hinweis"), str) else "",
            "gbif_matched_as": variant or "",
            "kg_taxon": kg_taxon or "",
            "taxon_check": agree,
            "taxon_source": source,
            "scientificName": (final or {}).get("scientificName", "") if final else "",
            "taxonRank": ((final or {}).get("rank", "") or "").lower() if final else "",
            "gbif_usageKey": (final or {}).get("usageKey", "") if final else "",
            "acceptedSpecies": (final or {}).get("species", "") if final else "",
            "gbif_speciesKey": (final or {}).get("speciesKey", "") if final else "",
            "gbif_status": (final or {}).get("status", "") if final else "",
        })
    CACHE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    out = pd.DataFrame(rows)
    out.to_csv(out_csv, index=False, encoding="utf-8")
    print(out.taxon_check.value_counts().to_string())
    print(out.taxon_source.value_counts().to_string())


if __name__ == "__main__":
    main(*sys.argv[1:4])
