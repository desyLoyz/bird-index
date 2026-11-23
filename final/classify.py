"""Classify every resolved reference of a matching run and check it against the KG (evaluation only).

    python final/classify.py <references.csv> [<references.csv> ...]
"""
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

HTR = Path(r"C:\Users\totom\Projects\laubmann-htr")
sys.path.insert(0, str(HTR / "eval"))
import confirm as C  # noqa: E402
from validate_kg import load_matching_script  # noqa: E402

M = load_matching_script()
SYN = pd.read_csv(HTR / "matching/register-synonyms.csv", dtype=str, keep_default_na=False, encoding="utf-8-sig")


def owners_table(species):
    owners = defaultdict(set)
    for s in species:
        owners[M.compact_name(s)].add(s)
    for row in SYN.itertuples(index=False):
        owners[M.compact_name(row[1])].add(row[0])
    return owners


def kind(ref, owners):
    if not ref.status.startswith("ok"):
        return ref.status
    matched = M.compact_name(ref.matched_name)
    own = owners.get(matched, set())
    if ref.species in own:
        k = "exact" if matched == M.compact_name(ref.species) else "synonym"
    elif own:
        k = "other-index-species"
    else:
        k = "fuzzy"
    return k + ("+other-page" if str(ref.found_page).strip() else "")


_KG = {}


def kg():
    if not _KG:
        _KG["index"] = C.journal_page_index(C.JOURNAL_CORPUS)
        _KG["kg"] = C.SpeciesOnPage(C.kg_records())
    return _KG["index"], _KG["kg"]


def check_names(ref, owners):
    """the index name, its synonyms, and the register name unless it is another index species' name"""
    names = [ref.species] + [row[1] for row in SYN.itertuples(index=False) if row[0] == ref.species]
    if ref.matched_name and owners.get(M.compact_name(ref.matched_name), {ref.species}) <= {ref.species}:
        names.append(ref.matched_name)
    return names


def supported(names, volume, page, page_index, species_kg):
    on_page = page_index.get((volume, page), [])
    if not on_page:
        return None
    near = on_page + page_index.get((volume, page - 1), []) + page_index.get((volume, page + 1), [])
    return any(species_kg.present(n, uid) for n in names for uid in near)


def ref_support(refs, owners, shifts=(-30, -20, -10, 10, 20, 30)):
    """KG pages with a record of the species on the diary page or a neighbour; same for shifted pages (chance)"""
    page_index, species_kg = kg()
    cols = {"kg_pages": [], "kg_confirmed": [], "chance_pages": [], "chance_confirmed": []}
    for ref in refs.itertuples():
        pages = [int(C.re.match(r"\d+", p).group()) for p in str(ref.diary_pages).split() if C.re.match(r"\d+", p)]
        if not ref.status.startswith("ok"):
            pages = []
        names = check_names(ref, owners)
        real = [supported(names, int(ref.volume), p, page_index, species_kg) for p in pages]
        fake = [supported(names, int(ref.volume), p + s, page_index, species_kg) for p in pages for s in shifts]
        real, fake = [x for x in real if x is not None], [x for x in fake if x is not None]
        for key, val in zip(cols, (len(real), sum(real), len(fake) / len(shifts), sum(fake) / len(shifts))):
            cols[key].append(val)
    for key, val in cols.items():
        refs[key] = val
    return refs


def load(path):
    refs = pd.read_csv(path, keep_default_na=False, dtype={"found_page": str})
    owners = owners_table(refs.species.unique())
    refs["kind"] = [kind(r, owners) for r in refs.itertuples()]
    return ref_support(refs, owners)


def report(refs, label):
    g = refs[refs.kg_pages > 0].groupby("kind")
    t = pd.DataFrame({"refs": refs.groupby("kind").size(),
                      "kg_pages": g.kg_pages.sum(), "kg_confirmed": g.kg_confirmed.sum()}).fillna(0)
    t["chance_pages"], t["chance_confirmed"] = g.chance_pages.sum(), g.chance_confirmed.sum()
    t["confirmed_%"] = (100 * t.kg_confirmed / t.kg_pages).round(1)
    t["chance_%"] = (100 * t.chance_confirmed / t.chance_pages).round(1)
    t = t.drop(columns=["chance_pages", "chance_confirmed"])
    print(f"== {label}  resolved {refs.status.str.startswith('ok').sum()} of {len(refs)}")
    print(t.to_string())


if __name__ == "__main__":
    for p in sys.argv[1:]:
        r = load(p)
        r.to_csv(Path(p).with_name("references_classified.csv"), index=False, encoding="utf-8")
        report(r, p)
