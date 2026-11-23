"""Index of vol 35 (rectify-index.py output) -> corrected index: rows merged and page numbers restored
as listed in index-fixes.csv. The input sheet stays as it is; the fixes are written to a second sheet.

    python final/fix_index.py laubmann_index_v14.xlsx index-fixes.csv laubmann_index_v15.xlsx
"""
import re
import sys

import pandas as pd

SHEET = "Species Index"
REF = re.compile(r"\b([IVXLCDM]+)\s*,\s*(\d+)\b")


def refs_of(text):
    return [(m.group(1), int(m.group(2))) for m in REF.finditer(text if isinstance(text, str) else "")]


def join(refs):
    return "; ".join(f"{roman}, {page}" for roman, page in refs)


def main(index_in, fixes_csv, index_out):
    index = pd.read_excel(index_in, sheet_name=SHEET)
    extra = {name: df for name, df in pd.read_excel(index_in, sheet_name=None).items() if name != SHEET}
    fixes = pd.read_csv(fixes_csv, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    name = index["Bird Name (German)"].astype(str).str.strip()
    row_of = {n: i for i, n in zip(index.index, name)}

    for fix in fixes.itertuples(index=False):
        row = row_of[fix.species]
        refs = refs_of(index.at[row, "References"])
        if fix.action == "page":
            old = (fix.roman, int(fix.cited_page))
            assert old in refs, (fix.species, old)
            refs[refs.index(old)] = (fix.roman, int(fix.fixed_page))
        elif fix.action == "add":
            refs.append((fix.roman, int(fix.fixed_page)))
        elif fix.action == "merge":
            target = row_of[fix.fixed_page]
            index.at[target, "References"] = join(refs_of(index.at[target, "References"]) + refs)
            refs = None
        else:
            raise ValueError(fix.action)
        if refs is None:
            index = index.drop(index=row)
        else:
            index.at[row, "References"] = join(refs)

    with pd.ExcelWriter(index_out, engine="openpyxl") as writer:
        index.to_excel(writer, sheet_name=SHEET, index=False)
        fixes.to_excel(writer, sheet_name="Index-Fixes", index=False)
        for sheet, df in extra.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
    print(f"{index_out}: {len(index)} species, {sum(len(refs_of(r)) for r in index['References'])} references, {len(fixes)} fixes")


if __name__ == "__main__":
    main(*sys.argv[1:4])
