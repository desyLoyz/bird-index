#!/usr/bin/env python3
"""
Laubmann General Index Parser
==============================
Parses the ornithological index pages from the Laubmann diary JSON,
rectifying inconsistent reference notation into a clean structured table.

Output columns:
  Bird name (German, edited) | Original name (if different) | Family (Latin) | References (rectified)

No AI technologies used — pure regex + rule-based text processing.
Auto-installs missing dependencies on first run.
"""

import subprocess, sys

def _ensure(pkg, import_name=None):
    try:
        __import__(import_name or pkg)
    except ImportError:
        print(f"Installing {pkg} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

_ensure("pandas")
_ensure("openpyxl")

# ─────────────────────────────────────────────────────────────────────────────
import json, re
from pathlib import Path
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
JSON_PATH    = "Laubmann_35_gemini_edits-170.json"
OUTPUT_EXCEL = "laubmann_index.xlsx"
OUTPUT_CSV   = "laubmann_index.csv"

# Page key from which the index begins (inclusive)
INDEX_START_PAGE = "2d6d52f0-b55f-431d-81ac-e2b8ce255fb2_0004_R"

# Roman numeral map (I–XXXV)
ROMAN = {
    "I":1,"II":2,"III":3,"IV":4,"V":5,"VI":6,"VII":7,"VIII":8,
    "IX":9,"X":10,"XI":11,"XII":12,"XIII":13,"XIV":14,"XV":15,
    "XVI":16,"XVII":17,"XVIII":18,"XIX":19,"XX":20,"XXI":21,
    "XXII":22,"XXIII":23,"XXIV":24,"XXV":25,"XXVI":26,"XXVII":27,
    "XXVIII":28,"XXIX":29,"XXX":30,"XXXI":31,"XXXII":32,"XXXIII":33,
    "XXXIV":34,"XXXV":35,
}
# Sorted longest-first so regex matches greedily
ROMAN_PATTERN = "|".join(sorted(ROMAN.keys(), key=len, reverse=True))

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — TEXT CLEANING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def strip_all_tags(text: str) -> str:
    """Remove ALL HTML tags (including <font color=...>, <u>, etc.)."""
    return re.sub(r"<[^>]+>", "", text)

def clean_page_text(text: str) -> str:
    """
    Full cleaning pipeline for a single page's edited text:
      1. Remove HTML tags
      2. Remove markdown headings (# ...)
      3. Remove leading list markers (- at line start)
      4. Remove stray artefacts: {Stempel:}, [?], p. before page numbers
      5. Normalise whitespace
    """
    text = strip_all_tags(text)
    # Remove markdown heading lines
    text = re.sub(r"^\s*#[^\n]*\n?", "", text, flags=re.MULTILINE)
    # Remove leading list dashes
    text = re.sub(r"^\s*-\s+", "", text, flags=re.MULTILINE)
    # Remove annotation artefacts
    text = re.sub(r"\{[^}]*\}", "", text)        # {Stempel:} etc.
    text = re.sub(r"\[\??\]", "", text)           # [?] or []
    text = re.sub(r"\bp\.\s*(?=\d)", "", text)    # "p. 169" → "169"
    # Collapse multiple blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — FAMILY HEADER DETECTION
# ─────────────────────────────────────────────────────────────────────────────

# A family header is a Latin name (capitalised, may contain spaces) followed
# by an optional dot, standing alone on a line (possibly with trailing punct).
# Examples: "Corvidae.", "Fringillidae", "Alaudidae."
FAMILY_RE = re.compile(
    r"^\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s*\.\s*$",
    re.MULTILINE
)

def extract_family_from_page(text: str) -> str | None:
    """Return the first Latin family name found on the page, or None."""
    m = FAMILY_RE.search(text)
    return m.group(1).strip() if m else None

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — REFERENCE STRING RECTIFICATION
# ─────────────────────────────────────────────────────────────────────────────

def rectify_references(raw_refs: str) -> str:
    """
    Convert a messy reference string into canonical "VOL, PAGE; VOL, PAGE; …"
    format.

    Handles:
      • Dots used as separators instead of semicolons  (I, 153. II, 161)
      • Missing comma between volume and page          (I 153)
      • Stray characters: 'l', trailing punctuation
      • Line-wrapped references (newlines inside a ref block)
      • Comma used as separator instead of semicolon   (XXI, 170, XXII, 171)
      • Malformed page numbers like 19[?] → kept as-is with note
    """
    # Flatten newlines within the reference block
    s = re.sub(r"\s*\n\s*", " ", raw_refs).strip()

    # Remove stray single letters that are OCR noise (e.g. ";l XVI" → "; XVI")
    s = re.sub(r"(?<=\d)\s*[a-km-z]\s+(?=[IVXLCDM])", " ", s)

    # Normalise all separators between references to " ; "
    # A new reference starts when we see a Roman numeral after a separator
    # Strategy: tokenise into (volume, page) pairs, then reassemble.

    # First, unify dots-as-separators: replace ". ROMAN" with "; ROMAN"
    # but only when the dot follows a page number (digits)
    s = re.sub(
        r"(\d)\s*[.]\s*(?=(?:" + ROMAN_PATTERN + r")(?:\s*,|\s+\d))",
        r"\1; ",
        s
    )

    # Replace comma-as-separator between two references:
    # pattern: digit , ROMAN  →  digit ; ROMAN
    s = re.sub(
        r"(\d)\s*,\s*(?=(?:" + ROMAN_PATTERN + r")(?:\s*[,;]|\s+\d))",
        r"\1; ",
        s
    )

    # Ensure volume and page are separated by ", " (handle missing comma)
    s = re.sub(
        r"(?<!\d)(" + ROMAN_PATTERN + r")\s+(\d)",
        r"\1, \2",
        s
    )

    # Remove the species-name separator at the very start (": " or ".- " etc.)
    s = re.sub(r"^[\s:.\-]+", "", s)

    # Normalise multiple spaces / semicolons
    s = re.sub(r";\s*;", ";", s)
    s = re.sub(r"\s{2,}", " ", s)
    s = re.sub(r";\s*$", "", s).strip()

    # Final pass: ensure every ";" is followed by exactly one space
    s = re.sub(r"\s*;\s*", "; ", s)

    return s

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — SPECIES ENTRY DETECTION & PARSING
# ─────────────────────────────────────────────────────────────────────────────

# A species entry looks like:
#   <BirdName>[optional colon or dot or dash]: <references>
# The name may contain letters, spaces, brackets ([x]), dots (Sib.)
# The separator between name and refs can be:  ":  "  ".-"  ". "  ":"
# References always contain at least one Roman numeral followed by a number.

SPECIES_ENTRY_RE = re.compile(
    # Bird name: starts at line start (after optional whitespace),
    # must begin with uppercase, may contain word chars, spaces, brackets, dots
    r"^((?:[A-ZÄÖÜ\[])[^\n:]{1,60}?)"
    # Separator: colon, dot-dash, dot, or just whitespace before Roman numeral
    r"\s*[:.\-]+\s*"
    # Reference block: everything until a blank line or end of string
    r"((?:(?:" + ROMAN_PATTERN + r")[\s,;.\d\[\]?lI]+)+)",
    re.MULTILINE
)

# Alternative: name ends with ":" on same line, refs may span multiple lines
SPECIES_BLOCK_RE = re.compile(
    r"^((?:[A-ZÄÖÜ][^\n:]{1,60}?))\s*:+\s*\n?"
    r"((?:.+\n?)+?)(?=\n\n|\Z|^[A-ZÄÖÜ][^\n:]{1,60}?\s*:)",
    re.MULTILINE
)

def parse_species_entries(text: str):
    """
    Extract all (bird_name, raw_references) pairs from a cleaned page text.
    Returns list of (name_raw, refs_raw) tuples.
    """
    entries = []
    seen_spans = []

    # Primary strategy: find "Name: refs" patterns where refs contain Roman nums
    ref_block_re = re.compile(
        r"^((?:[A-ZÄÖÜ\[][^\n:]{1,60}?))\s*[:.\-]+\s*"
        r"((?:(?:" + ROMAN_PATTERN + r")\s*,\s*\d[^\n]*(?:\n(?![ \t]*\n)[^\n]+)*?))"
        r"(?=\n\n|\n(?:[A-ZÄÖÜ\[]|\Z)|\Z)",
        re.MULTILINE
    )

    for m in ref_block_re.finditer(text):
        name = m.group(1).strip()
        refs = m.group(2).strip()
        # Skip if this looks like a family header (single Latin word + dot)
        if re.fullmatch(r"[A-Z][a-z]+", name):
            continue
        entries.append((name, refs, m.start()))
        seen_spans.append((m.start(), m.end()))

    return entries

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — NAME NORMALISATION
# ─────────────────────────────────────────────────────────────────────────────

# Known aliases / original names that differ from the edited heading
# (e.g. "Hausmeister" is the original name on the page for "Haussperling")
KNOWN_ALIASES = {
    "Hausmeister": "Haussperling (Hausmeister)",
}

def normalise_name(raw: str):
    """
    Clean a raw bird name string.
    Returns (edited_name, original_name_if_different).
    """
    # Remove trailing colons, dots, dashes
    name = re.sub(r"[\s:.\-]+$", "", raw).strip()
    # Remove double colons
    name = re.sub(r":+", "", name).strip()

    original = None
    if name in KNOWN_ALIASES:
        original = name
        name = KNOWN_ALIASES[name]

    return name, original

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 — MAIN PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def main():
    json_file = Path(JSON_PATH)
    if not json_file.exists():
        raise FileNotFoundError(
            f"'{JSON_PATH}' not found.\n"
            "Place the JSON file in the same folder as this script."
        )

    with open(json_file, encoding="utf-8") as fh:
        data = json.load(fh)

    # Collect ordered page keys (excluding metadata keys)
    all_page_keys = [k for k in data.keys() if not k.startswith("_")]

    # Find start index
    try:
        start_idx = all_page_keys.index(INDEX_START_PAGE)
    except ValueError:
        print(f"Warning: start page '{INDEX_START_PAGE}' not found. "
              "Processing all pages.")
        start_idx = 0

    index_pages = all_page_keys[start_idx:]
    print(f"Processing {len(index_pages)} index pages …")

    # ── Pass 1: collect all pages, track current family ──────────────────────
    records = []
    current_family = None
    skipped_names  = []

    for page_key in index_pages:
        entry = data.get(page_key, {})
        raw_text = entry.get("edited") or entry.get("original", "")
        if not raw_text:
            continue

        cleaned = clean_page_text(raw_text)

        # Update family if this page has a header
        page_family = extract_family_from_page(cleaned)
        if page_family:
            current_family = page_family

        # Extract species entries
        species_entries = parse_species_entries(cleaned)

        for (name_raw, refs_raw, _pos) in species_entries:
            name_edited, name_original = normalise_name(name_raw)

            # Skip empty or suspiciously short names
            if len(name_edited) < 3:
                skipped_names.append((page_key, name_raw))
                continue

            refs_clean = rectify_references(refs_raw)

            records.append({
                "Bird Name (German)"  : name_edited,
                "Original Name"       : name_original or "",
                "Family (Latin)"      : current_family or "",
                "References"          : refs_clean,
                "Source Page"         : page_key,
            })

    # ── Build DataFrame ───────────────────────────────────────────────────────
    df = pd.DataFrame(records, columns=[
        "Bird Name (German)", "Original Name", "Family (Latin)",
        "References", "Source Page"
    ])

    # De-duplicate: if the same bird appears on multiple pages (line-wrapped),
    # merge their references
    def merge_refs(series):
        combined = "; ".join(s for s in series if s)
        # Re-rectify the merged string to remove duplicate separators
        return rectify_references(combined)

    df_merged = (
        df.groupby(["Bird Name (German)", "Original Name", "Family (Latin)"],
                   sort=False, as_index=False)
          .agg({"References": merge_refs, "Source Page": lambda s: ", ".join(s.unique())})
    )

    # Sort by family, then bird name
    df_merged = df_merged.sort_values(
        ["Family (Latin)", "Bird Name (German)"]
    ).reset_index(drop=True)

    # ── Print preview ─────────────────────────────────────────────────────────
    print(f"\nTotal species entries parsed: {len(df_merged)}\n")
    pd.set_option("display.max_colwidth", 80)
    pd.set_option("display.max_rows", 200)
    print(df_merged[["Bird Name (German)", "Original Name",
                      "Family (Latin)", "References"]].to_string(index=True))

    if skipped_names:
        print(f"\n[!] {len(skipped_names)} entries skipped (name too short):")
        for pk, nm in skipped_names[:10]:
            print(f"    page={pk}  name={repr(nm)}")

    # ── Export ────────────────────────────────────────────────────────────────
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        df_merged.to_excel(writer, sheet_name="Species Index", index=True)
        # Also export raw (pre-merge) for debugging
        df.to_excel(writer, sheet_name="Raw Entries", index=False)

    df_merged.to_csv(OUTPUT_CSV, index=True, encoding="utf-8-sig")

    print(f"\n✓  Excel  →  {OUTPUT_EXCEL}")
    print(f"✓  CSV    →  {OUTPUT_CSV}")

if __name__ == "__main__":
    main()