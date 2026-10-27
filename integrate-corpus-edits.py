#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Integrate validated register-page edits into the master corpus JSON.

This script takes a base corpus JSON file and one or more validation-export JSON
files produced during manual text review. Each validation file is keyed by page_id,
so the script matches those IDs against the corpus entries, strips page headers and
markers from the validated text, and overwrites the corresponding page content in
memory before writing a merged output file such as corpus-incl-vol-1-edits.json.
This keeps the workflow reusable for successive validation rounds across multiple
volumes.
"""

import argparse
import json
import re
from pathlib import Path


def load_json(path):
    with Path(path).open("r", encoding="utf-8") as fh:
        return json.load(fh)


def ensure_newlines(text):
    if text is None:
        return ""
    return str(text).replace("\r\n", "\n").replace("\r", "\n")


def strip_page_header(text):
    text = ensure_newlines(text).strip()
    # The validation export begins with a page marker like:
    # # 900847d2-aabe-..._0003_R
    text = re.sub(r"^#\s*[^\n]+\n+", "", text, count=1)
    return text.strip()


def record_text(record):
    if not isinstance(record, dict):
        return None
    for key in ("edited", "modified", "original"):
        if key in record and record[key] is not None:
            if key == "original":
                return strip_page_header(record[key])
            return strip_page_header(record[key])
    return None


def apply_edit_to_page(page, edited_text):
    full_text = strip_page_header(edited_text)
    if not isinstance(page, dict):
        return page

    regions = page.get("regions") or []
    if not regions:
        page["regions"] = [{"text": full_text}]
        return page

    preserved = dict(regions[0]) if isinstance(regions[0], dict) else {}
    preserved["text"] = full_text
    preserved["line_count"] = max(1, full_text.count("\n") + 1)
    page["regions"] = [preserved]
    return page


def iter_validated_pages(edit_doc):
    for key, value in edit_doc.items():
        if key.startswith("_"):
            continue
        if isinstance(value, dict):
            yield key, value


def integrate_edits(corpus_pages, edit_files):
    page_index = {
        page.get("page_id"): page
        for page in corpus_pages
        if isinstance(page, dict) and page.get("page_id")
    }

    for edit_file in edit_files:
        doc = load_json(Path(edit_file))
        for page_id, record in iter_validated_pages(doc):
            if page_id not in page_index:
                continue
            page = page_index[page_id]
            text = record_text(record)
            if text is None:
                continue
            apply_edit_to_page(page, text)
    return corpus_pages


def build_output_path(base_path: Path, volume: int | None):
    stem = base_path.stem
    suffix = base_path.suffix
    if volume is not None:
        out_name = f"{stem}-incl-vol-{volume}-edits{suffix}"
    else:
        out_name = f"{stem}-incl-edits{suffix}"
    return base_path.with_name(out_name)


def main():
    parser = argparse.ArgumentParser(description="Integrate validated page edits into a corpus JSON file.")
    parser.add_argument("--corpus", required=True, help="Base corpus JSON file, e.g. corpus.json")
    parser.add_argument("--edit-file", action="append", default=[], help="Validation JSON file(s) to merge (repeatable)")
    parser.add_argument("--volume", type=int, help="Volume number used in the output file name")
    parser.add_argument("--output", help="Explicit output path; overrides the auto-generated file name")
    args = parser.parse_args()

    corpus_path = Path(args.corpus)
    corpus_data = load_json(corpus_path)
    if not isinstance(corpus_data, list):
        raise ValueError(f"{corpus_path} is not a JSON list of corpus pages.")

    edit_files = [Path(p) for p in args.edit_file]
    if not edit_files:
        raise ValueError("At least one --edit-file is required.")

    merged = integrate_edits(corpus_data, edit_files)

    output_path = Path(args.output) if args.output else build_output_path(corpus_path, args.volume)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fh:
        json.dump(merged, fh, ensure_ascii=False, indent=2)
        fh.write("\n")

    print(f"Wrote merged corpus to: {output_path}")
    print(f"Applied {len(edit_files)} validation file(s).")
    print(f"Matched pages: {sum(1 for p in merged if isinstance(p, dict) and p.get('page_id'))}")


if __name__ == "__main__":
    main()
