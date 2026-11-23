# Laubmann registers: final dataset (2026-11-23)

Alfred Laubmann indexed his own field journals twice. Each volume 1–34 ends with a small register
(`Amsel: 3; 7; 10;` = diary pages), and volume 35 is a general index that points from each species to a
volume and a page of that volume's small register (`Amsel: I, 167; II, 172; …`). This dataset resolves every
reference of the general index to an entry of the small register it cites. A resolved reference gives the
diary pages on which Laubmann himself listed the species. No reference depends on the knowledge graph (KG);
the KG is checked against the registers afterwards.

Built by `final/build.py` on branch `final-dataset` of bird-index, from:

| input | version |
|---|---|
| general index, vol 35 | `Laubmann_35_gemini_edits (18).json` → `rectify-index.py` → index v14 → `index-fixes.csv` → **v15** |
| small registers, vols 1–34 | Laubmann-HTR reading (Gemini + TrOCR + DeepOCR), `corpus_2026-10-28_htr_registers_trocr_deepocr_noalias.json` |
| Nomenclator | Hellmayr & Laubmann 1916, BHL scan, OCR text from archive.org `nonmenclatorderv00hell` (`work/fts.txt`) |
| KG | `kg_exports_2026-10-06_regions` (85,896 occurrences) |

## Numbers

| | |
|---|---|
| species in the general index | 367 |
| references (species × volume) | 6,941 |
| resolved | **6,663 (95.99 %)**: 5,696 exact name, 387 synonym, 326 spelling variant, 254 on another register page of the volume |
| not resolved | 192 name not on the cited page, 49 entry without page numbers (`Wanderfalke: –`), 28 register page not in the corpus, 9 page without register entries |
| species × diary page links | 61,132 on 4,954 diary pages |
| KG occurrences with a diary page number | 84,218 |
| … species listed by a register on the same page | **69,509 (82.5 %)**, by chance 21.3 % (pages shifted by ±10/20/30) |
| … on the same or a neighbouring page | 74,231 (88.1 %), by chance 34.4 % |

## Rules

1. **Matching** (`add-volume-references.py` v3). A reference takes the entry with the species' name or a
   register synonym (`register-synonyms.csv`) on the cited page; failing that, the same name on another
   register page of the volume, nearest first; failing that, the most similar entry on the cited page. That
   last step may not take an entry that carries the name of another index species, and the species must be
   more similar to the entry than every other index species is. Each entry goes to one species.
2. **Scan-checked readings** (`register-readings.csv`, 5 entries). Where the reader misread a headword as the
   name of another index species, the headword was read on the scan and corrected there: two `Gänsesäger` that
   are `Gänsegeier` (V 155, XIII 171) and three `Tannenmeise` that are `Nonnenmeise` (V 151, VII 172, XI 179;
   on XI 179 both meises were read as `Tannenmeise`). `register_entry_htr` keeps the reader's text.
3. **Index fixes** (`index_fixes.csv`, 14). The row `Sylviadae` is a family heading inside the Amsel entry of
   vol 35; its 28 references belong to Amsel, which is listed on all 28 cited register pages. Thirteen page
   numbers lost their last digits in the vol-35 transcription (`XIV 18[?]`, line ends); each was restored from
   the page of the small register that lists the species. `register_page_cited` keeps the transcribed number.
4. **Other register page** (254 references) is not written back into the index: Laubmann's index may itself
   be wrong. `index_correction_candidates.csv` lists them (216 differ in one digit).
5. **Taxa** (`species.csv`). The scientific name is the KG's GBIF backbone taxon for the index name, its
   synonyms and the names the small registers use; `taxa-overrides.csv` replaces it for 28 species (hybrids
   and pairs get none). The Nomenclator 1916 entry (`nomenclator_id`, from `nomenclator-match.py` v9 plus 22
   synonyms in `nomenclator_synonyms.csv`) is the independent check: same species for 261, different for 8
   (old genera such as *Pisobia minuta*, where GBIF picks the wrong species), no Nomenclator entry for 57, an ambiguous one for 3 (Gimpel, Weidenmeise, Brachschwalbe),
   Nomenclator name not resolved by GBIF for 38. The Nomenclator name serves only as this check because GBIF
   resolves many 1916 names to the wrong species (*Serinus canarius germanicus* → canary).

## Files

`data/`
- `references.csv`: one row per general-index reference: cited and resolved register page, register scan,
  entry as matched and as read, match kind, diary pages, index fix, taxon.
- `diary_pages.csv`: one row per species × volume × diary page. This is the register layer of the dataset.
- `species.csv`: index name, names in the small registers, Nomenclator entry, taxon, check, counts.
- `kg_register_check.csv`: every present KG occurrence with `register_check`
  (`listed on page` / `listed on neighbour page` / `not listed` / `no diary page number`); columns `index` and
  `registers` give the two routes (general index, every small-register entry). Summary in `kg_register_check_summary.csv`.
- `index_fixes.csv`, `register_readings_checked.csv`, `index_correction_candidates.csv`, `unresolved_references.csv`,
  `summary.json`.

`dwca/`
- `laubmann_dwca_all.zip`: the KG archive with `identificationVerificationStatus` added to every occurrence.
- `laubmann_dwca_listed_on_page.zip`: only the 69,509 occurrences whose species the registers list on their
  own diary page, with their events and facts; `eml.xml` is that of the full archive.

`work/`: index v15 and the Nomenclator workbook (`laubmann_index_v15_nomenclator.xlsx`, sheets Nomenclator and
Parser-Report), taxa before merging, the KG check per route.

## Limits

- A listing confirms that the species is mentioned on that diary page, not the date, place, count or observer
  of a record. Common species are listed on most pages, so for them it is weak evidence.
- `not listed` is not an error flag: about half of such records are species lists and tables that Laubmann did
  not index.
- To check by a person: 11 rows of `species.csv` whose `taxon_check` says *check*; the 254 index correction
  candidates; Gimpel and Weidenmeise (subspecies in the Nomenclator).
- Superscript numbers in a register entry count as diary pages, because Laubmann adds pages that way. In a few
  entries the reader turned other marks into pages (Schnatterente XXX `111<sup>177</sup>`, a cross-reference
  `(p. 135)` under Firngrasmücke IV); about ten such pages are in `diary_pages.csv`.
