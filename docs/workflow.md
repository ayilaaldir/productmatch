# Using ProductMatch

Open [ProductMatch.xlsx](../ProductMatch.xlsx) in Excel 365 desktop with Automatic calculation. Begin at START HERE. The workbook is macro-free; Python is only needed to refresh fuzzy candidates or export saved CSV results.

## Review a supplier feed

Use DASHBOARD B5 for category, D5 for brand, F5 for outcome, and H5 for case-insensitive search. `All` resets a selector; clearing H5 resets search. Selectors affect workload metrics and both charts. Search affects only the first-12-results view. Overall input counts remain unchanged.

Filter the CLEANING issue queue, then enter overrides in CORRECTIONS using the stable supplier source ID and a reason. Blank overrides retain original values. Keep MASTER RAW and SUPPLIER RAW intact; correcting a record should not erase the evidence of its original state.

Exact supplier retransmissions automatically exclude later raw rows. Suspected duplicates need a decision in DUPLICATE REVIEW. `Keep separate` requires reason, reviewer, and date. `Confirm duplicate` also requires an existing, lexically earlier supplier source ID with the same complete brand/model/variant identity. No raw rows are deleted.

In REVIEW DECISIONS, choose `Accept suggested match`, `Select different master`, `Reject suggestion`, or `Leave pending`. To accept, copy the intended master ID into column C, **Approved master ID**. This is a stored value, independent of the changing suggestion. Accepted decisions require a valid active master ID, reviewer, and date; fuzzy acceptance also requires notes. Ambiguous suggestions need explicit alternative selection. Attribute conflicts and invalid required fields continue to block export until resolved.

FINAL CLEAN CATALOG contains accepted, nonduplicate supplier mappings. EXCEPTIONS contains pending, unmatched, and rejected supplier rows. These are formula-driven workbook views. CSV files update only when exported again.

The four editable tables—REVIEW DECISIONS, CORRECTIONS, DUPLICATE REVIEW, and RESEARCH—store current state. Save dated workbook versions if a historical record of decisions is needed.

## Matching, duplicates, and research

Matching follows brand-scoped exact nonblank MPN, exact brand/model/full variant, brand/model fallback, exact normalized name/attributes, then fuzzy candidate. Supplier SKU is not equated with internal SKU. No verified GTIN was available; GTIN matching is not claimed. Known missing-GTIN placeholders normalize to blank. Identifiers remain text.

Color, capacity, size/hardware version, and pack quantity remain part of identity and conflict checks. The supplier retransmission key also includes clean supplier SKU, normalized name, price, and currency. The earliest raw row survives. Master duplicates use full brand/model/variant identity with the earliest raw row surviving. Different supplier SKUs with the same complete identity remain suspected duplicates until reviewed.

RESEARCH stores exact model, attribute, proposed value, evidence, source URL/title, access date, verification status, and application approval. Only `Specification` with `Verified`, `Yes`, exact model identity, complete evidence, and a unique approved field key applies automatically. Other attributes remain an evidence queue. [Research provenance](../data/README.md) explains unresolved fields and source caveats.

## Refresh and export

Run these commands from the repository root with Python 3.10 or later. No third-party Python packages are needed.

Save and close the workbook, then refresh in place:

```bash
python scripts/refresh_productmatch.py --workbook ProductMatch.xlsx --in-place
```

The command creates `ProductMatch.xlsx.bak`, changes CANDIDATES and calculation properties, and checks that the four manual-input sheets remain byte-identical. To preserve the original file instead:

```bash
python scripts/refresh_productmatch.py --workbook ProductMatch.xlsx --output .artifacts/refreshed.xlsx
```

The score is **0.6 × sorted-token `difflib.SequenceMatcher` ratio + 0.4 × token-set Jaccard**. Candidates are pooled by canonical category and brand; a missing brand broadens within category. Only the top three are retained. Scores are similarities, not statistical probabilities or trained-model confidence.

Changed supplier signatures or changed/appended/deleted master signatures hide stale fuzzy suggestions and display `Refresh required`. Exact native tiers continue to work. SETTINGS B7:B8 changes the minimum score and ambiguity gap against stored scores without recomputing similarities. Refresh after reordering MASTER RAW as well: a different earliest-row duplicate survivor can change candidates even when signatures still show Current.

After refresh, reopen, calculate, and save in Excel. Then export:

```bash
python scripts/refresh_productmatch.py --workbook ProductMatch.xlsx --export-cached exports
```

This reads saved formula caches and cannot calculate formulas. Saving a workbook with stale caches produces stale CSVs. Research and review dates export as ISO dates. In Excel, use **Data → From Text/CSV** and set identifier columns to **Text**; double-clicking a CSV can lose leading zeros through type inference.

## Add records within capacity

Capacity is 300 master, 350 supplier, and 100 research records, beginning at row 7. Paste values into blank reserved raw rows without overwriting headers or calculated columns. Use unique text IDs such as `MR0201` / `M0201` and `S0251`. Add each new supplier ID to an unused REVIEW DECISIONS row and to CORRECTIONS or DUPLICATE REVIEW when needed. Refresh candidates after adding records or changing matching attributes.

Sort complete tables, never an isolated column. Decisions are keyed by source IDs rather than editable-table row positions. Extend aliases and categories only within the reserved settings ranges; new dashboard categories and brands also need updated selector lists. Beyond capacity, extend bounded formulas, tables, validation lists, and scripts together.

Required completeness covers ten fields: name, brand, category, model, color, capacity, size/version, pack, price, and currency. Explicit `N/A` counts as populated. Before uses raw populated cells; after uses accepted master enrichment for accepted rows and cleaned supplier cells for the rest. Both use ten times the in-scope supplier row count. This measures presence, not validity. Dashboard exception counts are records, not individual defects; duplicate exclusions remain part of the selectable population.
