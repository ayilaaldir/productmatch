# Verification notes

Two verification sessions are recorded separately. The original report describes the delivered workbook's compatible-engine behavior. Repository preparation checks use the existing saved results and do not imply a new formula-engine or native-application test.

## Original verification: 11 November 2026

The preserved [report](testing/2026-11-11-report.md), [14 test-group results](testing/2026-11-11-results.json), and [baseline metrics](testing/baseline_metrics.json) document full headless LibreOfficeDev 26.8.0.0.alpha0 recalculation plus independent Python and OOXML checks. Mutation scenarios used disposable copies with formula caches discarded before calculation.

Those scenarios covered selectors/search and chart sources, accept/reject flow, invalid IDs, variant conflicts, reviewer dates, research eligibility, stale fuzzy candidates, stable keyed decisions after reordering, correction overrides, duplicate resolution, and capacity boundaries. No native Excel 365 or native Google Sheets UI test was performed. Original screenshots are evaluated workbook-layout renders, not application captures.

## Repository preparation: 20 November 2026

Rerun from the repository root with Python 3.11.9 on Windows:

```bash
python scripts/verify_productmatch.py --cached-only
python scripts/check_repository.py
```

The [saved-result verification](testing/2026-11-20-workbook.json) and [repository checks](testing/2026-11-20-repository.json) record the executed groups and workbook SHA-256. The checks cover:

- Saved formula error scan, eight reconciliation cells, raw data preservation, selected-master output attributes, and demonstrated duplicate/variant outcomes.
- Macro-free archive, 17 worksheets, 13 table filters, two linked charts, validation XML, and frozen working headers.
- Candidate refresh on a disposable copy: every archive part except CANDIDATES and workbook calculation properties remains byte-identical, including all four manual-input sheets. In-place refresh also creates an exact backup and succeeds on Windows.
- Export CLI agreement with all four delivered CSV datasets, including text identifiers such as `0010001`.
- Generation of the original raw data, research register, constructed labels, and builder inputs; fuzzy candidate scores and signatures agree when new generation timestamps are excluded.
- Independent evaluation from constructed labels: 204 correct suggestions out of 217, 215 positive cases, 124 accepted mappings, 15 excluded duplicates, and 111 exceptions. Research dates remain 11 November 2026.
- Relative documentation and image paths resolve within the repository.

Repository changes also update script/document paths in the workbook's instruction text. The workbook was compared with the original: formulas, formula caches, raw records, manual inputs, styles, tables, validation, charts, and relationship parts are preserved. The original datasets, CSV exports, research register, screenshots, and supplied case-study PDF are retained.

Python scripts compiled and JavaScript sources passed Node syntax checks. The updated workbook instructions were imported, inspected, rendered, and visually checked with the bundled artifact runtime. The three portfolio screenshots and original one-page A4 case-study PDF were visually reviewed. The PDF generator ran with ReportLab 4.4.9 using portable core fonts; its separate output was rendered with Poppler and visually checked. Packaging produced a valid ZIP with checksums in an ignored output directory. These supporting checks did not rebuild or replace the delivered workbook or PDF.

## Checks not rerun

LibreOffice is not installed in the publication environment, so full formula recalculation and input-mutation scenarios were not rerun. Native Excel 365 and Google Sheets UI behavior remains unverified. A fresh-clone full rebuild with the specialized artifact authoring runtime was not executed. Research sources were not revisited; original access dates and source caveats are preserved.

The verifier now writes new reports and disposable exports under `.artifacts/` instead of overwriting the workbook, delivered CSVs, or historical test evidence. [Development notes](development.md) explain how to run compatible-engine tests and the separate authoring workflow.
