# ProductMatch — Product Reconciliation & Data Quality Workspace

**Ridla A. Gani** · Excel 365 portfolio project · November 2026

## Problem

A fictional electronics and office-accessories retailer needs to reconcile 250 supplier records against 200 internal catalog records. Inconsistent names, brand aliases, missing fields, retransmissions and similar-looking variants make a simple SKU lookup unreliable.

## Approach

I built a macro-free Excel workspace that preserves both raw feeds, standardizes product attributes, detects exact and suspected duplicates, and prioritizes manufacturer identifiers before brand/model, normalized attributes and fuzzy name suggestions. Color, capacity, size/hardware version and pack quantity remain part of the identity checks.

Reviewers retain control through validated, keyed decisions. The approved master ID is stored as a value, so a refreshed suggestion cannot silently remap an accepted record. Conflicting variants, invalid inputs and incomplete approvals remain outside the final export. Category, brand and review-outcome selectors drive the dashboard and native charts; product search supports exception review.

I researched 15 real product models using official manufacturer pages and specification sheets. Each field records the exact model, proposed value, evidence, URL, access date and verification status. Approved specification values flow into the catalog; genuine source and hardware-version contradictions remain unresolved.

## Tools

Excel tables, INDEX/MATCH and VLOOKUP, COUNTIFS, SUMPRODUCT, text formulas, validation, conditional formatting, hyperlinks and native charts; Python standard library for reproducible synthetic data, refreshable top-three fuzzy candidates, CSV exports and independent QA; JavaScript with `@oai/artifact-tool` for workbook authoring and rendering.

## Measured demonstration results

| Measure | Delivered baseline |
|---|---:|
| Master input / active master products | 200 / 198 |
| Supplier input records | 250 |
| Accepted mappings exported | 124 |
| Awaiting review / unmatched / rejected | 88 / 18 / 5 |
| Confirmed supplier duplicate rows excluded | 15 |
| Remaining exception records | 111 |
| Suspected duplicate rows awaiting resolution | 27 |
| Required-field completeness, before / after | 98.52% / 98.72% |
| Verified research fields / total researched or queued | 15 / 22 |

The supplier partition reconciles exactly: **124 + 88 + 18 + 5 + 15 = 250**. Duplicate exclusions retain the raw rows. The output grain is one accepted supplier source row per mapping, allowing distinct supplier SKUs to map to one canonical product after review.

On the deliberately constructed nonduplicate evaluation set, **204 of 217 suggestions** identify the expected master product (**94.0% suggested-mapping precision**). The same 204 correct suggestions cover **215 known positive cases** (**94.9% positive suggestion recall**). Thirteen suggestions are incorrect; eleven positive cases lack a correct suggestion. These limitations justify human review and variant checks. The 124 seeded accepted decisions agree with the constructed expected IDs; this is a workflow-control check, not a production accuracy claim.

Completeness measures populated cells across ten required fields per supplier record, using the same 2,500-cell denominator before and after. The five-cell difference comes from demonstrated corrections and accepted master enrichment. It does not measure accuracy, productivity or business impact.

## Validation and limits

The supplied test report separates programmatic checks and full headless LibreOffice recalculation from native application testing. Microsoft Excel 365 and Google Sheets were not available for native UI verification. The workbook targets Excel 365; a Sheets conversion and testing guide is included.

The commercial dataset, supplier SKUs, prices and reviewer decisions are synthetic. Real model facts and research evidence are distinguished explicitly. Fuzzy scores are refreshable Python snapshots, not probabilities or live Excel similarity calculations. Normal review, dashboard and output formulas need no Python. Capacity is 300 master, 350 supplier and 100 research records. Only the Specification research field applies automatically in this version. Suspected duplicates require complete model/variant keys and can miss incomplete records. Manual tables retain current state rather than a complete automatic history.

No client engagement, time saving, production accuracy improvement or financial impact is claimed.

See the [operating guide](workflow.md), [verification notes](verification.md), and [research register](../data/reference/research_evidence.csv) for review prerequisites, test boundaries, and source evidence.
