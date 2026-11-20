# Data and provenance

The commercial scenario is synthetic. All prices, supplier offers, supplier/internal SKUs, corrections, decisions, and trading outcomes are demonstration data. Fifteen real models provide a separate research exercise; fictional products use `DEMO-` models and the Northstar, Vertex, Luma, and Oakline brands.

| Location | Contents |
| --- | --- |
| [raw/master_raw.csv](raw/master_raw.csv) | 200 original master rows, including two duplicate records |
| [raw/supplier_raw.csv](raw/supplier_raw.csv) | 250 original supplier rows, including 15 exact retransmissions |
| [reference/research_evidence.csv](reference/research_evidence.csv) | 22 field-level evidence or queued records; 15 verified Specification fields |
| [evaluation/ground_truth.csv](evaluation/ground_truth.csv) | Deliberately constructed counterpart labels and duplicate exclusions |
| [evaluation/scenario_issues.csv](evaluation/scenario_issues.csv) | Annotated issue manifest for evaluation and walkthroughs |
| [generated/demo_data.json](generated/demo_data.json) | Builder inputs: raw data, research, aliases, and seeded manual state |
| [generated/candidate_snapshot.json](generated/candidate_snapshot.json) | Delivered top-three fuzzy candidates, signatures, and generation timestamp |

Evaluation labels never feed operational matching or approval formulas. The issue manifest is an included annotated fixture; the generator recreates raw data, research, labels, and builder inputs, not this separate annotation file.

Research records retain the original **11 November 2026** access dates. They cite Logitech, Kingston, TP-Link, SanDisk, Crucial/Micron, Samsung, Epson, and Brother official pages or specification sheets. URLs, exact models, evidence notes, and direct-fetch caveats remain in the register. The repository preparation did not revisit those sources or change their access dates.

TL-SG105 power depends on an unspecified hardware revision. SanDisk's SDHC/SDXC label conflict remains unapplied. M185 GTIN was not found in the reviewed page, and four other GTIN fields remain pending. SanDisk used official indexed specifications after a direct-request timeout; Crucial used an official indexed table after a direct PDF request redirected to a reseller mirror.

The [exports](../exports/) are snapshots of the delivered workbook's saved results. The clean catalog has one accepted supplier source row per mapping, rather than one row per master product. Several distinct supplier SKUs may map to one master after review. Read CSV identifiers as text; opening a CSV by double-click can cause spreadsheet software to infer numbers and discard leading zeros.
