# Development and reproduction

Normal workbook use, Python maintenance, and full authoring have different requirements. Opening and reviewing the delivered workbook does not require the builder or a development environment.

## Python maintenance and checks

Refresh, export, data generation, packaging, and verification use **Python 3.10+ and its standard library**. Run commands from the repository root. See the [operating guide](workflow.md) for refresh/export prerequisites.

```bash
python scripts/verify_productmatch.py --cached-only
python scripts/check_repository.py
```

The verifier checks saved results, raw inputs, output attributes, native XML features, and refresh preservation. It writes reports and exported comparisons under `.artifacts/verification/`; it does not alter the delivered workbook, CSVs, or historical reports. Cached-only checks do not establish that formulas react correctly to input changes.

To rerun full recalculation and disposable mutation scenarios, install LibreOffice and make `soffice` available, or provide its executable explicitly:

```bash
python scripts/verify_productmatch.py --soffice /path/to/soffice
```

Use the installation's actual executable path on Windows. `--quick` recalculates only the baseline and skips mutation scenarios. Full tests cover selectors/search, accept/reject, invalid IDs, variants, review dates, research eligibility, fuzzy staleness, keyed reordering, corrections, duplicate resolution, and capacity. Each scenario runs in a temporary copy with its own LibreOffice profile. `--resume-directory` is for resuming existing compatible-engine scenario outputs from the same source hash.

Verification can write evaluated caches into a **new** workbook using `--recalculated-output .artifacts/recalculated.xlsx`. It imports calculated caches while preserving source native structures. It never replaces the input. Formula recalculation and native UI testing are distinct; see [testing notes](verification.md).

## Data generation and full authoring

The original builder uses **JavaScript, Node.js 18+, and `@oai/artifact-tool`**, supplied by the specialized Codex primary authoring runtime. This dependency is not needed for refresh or CSV export. A fresh-clone full rebuild has **not** been verified, and the builder is not presented as an ordinary npm-install workflow.

The generator reproduces the original scenario and recorded research register; it does not revisit websites or establish new research access dates. Generate into a separate tree:

```bash
python scripts/generate_demo_data.py --output-dir .artifacts/rebuild
python scripts/refresh_productmatch.py --data .artifacts/rebuild/data/generated/demo_data.json --output .artifacts/rebuild/data/generated/candidate_snapshot.json
```

In a compatible authoring runtime where the artifact package resolves:

```bash
node scripts/build_productmatch.mjs --data-root .artifacts/rebuild --output-dir .artifacts/rebuild
python scripts/verify_productmatch.py --workbook .artifacts/rebuild/ProductMatch.xlsx --report-dir .artifacts/rebuild/verification --recalculated-output .artifacts/rebuild/ProductMatch_evaluated.xlsx
node scripts/render_portfolio.mjs --workbook .artifacts/rebuild/ProductMatch_evaluated.xlsx --values .artifacts/rebuild/verification/preview_values.json --output-dir .artifacts/rebuild/previews
```

These are the separated authoring steps, not a claim that the full sequence was rerun for publication. The verification command requires LibreOffice; authoring-engine caches alone are not accepted as proof of formula correctness. The builder defaults to `.artifacts/rebuild` and supports the older `PRODUCTMATCH_OUTPUT` environment variable as an output-directory fallback. Inputs come from `--data-root` or this repository's `data/generated/`.

The renderer loads evaluated values into an in-memory visual copy and writes images. It never exports that values-only copy as the workbook deliverable. Provided screenshots remain the original evaluated-baseline renders.

## Supporting artifacts

The supplied [PDF](ProductMatch_case_study.pdf) is the original portfolio attachment. Its source script uses optional **ReportLab 4.x**, with portable PDF core fonts:

```bash
python -m pip install -r requirements-authoring.txt
python scripts/create_case_study_pdf.py --output .artifacts/case-study/ProductMatch_case_study.pdf
```

The script writes a separate PDF; the delivered original is retained. Font changes can alter pagination, so visually inspect any regenerated PDF before sharing it.

For a local shareable archive after checks:

```bash
python scripts/package_project.py
```

Packaging uses Git's nonignored file list, creates SHA-256 checksums and a ZIP under `.artifacts/package`, and leaves documentation and test dates untouched. Backups, local reports, credentials, caches, and redundant delivery archives are ignored by Git.

The repository has no added open-source license. Existing source provenance, dependency attribution, and research caveats have been retained.
