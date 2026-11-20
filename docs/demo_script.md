# Two-minute demonstration

**Author / presenter: Ridla A. Gani**

| Time | On-screen action | Narration |
|---|---|---|
| 0:00–0:15 | Open DASHBOARD, show overall counts. | “I built ProductMatch to reconcile a supplier feed against an internal product catalog. This demonstration contains 200 master rows and 250 supplier rows. Trading data is synthetic, with 15 real models included for source-based research.” |
| 0:15–0:30 | Select Storage in B5; return to All. Type M185 in H5; clear search. | “Category, brand and review-status selectors update the workload metrics and charts. Search narrows the product and exception list. Overall counts remain visible, and completeness percentages have an explicit denominator.” |
| 0:30–0:48 | Open CLEANING, show raw/clean names and aliases. Show S0201. | “The original feeds remain intact. Formulas standardize whitespace and aliases, preserve text identifiers, and flag missing or invalid fields. Exact retransmissions are excluded by a documented survivor rule, while suspected duplicates remain reviewable.” |
| 0:48–1:10 | PRODUCT MATCHING: inspect S0181 and S0191. | “Matching prioritizes nonblank manufacturer identifiers, then brand, model and variant attributes. Fuzzy name similarity is only a candidate suggestion. Different colors, capacity, hardware versions or pack sizes trigger review even when names look almost identical.” |
| 1:10–1:30 | REVIEW DECISIONS: S0170, choose Accept suggested match; confirm column C is M0170; enter presenter name and today's date. Show final output; then restore Leave pending for baseline. | “Decisions are stored separately using stable supplier row IDs. Accepting a valid mapping updates the final catalog and dashboard. Reviewers can select another existing master product or reject a suggestion. Invalid selections and unresolved conflicts cannot become approved output.” |
| 1:30–1:48 | RESEARCH: show R001, evidence note, URL and date; toggle H7 No/Yes. | “Research is recorded at field level, with the exact model, source and supporting evidence. Only eligible approved Specification fields flow into the catalog in this version. Genuine source conflicts remain unresolved instead of being marked verified.” |
| 1:48–2:00 | QA EVALUATION, then DASHBOARD. | “The workspace reconciles accepted mappings, exceptions and duplicate exclusions. Evaluation uses constructed ground truth. Native formulas update immediately; fuzzy scores and CSV exports have explicit refresh steps. The delivered test report distinguishes compatible-engine checks from native Excel and Sheets testing.” |

Practice the actions once before recording. Use a disposable copy so the portfolio baseline remains unchanged. The case study and test report contain measured baseline results; do not describe the simulated scenario as a real client engagement.
