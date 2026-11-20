# Google Sheets conversion and testing guide

No native Google Sheets version was created or tested. The Excel workbook is the delivered primary version.

1. Upload `ProductMatch.xlsx` to Google Drive, open with Google Sheets and choose **File → Save as Google Sheets**. Keep an unchanged `.xlsx` backup.
2. Confirm all 17 tabs imported. Excel tables may become ordinary filtered ranges. Reapply **Data → Create a filter** over each data region (headers row 6; reserved supplier rows through 356, master rows through 306, research through 106).
3. Check File → Settings → Calculation and ensure changes recalculate. Use the English function names from the workbook. Locale should preserve formula syntax; imported prices are numeric IDR values, identifiers must remain Plain text.
4. Repair any imported navigation links and validations manually, or paste [ProductMatch_Sheets.gs](../scripts/ProductMatch_Sheets.gs) into Extensions → Apps Script and run `repairProductMatch`. It uses only the current spreadsheet; authorize only when you choose to run it. This optional script has not been tested against a live Google Sheets document.
5. Check DASHBOARD B5/D5/F5 dropdowns and H5 search. Choose Storage and verify that both charts and filtered metrics change, while the four overall cards remain stable. Restore All and clear H5.
6. If charts fail to import, recreate a bar chart from DASHBOARD L6:M11 and another from O6:P15. Place them in the reserved chart areas. These cells remain formula-driven.
7. Test S0170 in REVIEW DECISIONS: Accept suggested match + column C = M0170 + reviewer + date should enter final output; Reject suggestion should remove it; Leave pending should retain the exception. Test S0181's explicit correct variant selection and S0191's incompatible variant. Restore baseline values after testing.
8. Toggle RESEARCH R001's Apply approved H7 to No: final specification should become blank for M0001; restore Yes. Verification statuses alone do not mean an attribute is applied.
9. Append a record within capacity and run the formula/duplicate/decision checks in the [historical test report](testing/2026-11-11-report.md). Check QA EVALUATION reconciliation totals.
10. Inspect for `#REF!`, `#VALUE!`, `#NAME?`, `#DIV/0!`, `#N/A`, broken validation and missing conditional formatting. Compare baseline metrics with [baseline_metrics.json](testing/baseline_metrics.json).

The Python fuzzy refresh patches `.xlsx` files, not a live Google Sheet. To refresh a converted copy: download an Excel copy, save/recalculate it in a compatible engine, run the Python snapshot refresh, recalculate again, and import the refreshed workbook as a new Google Sheet. Copying only CANDIDATES A7:I356 and L7:M306 values into the original Sheet preserves its manual inputs, but compare signatures and outcomes afterward. Native Sheets changes are not synchronized back to the original Excel file automatically.

Keep a dated record of the checks performed. Do not claim identical cross-platform behavior until these native Sheets tests pass. This guide deliberately distinguishes an import workflow from a tested native deliverable.
