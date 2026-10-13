/* WHICH PART OF THE WORKBOOK THIS PAGE IS FOR.
 *
 * The deal list workbook (the PWLB report tool's template) has typed-in sheets
 * and worked-out ones. This page is for the typing: the deals and the
 * projection grid. The calculations, dashboards and charts are the report
 * tool's own (Phil, 13 October 2026: "we don't need to replicate the formulas
 * - we produce the existing system charts in reports with raw data. we just
 * need the deal recording and projection grid"). The other sheets stay in the
 * file untouched, so it still opens in Excel as it always did.
 */
export const INPUT_SHEETS = ["Settings", "Term investments", "Call accounts and MMFs", "Loans", "Pooled funds", "Projections"];

/* Read but kept out of sight: the dropdown words the input sheets' lists name. */
export const SUPPORT_SHEETS = ["Lists"];

/* Opens on this sheet when the file's own active sheet is not one of the above. */
export const FIRST_SHEET = "Term investments";
