# Working on the deal list page

Read README.md first: what this is, how it is built, how to run and test it.

## Where it fits (Phil, 13 October 2026)

A test of keeping the deal list in PWLBtoday. The report tool already reads a
council's Excel deal list and keeps a dated snapshot of its rows, read again
whenever the user asks. The idea is one place for the deals, with the reports,
and later benchmarking and counterparty monitoring, each reading it as and when
("this tooling could be a separate part of PWLBtoday that supports reports,
benchmarking and counterparty monitoring (the latter are way down the line)").
So the workbook engine (src/xlsx, src/univer) is kept apart from the page
(src/main.js, index.html).

Files in and out are not the product: "I didn't think we need export of
upload. Users can just paste data in" (Phil, 13 October 2026). Open and
Save as .xlsx are here only because this test keeps nothing anywhere; in
PWLBtoday they give way to the account keeping the workbook, with users
typing or pasting rows in.

Only the typed-in sheets are needed: "we don't need to replicate the formulas -
we produce the existing system charts in reports with raw data. we just need
the deal recording and projection grid". A first version that recalculated all
102,123 formulas in Univer agreed with Excel on every visible sheet but took
150 seconds; it was dropped for that reason and because it is not needed.

## Rules

- Univer's free (Apache 2.0) packages only. Its Excel import/export, charts
  and print are paid ("Pro"): do not add @univerjs-pro packages.
- The page talks to no server. test/shots.cjs prints any request that leaves
  the page's own origin; it must be empty.
- Saving writes the user's own package back with only the input sheets'
  <sheetData> (and table refs) replaced. Never rebuild the workbook from
  scratch: the report tool and Excel rely on everything else in it.
- Dates are day first. Univer has no UK English; boot.js and worker.js make
  numfmt's "en" day-first. Keep both in step.
- The repository is public: only the made-up example and blank workbooks.
  Never a council's own file.
- Before saying something works, open the page and use it (the tests in
  test/ drive the real page in Chromium).

## Univer, as found (1.0.4)

- Conditional formats: the first rule in Univer's list wins (Excel: priority 1).
- A list rule's formula may be a defined name (`=TermTypes`).
- Date rules take `yyyy-mm-dd`, not Excel serials.
- Links are rich text in the cell (`p` with a customRange), not a resource.
- `createUniver({plugins})` takes a plugin with its config, e.g. the table UI
  with `hideAnchor`, which hides the name tag drawn over the row above a table.
- Univer works a conditional format out for every cell of its range (Excel only
  for cells on screen), so checks.js keeps the ranges to the rows in use plus
  200, growing as rows are typed.
- Pasted text ("7,500,000", "01/10/2026", "4.25%") already becomes numbers
  and dates in Univer (day first once numfmt's "en" is). But it reads "4.10%"
  as 0.040999999999999995: numbers.js rounds every typed or pasted number to
  15 significant digits, as Excel keeps them.
- The facade's getValue()/getCellData() give the formatted text for a number
  ("7,500,000"); tests read stored values from workbook.save().
- The calculation engine runs in a worker (workerURL); without it the page
  froze for up to 1.4s at a time while the checks were worked out.
