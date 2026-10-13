# Deal list in the browser (test project)

The PWLB report tool's deal list workbook, opened in a web page that looks and
behaves like Excel: the deals and the projection grid can be typed in, and the
workbook is saved back as the same `.xlsx` the report tool reads.

It is a test of one idea: that the deal list could live in PWLBtoday rather than
on a council's own computer, with the reports (and, later, benchmarking and
counterparty monitoring) each reading it when they need it.

## What it does

- Opens a deal list workbook (`.xlsx`): the example, the blank template, or the
  user's own.
- Shows the typed-in sheets as Excel does: Settings, Term investments, Call
  accounts and MMFs, Loans, Pooled funds and Projections, with the workbook's
  own fonts, colours, borders, number formats, column widths, merged cells,
  frozen headers, tab colours and hidden gridlines.
- Keeps the workbook's rules: dropdown lists (from the hidden Lists sheet),
  the date, amount and rate rules with their prompts, the conditional formats
  that flag missing or doubtful entries, the Excel tables with their filter
  buttons, and the protection on Settings (only the As at date, the pooling
  and the period start can be typed over; the sheets cannot be added to,
  renamed or removed).
- Reads typing the British way: `01/04/2026` is 1 April, `4.1%` is 0.041,
  `5,000,000` is a number.
- Saves the workbook as it came, with only those sheets' cells replaced and
  each table stretched over any rows typed beneath it. The calculated sheets
  (Dashboard, Position, Benchmark and the hidden maths) are left as they are;
  Excel works them out again when it opens the file.

It does not show the calculated sheets or the charts: the report tool works
those out from the raw rows itself.

## How it is built

- [Univer](https://github.com/dream-num/univer) 1.0, open source (Apache 2.0),
  only its free parts. Univer's own Excel import and export are paid, so this
  project reads and writes the `.xlsx` itself (`src/xlsx/`).
- No server: the page runs entirely in the browser and makes no requests
  beyond its own files.

```
src/
  config.js          which sheets are shown, and which are read but hidden
  main.js            the page: open, save, the sample workbooks
scripts/artifact.mjs the build, made publishable on claude.ai
  xlsx/zip.js        reading and writing the zip package
  xlsx/read.js       the .xlsx as a plain model: cells, styles, rules, tables, names
  xlsx/write.js      saving: the original package with the typed-in sheets replaced
  xlsx/ref.js        A1 references
  univer/boot.js     Univer with its free presets; dates day-first
  univer/worker.js   the calculation engine, off the main thread
  univer/build.js    the model as Univer's workbook data
  univer/protect.js  Excel's sheet protection: locked and unlocked cells
  univer/checks.js   the checking rules follow the rows typed
public/              the example and blank workbooks (made-up figures only)
test/                browser tests (Playwright)
```

## Running it

```
npm install
npm run build
npx vite preview --port 8302      # or serve dist/ any other way
```

To publish it as a page on claude.ai, `node scripts/artifact.mjs <outdir>` turns
`dist/` into the page and the files beside it (the sample workbooks go as
base64 text, which the page falls back to, since claude.ai serves no `.xlsx`);
saving there goes through the viewer's `downloads` capability, which asks the
person to confirm.

Tests, with the page served on 8302 (the round trip also needs the report
tool's page, from the pwlbtoday-tms repository, served on 8201):

```
node test/typing.cjs      # UK dates, rates, amounts; protected cells
node test/roundtrip.cjs   # saved, the report tool reads it as before, edits and all
node test/shots.cjs       # a screenshot of each sheet
node test/longtasks.cjs   # how long the page is busy while opening
```

## Known limits

- Formatting changed in the page (bold, a fill) is not saved: each cell keeps
  the workbook's own style for its place.
- Rows inserted or deleted in the page are saved as the rows now stand; the
  file's dropdown and checking rules keep their original ranges (to row 2,000).
- Other spreadsheet programs that do not recalculate on opening (LibreOffice by
  default) show the calculated sheets' old values until recalculated.
- Typing at machine speed into the first cell edited can drop a key; people
  typing (tested down to 20ms between keys) are unaffected.

Never commit a real council's workbook here: this repository is public.
