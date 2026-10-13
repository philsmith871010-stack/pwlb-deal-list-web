/* Univer, with only its free (Apache-2.0) parts: the sheet, number formats,
   conditional formatting, data validation, filters, links and Excel tables.
   The calculation engine runs in a worker (worker.js). Nothing here talks to a
   server. */
import { createUniver, LocaleType, mergeLocales, numfmt } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import sheetsCoreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import { UniverSheetsConditionalFormattingPreset } from "@univerjs/preset-sheets-conditional-formatting";
import cfEnUS from "@univerjs/preset-sheets-conditional-formatting/locales/en-US";
import { UniverSheetsDataValidationPreset } from "@univerjs/preset-sheets-data-validation";
import dvEnUS from "@univerjs/preset-sheets-data-validation/locales/en-US";
import { UniverSheetsFilterPreset } from "@univerjs/preset-sheets-filter";
import filterEnUS from "@univerjs/preset-sheets-filter/locales/en-US";
import { UniverSheetsHyperLinkPreset } from "@univerjs/preset-sheets-hyper-link";
import linkEnUS from "@univerjs/preset-sheets-hyper-link/locales/en-US";
import { UniverSheetsTablePlugin, UniverSheetsTableUIPlugin } from "@univerjs/preset-sheets-table";
import tableEnUS from "@univerjs/preset-sheets-table/locales/en-US";
import "@univerjs/preset-sheets-core/lib/index.css";
import "@univerjs/preset-sheets-conditional-formatting/lib/index.css";
import "@univerjs/preset-sheets-data-validation/lib/index.css";
import "@univerjs/preset-sheets-filter/lib/index.css";
import "@univerjs/preset-sheets-hyper-link/lib/index.css";
import "@univerjs/preset-sheets-table/lib/index.css";

/* DATES THE BRITISH WAY. Univer has no UK English, and its date reader takes
   plain "en" as month first: 01/04/2026 typed into a deal would become 4
   January. The number-format library inside it knows day-first English; its
   "en" is made day-first before anything is read or typed. */
numfmt.addLocale({ group: ",", decimal: ".", preferMDY: false }, "en");

export function bootUniver(container){
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: mergeLocales(sheetsCoreEnUS, cfEnUS, dvEnUS, filterEnUS, linkEnUS, tableEnUS) },
    presets: [
      UniverSheetsCorePreset({ container, workerURL: new Worker(new URL("./worker.js", import.meta.url), { type: "module" }) }),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsDataValidationPreset(),
      UniverSheetsFilterPreset(),
      UniverSheetsHyperLinkPreset(),
    ],
    /* Excel tables, without the name tag Univer draws over the row above each one */
    plugins: [UniverSheetsTablePlugin, [UniverSheetsTableUIPlugin, { hideAnchor: true }]],
  });
  return { univer, univerAPI };
}
