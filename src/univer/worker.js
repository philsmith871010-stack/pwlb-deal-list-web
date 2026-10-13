/* The calculation engine, in a worker of its own, so the checks on each row are
   worked out without holding up typing or scrolling. */
import { createUniver, numfmt } from "@univerjs/presets";
import { UniverSheetsCoreWorkerPreset } from "@univerjs/preset-sheets-core/worker";
import { UniverSheetsFilterWorkerPreset } from "@univerjs/preset-sheets-filter/worker";

numfmt.addLocale({ group: ",", decimal: ".", preferMDY: false }, "en");   /* dates day first, as on the page (boot.js) */
createUniver({ presets: [UniverSheetsCoreWorkerPreset(), UniverSheetsFilterWorkerPreset()] });
