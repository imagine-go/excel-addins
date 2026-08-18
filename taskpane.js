/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */


/**
 * Excel Add-in (Office.js)
 * -------------------------------------------------------------
 * 1) Creates a new worksheet in the same workbook.
 * 2) Copies column A of the source sheet into column A of the new sheet.
 * 3) Writes H + I of the source sheet into column B of the new sheet.
 *
 * Entry point: copyAndCombine()
 */

const CONFIG = {
  // Base name for the new sheet. A suffix is added if the name is taken.
  newSheetBaseName: "Summary",

  // Base name for the hidden worksheet that holds the scatter charts' (X, Y)
  // source data. See the "hidden chart-data worksheet" comment in
  // copyAndCombine() for why this lives on its own sheet.
  chartDataSheetBaseName: "Summary Chart Data",

  // Width, in points, of column C on the Summary sheet. autofitColumns()
  // sizes columns to their text content, which would leave column C
  // (empty except for the merged chart cells) too narrow to show a
  // scatter chart, so it gets a fixed width instead.
  chartColumnWidth: 320,

  // If true, row 1 is treated as a header: A1 is copied as-is and
  // B1 gets `headerForB` instead of a computed value.
  hasHeaderRow: true,
  headerForB: "H + I",

  // "sum"    -> numeric addition (blanks count as 0)
  // "concat" -> text concatenation joined by `separator`
  combineMode: "sum",
  separator: " ",

  // If true, column B gets live formulas (=Sheet1!H2+Sheet1!I2) instead of
  // static values, so it updates when the source changes.
  useFormulas: true,
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "This is a test of something";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = copyAndCombine;
  }
});

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const sheet = context.workbook.worksheets.getActiveWorksheet();

//       const range = sheet.getRange("D1");

//       range.values = [["hello from Javascript "]];

//       range.format.fill.color = "yellow";
//       range.format.font.bold = true;

//       await context.sync();
//       console.log("set cell value");
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }



async function copyAndCombine() {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const source = sheets.getActiveWorksheet();
      source.load("name");
      sheets.load("items/name");

      const used = source.getUsedRange();
      used.load(["rowIndex", "rowCount"]);
      await context.sync();

      // Total rows from row 1 through the last used row.
      const rowCount = used.rowIndex + used.rowCount;
      if (rowCount < 1) {
        throw new Error("The source sheet appears to be empty.");
      }

      // Column indexes are 0-based: A = 0, H = 7, I = 8.
      const colA = source.getRangeByIndexes(0, 0, rowCount, 1);
      const colHI = source.getRangeByIndexes(0, 7, rowCount, 2);
      colA.load("values");
      colHI.load("values");
      await context.sync();

      const aValues = colA.values;
      const hiValues = colHI.values;
      const existingNames = sheets.items.map((s) => s.name);
      const sheetName = uniqueSheetName(CONFIG.newSheetBaseName, existingNames);

      const output = [];
      for (let r = 0; r < rowCount; r++) {
        const isHeader = CONFIG.hasHeaderRow && r === 0;
        let b;
        const bNumeric = isHeader ? null : combine(hiValues[r][0], hiValues[r][1]);

        if (isHeader) {
          b = CONFIG.headerForB;
        } else if (CONFIG.useFormulas) {
          const excelRow = r + 1; // 0-based index -> 1-based row
          b = `='${escapeSheetRef(source.name)}'!H${excelRow}+'${escapeSheetRef(
            source.name
          )}'!I${excelRow}`;
        } else {
          b = bNumeric;
        }

        // bNumeric is always the plain computed H+I number, even when `b`
        // itself is a formula string; it feeds the scatter chart below.
        output.push({ a: aValues[r][0], b, bNumeric });
      }

      // Sort the whole table by column A: text before the last '_', then
      // the numeric value after it. The header row (if any) stays in place.
      const header = CONFIG.hasHeaderRow ? output[0] : null;
      const dataRows = CONFIG.hasHeaderRow ? output.slice(1) : output;

      dataRows.forEach((row) => {
        const key = splitSortKey(row.a);
        row.text = key.text;
        row.num = key.num;
      });

      dataRows.sort((row1, row2) => {
        const textCompare = row1.text.localeCompare(row2.text);
        if (textCompare !== 0) return textCompare;
        return row1.num - row2.num;
      });

      const sortedOutput = header
        ? [[header.a, header.b], ...dataRows.map((row) => [row.a, row.b])]
        : dataRows.map((row) => [row.a, row.b]);

      // Create the new sheet and write both columns in a single range set.
      const target = sheets.add(sheetName);
      const destination = target.getRangeByIndexes(0, 0, rowCount, 2);

      if (CONFIG.useFormulas) {
        destination.formulas = sortedOutput;
      } else {
        destination.values = sortedOutput;
      }

      // Column C is left blank; each text-group's cells in it get merged
      // into one below, and hold that group's scatter chart.
      const columnCOutput = header
        ? [["Chart"], ...Array(dataRows.length).fill([""])]
        : Array(rowCount).fill([""]);
      target.getRangeByIndexes(0, 2, rowCount, 1).values = columnCOutput;

      if (CONFIG.hasHeaderRow) {
        target.getRangeByIndexes(0, 0, 1, 3).format.font.bold = true;
      }
      // Autofit A:B to their text content, but give column C a fixed width
      // instead of autofitting it — it holds charts, not text, so sizing
      // it off content would leave it too narrow to display them.
      target.getRangeByIndexes(0, 0, rowCount, 2).format.autofitColumns();
      target.getRangeByIndexes(0, 2, rowCount, 1).format.columnWidth = CONFIG.chartColumnWidth;
      target.activate();

      // --- Hidden chart-data worksheet -------------------------------
      // Excel.Chart needs a real Range as its data source (there's no API
      // to hand a chart raw JS arrays), so the (x, y) pairs each group's
      // scatter chart is built from — x = numeric part of column A, y =
      // the row's column B value — have to live in actual cells somewhere.
      // Rather than parking them in extra columns on the visible "Summary"
      // sheet (which would need to be hidden AND would require every
      // chart's `plotVisibleOnly` set to false, since Excel excludes
      // hidden cells from charts by default), they get their own worksheet
      // that is created and immediately hidden. It has no UI of its own;
      // it only exists so the charts on the Summary sheet have something
      // to point at. Do not delete it — deleting a chart's source range
      // breaks the chart (its series turn into #REF! errors).
      const chartDataSheetName = uniqueSheetName(CONFIG.chartDataSheetBaseName, existingNames.concat([sheetName]));
      const chartDataSheet = sheets.add(chartDataSheetName);
      const chartDataRows = dataRows.map((row) => [row.num, row.bNumeric]);
      chartDataSheet.getRangeByIndexes(0, 0, chartDataRows.length, 2).values = chartDataRows;
      chartDataSheet.visibility = Excel.SheetVisibility.hidden;

      // For each contiguous text-group (rows are already sorted by text),
      // merge its column C cells into one and drop an XY-scatter chart of
      // Y (column B) vs X (numeric part of column A) into that space. The
      // chart's source data comes from the hidden sheet above, at the same
      // group row range but without the header offset (that sheet has no
      // header row of its own).
      const headerOffset = CONFIG.hasHeaderRow ? 1 : 0;
      let groupStart = 0;
      for (let i = 1; i <= dataRows.length; i++) {
        const atBoundary = i === dataRows.length || dataRows[i].text !== dataRows[groupStart].text;
        if (!atBoundary) continue;

        const groupEnd = i - 1; // inclusive
        const rowSpan = groupEnd - groupStart + 1;

        const chartRange = target.getRangeByIndexes(groupStart + headerOffset, 2, rowSpan, 1);
        if (rowSpan > 1) {
          chartRange.merge();
        }

        const dataRange = chartDataSheet.getRangeByIndexes(groupStart, 0, rowSpan, 2);
        const chart = target.charts.add(Excel.ChartType.xyscatter, dataRange, Excel.ChartSeriesBy.columns);
        chart.setPosition(chartRange, chartRange);

        groupStart = i;
      }

      await context.sync();
      console.log(`Wrote ${rowCount} rows to "${sheetName}".`);
    });
  } catch (error) {
    console.error(error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", JSON.stringify(error.debugInfo));
    }
  }
}

/** Combines two cell values according to CONFIG.combineMode. */
function combine(h, i) {
  const hEmpty = h === null || h === "";
  const iEmpty = i === null || i === "";
  if (hEmpty && iEmpty) return "";

  if (CONFIG.combineMode === "concat") {
    if (hEmpty) return String(i);
    if (iEmpty) return String(h);
    return String(h) + CONFIG.separator + String(i);
  }

  // "sum": coerce to numbers, treat blanks and non-numerics as 0.
  return toNumber(h) + toNumber(i);
}

function toNumber(v) {
  if (typeof v === "number") return v;
  if (v === null || v === "" || typeof v === "boolean") return 0;
  const n = parseFloat(String(v).replace(/[$,\s]/g, ""));
  return isNaN(n) ? 0 : n;
}

/**
 * Splits a column-A cell on its LAST underscore (there may be several).
 * "foo_bar_12" -> { text: "foo_bar", num: 12 }
 * If the part after the last '_' isn't a valid number (or there's no '_'
 * at all), the whole cell is kept as-is in `text` with num 0.
 */
function splitSortKey(value) {
  const str = String(value);
  const idx = str.lastIndexOf("_");
  if (idx !== -1) {
    const suffix = str.slice(idx + 1);
    const num = Number(suffix);
    if (suffix !== "" && !isNaN(num)) {
      return { text: str.slice(0, idx), num };
    }
  }
  return { text: str, num: 0 };
}

/** Excel sheet names: max 31 chars, no : \ / ? * [ ] characters. */
function uniqueSheetName(base, existing) {
  const clean = base.replace(/[:\\/?*[\]]/g, "").slice(0, 31) || "Sheet";
  const taken = new Set(existing.map((n) => n.toLowerCase()));
  if (!taken.has(clean.toLowerCase())) return clean;

  for (let n = 2; n < 1000; n++) {
    const suffix = ` (${n})`;
    const candidate = clean.slice(0, 31 - suffix.length) + suffix;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
  throw new Error("Could not generate a unique sheet name.");
}

/** Single quotes inside a sheet name must be doubled in a formula reference. */
function escapeSheetRef(name) {
  return name.replace(/'/g, "''");
}

// Wire up to a button in your task pane (id="run").
Office.onReady(() => {
  const button = document.getElementById("run");
  if (button) button.onclick = copyAndCombine;
});
