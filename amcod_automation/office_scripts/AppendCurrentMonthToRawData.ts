/**
 * AppendCurrentMonthToRawData.ts (Office Script for Excel)
 *
 * DESIGN B, step 3: takes the current month's rows produced by the Power
 * Query "Debits_CurrentMonth" (loaded to a staging table) and appends them
 * to the bottom of the RawData table AS VALUES, so they become permanent
 * history that the next month's duplicated file carries forward. This is the
 * automated version of Marcia's Stage 3 "scroll to the bottom and paste the
 * new month's block."
 *
 * Why values (not a live query load): RawData mixes formula columns (DSO,
 * TDSO Gap, %) with data columns, and history must survive being duplicated
 * into next month's file without re-reading old Debits files. Writing the
 * month's data as values into the RawData table both preserves it and lets
 * the table's calculated columns (DSO etc.) auto-fill for the new rows.
 *
 * WHAT IT DOES
 *   1. Reads the staging table (STAGING_TABLE).
 *   2. Guards against double-appending the same month (re-run safe).
 *   3. Adds the rows to the RawData table, writing only the mapped data
 *      columns; calculated columns are left for Excel to auto-fill.
 *
 * SETUP
 *   - STAGING_TABLE / RAWDATA_TABLE: set to the actual table names.
 *   - COLUMN_MAP: left = staging table header, right = RawData header. The
 *     RawData header strings must match EXACTLY (including the "€" and
 *     "(Live)" text). Fix any that differ in your workbook - that's the only
 *     edit this script normally needs.
 *   - CALCULATED_COLUMNS: RawData headers that are formulas (DSO, TDSO Gap,
 *     the % columns). The script never writes these; it re-asserts their
 *     formula onto the new rows so they calculate instead of coming through
 *     blank.
 */

const STAGING_TABLE = "Staging_CurrentMonth";
const RAWDATA_TABLE = "RawData";

// staging header  ->  RawData header (edit the right-hand strings to match)
const COLUMN_MAP: { from: string; to: string }[] = [
  { from: "Month", to: "Month" },
  { from: "Country", to: "Country" },
  { from: "Customer", to: "Customer" },
  { from: "Onboard Date", to: "Onboard Date" },
  { from: "Payment (days)", to: "Payment (days)" },
  { from: "Total AR € (Live)", to: "Total AR € (Live)" },
  { from: "Overdue € (Live)", to: "Overdue € (Live)" },
  { from: ">60 days € (Live)", to: ">60 days € (Live)" },
  { from: "Gross Sales (Live)", to: "Gross Sales (Live)" },
  { from: ">90 days (Live)", to: ">90 days (Live)" },
  { from: "Total UAC € (Live)", to: "Total UAC € (Live)" },
  { from: "Total Payments € (Live)", to: "Total Payments € (Live)" },
];

// RawData formula columns - never written, formula re-asserted on new rows.
const CALCULATED_COLUMNS = ["DSO", "TDSO", "TDSO Gap", ">60 days (%)", ">90 days (%)"];

function main(workbook: ExcelScript.Workbook): { appendedRows: number; month: string } {
  const staging = getTable(workbook, STAGING_TABLE);
  const rawData = getTable(workbook, RAWDATA_TABLE);

  const stagingHeaders = staging.getHeaderRowRange().getValues()[0].map((h) => String(h));
  const stagingBody = staging.getRangeBetweenHeaderAndTotal().getValues();
  if (stagingBody.length === 0) {
    console.log("Staging table is empty - nothing to append.");
    return { appendedRows: 0, month: "" };
  }

  const rawHeaders = rawData.getHeaderRowRange().getValues()[0].map((h) => String(h));

  // Resolve the "Month" column so we can dedupe and report.
  const stagingMonthIdx = requireHeader(stagingHeaders, "Month", STAGING_TABLE);
  const rawMonthIdx = requireHeader(rawHeaders, mapTo("Month"), RAWDATA_TABLE);

  // The month being appended (all staging rows share one month).
  const monthValue = stagingBody[0][stagingMonthIdx];
  const monthLabel = String(monthValue);

  // Guard: if RawData already contains this month, do not append again.
  const rawBody = rawData.getRangeBetweenHeaderAndTotal().getValues();
  const alreadyPresent = rawBody.some((r) => sameMonth(r[rawMonthIdx], monthValue));
  if (alreadyPresent) {
    console.log(`Month ${monthLabel} already present in ${RAWDATA_TABLE} - skipping append (re-run safe).`);
    return { appendedRows: 0, month: monthLabel };
  }

  // Build the rows to add, sized to the full RawData width, mapped columns
  // filled and everything else left null for Excel to fill / calc.
  const width = rawHeaders.length;
  const newRows: (string | number | boolean | null)[][] = stagingBody.map((srcRow) => {
    const out: (string | number | boolean | null)[] = new Array(width).fill(null);
    for (const { from, to } of COLUMN_MAP) {
      const sIdx = stagingHeaders.indexOf(from);
      const rIdx = rawHeaders.indexOf(to);
      if (sIdx === -1) {
        throw new Error(`Staging column "${from}" not found in ${STAGING_TABLE}.`);
      }
      if (rIdx === -1) {
        throw new Error(`RawData column "${to}" not found in ${RAWDATA_TABLE}. Fix COLUMN_MAP.`);
      }
      out[rIdx] = srcRow[sIdx] as string | number | boolean | null;
    }
    return out;
  });

  const firstNewRowIndex = rawBody.length; // 0-based within the table body
  rawData.addRows(-1, newRows);

  // Re-assert calculated-column formulas onto the appended rows, in case
  // addRows wrote nulls into them instead of auto-filling.
  reassertFormulas(rawData, rawHeaders, firstNewRowIndex, newRows.length);

  console.log(`Appended ${newRows.length} rows for ${monthLabel} to ${RAWDATA_TABLE}.`);
  return { appendedRows: newRows.length, month: monthLabel };
}

function getTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const t = workbook.getTable(name);
  if (!t) {
    throw new Error(`Table "${name}" not found. Check the table name.`);
  }
  return t;
}

function requireHeader(headers: string[], name: string, tableName: string): number {
  const idx = headers.indexOf(name);
  if (idx === -1) {
    throw new Error(`Column "${name}" not found in ${tableName}.`);
  }
  return idx;
}

function mapTo(stagingHeader: string): string {
  const entry = COLUMN_MAP.find((m) => m.from === stagingHeader);
  return entry ? entry.to : stagingHeader;
}

function sameMonth(a: string | number | boolean, b: string | number | boolean): boolean {
  // Excel dates come through as serial numbers; compare loosely so a date
  // and its string form still match.
  if (typeof a === "number" && typeof b === "number") {
    return Math.round(a) === Math.round(b);
  }
  return String(a) === String(b);
}

/**
 * For each calculated column, copy the formula from an existing row (the row
 * just above the appended block) down onto the new rows, so DSO/TDSO/% fill
 * in rather than staying null.
 */
function reassertFormulas(
  table: ExcelScript.Table,
  rawHeaders: string[],
  firstNewRowIndex: number,
  count: number
) {
  if (firstNewRowIndex === 0) {
    // No prior row to copy a formula from - nothing to re-assert.
    return;
  }
  const bodyRange = table.getRangeBetweenHeaderAndTotal();
  for (const colName of CALCULATED_COLUMNS) {
    const cIdx = rawHeaders.indexOf(colName);
    if (cIdx === -1) {
      continue; // column not in this workbook - skip
    }
    const templateCell = bodyRange.getCell(firstNewRowIndex - 1, cIdx);
    const formula = templateCell.getFormula();
    if (!formula || formula[0] !== "=") {
      continue; // not actually a formula column here
    }
    for (let r = 0; r < count; r++) {
      bodyRange.getCell(firstNewRowIndex + r, cIdx).setFormula(formula);
    }
  }
}
