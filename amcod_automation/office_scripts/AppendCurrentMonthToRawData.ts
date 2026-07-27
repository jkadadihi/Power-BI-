/**
 * AppendCurrentMonthToRawData.ts (Office Script for Excel)
 *
 * Appends the current month's rows - produced by the Power Query
 * "Debits_CurrentMonth" and loaded to the staging table - to the bottom of
 * the RawData table AS VALUES, so they become permanent history that next
 * month's file carries forward. This is the automated version of Marcia's
 * Stage 3 "scroll to the bottom and paste the new month's block."
 *
 * WHY VALUES: RawData mixes formula columns (DSO, TDSO, TDSO Gap, the %
 * columns) with data columns, so Power Query cannot load into it directly -
 * it would own and destroy every column. The query lands in staging; this
 * script carries the rows across and leaves the formula columns alone.
 *
 * HEADER MATCHING: staging and RawData spell some headers differently
 * ("Onboard Date" vs "Onboard Dt", "Gross Sales (Live)" vs
 * "Gross Sales € (Live)", ">60 days" vs "> 60 days"). Exact-text matching
 * silently dropped those columns, so matching here is normalized - spaces,
 * "€" and case are ignored - plus an explicit alias list for headers whose
 * words genuinely differ. Every match and non-match is logged so a future
 * mismatch is visible instead of silent.
 *
 * RE-RUN SAFE: if the month is already in RawData the script skips.
 */

const STAGING_TABLE = "Staging_CurrentMonth";
const RAWDATA_TABLE = "RawData";

// RawData columns that are formulas - never written with data. Their formula
// is copied down from the row above onto the appended rows.
const CALCULATED_COLUMNS = ["TDSO", "DSO", "TDSO Gap", "> 60 days (%)", "> 90 days (%)"];

// Headers whose words differ between the two tables and so cannot be matched
// by normalization alone. Left = staging, right = RawData.
const HEADER_ALIASES: { staging: string; rawData: string }[] = [
  { staging: "Onboard Date", rawData: "Onboard Dt" },
  { staging: "Payment (days)", rawData: "Payment Term (days)" },
];

// Explicit formats, used when RawData has no existing row to copy from.
const DATE_COLUMNS = ["Month"];
const PERCENT_COLUMNS = ["> 60 days (%)", "> 90 days (%)"];
const TEXT_COLUMNS = ["Country", "Customer", "Onboard Dt"];
const INTEGER_COLUMNS = ["Payment Term (days)", "TDSO", "DSO", "TDSO Gap"];
const DATE_FORMAT = "m/d/yyyy";
const PERCENT_FORMAT = "0.0%";
const INTEGER_FORMAT = "0";
const TEXT_FORMAT = "@";
const CURRENCY_FORMAT = "#,##0";

function main(workbook: ExcelScript.Workbook): { appendedRows: number; month: string } {
  const staging = workbook.getTable(STAGING_TABLE);
  if (!staging) {
    throw new Error(`Table "${STAGING_TABLE}" not found.`);
  }
  const rawData = workbook.getTable(RAWDATA_TABLE);
  if (!rawData) {
    throw new Error(`Table "${RAWDATA_TABLE}" not found.`);
  }

  const stagingHeaders = staging.getHeaderRowRange().getValues()[0].map((h) => String(h));
  const stagingBody = staging.getRangeBetweenHeaderAndTotal().getValues();
  if (stagingBody.length === 0) {
    console.log("Staging table is empty - nothing to append.");
    return { appendedRows: 0, month: "" };
  }

  const rawHeaders = rawData.getHeaderRowRange().getValues()[0].map((h) => String(h));

  // staging column index -> RawData column index
  const columnMap = buildColumnMap(stagingHeaders, rawHeaders);

  const stagingMonthIdx = stagingHeaders.indexOf("Month");
  const rawMonthIdx = rawHeaders.indexOf("Month");
  if (stagingMonthIdx === -1 || rawMonthIdx === -1) {
    throw new Error('Column "Month" not found in one of the tables.');
  }

  const monthValue = stagingBody[0][stagingMonthIdx];
  const monthLabel = String(monthValue);

  const rawBody = rawData.getRangeBetweenHeaderAndTotal().getValues();
  // An empty table still reads back as a single blank row - treat that as empty.
  const hasExistingRows =
    rawBody.length > 0 && rawBody.some((r) => String(r[rawMonthIdx]).trim() !== "");

  if (hasExistingRows && rawBody.some((r) => sameMonth(r[rawMonthIdx], monthValue))) {
    console.log(`Month ${monthLabel} already present in ${RAWDATA_TABLE} - skipping (re-run safe).`);
    return { appendedRows: 0, month: monthLabel };
  }

  const width = rawHeaders.length;
  const newRows: (string | number | boolean | null)[][] = stagingBody.map((srcRow) => {
    const out: (string | number | boolean | null)[] = new Array(width).fill(null);
    columnMap.forEach((rIdx, sIdx) => {
      if (rIdx !== -1) {
        out[rIdx] = srcRow[sIdx] as string | number | boolean | null;
      }
    });
    return out;
  });

  const firstNewRowIndex = hasExistingRows ? rawBody.length : 0;
  if (hasExistingRows) {
    rawData.addRows(-1, newRows);
  } else {
    // Reuse the existing blank row, then grow the table for the rest, so no
    // stray empty row is left sitting above the data.
    if (newRows.length > 1) {
      rawData.addRows(-1, newRows.slice(1));
    }
    rawData.getRangeBetweenHeaderAndTotal().setValues(newRows);
  }

  applyFormatsAndFormulas(rawData, rawHeaders, firstNewRowIndex, newRows.length);

  console.log(`Appended ${newRows.length} rows for ${monthLabel} to ${RAWDATA_TABLE}.`);
  return { appendedRows: newRows.length, month: monthLabel };
}

/**
 * Normalizes a header for comparison: lowercase, no whitespace, no "€".
 * Turns "> 60 days € (Live)" and ">60 days (Live)" into the same key.
 */
function normalizeHeader(header: string): string {
  return header.toLowerCase().replace(/\s+/g, "").replace(/€/g, "");
}

/**
 * Maps each staging column index to a RawData column index (-1 if none),
 * using the alias list first, then normalized matching. Logs the result of
 * every column so mismatches are never silent.
 */
function buildColumnMap(stagingHeaders: string[], rawHeaders: string[]): number[] {
  const normalizedRaw = rawHeaders.map(normalizeHeader);

  return stagingHeaders.map((stagingHeader) => {
    const alias = HEADER_ALIASES.find((a) => normalizeHeader(a.staging) === normalizeHeader(stagingHeader));
    const targetKey = alias ? normalizeHeader(alias.rawData) : normalizeHeader(stagingHeader);
    const rIdx = normalizedRaw.indexOf(targetKey);

    if (rIdx === -1) {
      console.log(`WARNING: staging column "${stagingHeader}" has no RawData match - it will be skipped.`);
    } else if (rawHeaders[rIdx] !== stagingHeader) {
      console.log(`Matched "${stagingHeader}" -> "${rawHeaders[rIdx]}"`);
    }
    return rIdx;
  });
}

function sameMonth(a: string | number | boolean, b: string | number | boolean): boolean {
  if (typeof a === "number" && typeof b === "number") {
    return Math.round(a) === Math.round(b);
  }
  return String(a) === String(b);
}

/**
 * Copies number formats (and formulas, for calculated columns) from the row
 * above onto the appended rows. Falls back to explicit formats when there is
 * no row above, so dates don't render as raw serial numbers on a first load.
 */
function applyFormatsAndFormulas(
  table: ExcelScript.Table,
  rawHeaders: string[],
  firstNewRowIndex: number,
  rowCount: number
) {
  const bodyRange = table.getRangeBetweenHeaderAndTotal();
  const templateRowIndex = firstNewRowIndex - 1;
  const missingFormulas: string[] = [];

  for (let c = 0; c < rawHeaders.length; c++) {
    const header = rawHeaders[c];
    const isCalculated = CALCULATED_COLUMNS.indexOf(header) !== -1;

    let numberFormat = "";
    let formula = "";
    if (templateRowIndex >= 0) {
      const templateCell = bodyRange.getCell(templateRowIndex, c);
      numberFormat = templateCell.getNumberFormat();
      if (isCalculated) {
        formula = templateCell.getFormula();
      }
    }
    if (!numberFormat || numberFormat === "General") {
      numberFormat = defaultFormatFor(header);
    }
    if (isCalculated && (!formula || formula[0] !== "=")) {
      missingFormulas.push(header);
    }

    for (let r = 0; r < rowCount; r++) {
      const cell = bodyRange.getCell(firstNewRowIndex + r, c);
      cell.setNumberFormat(numberFormat);
      if (isCalculated && formula && formula[0] === "=") {
        cell.setFormula(formula);
      }
    }
  }

  if (missingFormulas.length > 0) {
    console.log(
      `WARNING: no prior row to copy formulas from, so these stayed blank: ${missingFormulas.join(", ")}. ` +
        `Keep at least one earlier month in RawData, or fill these formulas down once by hand.`
    );
  }
}

function defaultFormatFor(header: string): string {
  if (DATE_COLUMNS.indexOf(header) !== -1) {
    return DATE_FORMAT;
  }
  if (PERCENT_COLUMNS.indexOf(header) !== -1) {
    return PERCENT_FORMAT;
  }
  if (TEXT_COLUMNS.indexOf(header) !== -1) {
    return TEXT_FORMAT;
  }
  if (INTEGER_COLUMNS.indexOf(header) !== -1) {
    return INTEGER_FORMAT;
  }
  return CURRENCY_FORMAT;
}
