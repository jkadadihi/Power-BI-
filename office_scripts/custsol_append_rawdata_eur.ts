/*
 * Runs on the DAILY REPORTING WORKBOOK. Takes the JSON produced by
 * custsol_read_daily.ts (run against that day's source file) and appends new
 * rows to the "Raw Data EUR" table.
 *
 * NOTE: "Raw Data EUR" is expected to be a real Excel Table (not just a
 * formatted range). Appending via table.addRows() makes any formula columns
 * in F:P auto-fill down to the new rows the same way they would if you
 * dragged a formula down manually — so this script never needs to know
 * what those formulas are.
 *
 * Table name confirmed as "RawData" (Table Design > Table Name).
 */
const RAW_DATA_EUR_TABLE_NAME = "RawData";

function main(workbook: ExcelScript.Workbook, sourceJson: string): string {
  const parsed = JSON.parse(sourceJson) as {
    rows: Record<string, number | string | null>[];
    totalRows: number;
    missingEuroRows: number;
  };

  const table = workbook.getTable(RAW_DATA_EUR_TABLE_NAME);
  if (!table) {
    throw new Error(`Table "${RAW_DATA_EUR_TABLE_NAME}" not found in this workbook. Check the name in Table Design.`);
  }

  const headerNames = table.getHeaderRowRange().getValues()[0].map(h => String(h).trim());

  function colIndex(name: string): number {
    const idx = headerNames.indexOf(name);
    if (idx === -1) throw new Error(`Expected column "${name}" not found in Raw Data EUR table.`);
    return idx;
  }

  const destIdx = {
    date: colIndex("Month"),
    country: colIndex("Country"),
    customer: colIndex("Customer"),
    goLive: colIndex("Go Live"),
    paymentTerm: colIndex("Payment Term (days)"),
    totalAREuroLive: colIndex("Total AR € (Live)"),
    grossSalesEuroLive: colIndex("Gross Sales € (Live)"),
    overdueEuroLive: colIndex("Overdue € (Live)"),
    gt60EuroLive: colIndex("> 60 days € (Live)"),
    gt90EuroLive: colIndex("> 90 days € (Live)"),
    totalAR: colIndex("Total AR"),
    overdue: colIndex("Overdue"),
    gt60days: colIndex("> 60 days"),
    grossSales: colIndex("Gross Sales"),
    gt90days: colIndex("> 90 days"),
    totalUACEuroLive: colIndex("Total UAC € (Live)"),
    totalUAC: colIndex("Total UAC"),
    totalPaymentsEuroLive: colIndex("Total Payments € (Live)"),
    totalPayments: colIndex("Total Payments")
  };

  // Duplicate guard: if this script ever runs twice against the same file
  // (retry, manual re-run), don't add the same Date+Country+Customer combo
  // twice. Build a lookup of what's already in the table before appending.
  const existingValues = table.getRangeBetweenHeaderAndTotal().getValues();
  const existingKeys = new Set<string>(
    existingValues.map(row => dedupeKey(row[destIdx.date], row[destIdx.country], row[destIdx.customer]))
  );

  const columnCount = headerNames.length;
  const newRows: (number | string | null)[][] = [];
  let skippedDuplicates = 0;

  parsed.rows.forEach(r => {
    const key = dedupeKey(r.reportDate, r.countryCode, r.customerName);
    if (existingKeys.has(key)) {
      skippedDuplicates++;
      return;
    }
    existingKeys.add(key); // guard against duplicates within the same incoming batch too

    const rowValues: (number | string | null)[] = new Array(columnCount).fill(null);
    rowValues[destIdx.date] = r.reportDate as string;
    rowValues[destIdx.country] = r.countryCode as string;
    rowValues[destIdx.customer] = r.customerName as string;
    rowValues[destIdx.goLive] = r.goLiveDate as string;
    rowValues[destIdx.paymentTerm] = r.paymentTerms as number;

    rowValues[destIdx.totalAREuroLive] = round0(r.totalAREuro);
    rowValues[destIdx.grossSalesEuroLive] = round0(r.grossSalesEuro);
    rowValues[destIdx.overdueEuroLive] = round0(r.overdueAREuro);
    rowValues[destIdx.gt60EuroLive] = round0(r.gt60DaysEuro);
    rowValues[destIdx.gt90EuroLive] = round0(r.gt90DaysEuro);

    rowValues[destIdx.totalAR] = round0(r.ttlar);
    rowValues[destIdx.overdue] = round0(r.overdue);
    rowValues[destIdx.gt60days] = round0(r.gt60days);
    rowValues[destIdx.grossSales] = round0(r.grossSales);
    rowValues[destIdx.gt90days] = round0(r.gt90Days);

    rowValues[destIdx.totalUACEuroLive] = round0(r.totalUACEuro);
    rowValues[destIdx.totalUAC] = round0(r.totalUAC);
    rowValues[destIdx.totalPaymentsEuroLive] = round0(r.totalPaymentsEuro);
    rowValues[destIdx.totalPayments] = round0(r.totalPayments);

    newRows.push(rowValues);
  });

  // Diagnostic: surface the exact dimensions so a mismatch is unambiguous
  // instead of the opaque "doesn't match the size" Office JS error.
  const firstRowWidth = newRows.length > 0 ? newRows[0].length : -1;
  if (newRows.length > 0 && firstRowWidth !== columnCount) {
    throw new Error(
      `Row width ${firstRowWidth} != table column count ${columnCount}. ` +
      `Headers seen (${columnCount}): [ ${headerNames.join(" | ")} ]`
    );
  }

  // addRows appends below existing data and extends any formula columns
  // (F:P) down automatically, same as Excel does when you fill a table down.
  // Guard against an empty array: addRows([]) throws a dimension-mismatch
  // error, so if every incoming row was skipped as a duplicate there is
  // simply nothing to add — that is a valid no-op, not a failure.
  if (newRows.length > 0) {
    table.addRows(-1, newRows);
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  }

  return JSON.stringify({
    rowsAppended: newRows.length,
    skippedDuplicates: skippedDuplicates,
    missingEuroRows: parsed.missingEuroRows,
    needsReview: parsed.missingEuroRows > 0
  });
}

function round0(value: number | string | null): number | null {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return isFinite(n) ? Math.round(n) : null;
}

// Same date + country + customer = the same source row, regardless of which
// run produced it. Values are normalized (trimmed, lowercased) so formatting
// differences (e.g. a date read as a string vs. a serial number) don't cause
// false negatives.
function dedupeKey(date: unknown, country: unknown, customer: unknown): string {
  return [date, country, customer]
    .map(v => String(v ?? "").trim().toLowerCase())
    .join("|");
}
