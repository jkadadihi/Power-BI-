/*
 * Runs on the SOURCE file (the daily "AMKAD_Daily_Report" file that lands in the
 * Cust_Sol Daily File folder). Reads every data row, returns it as JSON so Power
 * Automate can hand it to custsol_append_rawdata_eur.ts without needing two
 * workbooks open in the same script (Office Scripts can't cross-open files).
 *
 * Column lookup is done by matching header text, not fixed letters, so this
 * keeps working if the source report ever reorders its columns.
 */
function main(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheets()[0];
  const usedRange = sheet.getUsedRange();
  const values = usedRange.getValues();

  if (values.length < 2) {
    return JSON.stringify({ rows: [], totalRows: 0, missingEuroRows: 0 });
  }

  const headerRow = values[0].map(h => String(h).trim());

  function col(name: string): number {
    const idx = headerRow.indexOf(name);
    if (idx === -1) throw new Error(`Expected column "${name}" not found in source report header row.`);
    return idx;
  }

  const idx = {
    reportDate: col("Report Date"),
    countryCode: col("countrycode"),
    customerName: col("Customer Reporting Name"),
    goLiveDate: col("Go Live Date"),
    paymentTerms: col("Payment Terms"),
    totalPayments: col("Total Payments"),
    totalUAC: col("Total UAC"),
    ttlar: col("TTLAR"),
    overdue: col("Overdue"),
    gt60days: col("GT60days"),
    grossSales: col("Gross Sales"),
    gt90Days: col("GT90Days"),
    gt60DaysEuro: col("GT60Days Euro"),
    grossSalesEuro: col("Gross Sales Euro"),
    totalAREuro: col("Total AREuro"),
    overdueAREuro: col("Overdue AREuro"),
    gt90DaysEuro: col("GT90Days Euro"),
    totalUACEuro: col("Total UAC Euro"),
    totalPaymentsEuro: col("Total Payments") // last column, same header text as source col G
  };

  // Source has two columns both literally headed "Total Payments" (native at G,
  // Euro at the very last column V). indexOf() above would only ever find the
  // first one, so find the SECOND occurrence explicitly for the Euro version.
  const lastTotalPaymentsIdx = headerRow.lastIndexOf("Total Payments");
  idx.totalPaymentsEuro = lastTotalPaymentsIdx;

  const rows: Record<string, unknown>[] = [];
  let missingEuroRows = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!row[idx.customerName]) continue; // skip blank trailing rows

    const euroFieldsPresent =
      row[idx.gt60DaysEuro] !== "" && row[idx.gt60DaysEuro] !== null &&
      row[idx.totalAREuro] !== "" && row[idx.totalAREuro] !== null;

    if (!euroFieldsPresent) missingEuroRows++;

    rows.push({
      reportDate: row[idx.reportDate],
      countryCode: row[idx.countryCode],
      customerName: row[idx.customerName],
      goLiveDate: row[idx.goLiveDate],
      paymentTerms: row[idx.paymentTerms],
      totalPayments: row[idx.totalPayments],
      totalUAC: row[idx.totalUAC],
      ttlar: row[idx.ttlar],
      overdue: row[idx.overdue],
      gt60days: row[idx.gt60days],
      grossSales: row[idx.grossSales],
      gt90Days: row[idx.gt90Days],
      gt60DaysEuro: euroFieldsPresent ? row[idx.gt60DaysEuro] : null,
      grossSalesEuro: euroFieldsPresent ? row[idx.grossSalesEuro] : null,
      totalAREuro: euroFieldsPresent ? row[idx.totalAREuro] : null,
      overdueAREuro: euroFieldsPresent ? row[idx.overdueAREuro] : null,
      gt90DaysEuro: euroFieldsPresent ? row[idx.gt90DaysEuro] : null,
      totalUACEuro: euroFieldsPresent ? row[idx.totalUACEuro] : null,
      totalPaymentsEuro: euroFieldsPresent ? row[idx.totalPaymentsEuro] : null
    });
  }

  return JSON.stringify({
    rows: rows,
    totalRows: rows.length,
    missingEuroRows: missingEuroRows
  });
}
