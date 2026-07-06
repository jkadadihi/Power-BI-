function main(workbook: ExcelScript.Workbook) {
  // Intentionally minimal. Opening this workbook via "Run script" is what
  // triggers "Refresh data when opening the file" on the RawData query.
  // This script exists only to create that open event on a schedule,
  // well ahead of the main report flow, so the Power Query refresh has
  // time to finish before the report script ever reads RawData.
  //
  // Deployed as its own Power Automate flow ("AMKAD - Trigger Daily
  // Refresh"), scheduled to run before the main "AMKAD AR Report" flow.
  // See docs/MAINTENANCE_GUIDE.md, section 4E.
  const sheet = workbook.getActiveWorksheet();
  console.log(`Refresh-trigger ping at ${new Date().toISOString()} on sheet ${sheet.getName()}`);
}
