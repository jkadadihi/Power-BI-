/**
 * ValidationDashboard.ts (Office Script for Excel)
 *
 * Replaces AMCOD Stage 7 ("DSO Validation"): Marcia manually checks DSO,
 * TDSO, and SPR across multiple tabs to make sure they agree, and
 * investigates by hand when they don't. Priority 4 in the automation plan.
 *
 * Setup (one time): add a sheet named "Validation_Config" to the Monthly
 * Workbook with one row per cross-check, columns:
 *   metric | sheet_a | cell_a | sheet_b | cell_b | tolerance_pct
 * Example row: DSO | DSO Calc | B5 | SPR | C12 | 0.5
 * (meaning: DSO Calc!B5 and SPR!C12 should agree within 0.5%).
 *
 * The script reads every configured pair, compares the values, and writes
 * a pass/fail row to a "Validation Dashboard" sheet (created if missing) so
 * the whole reconciliation is visible at a glance instead of tab-hopping.
 *
 * Run this from Power Automate (Excel: "Run script" action) right after the
 * month's pivots/calculations are refreshed, so Marcia sees a red flag on
 * anything that needs investigating rather than finding it by eye.
 */

interface ValidationResult {
  metric: string;
  sheetA: string;
  cellA: string;
  valueA: number;
  sheetB: string;
  cellB: string;
  valueB: number;
  diffPct: number;
  tolerancePct: number;
  passed: boolean;
}

function main(workbook: ExcelScript.Workbook): ValidationResult[] {
  const configSheet = workbook.getWorksheet("Validation_Config");
  if (!configSheet) {
    throw new Error(
      "Missing 'Validation_Config' sheet. Add it with columns: " +
        "metric | sheet_a | cell_a | sheet_b | cell_b | tolerance_pct"
    );
  }

  const configRange = configSheet.getUsedRange();
  const configValues = configRange.getValues();
  // Row 0 is the header row.
  const rows = configValues.slice(1);

  const results: ValidationResult[] = [];

  for (const row of rows) {
    const [metric, sheetAName, cellA, sheetBName, cellB, toleranceRaw] = row;
    if (!metric) {
      continue;
    }

    const sheetA = workbook.getWorksheet(sheetAName as string);
    const sheetB = workbook.getWorksheet(sheetBName as string);
    if (!sheetA || !sheetB) {
      console.log(`Skipping "${metric}": sheet not found (${sheetAName} / ${sheetBName})`);
      continue;
    }

    const valueA = Number(sheetA.getRange(cellA as string).getValue());
    const valueB = Number(sheetB.getRange(cellB as string).getValue());
    const tolerancePct = Number(toleranceRaw) || 0;

    const denominator = valueA !== 0 ? Math.abs(valueA) : 1;
    const diffPct = (Math.abs(valueA - valueB) / denominator) * 100;

    results.push({
      metric: String(metric),
      sheetA: sheetAName as string,
      cellA: cellA as string,
      valueA,
      sheetB: sheetBName as string,
      cellB: cellB as string,
      valueB,
      diffPct,
      tolerancePct,
      passed: diffPct <= tolerancePct,
    });
  }

  writeDashboard(workbook, results);
  return results;
}

function writeDashboard(workbook: ExcelScript.Workbook, results: ValidationResult[]) {
  let dashboard = workbook.getWorksheet("Validation Dashboard");
  if (!dashboard) {
    dashboard = workbook.addWorksheet("Validation Dashboard");
  } else {
    dashboard.getUsedRange()?.clear(ExcelScript.ClearApplyTo.contents);
  }

  const header = [
    "Metric", "Source A", "Value A", "Source B", "Value B",
    "Diff %", "Tolerance %", "Status",
  ];
  dashboard.getRangeByIndexes(0, 0, 1, header.length).setValues([header]);

  const dataRows = results.map((r) => [
    r.metric,
    `${r.sheetA}!${r.cellA}`,
    r.valueA,
    `${r.sheetB}!${r.cellB}`,
    r.valueB,
    Number(r.diffPct.toFixed(2)),
    r.tolerancePct,
    r.passed ? "OK" : "CHECK",
  ]);

  if (dataRows.length > 0) {
    const body = dashboard.getRangeByIndexes(1, 0, dataRows.length, header.length);
    body.setValues(dataRows);

    const statusColumn = dashboard.getRangeByIndexes(1, header.length - 1, dataRows.length, 1);
    statusColumn.getFormat().getFont().setBold(true);
    results.forEach((r, i) => {
      const cell = dashboard.getRangeByIndexes(1 + i, header.length - 1, 1, 1);
      cell.getFormat().getFill().setColor(r.passed ? "#C6EFCE" : "#FFC7CE");
    });
  }

  dashboard.getRangeByIndexes(0, 0, 1, header.length).getFormat().getFont().setBold(true);
  dashboard.getUsedRange()?.getFormat().autofitColumns();
}
