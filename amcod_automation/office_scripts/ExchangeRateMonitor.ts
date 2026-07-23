/**
 * ExchangeRateMonitor.ts (Office Script for Excel)
 *
 * Addresses the exchange-rate maintenance risk called out separately from
 * the main AMCOD stages: if exchange rates aren't loaded for the month,
 * daily reporting files go blank downstream, and the fix (recalculating
 * rates by hand after the fact) is worse than the ten minutes it takes to
 * load them on time. Priority 5 in the automation plan.
 *
 * This script does NOT load rates - it only checks whether the current
 * month's rates are present and non-blank, so a missed upload is caught
 * *before* the deadline instead of discovered in broken downstream reports.
 *
 * Setup: point RATE_SHEET_NAME / RATE_TABLE_NAME at the sheet+table where
 * monthly exchange rates are maintained. Expected table columns:
 *   Currency | <one column per month, e.g. "Jun-26", "Jul-26">
 * The script checks the column matching the current (or a supplied) month
 * and flags any currency with a blank/zero/non-numeric rate.
 *
 * Run this from Power Automate on a recurring schedule in the days leading
 * up to the exchange-rate deadline (e.g. daily from the 20th of the month),
 * and have the flow message Marcia on Teams/email only when
 * missingCurrencies.length > 0 - no news is no message.
 */

interface RateCheckResult {
  monthColumn: string;
  totalCurrencies: number;
  missingCurrencies: string[];
  isComplete: boolean;
}

const RATE_SHEET_NAME = "Exchange Rates";
const RATE_TABLE_NAME = "ExchangeRateTable";

function main(workbook: ExcelScript.Workbook, monthColumnOverride?: string): RateCheckResult {
  const sheet = workbook.getWorksheet(RATE_SHEET_NAME);
  if (!sheet) {
    throw new Error(`Missing "${RATE_SHEET_NAME}" sheet.`);
  }

  const table = sheet.getTables().find((t) => t.getName() === RATE_TABLE_NAME) ?? sheet.getTables()[0];
  if (!table) {
    throw new Error(`No table found on "${RATE_SHEET_NAME}" (expected "${RATE_TABLE_NAME}").`);
  }

  const headerRange = table.getHeaderRowRange();
  const headers = headerRange.getValues()[0].map((h) => String(h));

  const monthColumn = monthColumnOverride ?? currentMonthLabel();
  const columnIndex = headers.indexOf(monthColumn);
  if (columnIndex === -1) {
    throw new Error(
      `Column "${monthColumn}" not found in ${RATE_TABLE_NAME}. ` +
        `Add a column for this month before running the check.`
    );
  }

  const currencyColumnIndex = headers.indexOf("Currency");
  if (currencyColumnIndex === -1) {
    throw new Error(`Expected a "Currency" column in ${RATE_TABLE_NAME}.`);
  }

  const bodyRange = table.getRangeBetweenHeaderAndTotal();
  const values = bodyRange.getValues();

  const missingCurrencies: string[] = [];
  for (const row of values) {
    const currency = String(row[currencyColumnIndex]);
    const rate = row[columnIndex];
    const isBlank = rate === "" || rate === null || rate === undefined;
    const isNonNumeric = typeof rate !== "number";
    const isZero = typeof rate === "number" && rate === 0;
    if (isBlank || isNonNumeric || isZero) {
      missingCurrencies.push(currency);
    }
  }

  return {
    monthColumn,
    totalCurrencies: values.length,
    missingCurrencies,
    isComplete: missingCurrencies.length === 0,
  };
}

function currentMonthLabel(): string {
  // Mirrors the "MMM-yy" convention used elsewhere in this repo's Power
  // Query scripts (e.g. "Jul-26"). Office Scripts can't take a live
  // system clock as a hard dependency in every environment, so prefer
  // passing monthColumnOverride explicitly from the Power Automate flow
  // (see power_automate/exchange_rate_monitor_flow.md) when precision
  // matters; this fallback covers manual runs.
  const now = new Date();
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const yy = String(now.getFullYear()).slice(-2);
  return `${months[now.getMonth()]}-${yy}`;
}
