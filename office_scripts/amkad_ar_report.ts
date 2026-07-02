function main(workbook: ExcelScript.Workbook) {
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  type Agg = {
    totalAR: number;
    overdue: number;
    gt60: number;
    gt90: number;
    grossSales: number;
    uac: number;
    payments: number;
  };

  type CustomerCountryAgg = {
    country: string;
    customer: string;
    months: { [month: string]: Agg };
  };

  type FieldMapItem = {
    sourceColumn: string;
    required: boolean;
    defaultValue: number;
  };

  type FieldMap = {
    [standardField: string]: FieldMapItem;
  };

  type ConfigMap = {
    [key: string]: string;
  };

  const sourceTableName = "RawData";

  function blankAgg(): Agg {
    return {
      totalAR: 0,
      overdue: 0,
      gt60: 0,
      gt90: 0,
      grossSales: 0,
      uac: 0,
      payments: 0
    };
  }

  function addAgg(target: Agg, source: Agg) {
    target.totalAR += source.totalAR;
    target.overdue += source.overdue;
    target.gt60 += source.gt60;
    target.gt90 += source.gt90;
    target.grossSales += source.grossSales;
    target.uac += source.uac;
    target.payments += source.payments;
  }

  function num(value: string | number | boolean | null | undefined): number {
    if (value === null || value === undefined) return 0;
    if (typeof value === "number") return value;

    const cleaned = value
      .toString()
      .replace(/€/g, "")
      .replace(/\$/g, "")
      .replace(/,/g, "")
      .trim();

    const parsed = Number(cleaned);
    return isFinite(parsed) ? parsed : 0;
  }

  function text(value: string | number | boolean | null | undefined): string {
    if (value === null || value === undefined) return "";
    return value.toString().trim();
  }

  function boolFromText(value: string | number | boolean | null | undefined): boolean {
    if (value === null || value === undefined) return false;
    const v = value.toString().trim().toLowerCase();
    return v === "true" || v === "yes" || v === "1";
  }

  function escapeHtml(value: string): string {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function normalizeMonth(value: string | number | boolean): string {
    if (typeof value === "number") {
      if (value <= 0) return "";
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const date = new Date(excelEpoch.getTime() + value * 86400000);
      return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
    }

    const rawText = value.toString().trim();

    if (/^\d{4}-\d{2}$/.test(rawText)) {
      return rawText;
    }

    const parsed = new Date(rawText);

    if (!isNaN(parsed.getTime())) {
      return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`;
    }

    return rawText;
  }

  // Returns YYYY-MM-DD for the full snapshot date so we can identify the
  // latest daily extract within each month. If only YYYY-MM is available,
  // returns YYYY-MM-01 as a fallback (all rows share the same pseudo-date).
  function normalizeDate(value: string | number | boolean): string {
    if (typeof value === "number") {
      if (value <= 0) return "";
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const date = new Date(excelEpoch.getTime() + value * 86400000);
      return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
    }

    const rawText = value.toString().trim();

    if (/^\d{4}-\d{2}-\d{2}$/.test(rawText)) return rawText;

    // Month-only — no day info; all rows in this month share the same pseudo-date
    if (/^\d{4}-\d{2}$/.test(rawText)) return rawText + "-01";

    const parsed = new Date(rawText);
    if (!isNaN(parsed.getTime())) {
      return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
    }

    return "";
  }

  // Rejects Excel epoch garbage (1899-12) and far-future/past noise.
  function isValidMonth(month: string): boolean {
    if (!month || !/^\d{4}-\d{2}$/.test(month)) return false;
    const year = Number(month.split("-")[0]);
    return year >= 2000 && year <= 2099;
  }

  function formatCurrency(value: number, currencySymbol: string): string {
    if (!isFinite(value)) return `${currencySymbol}0`;
    return `${currencySymbol}${Math.round(value).toLocaleString("en-US")}`;
  }

  function formatPercent(value: number): string {
    if (!isFinite(value)) return "0.0%";
    return `${(value * 100).toFixed(1)}%`;
  }

  function pct(numerator: number, denominator: number): number {
    return denominator !== 0 ? numerator / denominator : 0;
  }

  function trendIcon(current: number, previous: number): string {
    if (!isFinite(current) || !isFinite(previous)) {
      return `<span style="color:#888;font-weight:bold;">→</span>`;
    }

    if (current > previous) {
      return `<span style="color:#d40511;font-weight:bold;font-size:18px;">▲</span>`;
    }

    if (current < previous) {
      return `<span style="color:#00843d;font-weight:bold;font-size:18px;">▼</span>`;
    }

    return `<span style="color:#888;font-weight:bold;">→</span>`;
  }

  function statusBadge(label: string, color: string): string {
    return `<span style="display:inline-block;background:${color};color:white;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:bold;">${label}</span>`;
  }

  function getTableOrNull(tableName: string): ExcelScript.Table | null {
    try {
      return workbook.getTable(tableName);
    } catch (e) {
      return null;
    }
  }

  function getTableValuesSafe(table: ExcelScript.Table): (string | number | boolean)[][] {
    const range = table.getRangeBetweenHeaderAndTotal();
    const rowCount = range.getRowCount();
    const colCount = range.getColumnCount();
    const output: (string | number | boolean)[][] = [];
    const chunkSize = 3000;

    for (let start = 0; start < rowCount; start += chunkSize) {
      const take = Math.min(chunkSize, rowCount - start);
      const chunk = range
        .getCell(start, 0)
        .getResizedRange(take - 1, colCount - 1)
        .getValues() as (string | number | boolean)[][];

      output.push(...chunk);
    }

    return output;
  }

  function readReportConfig(): ConfigMap {
    const table = getTableOrNull("ReportConfigTbl");
    if (!table) return {};

    const values = getTableValuesSafe(table);
    const config: ConfigMap = {};

    values.forEach(row => {
      const key = text(row[0]);
      const value = text(row[1]);
      if (key) config[key] = value;
    });

    return config;
  }

  function readFieldMapping(): FieldMap {
    const table = getTableOrNull("FieldMappingTbl");
    if (!table) return {};

    const values = getTableValuesSafe(table);
    const mapping: FieldMap = {};

    values.forEach(row => {
      const standardField = text(row[0]);
      const sourceColumn = text(row[1]);
      const required = boolFromText(row[2]);
      const defaultValue = num(row[3]);

      if (standardField) {
        mapping[standardField] = {
          sourceColumn,
          required,
          defaultValue
        };
      }
    });

    return mapping;
  }

  const config = readReportConfig();
  const mapping = readFieldMapping();

  function getConfigString(key: string, fallback: string): string {
    const value = config[key];
    if (value === undefined || value === null || value.toString().trim() === "") return fallback;
    return value.toString().trim();
  }

  function getConfigNumber(key: string, fallback: number): number {
    const value = config[key];
    if (value === undefined || value === null || value.toString().trim() === "") return fallback;
    const parsed = Number(value);
    return isFinite(parsed) ? parsed : fallback;
  }

  function getConfigBool(key: string, fallback: boolean): boolean {
    const value = config[key];
    if (value === undefined || value === null || value.toString().trim() === "") return fallback;
    return boolFromText(value);
  }

  function getMapSource(standardField: string, fallback: string): string {
    const mapItem = mapping[standardField];
    if (!mapItem || !mapItem.sourceColumn) return fallback;
    return mapItem.sourceColumn;
  }

  function getMapRequired(standardField: string, fallback: boolean): boolean {
    const mapItem = mapping[standardField];
    if (!mapItem) return fallback;
    return mapItem.required;
  }

  function getMapDefault(standardField: string, fallback: number): number {
    const mapItem = mapping[standardField];
    if (!mapItem) return fallback;
    return mapItem.defaultValue;
  }

  const reportName = getConfigString("ReportName", "AMKAD AR Report");
  const businessArea = getConfigString("BusinessArea", "DHL Express AMKAD Finance");
  const currencySymbol = getConfigString("CurrencySymbol", "€");

  const showUAC = getConfigBool("ShowUAC", true);
  const showSimulator = getConfigBool("ShowSimulator", true);
  const showYoY = getConfigBool("ShowYoY", true);
  const showMoM = getConfigBool("ShowMoM", true);
  const showCountry = getConfigBool("ShowCountry", true);
  const showCustomer = getConfigBool("ShowCustomer", true);
  const showTrends = getConfigBool("ShowTrends", true);
  const showDataQuality = getConfigBool("ShowDataQuality", true);

  const highOverduePct = getConfigNumber("HighOverduePct", 0.30);
  const watchOverduePct = getConfigNumber("WatchOverduePct", 0.20);
  const highGT60Pct = getConfigNumber("HighGT60Pct", 0.15);
  const watchGT60Pct = getConfigNumber("WatchGT60Pct", 0.08);
  const highGT90Pct = getConfigNumber("HighGT90Pct", 0.08);
  const watchGT90Pct = getConfigNumber("WatchGT90Pct", 0.04);
  const highUACPct = getConfigNumber("HighUACPct", 0.10);
  const watchUACPct = getConfigNumber("WatchUACPct", 0.05);

  function countryStatus(overdueValuePct: number, gt60ValuePct: number, gt90ValuePct: number): string {
    if (overdueValuePct >= highOverduePct || gt60ValuePct >= highGT60Pct || gt90ValuePct >= highGT90Pct) {
      return statusBadge("High", "#d40511");
    }

    if (overdueValuePct >= watchOverduePct || gt60ValuePct >= watchGT60Pct || gt90ValuePct >= watchGT90Pct) {
      return statusBadge("Watch", "#ed7d31");
    }

    return statusBadge("Low", "#00843d");
  }

  function uacStatus(uacValuePct: number): string {
    if (uacValuePct >= highUACPct) return statusBadge("High", "#d40511");
    if (uacValuePct >= watchUACPct) return statusBadge("Watch", "#ed7d31");
    return statusBadge("Low", "#00843d");
  }

  const rawTable = getTableOrNull(sourceTableName);

  if (!rawTable) {
    return {
      Error: `Source table "${sourceTableName}" was not found.`
    };
  }

  const headers = rawTable.getHeaderRowRange().getTexts()[0].map(h => text(h));
  const values = getTableValuesSafe(rawTable);

  const requiredRawDataColumns = [
    "Month",
    "Country",
    "Customer",
    "Go Live",
    "Payment Term (days)",
    "TDSO",
    "DSO",
    "TDSO Gap",
    "Total AR € ()",
    "Gross Sales € ()",
    "Overdue € ()",
    "> 60 days € ()",
    "> 90 days € ()",
    "Bad Debt Provision € ()",
    "Total AR € (Live)",
    "Gross Sales € (Live)",
    "Overdue € (Live)",
    "> 60 days € (Live)",
    "> 90 days € (Live)",
    "Total AR",
    "Overdue",
    "> 60 days",
    "Gross Sales",
    "> 90 days",
    "Total UAC € (Live)",
    "Total UAC",
    "Total Payments € (Live)",
    "Total Payments"
  ];

  const missingRawDataColumns = requiredRawDataColumns.filter(columnName => headers.indexOf(columnName) === -1);

  if (missingRawDataColumns.length > 0) {
    return {
      Error: `Missing required RawData table columns: ${missingRawDataColumns.join("; ")}.`
    };
  }

  function colBySourceName(sourceColumn: string): number {
    return headers.indexOf(sourceColumn);
  }

  function colByStandardField(standardField: string, fallbackSourceColumn: string): number {
    const sourceColumn = getMapSource(standardField, fallbackSourceColumn);
    return colBySourceName(sourceColumn);
  }

  const columnIndexByStandardField: { [standardField: string]: number } = {
    Month: colByStandardField("Month", "Month"),
    Country: colByStandardField("Country", "Country"),
    Customer: colByStandardField("Customer", "Customer"),
    TotalAR: colByStandardField("TotalAR", "Total AR € (Live)"),
    Overdue: colByStandardField("Overdue", "Overdue € (Live)"),
    GT60: colByStandardField("GT60", "> 60 days € (Live)"),
    GT90: colByStandardField("GT90", "> 90 days € (Live)"),
    GrossSales: colByStandardField("GrossSales", "Gross Sales € (Live)"),
    UAC: colByStandardField("UAC", "Total UAC € (Live)"),
    Payments: colByStandardField("Payments", "Total Payments € (Live)")
  };

  const requiredFields = [
    "Month",
    "Country",
    "Customer",
    "TotalAR",
    "Overdue"
  ];

  const missingRequiredFields: string[] = [];

  requiredFields.forEach(field => {
    const isRequired = getMapRequired(field, true);
    const colIndex = columnIndexByStandardField[field];

    if (isRequired && colIndex === -1) {
      missingRequiredFields.push(`${field} mapped to "${getMapSource(field, field)}"`);
    }
  });

  Object.keys(mapping).forEach(field => {
    const isRequired = mapping[field].required;
    const sourceColumn = mapping[field].sourceColumn;
    const colIndex = columnIndexByStandardField[field];

    if (isRequired && colIndex === -1) {
      const alreadyIncluded = missingRequiredFields.some(x => x.startsWith(`${field} `));

      if (!alreadyIncluded) {
        missingRequiredFields.push(`${field} mapped to "${sourceColumn}"`);
      }
    }
  });

  if (missingRequiredFields.length > 0) {
    return {
      Error: `Missing required mapped columns: ${missingRequiredFields.join("; ")}. Update Field_Mapping or RawData headers.`
    };
  }

  function readCellAsText(row: (string | number | boolean)[], standardField: string): string {
    const colIndex = columnIndexByStandardField[standardField];

    if (colIndex === -1) {
      return "";
    }

    return text(row[colIndex]);
  }

  function readCellAsNumber(row: (string | number | boolean)[], standardField: string): number {
    const colIndex = columnIndexByStandardField[standardField];

    if (colIndex === -1) {
      return getMapDefault(standardField, 0);
    }

    const rawValue = row[colIndex];

    if (rawValue === null || rawValue === undefined || rawValue.toString().trim() === "") {
      return getMapDefault(standardField, 0);
    }

    return num(rawValue);
  }

  // ─────────────────────────────────────────────────────────────────────────
  // PASS 1 — MTD snapshot detection
  //
  // The RawData table may contain multiple daily extracts for the same month
  // (e.g., June 1 through June 16 all tagged as "2026-06" or with full dates
  // like "2026-06-01" … "2026-06-16").  Summing all of those rows would
  // inflate every KPI by ~N× (where N = number of daily snapshots).
  //
  // We parse the full date from the Month cell so that, per calendar month,
  // we only retain the rows belonging to the LATEST snapshot date.  This
  // gives us a true MTD point-in-time view instead of a running accumulation.
  // ─────────────────────────────────────────────────────────────────────────
  const latestDatePerMonth: { [month: string]: string } = {};

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const monthRaw = row[columnIndexByStandardField.Month];
    if (monthRaw === null || monthRaw === undefined) continue;

    const month = normalizeMonth(monthRaw as string | number | boolean);
    if (!isValidMonth(month)) continue;

    const fullDate = normalizeDate(monthRaw as string | number | boolean);
    if (!fullDate) continue;

    if (!latestDatePerMonth[month] || fullDate > latestDatePerMonth[month]) {
      latestDatePerMonth[month] = fullDate;
    }
  }

  // ─────────────────────────────────────────────────────────────────────────
  // PASS 2 — Aggregate only the latest-snapshot rows per month
  // ─────────────────────────────────────────────────────────────────────────
  const customerData: { [customer: string]: { [month: string]: Agg } } = {};
  const countryData: { [country: string]: { [month: string]: Agg } } = {};
  const customerCountryData: { [key: string]: CustomerCountryAgg } = {};
  const monthTotals: { [month: string]: Agg } = {};
  const monthSet: { [month: string]: boolean } = {};

  let negativeTotalARRows = 0;
  let negativeOverdueRows = 0;
  let blankCustomerRows = 0;
  let blankCountryRows = 0;
  let gt60GreaterThanARRows = 0;
  let gt90GreaterThanARRows = 0;
  let snapshotSkippedRows = 0;

  for (let r = 0; r < values.length; r++) {
    const row = values[r];

    const monthRaw = row[columnIndexByStandardField.Month];
    const countryRaw = readCellAsText(row, "Country");
    const customerRaw = readCellAsText(row, "Customer");

    if (monthRaw === null || monthRaw === undefined) continue;

    const country = countryRaw.trim();
    const customer = customerRaw.trim();

    if (!country) blankCountryRows++;
    if (!customer) blankCustomerRows++;

    if (!country || !customer) continue;

    const month = normalizeMonth(monthRaw as string | number | boolean);

    if (!isValidMonth(month)) continue;

    // Skip rows that belong to an earlier daily snapshot within this month.
    // This is the MTD fix: only the latest extract date is used.
    const rowDate = normalizeDate(monthRaw as string | number | boolean);
    if (rowDate && latestDatePerMonth[month] && rowDate !== latestDatePerMonth[month]) {
      snapshotSkippedRows++;
      continue;
    }

    const rowAgg: Agg = {
      totalAR: readCellAsNumber(row, "TotalAR"),
      overdue: readCellAsNumber(row, "Overdue"),
      gt60: readCellAsNumber(row, "GT60"),
      gt90: readCellAsNumber(row, "GT90"),
      grossSales: readCellAsNumber(row, "GrossSales"),
      uac: readCellAsNumber(row, "UAC"),
      payments: readCellAsNumber(row, "Payments")
    };

    if (rowAgg.totalAR < 0) negativeTotalARRows++;
    if (rowAgg.overdue < 0) negativeOverdueRows++;
    if (rowAgg.gt60 > rowAgg.totalAR && rowAgg.totalAR > 0) gt60GreaterThanARRows++;
    if (rowAgg.gt90 > rowAgg.totalAR && rowAgg.totalAR > 0) gt90GreaterThanARRows++;

    if (!customerData[customer]) customerData[customer] = {};
    if (!customerData[customer][month]) customerData[customer][month] = blankAgg();
    addAgg(customerData[customer][month], rowAgg);

    if (!countryData[country]) countryData[country] = {};
    if (!countryData[country][month]) countryData[country][month] = blankAgg();
    addAgg(countryData[country][month], rowAgg);

    const ccKey = `${country}||${customer}`;

    if (!customerCountryData[ccKey]) {
      customerCountryData[ccKey] = {
        country,
        customer,
        months: {}
      };
    }

    if (!customerCountryData[ccKey].months[month]) {
      customerCountryData[ccKey].months[month] = blankAgg();
    }

    addAgg(customerCountryData[ccKey].months[month], rowAgg);

    if (!monthTotals[month]) monthTotals[month] = blankAgg();
    addAgg(monthTotals[month], rowAgg);

    monthSet[month] = true;
  }

  const months = Object.keys(monthSet).sort().reverse();

  if (months.length === 0) {
    return {
      Error: "No valid data rows were found after applying mapping rules."
    };
  }

  const latestMonth = months[0];
  const previousMonth = months[1] || "";
  const latestThreeMonths = months.slice(0, 3);
  const latestSnapshotDate = latestDatePerMonth[latestMonth] || latestMonth;

  const latestAgg = monthTotals[latestMonth] || blankAgg();

  const overduePctValue = pct(latestAgg.overdue, latestAgg.totalAR);
  const gt60PctValue = pct(latestAgg.gt60, latestAgg.totalAR);
  const gt90PctValue = pct(latestAgg.gt90, latestAgg.totalAR);
  const uacPctValue = pct(latestAgg.uac, latestAgg.totalAR);
  const paymentsVsGrossSalesValue = pct(latestAgg.payments, latestAgg.grossSales);
  const netOverdueEstimateValue = latestAgg.overdue - latestAgg.uac;

  const topCustomers = Object.keys(customerData)
    .filter(c => customerData[c][latestMonth])
    .sort((a, b) => {
      const aVal = customerData[a][latestMonth]?.totalAR || 0;
      const bVal = customerData[b][latestMonth]?.totalAR || 0;
      return bVal - aVal;
    })
    .slice(0, 15);

  function buildCustomerMoMTable(metric: "gt60" | "gt90", title: string): string {
    if (!showMoM) return "";

    let html = `
      <div class="customer-table-wrapper">
        <div class="customer-table-title">${escapeHtml(title)}</div>
        <table class="customer-matrix searchable-table">
          <thead>
            <tr>
              <th>Customer</th>
              <th>Trend</th>
    `;

    latestThreeMonths.forEach(m => {
      html += `<th>${escapeHtml(m)}</th>`;
    });

    html += `
            </tr>
          </thead>
          <tbody>
    `;

    topCustomers.forEach(customer => {
      const current = customerData[customer][latestMonth];
      const previous = previousMonth ? customerData[customer][previousMonth] : undefined;

      const currentPct = current ? pct(current[metric], current.totalAR) : 0;
      const previousPct = previous ? pct(previous[metric], previous.totalAR) : 0;

      html += `
        <tr data-customer="${escapeHtml(customer)}">
          <td class="customer-name">${escapeHtml(customer)}</td>
          <td class="trend-cell">${trendIcon(currentPct, previousPct)}</td>
      `;

      latestThreeMonths.forEach(month => {
        const agg = customerData[customer][month];
        const valuePct = agg ? pct(agg[metric], agg.totalAR) : 0;
        html += `<td>${formatPercent(valuePct)}</td>`;
      });

      html += `</tr>`;
    });

    html += `
          </tbody>
        </table>
      </div>
    `;

    return html;
  }

  const gt90CustomerTable = buildCustomerMoMTable(
    "gt90",
    "> 90 days (%) — Customer Month-over-Month View"
  );

  const gt60CustomerTable = buildCustomerMoMTable(
    "gt60",
    "> 60 days (%) — Customer Month-over-Month View"
  );

  const [latestYearText, latestMonthText] = latestMonth.split("-");
  const latestYear = Number(latestYearText);
  const monthNumber = latestMonthText;

  const yoyMonths = [
    `${latestYear}-${monthNumber}`,
    `${latestYear - 1}-${monthNumber}`,
    `${latestYear - 2}-${monthNumber}`
  ];

  function metricValue(monthTotal: Agg, metric: string): string {
    if (metric === "Total AR") return formatCurrency(monthTotal.totalAR, currencySymbol);
    if (metric === "Total Overdue") return formatCurrency(monthTotal.overdue, currencySymbol);
    if (metric === "Overdue %") return formatPercent(pct(monthTotal.overdue, monthTotal.totalAR));
    if (metric === "GT60 %") return formatPercent(pct(monthTotal.gt60, monthTotal.totalAR));
    if (metric === "GT90 %") return formatPercent(pct(monthTotal.gt90, monthTotal.totalAR));
    if (metric === "Gross Sales") return formatCurrency(monthTotal.grossSales, currencySymbol);
    if (metric === "Total UAC") return formatCurrency(monthTotal.uac, currencySymbol);
    if (metric === "UAC % of AR") return formatPercent(pct(monthTotal.uac, monthTotal.totalAR));
    return "";
  }

  function metricNumber(monthTotal: Agg, metric: string): number {
    if (metric === "Total AR") return monthTotal.totalAR;
    if (metric === "Total Overdue") return monthTotal.overdue;
    if (metric === "Overdue %") return pct(monthTotal.overdue, monthTotal.totalAR);
    if (metric === "GT60 %") return pct(monthTotal.gt60, monthTotal.totalAR);
    if (metric === "GT90 %") return pct(monthTotal.gt90, monthTotal.totalAR);
    if (metric === "Gross Sales") return monthTotal.grossSales;
    if (metric === "Total UAC") return monthTotal.uac;
    if (metric === "UAC % of AR") return pct(monthTotal.uac, monthTotal.totalAR);
    return 0;
  }

  const yoyTotals = yoyMonths.map(m => monthTotals[m] || blankAgg());

  const sameMonthYoYTable = showYoY ? `
    <div class="customer-table-wrapper">
      <div class="customer-table-title">Same-Month YoY Comparison (${escapeHtml(yoyMonths[0])} vs prior years)</div>
      <table class="customer-matrix searchable-table">
        <thead>
          <tr>
            <th>Metric</th>
            <th>${escapeHtml(yoyMonths[0])}</th>
            <th>${escapeHtml(yoyMonths[1])}</th>
            <th>${escapeHtml(yoyMonths[2])}</th>
            <th>YoY Trend</th>
          </tr>
        </thead>
        <tbody>
          ${["Total AR", "Total Overdue", "Overdue %", "GT60 %", "GT90 %", "Gross Sales", "Total UAC", "UAC % of AR"].map(metric => `
            <tr>
              <td class="customer-name">${escapeHtml(metric)}</td>
              <td>${metricValue(yoyTotals[0], metric)}</td>
              <td>${metricValue(yoyTotals[1], metric)}</td>
              <td>${metricValue(yoyTotals[2], metric)}</td>
              <td>${trendIcon(metricNumber(yoyTotals[0], metric), metricNumber(yoyTotals[1], metric))}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>` : "";

  const countrySummaryTable = showCountry ? `
    <div class="customer-table-wrapper">
      <div class="customer-table-title">Country Summary — ${escapeHtml(latestMonth)}</div>
      <table class="customer-matrix searchable-table" id="countrySummaryTable">
        <thead>
          <tr>
            <th>Country</th>
            <th>Total AR</th>
            <th>Overdue</th>
            <th>Overdue %</th>
            <th>GT60 %</th>
            <th>GT90 %</th>
            <th>Gross Sales</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          ${Object.keys(countryData).filter(c => countryData[c][latestMonth]).map(country => {
    const agg = countryData[country][latestMonth];
    return `<tr data-country="${escapeHtml(country)}">
              <td class="customer-name">${escapeHtml(country)}</td>
              <td>${formatCurrency(agg.totalAR, currencySymbol)}</td>
              <td>${formatCurrency(agg.overdue, currencySymbol)}</td>
              <td>${formatPercent(pct(agg.overdue, agg.totalAR))}</td>
              <td>${formatPercent(pct(agg.gt60, agg.totalAR))}</td>
              <td>${formatPercent(pct(agg.gt90, agg.totalAR))}</td>
              <td>${formatCurrency(agg.grossSales, currencySymbol)}</td>
              <td>${countryStatus(pct(agg.overdue, agg.totalAR), pct(agg.gt60, agg.totalAR), pct(agg.gt90, agg.totalAR))}</td>
            </tr>`;
  }).join("")}
        </tbody>
      </table>
    </div>` : "";

  const topRiskCustomersTable = showCustomer ? `
    <div class="customer-table-wrapper">
      <div class="customer-table-title">Top 10 Customers by GT90 Amount — ${escapeHtml(latestMonth)}</div>
      <table class="customer-matrix searchable-table" id="topRiskTable">
        <thead>
          <tr>
            <th>#</th>
            <th>Country</th>
            <th>Customer</th>
            <th>Total AR</th>
            <th>Overdue</th>
            <th>GT60</th>
            <th>GT90</th>
            <th>GT90 %</th>
          </tr>
        </thead>
        <tbody>
          ${Object.keys(customerCountryData).filter(k => customerCountryData[k].months[latestMonth]).sort((a, b) => customerCountryData[b].months[latestMonth].gt90 - customerCountryData[a].months[latestMonth].gt90).slice(0, 10).map((key, i) => {
    const item = customerCountryData[key];
    const agg = item.months[latestMonth];
    return `<tr><td>${i + 1}</td><td>${escapeHtml(item.country)}</td><td class="customer-name">${escapeHtml(item.customer)}</td><td>${formatCurrency(agg.totalAR, currencySymbol)}</td><td>${formatCurrency(agg.overdue, currencySymbol)}</td><td>${formatCurrency(agg.gt60, currencySymbol)}</td><td>${formatCurrency(agg.gt90, currencySymbol)}</td><td>${formatPercent(pct(agg.gt90, agg.totalAR))}</td></tr>`;
  }).join("")}
        </tbody>
      </table>
    </div>` : "";

  const topUACCustomersTable = showUAC ? `
    <div class="customer-table-wrapper">
      <div class="customer-table-title">Top 10 Customers by UAC — ${escapeHtml(latestMonth)}</div>
      <table class="customer-matrix searchable-table" id="topUACTable">
        <thead>
          <tr>
            <th>#</th>
            <th>Country</th>
            <th>Customer</th>
            <th>Total AR</th>
            <th>UAC</th>
            <th>UAC %</th>
            <th>Overdue</th>
            <th>Net Overdue</th>
            <th>Payments</th>
          </tr>
        </thead>
        <tbody>
          ${Object.keys(customerCountryData).filter(k => customerCountryData[k].months[latestMonth]).sort((a, b) => customerCountryData[b].months[latestMonth].uac - customerCountryData[a].months[latestMonth].uac).slice(0, 10).map((key, i) => {
    const item = customerCountryData[key];
    const agg = item.months[latestMonth];
    return `<tr><td>${i + 1}</td><td>${escapeHtml(item.country)}</td><td class="customer-name">${escapeHtml(item.customer)}</td><td>${formatCurrency(agg.totalAR, currencySymbol)}</td><td>${formatCurrency(agg.uac, currencySymbol)}</td><td>${formatPercent(pct(agg.uac, agg.totalAR))}</td><td>${formatCurrency(agg.overdue, currencySymbol)}</td><td>${formatCurrency(agg.overdue - agg.uac, currencySymbol)}</td><td>${formatCurrency(agg.payments, currencySymbol)}</td></tr>`;
  }).join("")}
        </tbody>
      </table>
    </div>` : "";

  const uacCountryTable = showUAC && showCountry ? `
    <div class="customer-table-wrapper">
      <div class="customer-table-title">Country UAC Summary — ${escapeHtml(latestMonth)}</div>
      <table class="customer-matrix searchable-table" id="uacCountryTable">
        <thead>
          <tr>
            <th>Country</th>
            <th>Total AR</th>
            <th>UAC</th>
            <th>UAC %</th>
            <th>Overdue</th>
            <th>Payments</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          ${Object.keys(countryData).filter(c => countryData[c][latestMonth]).map(country => {
    const agg = countryData[country][latestMonth];
    const uacP = pct(agg.uac, agg.totalAR);
    return `<tr><td class="customer-name">${escapeHtml(country)}</td><td>${formatCurrency(agg.totalAR, currencySymbol)}</td><td>${formatCurrency(agg.uac, currencySymbol)}</td><td>${formatPercent(uacP)}</td><td>${formatCurrency(agg.overdue, currencySymbol)}</td><td>${formatCurrency(agg.payments, currencySymbol)}</td><td>${uacStatus(uacP)}</td></tr>`;
  }).join("")}
        </tbody>
      </table>
    </div>` : "";

  const uacExceptionsTable = showUAC ? `
    <div class="customer-table-wrapper">
      <div class="customer-table-title">UAC Exceptions / Action Required — ${escapeHtml(latestMonth)}</div>
      <table class="customer-matrix searchable-table" id="uacExceptionsTable">
        <thead>
          <tr>
            <th>Country</th>
            <th>Customer</th>
            <th>UAC</th>
            <th>UAC %</th>
            <th>Overdue</th>
            <th>Net Overdue</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          ${Object.keys(customerCountryData).filter(k => customerCountryData[k].months[latestMonth]).slice(0, 15).map(key => {
    const item = customerCountryData[key];
    const agg = item.months[latestMonth];
    return `<tr><td>${escapeHtml(item.country)}</td><td class="customer-name">${escapeHtml(item.customer)}</td><td>${formatCurrency(agg.uac, currencySymbol)}</td><td>${formatPercent(pct(agg.uac, agg.totalAR))}</td><td>${formatCurrency(agg.overdue, currencySymbol)}</td><td>${formatCurrency(agg.overdue - agg.uac, currencySymbol)}</td><td>Review</td></tr>`;
  }).join("")}
        </tbody>
      </table>
    </div>` : "";

  // Build a real, ranked list of items that breach the High thresholds.
  type ActionItem = { country: string; customer: string; totalAR: number; overdue: number; overduePct: number; gt90Pct: number; uacPct: number; reason: string; severity: number };

  const actionItems: ActionItem[] = [];

  Object.keys(customerCountryData)
    .filter(k => customerCountryData[k].months[latestMonth])
    .forEach(key => {
      const item = customerCountryData[key];
      const agg = item.months[latestMonth];
      if (agg.totalAR <= 0) return;

      const overdueP = pct(agg.overdue, agg.totalAR);
      const gt60P = pct(agg.gt60, agg.totalAR);
      const gt90P = pct(agg.gt90, agg.totalAR);
      const uacP = pct(agg.uac, agg.totalAR);

      const reasons: string[] = [];
      if (overdueP >= highOverduePct) reasons.push("Overdue % over threshold");
      if (gt60P >= highGT60Pct) reasons.push("GT60 % over threshold");
      if (gt90P >= highGT90Pct) reasons.push("GT90 % over threshold");
      if (uacP >= highUACPct) reasons.push("High unapplied cash");

      if (reasons.length === 0) return;

      actionItems.push({
        country: item.country,
        customer: item.customer,
        totalAR: agg.totalAR,
        overdue: agg.overdue,
        overduePct: overdueP,
        gt90Pct: gt90P,
        uacPct: uacP,
        reason: reasons.join("; "),
        severity: agg.overdue
      });
    });

  actionItems.sort((a, b) => b.severity - a.severity);
  const topActionItems = actionItems.slice(0, 20);

  const actionRequiredRows = topActionItems.length > 0
    ? topActionItems.map((item, i) => `<tr data-country="${escapeHtml(item.country)}" data-customer="${escapeHtml(item.customer)}">
              <td>${i + 1}</td>
              <td>${escapeHtml(item.country)}</td>
              <td class="customer-name">${escapeHtml(item.customer)}</td>
              <td>${formatCurrency(item.totalAR, currencySymbol)}</td>
              <td>${formatCurrency(item.overdue, currencySymbol)}</td>
              <td>${formatPercent(item.overduePct)}</td>
              <td>${formatPercent(item.gt90Pct)}</td>
              <td>${formatPercent(item.uacPct)}</td>
              <td style="text-align:left;">${escapeHtml(item.reason)}</td>
            </tr>`).join("")
    : `<tr><td colspan="9">No customers currently breach the High-risk thresholds for ${escapeHtml(latestMonth)}.</td></tr>`;

  const actionRequiredTable = `
    <div class="customer-table-wrapper">
      <div class="customer-table-title">Action Required — Customers Breaching High-Risk Thresholds (${escapeHtml(latestMonth)})</div>
      <table class="customer-matrix searchable-table" id="actionRequiredTable">
        <thead>
          <tr>
            <th>#</th>
            <th>Country</th>
            <th>Customer</th>
            <th>Total AR</th>
            <th>Overdue</th>
            <th>Overdue %</th>
            <th>GT90 %</th>
            <th>UAC %</th>
            <th>Reason</th>
          </tr>
        </thead>
        <tbody>
          ${actionRequiredRows}
        </tbody>
      </table>
    </div>`;

  const dataQualityTable = showDataQuality ? `
    <div class="customer-table-wrapper">
      <div class="customer-table-title">Data Quality Checks</div>
      <table class="customer-matrix searchable-table" id="dataQualityTable">
        <tbody>
          <tr><td class="customer-name">Report name</td><td>${escapeHtml(reportName)}</td></tr>
          <tr><td class="customer-name">Business area</td><td>${escapeHtml(businessArea)}</td></tr>
          <tr><td class="customer-name">Source table</td><td>${escapeHtml(sourceTableName)}</td></tr>
          <tr><td class="customer-name">Latest month detected</td><td>${escapeHtml(latestMonth)}</td></tr>
          <tr><td class="customer-name">Latest snapshot date used</td><td>${escapeHtml(latestSnapshotDate)}</td></tr>
          <tr><td class="customer-name">Rows processed (total)</td><td>${values.length.toLocaleString("en-US")}</td></tr>
          <tr><td class="customer-name">Rows skipped (earlier snapshots)</td><td>${snapshotSkippedRows.toLocaleString("en-US")}</td></tr>
          <tr><td class="customer-name">Blank customer rows skipped</td><td>${blankCustomerRows.toLocaleString("en-US")}</td></tr>
          <tr><td class="customer-name">Blank country rows skipped</td><td>${blankCountryRows.toLocaleString("en-US")}</td></tr>
          <tr><td class="customer-name">Negative Total AR rows</td><td>${negativeTotalARRows.toLocaleString("en-US")}</td></tr>
          <tr><td class="customer-name">Negative Overdue rows</td><td>${negativeOverdueRows.toLocaleString("en-US")}</td></tr>
          <tr><td class="customer-name">Rows where GT60 > AR</td><td>${gt60GreaterThanARRows.toLocaleString("en-US")}</td></tr>
          <tr><td class="customer-name">Rows where GT90 > AR</td><td>${gt90GreaterThanARRows.toLocaleString("en-US")}</td></tr>
        </tbody>
      </table>
    </div>` : "";

  const sameMonthGT90CustomerTable = buildCustomerMoMTable("gt90", "> 90 days (%) — Same-Month YoY Customer View");
  const sameMonthGT60CustomerTable = buildCustomerMoMTable("gt60", "> 60 days (%) — Same-Month YoY Customer View");

  const countryOptions = Object.keys(countryData).sort();
  const customerOptions = Object.keys(customerData).sort();

  const trendData = months.slice().reverse()
    // Drop months whose latest snapshot has no AR — these are missing/empty
    // extracts (e.g. a month where no snapshot was captured) and would
    // otherwise show as a false crash to zero in the trend charts.
    .filter(month => (monthTotals[month] || blankAgg()).totalAR > 0)
    .map(month => {
    const agg = monthTotals[month] || blankAgg();
    return {
      month,
      totalAR: agg.totalAR,
      overdue: agg.overdue,
      overduePct: pct(agg.overdue, agg.totalAR),
      gt60Pct: pct(agg.gt60, agg.totalAR),
      gt90Pct: pct(agg.gt90, agg.totalAR),
      uac: agg.uac,
      uacPct: pct(agg.uac, agg.totalAR),
      payments: agg.payments,
      grossSales: agg.grossSales
    };
  });

  const countryDrilldown = Object.keys(countryData).filter(country => countryData[country][latestMonth]).map(country => {
    const agg = countryData[country][latestMonth];
    return {
      country,
      totalAR: agg.totalAR,
      overdue: agg.overdue,
      overduePct: pct(agg.overdue, agg.totalAR),
      gt60: agg.gt60,
      gt60Pct: pct(agg.gt60, agg.totalAR),
      gt90: agg.gt90,
      gt90Pct: pct(agg.gt90, agg.totalAR),
      grossSales: agg.grossSales,
      uac: agg.uac,
      uacPct: pct(agg.uac, agg.totalAR),
      payments: agg.payments
    };
  });

  const customerDrilldown = Object.keys(customerCountryData).filter(key => customerCountryData[key].months[latestMonth]).map(key => {
    const item = customerCountryData[key];
    const agg = item.months[latestMonth];
    return {
      country: item.country,
      customer: item.customer,
      totalAR: agg.totalAR,
      overdue: agg.overdue,
      overduePct: pct(agg.overdue, agg.totalAR),
      gt60: agg.gt60,
      gt60Pct: pct(agg.gt60, agg.totalAR),
      gt90: agg.gt90,
      gt90Pct: pct(agg.gt90, agg.totalAR),
      grossSales: agg.grossSales,
      uac: agg.uac,
      uacPct: pct(agg.uac, agg.totalAR),
      payments: agg.payments
    };
  });

  // Per-customer month-over-month (aggregated across countries) for the Movers board.
  const customerMoM = Object.keys(customerData)
    .filter(customer => customerData[customer][latestMonth])
    .map(customer => {
      const now = customerData[customer][latestMonth] || blankAgg();
      const prev = (previousMonth && customerData[customer][previousMonth]) || blankAgg();
      return {
        customer,
        arNow: now.totalAR,
        arPrev: prev.totalAR,
        overdueNow: now.overdue,
        overduePrev: prev.overdue,
        gt90Now: now.gt90,
        gt90Prev: prev.gt90,
        hasPrev: previousMonth ? !!customerData[customer][previousMonth] : false
      };
    });

  // Numeric data-quality metrics for the confidence banner + reconciliation strip.
  const dataQuality = {
    latestMonth: latestMonth,
    previousMonth: previousMonth,
    latestSnapshotDate: latestSnapshotDate,
    rowsTotal: values.length,
    rowsSkipped: snapshotSkippedRows,
    rowsUsed: values.length - snapshotSkippedRows,
    blankCustomerRows: blankCustomerRows,
    blankCountryRows: blankCountryRows,
    negativeARRows: negativeTotalARRows,
    negativeOverdueRows: negativeOverdueRows,
    gt60OverARRows: gt60GreaterThanARRows,
    gt90OverARRows: gt90GreaterThanARRows
  };

  return {
    ReportName: reportName,
    BusinessArea: businessArea,
    SourceSheet: sourceTableName,
    CurrencySymbol: currencySymbol,

    ShowUAC: showUAC,
    ShowSimulator: showSimulator,
    ShowYoY: showYoY,
    ShowMoM: showMoM,
    ShowCountry: showCountry,
    ShowCustomer: showCustomer,
    ShowTrends: showTrends,
    ShowDataQuality: showDataQuality,

    TotalAR: formatCurrency(latestAgg.totalAR, currencySymbol),
    TotalOverdue: formatCurrency(latestAgg.overdue, currencySymbol),
    OverduePct: formatPercent(overduePctValue),
    GT60Pct: formatPercent(gt60PctValue),
    GT90Pct: formatPercent(gt90PctValue),
    GrossSales: formatCurrency(latestAgg.grossSales, currencySymbol),

    TotalUAC: formatCurrency(latestAgg.uac, currencySymbol),
    UACPctOfAR: formatPercent(uacPctValue),
    TotalPayments: formatCurrency(latestAgg.payments, currencySymbol),
    PaymentsVsGrossSales: formatPercent(paymentsVsGrossSalesValue),
    NetOverdueEstimate: formatCurrency(netOverdueEstimateValue, currencySymbol),

    TotalARNumeric: latestAgg.totalAR,
    TotalOverdueNumeric: latestAgg.overdue,
    GT60Numeric: latestAgg.gt60,
    GT90Numeric: latestAgg.gt90,
    GrossSalesNumeric: latestAgg.grossSales,
    TotalUACNumeric: latestAgg.uac,
    TotalPaymentsNumeric: latestAgg.payments,

    LatestMonth: latestMonth,
    LatestSnapshotDate: latestSnapshotDate,

    GT90CustomerTable: gt90CustomerTable,
    GT60CustomerTable: gt60CustomerTable,
    SameMonthYoYTable: sameMonthYoYTable,
    SameMonthGT90CustomerTable: sameMonthGT90CustomerTable,
    SameMonthGT60CustomerTable: sameMonthGT60CustomerTable,
    CountrySummaryTable: countrySummaryTable,
    TopRiskCustomersTable: topRiskCustomersTable,
    TopUACCustomersTable: topUACCustomersTable,
    UACCountryTable: uacCountryTable,
    UACExceptionsTable: uacExceptionsTable,
    ActionRequiredTable: actionRequiredTable,
    DataQualityTable: dataQualityTable,

    CountryOptionsJson: JSON.stringify(countryOptions),
    CustomerOptionsJson: JSON.stringify(customerOptions),
    TrendDataJson: JSON.stringify(trendData),
    CountryDrilldownJson: JSON.stringify(countryDrilldown),
    CustomerDrilldownJson: JSON.stringify(customerDrilldown),
    CustomerMoMJson: JSON.stringify(customerMoM),
    DataQualityJson: JSON.stringify(dataQuality)
  };
}
