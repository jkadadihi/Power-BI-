/*
 * Offline test harness for the two Cust Sol Office Scripts.
 * Run via tests/test_custsol.sh — it compiles the .ts files to plain JS,
 * then this file executes them against mock workbooks built from the real
 * source/destination headers and the Amazon CA sample row.
 *
 * What it proves before any live Power Automate run:
 *   - reader maps every column by header name, skips blank rows,
 *     and resolves the duplicated "Total Payments" header correctly
 *     (first occurrence = native col G, last occurrence = Euro col V)
 *   - reader flags rows with blank Euro columns instead of failing
 *   - writer lands every value in the exact A:AD position Raw Data EUR expects
 *   - writer leaves F:P null so the table's own formulas auto-fill
 *   - writer's duplicate guard makes a same-day re-run a no-op
 */
const fs = require("fs");
const path = require("path");

const buildDir = process.argv[2];
if (!buildDir) {
  console.error("usage: node custsol_test.js <dir containing compiled custsol_*.js>");
  process.exit(2);
}

const readerSrc = fs.readFileSync(path.join(buildDir, "reader/custsol_read_daily.js"), "utf8");
const writerSrc = fs.readFileSync(path.join(buildDir, "writer/custsol_append_rawdata_eur.js"), "utf8");

const runReader = new Function("workbook", readerSrc + "\nreturn main(workbook);");
const runWriter = new Function("workbook", "sourceJson", "ExcelScript",
  writerSrc + "\nreturn main(workbook, sourceJson);");

const ExcelScriptMock = { CalculationType: { full: "full" } };

/* ---------- SOURCE FILE MOCK (headers exactly as in the daily report) ---------- */
const srcHeaders = [
  "Report Date","countrycode","currencycode","Go Live Date","Customer Reporting Name",
  "Payment Terms","Total Payments","Total UAC","TDSO","TTLAR","Overdue","GT60days",
  "Gross Sales","GT90Days","GT60Days Euro","Gross Sales Euro","Company Code",
  "Total AREuro","Overdue AREuro","GT90Days Euro","Total UAC Euro","Total Payments"
];
const amazonRow = [
  "07/13/2026","CA","CAD","2018-07","Amazon",
  45, 120.5, 33.2, 49, 2702.88, 2648.66, 2536.23,
  1301.99, 1400.89, 1563.15, 802.45, "CA13",
  1665.86, 1632.45, 863.41, 12.6, 77.9
];
const nokiaRowMissingEuro = [
  "07/13/2026","MX","MXN","2019-03","Nokia",
  30, 500, 10, 40, 1988.89, 97.82, 539.86,
  700.11, 90.71, "", "", "MX01",
  "", 0, "", "", ""
];
const blankRow = new Array(srcHeaders.length).fill("");

const sourceWorkbookMock = {
  getWorksheets: () => [{
    getUsedRange: () => ({ getValues: () => [srcHeaders, amazonRow, nokiaRowMissingEuro, blankRow] })
  }]
};

/* ---------- DESTINATION TABLE MOCK (Raw Data EUR headers, A:AD) ---------- */
const destHeaders = [
  "Date","Country","Customer","Go Live","Payment Term (days)",
  "TDSO","DSO","TDSO Gap","> 60 days (%)","> 90 days (%)",
  "Total AR € ()","Gross Sales € ()","Overdue € ()","> 60 days € ()","> 90 days € ()",
  "Bad Debt Provision € ()",
  "Total AR € (Live)","Gross Sales € (Live)","Overdue € (Live)","> 60 days € (Live)","> 90 days € (Live)",
  "Total AR","Overdue","> 60 days","Gross Sales","> 90 days",
  "Total UAC € (Live)","Total UAC","Total Payments € (Live)","Total Payments"
];
if (destHeaders.length !== 30) throw new Error("dest header count " + destHeaders.length + " != 30 (A:AD)");

const priorDayRow = new Array(30).fill(null);
priorDayRow[0] = "07/12/2026"; priorDayRow[1] = "CA"; priorDayRow[2] = "Amazon";

function makeDestWorkbook(dataRows) {
  const added = [];
  return {
    _added: added,
    _dataRows: dataRows,
    getTable: (name) => name === "RawData" ? {
      getHeaderRowRange: () => ({ getValues: () => [destHeaders] }),
      getRangeBetweenHeaderAndTotal: () => ({ getValues: () => dataRows }),
      addRows: (_pos, rows) => { rows.forEach(r => { added.push(r); dataRows.push(r); }); }
    } : null,
    getApplication: () => ({ calculate: () => {} })
  };
}

/* =========================== RUN =========================== */
let failures = 0;
function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) { failures++; console.log("FAIL  " + label + "  got=" + JSON.stringify(actual) + " want=" + JSON.stringify(expected)); }
  else console.log("ok    " + label);
}

console.log("--- reader ---");
const readerOut = JSON.parse(runReader(sourceWorkbookMock));
check("totalRows (blank row skipped)", readerOut.totalRows, 2);
check("missingEuroRows", readerOut.missingEuroRows, 1);
const am = readerOut.rows[0];
check("customer", am.customerName, "Amazon");
check("euro TotalAR passthrough", am.totalAREuro, 1665.86);
check("euro TotalPayments = LAST 'Total Payments' col (V not G)", am.totalPaymentsEuro, 77.9);
check("native TotalPayments = FIRST 'Total Payments' col (G)", am.totalPayments, 120.5);
check("missing-euro row nulls", readerOut.rows[1].totalAREuro, null);

console.log("--- writer: first run ---");
const wb1 = makeDestWorkbook([priorDayRow.slice()]);
const w1 = JSON.parse(runWriter(wb1, JSON.stringify(readerOut), ExcelScriptMock));
check("rowsAppended", w1.rowsAppended, 2);
check("skippedDuplicates", w1.skippedDuplicates, 0);
check("needsReview (euro gap flagged)", w1.needsReview, true);

const row = wb1._added[0];
const expect = {
  "A Date": [0, "07/13/2026"], "B Country": [1, "CA"], "C Customer": [2, "Amazon"],
  "D Go Live": [3, "2018-07"], "E Payment Term": [4, 45],
  "Q Total AR € (Live)": [16, 1666], "R Gross Sales € (Live)": [17, 802],
  "S Overdue € (Live)": [18, 1632], "T >60 € (Live)": [19, 1563], "U >90 € (Live)": [20, 863],
  "V Total AR": [21, 2703], "W Overdue": [22, 2649], "X >60 days": [23, 2536],
  "Y Gross Sales": [24, 1302], "Z >90 days": [25, 1401],
  "AA Total UAC € (Live)": [26, 13], "AB Total UAC": [27, 33],
  "AC Total Payments € (Live)": [28, 78], "AD Total Payments": [29, 121]
};
for (const [label, pair] of Object.entries(expect)) check(label, row[pair[0]], pair[1]);
check("F:P untouched (formula columns)", row.slice(5, 16).every(v => v === null), true);

console.log("--- writer: re-run same payload (duplicate guard) ---");
const w2 = JSON.parse(runWriter(wb1, JSON.stringify(readerOut), ExcelScriptMock));
check("rowsAppended on re-run", w2.rowsAppended, 0);
check("skippedDuplicates on re-run", w2.skippedDuplicates, 2);
check("table row count unchanged", wb1._dataRows.length, 3);

console.log(failures === 0 ? "\nALL TESTS PASSED" : "\n" + failures + " FAILURES");
process.exit(failures === 0 ? 0 : 1);
