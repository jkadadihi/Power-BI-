/*
 * Tests the real readSprCommentary() out of amkad_ar_report.ts against mock
 * SPR tables shaped exactly like the generated sheets (SPRIssues_<CC> /
 * SPRForecast_<CC>, per-country, pre-seeded with blank issue rows).
 *
 * The function is nested inside main(), so rather than mocking the whole
 * report we extract just that function's source and run it with the helpers
 * it closes over (workbook, text, getTableValuesSafe) supplied as mocks. That
 * exercises the actual shipped logic instead of a reimplementation of it.
 */
const fs = require("fs");
const path = require("path");

const src = fs.readFileSync(
  path.join(__dirname, "..", "office_scripts", "amkad_ar_report.ts"), "utf8");

// pull out `function readSprCommentary() { ... }` by brace matching
const start = src.indexOf("function readSprCommentary(");
if (start === -1) throw new Error("readSprCommentary not found");
let depth = 0, i = src.indexOf("{", start), end = -1;
for (let p = i; p < src.length; p++) {
  if (src[p] === "{") depth++;
  else if (src[p] === "}") { depth--; if (depth === 0) { end = p + 1; break; } }
}
let fnSrc = src.slice(start, end);
// strip TypeScript annotations that plain node cannot parse
fnSrc = fnSrc
  .replace(/function readSprCommentary\(\)[^{]*\{/, "function readSprCommentary() {")
  .replace(/:\s*ExcelScript\.Table\[\]/g, "")
  .replace(/let tables\s*=\s*\[\];/, "let tables = [];")
  .replace(/type SprRecord = \{[\s\S]*?\};/, "")
  .replace(/const byKey:[^=]*=/, "const byKey =")
  .replace(/function keyFor\(month: string, country: string, customer: string\): string/,
           "function keyFor(month, country, customer)")
  .replace(/function recordFor\(month: string, country: string, customer: string\): SprRecord/,
           "function recordFor(month, country, customer)")
  .replace(/const idxOf = \(label: string\): number =>/g, "const idxOf = (label) =>")
  .replace(/\(row: \(string \| number \| boolean\)\[\], i: number\): string =>/g, "(row, i) =>")
  .replace(/\(row: \(string \| number \| boolean\)\[\], i: number\): number =>/g, "(row, i) =>")
  .replace(/const out:[^=]*=/, "const out =")
  .replace(/let name = "";/, 'let name = "";')
  .replace(/: string\b/g, "").replace(/: number\b/g, "");

/* ---------- mocks ---------- */
function mkTable(name, headers, rows) {
  return {
    getName: () => name,
    getHeaderRowRange: () => ({ getValues: () => [headers] }),
    _rows: rows
  };
}
const ISSUE_H = ["Month", "Country", "Customer", "Issue", "Action"];
const FC_H = ["Month", "Country", "Customer", "Payments In House Pending",
  "Expected Payments", "Forecasted Overdue EOM", "Forecasted GT60 EOM",
  "Forecasted GT90 EOM"];

const tables = [
  // US: Nokia has 2 filled issues + 1 blank seeded row; Amazon all blank
  mkTable("SPRIssues_US", ISSUE_H, [
    ["2026-07", "US", "Nokia", "80k >90 EOM Pending Approval", "Escalated to Scott."],
    ["2026-07", "US", "Nokia", "48K >90 EOM Payment Debit", "Call set up this week."],
    ["2026-07", "US", "Nokia", "", ""],
    ["2026-07", "US", "Amazon", "", ""],
    ["2026-07", "US", "Amazon", "", ""],
    ["", "", "", "", ""],
  ]),
  mkTable("SPRForecast_US", FC_H, [
    ["2026-07", "US", "Nokia", 0, 85000, 247000, 647000, 223000],
    ["2026-07", "US", "Amazon", "", "", "", "", ""],
  ]),
  // MX: forecast only, no issues typed
  mkTable("SPRIssues_MX", ISSUE_H, [["2026-07", "MX", "Nokia", "", ""]]),
  mkTable("SPRForecast_MX", FC_H, [["2026-07", "MX", "Nokia", 1000, 0, 0, 0, 0]]),
  // an unrelated table must be ignored
  mkTable("RawData", ["Month", "Country", "Customer"], [["2026-07", "US", "Nokia"]]),
];

const workbook = { getTables: () => tables };
const text = v => (v === null || v === undefined) ? "" : v.toString().trim();
const getTableValuesSafe = t => t._rows;

const readSprCommentary = new Function(
  "workbook", "text", "getTableValuesSafe",
  fnSrc + "\nreturn readSprCommentary();"
);

/* ---------- assertions ---------- */
let fails = 0;
function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) { fails++; console.log("FAIL  " + label + "\n   got " + JSON.stringify(actual) + "\n  want " + JSON.stringify(expected)); }
  else console.log("ok    " + label);
}

const out = readSprCommentary(workbook, text, getTableValuesSafe);
const find = (c, cu) => out.find(r => r.country === c && r.customer === cu);

check("only records with content are returned", out.length, 2);
check("blank-only customer (Amazon) omitted", find("US", "Amazon"), undefined);
check("unrelated RawData table ignored",
      out.some(r => r.customer === "Nokia" && r.country === "US"), true);

const nokiaUS = find("US", "Nokia");
check("US Nokia: blank seeded issue row skipped", nokiaUS.issues.length, 2);
check("US Nokia: first issue text", nokiaUS.issues[0].issue, "80k >90 EOM Pending Approval");
check("US Nokia: first action text", nokiaUS.issues[0].action, "Escalated to Scott.");
check("US Nokia: forecast joined from the other table", nokiaUS.expectedPayments, 85000);
check("US Nokia: forecastGT60", nokiaUS.forecastGT60, 647000);
check("US Nokia: blank forecast cell reads as 0", nokiaUS.pendingApplication, 0);
check("US Nokia: month", nokiaUS.month, "2026-07");

const nokiaMX = find("MX", "Nokia");
check("MX Nokia kept: forecast only, no issues", nokiaMX.issues.length, 0);
check("MX Nokia forecast value", nokiaMX.pendingApplication, 1000);
check("same customer name in two countries stays separate",
      nokiaUS.country + "/" + nokiaMX.country, "US/MX");

console.log(fails === 0 ? "\nALL TESTS PASSED" : "\n" + fails + " FAILURES");
process.exit(fails === 0 ? 0 : 1);
