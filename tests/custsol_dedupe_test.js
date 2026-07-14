/*
 * Power Query's Table.Distinct(table, keys) keeps the FIRST row it encounters
 * per key and silently drops later ones. This test models that exact
 * semantic in JS to verify custsol_rawdata.m's dedupe is wired the right way
 * around: TodayRows must come first in Table.Combine, so a same-day re-run
 * keeps freshly re-read source data instead of a stale history copy.
 *
 * This can't execute the actual M code (no local M engine), so it's a
 * behavioral check on the ordering logic, not the M file itself. Re-verify
 * by eye against power_query/custsol_rawdata.m if that file's Combined/Result
 * lines ever change.
 */
function tableDistinct(rows, keyFields) {
  const seen = new Set();
  const out = [];
  for (const row of rows) {
    const key = keyFields.map(f => row[f]).join("|");
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(row);
  }
  return out;
}

let failures = 0;
function check(label, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) { failures++; console.log("FAIL  " + label + "  got=" + JSON.stringify(actual) + " want=" + JSON.stringify(expected)); }
  else console.log("ok    " + label);
}

const staleHistoryRow = { Date: "07/14/2026", Country: "CA", Customer: "Amazon", "Total AR": 100 };
const freshTodayRow    = { Date: "07/14/2026", Country: "CA", Customer: "Amazon", "Total AR": 105 };
const untouchedHistory  = { Date: "07/13/2026", Country: "CA", Customer: "Amazon", "Total AR": 90 };

// Mirrors: Combined = Table.Combine({TodayRows, RawDataTable})
const combinedCorrectOrder = [freshTodayRow, staleHistoryRow, untouchedHistory];
const resultCorrect = tableDistinct(combinedCorrectOrder, ["Date", "Country", "Customer"]);
check("correct order: fresh today value wins on same-day re-run", resultCorrect[0]["Total AR"], 105);
check("correct order: untouched history row still present", resultCorrect.length, 2);

// What the bug (RawDataTable first) would have produced -- documents why the fix mattered.
const combinedBuggyOrder = [staleHistoryRow, freshTodayRow, untouchedHistory];
const resultBuggy = tableDistinct(combinedBuggyOrder, ["Date", "Country", "Customer"]);
check("buggy order would have kept the STALE value (regression guard)", resultBuggy[0]["Total AR"], 100);

console.log(failures === 0 ? "\nALL TESTS PASSED" : "\n" + failures + " FAILURES");
process.exit(failures === 0 ? 0 : 1);
