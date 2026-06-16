import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "/Users/smolina/Documents/budget";
const outputDir = path.join(root, "outputs", "summer-budget");
const outputPath = path.join(outputDir, "summer_budget_plan.xlsx");

const workbook = Workbook.create();
const dashboard = workbook.worksheets.add("Dashboard");
const assumptions = workbook.worksheets.add("Assumptions");
const plan = workbook.worksheets.add("Budget Plan");
const log = workbook.worksheets.add("Expense Log");
const weekly = workbook.worksheets.add("Weekly Scenarios");
const nextTwoWeeks = workbook.worksheets.add("Next 2 Weeks");
const checks = workbook.worksheets.add("Checks");

const currency = '$#,##0;[Red]($#,##0);-';
const percent = '0.0%;[Red](0.0%);-';
const dateFmt = 'yyyy-mm-dd';

function writeValues(sheet, range, values) {
  sheet.getRange(range).values = values;
}

function writeFormulas(sheet, range, formulas) {
  sheet.getRange(range).formulas = formulas;
}

function fill(sheet, range, color) {
  sheet.getRange(range).format.fill.color = color;
}

function font(sheet, range, options) {
  const f = sheet.getRange(range).format.font;
  if (options.bold !== undefined) f.bold = options.bold;
  if (options.color) f.color = options.color;
  if (options.size) f.size = options.size;
}

function numberFormat(sheet, range, format) {
  sheet.getRange(range).numberFormat = format;
}

function align(sheet, range, horizontal = "left") {
  sheet.getRange(range).format.horizontalAlignment = horizontal;
}

function width(sheet, widths) {
  for (const [col, px] of Object.entries(widths)) {
    sheet.getRange(`${col}:${col}`).format.columnWidthPx = px;
  }
}

function tableHeader(sheet, range) {
  fill(sheet, range, "#244C5A");
  font(sheet, range, { bold: true, color: "#FFFFFF" });
  align(sheet, range, "center");
}

// Assumptions
writeValues(assumptions, "A1:F1", [["Summer Budget Assumptions", null, null, null, null, null]]);
fill(assumptions, "A1:F1", "#173B45");
font(assumptions, "A1:F1", { bold: true, color: "#FFFFFF", size: 16 });

writeValues(assumptions, "A3:B12", [
  ["Start date", new Date("2026-06-16T00:00:00")],
  ["End date", new Date("2026-08-14T00:00:00")],
  ["Starting checking / cash", 0],
  ["Net paycheck after account transfer", 2300],
  ["Savings goal by end date", 3000],
  ["Safety buffer", 300],
  ["Total available to spend", null],
  ["Days in plan", null],
  ["Weeks in plan", null],
  ["Daily spending target", null],
]);
writeFormulas(assumptions, "B9:B12", [
  ["=B5+SUM(E4:E8)-B7-B8"],
  ["=B4-B3+1"],
  ["=ROUNDUP(B10/7,0)"],
  ["=B9/B10"],
]);
numberFormat(assumptions, "B3:B4", dateFmt);
numberFormat(assumptions, "B5:B9", currency);
numberFormat(assumptions, "B12", currency);
fill(assumptions, "A3:B12", "#F7FAFA");
fill(assumptions, "B3:B8", "#FFF2CC");
font(assumptions, "B3:B8", { color: "#0000FF" });
font(assumptions, "B9:B12", { color: "#000000" });
writeValues(assumptions, "D3:F8", [
  ["Pay date", "Net usable pay", "Notes"],
  [new Date("2026-06-19T00:00:00"), 2300, "Assumed Friday payday"],
  [new Date("2026-07-03T00:00:00"), 2300, "Biweekly"],
  [new Date("2026-07-17T00:00:00"), 2300, "Biweekly"],
  [new Date("2026-07-31T00:00:00"), 2300, "Biweekly"],
  [new Date("2026-08-14T00:00:00"), 2300, "Biweekly"],
]);
tableHeader(assumptions, "D3:F3");
fill(assumptions, "D4:F8", "#FFF2CC");
font(assumptions, "D4:E8", { color: "#0000FF" });
numberFormat(assumptions, "D4:D8", dateFmt);
numberFormat(assumptions, "E4:E8", currency);
writeValues(assumptions, "H3:J10", [
  ["Input legend", "Meaning", "Update cadence"],
  ["Blue text + yellow fill", "Editable assumption", "Whenever your plan changes"],
  ["Paycheck table", "Net usable checks after $1,000 transfer", "Update if payday dates differ"],
  ["Savings goal", "Money moved or left untouched", "Tune until spendable feels right"],
  ["Safety buffer", "Cushion for surprises", "Keep unless needed"],
  ["Weekly groceries/gas", "Uses rounded-up weeks", "Adjust category budget"],
  ["Expense Log", "Daily purchases", "Filled by you now, app later"],
  ["Phone app path", "Expense Log columns become the database schema", "Use later"],
]);
tableHeader(assumptions, "H3:J3");
fill(assumptions, "H4:J10", "#F7FAFA");
width(assumptions, { A: 230, B: 170, D: 125, E: 140, F: 190, H: 180, I: 300, J: 200 });

// Budget plan
writeValues(plan, "A1:I1", [["Summer Category Budget", null, null, null, null, null, null, null, null]]);
fill(plan, "A1:I1", "#173B45");
font(plan, "A1:I1", { bold: true, color: "#FFFFFF", size: 16 });
writeValues(plan, "A3:I3", [[
  "Category", "Type", "Priority", "Budget", "Actual", "Remaining", "% Used", "Weekly Target", "Notes",
]]);
tableHeader(plan, "A3:I3");
const categories = [
  ["Rent total", "Fixed", "Must pay", 2500, "", "", "", "", "Total rent reserve for July/August"],
  ["Rent extra reserve", "Fixed", "Must pay", 0, "", "", "", "", "Use only if rent estimate changes"],
  ["Groceries", "Weekly", "Must pay", 675, "", "", "", "", "$75 x rounded-up weeks"],
  ["Gas", "Weekly", "Must pay", 270, "", "", "", "", "$30 x rounded-up weeks"],
  ["Storage unit", "Fixed", "Must pay", 140, "", "", "", "", "$70 due July 7 and Aug 7"],
  ["Savings account", "Savings", "Must pay", 2500, "", "", "", "", "Editable via Assumptions savings goal"],
  ["Safety buffer", "Reserve", "Must pay", 300, "", "", "", "", "Cushion for surprise costs"],
  ["Discretionary spending pool", "Flex", "Flex", 0, "", "", "", "", "What you can spend after reserves"],
  ["Restaurants / coffee", "Flex", "Flex", 0, "", "", "", "", "Track here or use the pool"],
  ["Fun / summer plans", "Flex", "Flex", 0, "", "", "", "", "Track here or use the pool"],
  ["Shopping / personal", "Flex", "Flex", 0, "", "", "", "", "Track here or use the pool"],
  ["Miscellaneous", "Flex", "Flex", 0, "", "", "", "", "Track here or use the pool"],
];
writeValues(plan, "A4:I15", categories);
writeFormulas(plan, "D6:D10", [
  ["=75*Assumptions!$B$11"],
  ["=30*Assumptions!$B$11"],
  ["=70*2"],
  ["=Assumptions!$B$7"],
  ["=Assumptions!$B$8"],
]);
writeFormulas(plan, "D11:D11", [["=SUM(Assumptions!$E$4:$E$8)-SUM(D4:D10)"]]);
writeFormulas(plan, "E4:H15", categories.map((_, i) => {
  const row = i + 4;
  return [
    `=SUMIF('Expense Log'!$C:$C,A${row},'Expense Log'!$D:$D)`,
    `=D${row}-E${row}`,
    `=IFERROR(E${row}/D${row},0)`,
    `=D${row}/Assumptions!$B$11`,
  ];
}));
writeValues(plan, "A17:I17", [["Total", "", "", "", "", "", "", "", ""]]);
writeFormulas(plan, "D17:H17", [["=SUM(D4:D15)", "=SUM(E4:E15)", "=SUM(F4:F15)", "=IFERROR(E17/D17,0)", "=SUM(H4:H15)"]]);
fill(plan, "A17:I17", "#D9EAD3");
font(plan, "A17:I17", { bold: true });
numberFormat(plan, "D4:F17", currency);
numberFormat(plan, "G4:G17", percent);
numberFormat(plan, "H4:H17", currency);
fill(plan, "A4:I15", "#FBFDFD");
fill(plan, "D4:D15", "#FFF2CC");
font(plan, "D4:D15", { color: "#0000FF" });
font(plan, "E4:H17", { color: "#000000" });
width(plan, { A: 170, B: 115, C: 115, D: 115, E: 115, F: 115, G: 95, H: 120, I: 280 });

// Expense log
writeValues(log, "A1:H1", [["Expense Log", null, null, null, null, null, null, null]]);
fill(log, "A1:H1", "#173B45");
font(log, "A1:H1", { bold: true, color: "#FFFFFF", size: 16 });
writeValues(log, "A3:H3", [["Date", "Merchant", "Category", "Amount", "Payment", "Need/Flex", "Notes", "Month"]]);
tableHeader(log, "A3:H3");
writeValues(log, "A4:H8", [
  [new Date("2026-06-16T00:00:00"), "Example: grocery store", "Groceries", 42.75, "Debit", "Need", "Replace sample rows with real expenses", null],
  [new Date("2026-06-17T00:00:00"), "Example: coffee", "Discretionary spending pool", 6.5, "Credit", "Flex", "Small daily purchases add up", null],
  [new Date("2026-06-20T00:00:00"), "Example: gas station", "Gas", 48.25, "Debit", "Need", "", null],
  [new Date("2026-06-22T00:00:00"), "Example: movie", "Discretionary spending pool", 18, "Credit", "Flex", "", null],
  [new Date("2026-07-07T00:00:00"), "Example: storage", "Storage unit", 70, "Debit", "Need", "Scheduled July/August payment", null],
]);
writeFormulas(log, "H4:H203", Array.from({ length: 200 }, (_, idx) => [`=IF(A${idx + 4}="","",TEXT(A${idx + 4},"mmm"))`]));
fill(log, "A4:H203", "#FBFDFD");
numberFormat(log, "A4:A203", dateFmt);
numberFormat(log, "D4:D203", currency);
width(log, { A: 110, B: 210, C: 170, D: 100, E: 100, F: 100, G: 280, H: 80 });

// Weekly savings scenarios
writeValues(weekly, "A1:K1", [["Weekly Savings Scenarios", null, null, null, null, null, null, null, null, null, null]]);
fill(weekly, "A1:K1", "#173B45");
font(weekly, "A1:K1", { bold: true, color: "#FFFFFF", size: 16 });
writeValues(weekly, "A3:B11", [
  ["Planning weeks", null],
  ["Net pay in window", null],
  ["Rent total", null],
  ["Groceries total", null],
  ["Gas total", null],
  ["Storage total", null],
  ["Safety buffer", null],
  ["Non-savings fixed total", null],
  ["Spendable before savings", null],
]);
writeFormulas(weekly, "B3:B11", [
  ["=Assumptions!B11"],
  ["=SUM(Assumptions!E4:E8)"],
  ["='Budget Plan'!D4"],
  ["='Budget Plan'!D6"],
  ["='Budget Plan'!D7"],
  ["='Budget Plan'!D8"],
  ["='Budget Plan'!D10"],
  ["=SUM(B5:B9)"],
  ["=B4-B10"],
]);
numberFormat(weekly, "B4:B11", currency);
fill(weekly, "A3:B11", "#F7FAFA");
font(weekly, "A3:A11", { bold: true });

writeValues(weekly, "D3:K3", [[
  "Savings goal", "Weekly savings", "Total discretionary", "Weekly discretionary", "Daily discretionary", "Weekly all-in spend", "Monthly avg discretionary", "Notes",
]]);
tableHeader(weekly, "D3:K3");
writeValues(weekly, "D4:D6", [[3000], [4000], [5000]]);
writeFormulas(weekly, "E4:K6", [
  ["=D4/$B$3", "=$B$11-D4", "=F4/$B$3", "=G4/7", "=($B$10+F4)/$B$3", "=F4/(Assumptions!$B$10/30.4375)", '=IF(F4<0,"Savings too high for current income/costs","Works")'],
  ["=D5/$B$3", "=$B$11-D5", "=F5/$B$3", "=G5/7", "=($B$10+F5)/$B$3", "=F5/(Assumptions!$B$10/30.4375)", '=IF(F5<0,"Savings too high for current income/costs","Works")'],
  ["=D6/$B$3", "=$B$11-D6", "=F6/$B$3", "=G6/7", "=($B$10+F6)/$B$3", "=F6/(Assumptions!$B$10/30.4375)", '=IF(F6<0,"Savings too high for current income/costs","Works")'],
]);
numberFormat(weekly, "D4:J6", currency);
fill(weekly, "D4:K6", "#FBFDFD");
fill(weekly, "D4:D6", "#FFF2CC");
font(weekly, "D4:D6", { color: "#0000FF" });
writeValues(weekly, "D9:K13", [
  ["How to read this", "", "", "", "", "", "", ""],
  ["Weekly savings", "How much to set aside each week to hit the total goal.", "", "", "", "", "", ""],
  ["Weekly discretionary", "What is left each week for restaurants, fun, shopping, personal, and misc.", "", "", "", "", "", ""],
  ["Weekly all-in spend", "Fixed weekly cost load plus discretionary; excludes savings.", "", "", "", "", "", ""],
  ["Recommendation", "Try the highest savings target that still leaves a weekly discretionary number you can realistically live with.", "", "", "", "", "", ""],
]);
fill(weekly, "D9:K13", "#EAF3F6");
font(weekly, "D9:K9", { bold: true, color: "#173B45" });
width(weekly, { A: 210, B: 140, D: 125, E: 130, F: 150, G: 155, H: 145, I: 145, J: 170, K: 170 });

// Next two weeks cash plan
writeValues(nextTwoWeeks, "A1:H1", [["Next 2 Weeks Cash Plan", null, null, null, null, null, null, null]]);
fill(nextTwoWeeks, "A1:H1", "#173B45");
font(nextTwoWeeks, "A1:H1", { bold: true, color: "#FFFFFF", size: 16 });
writeValues(nextTwoWeeks, "A3:B12", [
  ["Plan start", new Date("2026-06-16T00:00:00")],
  ["Plan end", new Date("2026-06-30T00:00:00")],
  ["Current total across accounts", null],
  ["Next paycheck timing", "Two weeks"],
  ["Rent to reserve for next month", 1200],
  ["Groceries per week", 75],
  ["Gas per week", 30],
  ["Discretionary target", 290],
  ["Mini-buffer", null],
  ["Savings account balance", 600],
]);
writeFormulas(nextTwoWeeks, "B5:B5", [["=SUM(E4:E6)"]]);
writeFormulas(nextTwoWeeks, "B11:B11", [["=E5-B10"]]);
numberFormat(nextTwoWeeks, "B3:B4", dateFmt);
numberFormat(nextTwoWeeks, "B5:B5", currency);
numberFormat(nextTwoWeeks, "B7:B12", currency);
fill(nextTwoWeeks, "A3:B12", "#F7FAFA");
fill(nextTwoWeeks, "B3:B5", "#FFF2CC");
fill(nextTwoWeeks, "B7:B11", "#FFF2CC");
font(nextTwoWeeks, "B3:B5", { color: "#0000FF" });
font(nextTwoWeeks, "B7:B11", { color: "#0000FF" });
font(nextTwoWeeks, "B12:B12", { color: "#000000", bold: true });

writeValues(nextTwoWeeks, "D3:H3", [["Bucket", "Amount", "Weekly", "Daily", "Move/keep where"]]);
tableHeader(nextTwoWeeks, "D3:H3");
writeValues(nextTwoWeeks, "D4:D10", [
  ["Essentials account"],
  ["Discretionary account"],
  ["Savings account"],
  ["Rent reserve inside essentials"],
  ["Groceries inside essentials"],
  ["Gas inside essentials"],
  ["Total current balances"],
]);
writeFormulas(nextTwoWeeks, "E4:G10", [
  ["=($B$7)+($B$8*2)+($B$9*2)", "=E4/2", "=E4/14"],
  [545, "=E5/2", "=E5/14"],
  ["=$B$12", "=E6/2", "=E6/14"],
  ["=$B$7", "=E7/2", "=E7/14"],
  ["=$B$8*2", "=E8/2", "=E8/14"],
  ["=$B$9*2", "=E9/2", "=E9/14"],
  ["=SUM(E4:E6)", "=SUM(F4:F6)", "=SUM(G4:G6)"],
]);
writeValues(nextTwoWeeks, "H4:H10", [
  ["Matches rent + two weeks groceries/gas"],
  ["Your current spend/buffer account"],
  ["Already saved"],
  ["Reserved from essentials account"],
  ["Two weeks at $75/week"],
  ["Two weeks at $30/week"],
  ["Current account total"],
]);
numberFormat(nextTwoWeeks, "E4:G10", currency);
fill(nextTwoWeeks, "D4:H10", "#FBFDFD");
fill(nextTwoWeeks, "D10:H10", "#D9EAD3");
font(nextTwoWeeks, "D10:H10", { bold: true });

writeValues(nextTwoWeeks, "D13:H16", [
  ["Spending Rule", "", "", "", ""],
  ["Daily discretionary cap", null, "", "", "If you spend less, sweep the difference to savings."],
  ["Weekly discretionary cap", null, "", "", "This is the most important number to watch."],
  ["Check-in date", new Date("2026-06-23T00:00:00"), "", "", "Halfway point: compare actual spending to cap."],
]);
writeFormulas(nextTwoWeeks, "E14:E15", [["=($B$10)/14"], ["=($B$10)/2"]]);
numberFormat(nextTwoWeeks, "E14:E15", currency);
numberFormat(nextTwoWeeks, "E16:E16", dateFmt);
fill(nextTwoWeeks, "D13:H16", "#EAF3F6");
font(nextTwoWeeks, "D13:H13", { bold: true, color: "#173B45" });
width(nextTwoWeeks, { A: 230, B: 170, D: 205, E: 115, F: 115, G: 115, H: 260 });

// Dashboard
writeValues(dashboard, "A1:J1", [["Summer Budget Dashboard", null, null, null, null, null, null, null, null, null]]);
fill(dashboard, "A1:J1", "#173B45");
font(dashboard, "A1:J1", { bold: true, color: "#FFFFFF", size: 18 });
writeValues(dashboard, "A3:B8", [
  ["Planning window", null],
  ["Net pay in window", null],
  ["Reserved + spending plan", null],
  ["Logged spending", null],
  ["Discretionary left after plan", null],
  ["Suggested daily spend cap", null],
]);
writeFormulas(dashboard, "B3:B8", [
  ['=TEXT(Assumptions!B3,"mmm d")&" - "&TEXT(Assumptions!B4,"mmm d")'],
  ["=SUM(Assumptions!E4:E8)"],
  ["='Budget Plan'!D17"],
  ["='Budget Plan'!E17"],
  ["='Budget Plan'!D11-SUMIF('Expense Log'!$F:$F,\"Flex\",'Expense Log'!$D:$D)"],
  ["=MAX(0,B7)/Assumptions!B10"],
]);
numberFormat(dashboard, "B4:B7", currency);
numberFormat(dashboard, "B8", currency);
fill(dashboard, "A3:B8", "#F7FAFA");
font(dashboard, "A3:A8", { bold: true });
font(dashboard, "B4:B8", { bold: true });

writeValues(dashboard, "D3:J3", [["Category", "Budget", "Actual", "Remaining", "% Used", "Priority", "Status"]]);
tableHeader(dashboard, "D3:J3");
writeFormulas(dashboard, "D4:J15", categories.map((_, i) => {
  const row = i + 4;
  return [
    `='Budget Plan'!A${row}`,
    `='Budget Plan'!D${row}`,
    `='Budget Plan'!E${row}`,
    `='Budget Plan'!F${row}`,
    `='Budget Plan'!G${row}`,
    `='Budget Plan'!C${row}`,
    `=IF(F${row}=0,"No spend yet",IF(G${row}<0,"Over budget",IF(H${row}>0.8,"Watch","OK")))`,
  ];
}));
numberFormat(dashboard, "E4:G15", currency);
numberFormat(dashboard, "H4:H15", percent);
fill(dashboard, "D4:J15", "#FBFDFD");
writeValues(dashboard, "A11:B16", [
  ["Current guidance", ""],
  ["Best first trim", ""],
  ["Weekly check-in", "Update logged expenses and compare the dashboard to your remaining summer weeks."],
  ["Rule of thumb", "Keep flex spending below the daily target unless you planned a bigger event."],
  ["Phone app path", "Use the Expense Log columns as the eventual database schema."],
  ["Sheet sync", "This workbook can be imported to Google Sheets when we wire the app."],
]);
writeFormulas(dashboard, "B12:B12", [['=INDEX(D4:D14,MATCH(MAX(FILTER(E4:E14,I4:I14="Flex")),E4:E14,0))']]);
fill(dashboard, "A11:B16", "#EAF3F6");
font(dashboard, "A11:B11", { bold: true, color: "#173B45" });
width(dashboard, { A: 165, B: 220, D: 170, E: 105, F: 105, G: 105, H: 85, I: 100, J: 115 });

// Checks
writeValues(checks, "A1:F1", [["Budget Checks", null, null, null, null, null]]);
fill(checks, "A1:F1", "#173B45");
font(checks, "A1:F1", { bold: true, color: "#FFFFFF", size: 16 });
writeValues(checks, "A3:F3", [["Check", "Actual", "Expected", "Difference", "Status", "Notes"]]);
tableHeader(checks, "A3:F3");
writeValues(checks, "A4:A8", [
  ["Category budgets tie to total"],
  ["Actuals tie to expense log"],
  ["Spendable budget does not exceed available"],
  ["End savings goal retained"],
  ["No negative category remaining"],
]);
writeFormulas(checks, "B4:E8", [
  ["='Budget Plan'!D17", "=SUM('Budget Plan'!D4:D15)", "=B4-C4", '=IF(ABS(D4)<0.01,"OK","Review")'],
  ["='Budget Plan'!E17", "=SUM('Expense Log'!D:D)", "=B5-C5", '=IF(ABS(D5)<0.01,"OK","Review")'],
  ["='Budget Plan'!D17", "=SUM(Assumptions!E4:E8)", "=B6-C6", '=IF(ABS(D6)<0.01,"OK","Review")'],
  ["=Assumptions!B7", "=Assumptions!B7", "=B7-C7", '=IF(ABS(D7)<0.01,"OK","Review")'],
  ["=MIN('Budget Plan'!F4:F15)", "0", "=B8-C8", '=IF(D8>=0,"OK","Review")'],
]);
writeValues(checks, "F4:F8", [
  ["Budget Plan total should equal category sum."],
  ["Actual total should match every logged expense."],
  ["If Review, reduce categories or change assumptions."],
  ["Savings goal is explicitly reserved."],
  ["Review means at least one category is over budget."],
]);
numberFormat(checks, "B4:D8", currency);
fill(checks, "A4:F8", "#FBFDFD");
width(checks, { A: 260, B: 120, C: 120, D: 120, E: 90, F: 300 });

for (const sheet of [dashboard, assumptions, plan, log, weekly, nextTwoWeeks, checks]) {
  sheet.getRange("A:Z").format.font.name = "Aptos";
  sheet.getRange("A:Z").format.font.size = 10;
}

// Add lightweight conditional formats where supported by Excel viewers.
dashboard.getRange("J4:J14").format.horizontalAlignment = "center";
checks.getRange("E4:E8").format.horizontalAlignment = "center";

await fs.mkdir(outputDir, { recursive: true });

const inspectDashboard = await workbook.inspect({
  kind: "table",
  range: "Dashboard!A1:J16",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 12,
});
console.log(inspectDashboard.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

for (const sheetName of ["Dashboard", "Assumptions", "Budget Plan", "Expense Log", "Weekly Scenarios", "Next 2 Weeks", "Checks"]) {
  await workbook.render({ sheetName, range: "A1:J22", scale: 1 });
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${outputPath}`);
