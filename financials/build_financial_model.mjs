import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const repoRoot = "/Users/dan/Documents/GitHub/Wallaboo";
const outputDir = path.join(repoRoot, "outputs");
const qaDir = "/private/tmp/wallaboo-model-qa";
const outputPath = path.join(outputDir, "Wallaboo_Financial_Model_v0.1.xlsx");

const colors = {
  navy: "#15324B",
  navy2: "#234E70",
  teal: "#2A7F75",
  gold: "#D6A84B",
  paleGold: "#FFF4D6",
  paleBlue: "#EAF2F8",
  paleTeal: "#E8F5F2",
  paleRed: "#FCE8E6",
  paleGray: "#F3F5F7",
  midGray: "#D4DCE3",
  darkGray: "#44515C",
  white: "#FFFFFF",
  black: "#000000",
  inputBlue: "#0000FF",
  linkGreen: "#008000",
  failRed: "#B42318",
  passGreen: "#1E7A46",
};

const years = [2027, 2028, 2029, 2030, 2031];
const cases = {
  Conservative: {
    units: [1000, 2500, 5000, 8000, 12000],
    msrp: [39, 39, 40, 41, 42],
    dtc: [0.70, 0.55, 0.45, 0.35, 0.30],
    retail: [0.20, 0.20, 0.20, 0.20, 0.20],
    landed: [14.00, 13.00, 12.00, 11.50, 11.00],
    development: [28000, 32000, 35000, 38000, 42000],
    marketing: [10000, 14000, 18000, 22000, 26000],
    overhead: [9000, 12000, 15000, 19000, 24000],
    founderComp: [0, 18000, 30000, 42000, 54000],
  },
  Base: {
    units: [2000, 5000, 10000, 17000, 26000],
    msrp: [39, 39, 40, 41, 42],
    dtc: [0.75, 0.60, 0.50, 0.42, 0.35],
    retail: [0.15, 0.20, 0.20, 0.20, 0.20],
    landed: [12.50, 11.50, 10.50, 9.75, 9.25],
    development: [30000, 38000, 45000, 52000, 60000],
    marketing: [14000, 22000, 32000, 45000, 60000],
    overhead: [10000, 15000, 22000, 32000, 45000],
    founderComp: [0, 24000, 48000, 72000, 90000],
  },
  Upside: {
    units: [3500, 9000, 18000, 32000, 50000],
    msrp: [39, 40, 41, 42, 43],
    dtc: [0.80, 0.68, 0.58, 0.50, 0.42],
    retail: [0.15, 0.17, 0.17, 0.18, 0.18],
    landed: [11.50, 10.25, 9.25, 8.50, 8.00],
    development: [34000, 45000, 60000, 78000, 100000],
    marketing: [20000, 35000, 55000, 80000, 115000],
    overhead: [12000, 18000, 28000, 45000, 65000],
    founderComp: [0, 30000, 60000, 96000, 120000],
  },
};

const workbook = Workbook.create();
const cover = workbook.worksheets.add("Cover");
const assumptions = workbook.worksheets.add("Assumptions");
const unitEcon = workbook.worksheets.add("Unit Economics");
const operating = workbook.worksheets.add("Operating Model");
const summary = workbook.worksheets.add("Scenario Summary");
const cash = workbook.worksheets.add("Cash & Funding");
const checks = workbook.worksheets.add("Checks");
const sources = workbook.worksheets.add("Sources");

function mergeTitle(sheet, address, text, fill = colors.navy) {
  sheet.getRange(address).merge();
  const cell = sheet.getRange(address.split(":")[0]);
  cell.values = [[text]];
  sheet.getRange(address).format = {
    fill,
    font: { bold: true, color: colors.white, size: 18 },
    verticalAlignment: "center",
    horizontalAlignment: "left",
  };
  sheet.getRange(address).format.rowHeight = 34;
}

function sectionHeader(sheet, address, text) {
  sheet.getRange(address).merge();
  const cell = sheet.getRange(address.split(":")[0]);
  cell.values = [[text]];
  sheet.getRange(address).format = {
    fill: colors.navy2,
    font: { bold: true, color: colors.white, size: 11 },
    verticalAlignment: "center",
  };
  sheet.getRange(address).format.rowHeight = 23;
}

function styleTableHeader(range) {
  range.format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: colors.navy },
  };
  range.format.rowHeight = 28;
}

function styleSubheader(range) {
  range.format = {
    fill: colors.paleBlue,
    font: { bold: true, color: colors.navy },
    borders: { bottom: { style: "thin", color: colors.midGray } },
  };
}

function styleCurrency(range) {
  range.format.numberFormat = '"$"#,##0;[Red]("$"#,##0);-';
  range.format.horizontalAlignment = "right";
}

function styleCurrencyOne(range) {
  range.format.numberFormat = '"$"0.00;[Red]("$"0.00);-';
  range.format.horizontalAlignment = "right";
}

function stylePercent(range) {
  range.format.numberFormat = "0.0%;[Red](0.0%);-";
  range.format.horizontalAlignment = "right";
}

function styleCount(range) {
  range.format.numberFormat = "#,##0;[Red](#,##0);-";
  range.format.horizontalAlignment = "right";
}

function styleInput(range, numberFormat = null) {
  range.format.fill = colors.paleGold;
  range.format.font = { color: colors.inputBlue };
  if (numberFormat) range.format.numberFormat = numberFormat;
}

function styleLinkedFormula(range, numberFormat = null) {
  range.format.font = { color: colors.linkGreen };
  if (numberFormat) range.format.numberFormat = numberFormat;
}

function borderTotal(range) {
  range.format.borders = { top: { style: "thin", color: colors.navy } };
  range.format.font = { bold: true, color: colors.black };
}

function addSourceComment(sheet, cellAddress, text) {
  workbook.comments.addThread({ cell: sheet.getRange(cellAddress) }, text);
}

for (const sheet of workbook.worksheets.items) {
  sheet.showGridLines = false;
}

await workbook.comments.setSelf({ displayName: "Dan Trezise" });

// Cover
mergeTitle(cover, "A1:H2", "MR WALLABOO GAMES — FINANCIAL MODEL");
cover.getRange("A3:H3").merge();
cover.getRange("A3").values = [["Illustrative five-year operating plan | WORKING v0.1 | 17 Jul 2026"]];
cover.getRange("A3:H3").format = { font: { italic: true, color: colors.darkGray }, fill: colors.paleGray };

sectionHeader(cover, "A5:H5", "DECISION SNAPSHOT");
cover.getRange("A6:B6").merge();
cover.getRange("A6").values = [["Recommendation"]];
cover.getRange("C6:D7").merge();
cover.getRange("C6").values = [["CONDITIONAL GO"]];
cover.getRange("C6:D7").format = { fill: colors.paleGold, font: { bold: true, color: colors.navy, size: 16 }, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "outside", style: "medium", color: colors.gold } };
cover.getRange("E6:F6").merge();
cover.getRange("E6").values = [["Base-case break-even"]];
cover.getRange("G6:H7").merge();
cover.getRange("G6").formulas = [["='Scenario Summary'!F7"]];
cover.getRange("G6:H7").format = { fill: colors.paleTeal, font: { bold: true, color: colors.passGreen, size: 16 }, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "outside", style: "medium", color: colors.teal } };
cover.getRange("A6:B7").format = { fill: colors.paleBlue, font: { bold: true, color: colors.navy }, verticalAlignment: "center" };
cover.getRange("E6:F7").format = { fill: colors.paleBlue, font: { bold: true, color: colors.navy }, verticalAlignment: "center" };

cover.getRange("A9:H9").values = [["KPI", "2027", "2028", "2029", "2030", "2031", "Five-year", "Status"]];
styleTableHeader(cover.getRange("A9:H9"));
cover.getRange("A10:A13").values = [["Net Revenue"], ["EBITDA"], ["Units Sold"], ["Ending Cash"]];
cover.getRange("B10:F10").formulas = [["='Operating Model'!B17", "='Operating Model'!C17", "='Operating Model'!D17", "='Operating Model'!E17", "='Operating Model'!F17"]];
cover.getRange("B11:F11").formulas = [["='Operating Model'!B30", "='Operating Model'!C30", "='Operating Model'!D30", "='Operating Model'!E30", "='Operating Model'!F30"]];
cover.getRange("B12:F12").formulas = [["='Operating Model'!B6", "='Operating Model'!C6", "='Operating Model'!D6", "='Operating Model'!E6", "='Operating Model'!F6"]];
cover.getRange("B13:F13").formulas = [["='Cash & Funding'!B9", "='Cash & Funding'!C9", "='Cash & Funding'!D9", "='Cash & Funding'!E9", "='Cash & Funding'!F9"]];
cover.getRange("G10").formulas = [["=SUM(B10:F10)"]];
cover.getRange("G11").formulas = [["=SUM(B11:F11)"]];
cover.getRange("G12").formulas = [["=SUM(B12:F12)"]];
cover.getRange("G13").formulas = [["=F13"]];
cover.getRange("H10:H13").values = [["Base case"], ["Positive in 2030"], ["Catalog scale"], ["Seed reserve modeled"]];
styleCurrency(cover.getRange("B10:G11"));
styleCount(cover.getRange("B12:G12"));
styleCurrency(cover.getRange("B13:G13"));
cover.getRange("A10:H13").format.borders = { bottom: { style: "thin", color: colors.midGray } };

sectionHeader(cover, "A15:H15", "HOW TO USE THIS MODEL");
cover.getRange("A16:H19").merge();
cover.getRange("A16").values = [["Blue/yellow cells are editable assumptions. Green formulas link to another sheet. Black formulas calculate outputs. The three scenarios are independent planning cases, not probabilities. Update the model only after a flagship bill of materials, supplier quotes, fulfillment quotes, and prelaunch conversion data exist. Crowdfunding receipts are product sales, not financing, and customer-paid shipping is excluded from revenue and expense."]];
cover.getRange("A16:H19").format = { fill: colors.paleGray, font: { color: colors.darkGray }, wrapText: true, verticalAlignment: "top", borders: { preset: "outside", style: "thin", color: colors.midGray } };

sectionHeader(cover, "A21:H21", "COLOR LEGEND");
cover.getRange("A22:B22").merge(); cover.getRange("A22").values = [["Editable assumptions"]]; styleInput(cover.getRange("A22:B22"));
cover.getRange("C22:D22").merge(); cover.getRange("C22").values = [["Cross-sheet links"]]; cover.getRange("C22:D22").format = { fill: colors.white, font: { color: colors.linkGreen } };
cover.getRange("E22:F22").merge(); cover.getRange("E22").values = [["Calculated formulas"]]; cover.getRange("E22:F22").format = { fill: colors.white, font: { color: colors.black } };
cover.getRange("G22:H22").merge(); cover.getRange("G22").values = [["Checks / cautions"]]; cover.getRange("G22:H22").format = { fill: colors.paleRed, font: { color: colors.failRed } };

cover.getRange("A24:H26").merge();
cover.getRange("A24").values = [["Important: This is an illustrative, unaudited planning model—not a valuation, securities offering, or legal/tax/customs opinion. The base case supports a lean publisher after a multi-year build; it does not support two market-rate founder salaries from launch."]];
cover.getRange("A24:H26").format = { fill: colors.paleGold, font: { bold: true, color: colors.navy }, wrapText: true, verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: colors.gold } };
cover.getRange("A1:H26").format.font.name = "Aptos";
cover.getRange("A1:A26").format.columnWidth = 24;
cover.getRange("B1:H26").format.columnWidth = 15;
cover.freezePanes.freezeRows(3);

// Assumptions
mergeTitle(assumptions, "A1:G2", "ASSUMPTIONS & SCENARIO DRIVERS");
assumptions.getRange("A3:G3").merge();
assumptions.getRange("A3").values = [["All hardcodes are illustrative and should be replaced with quotes, test data, or approved management assumptions."]];
assumptions.getRange("A3:G3").format = { fill: colors.paleGray, font: { italic: true, color: colors.darkGray } };
sectionHeader(assumptions, "A5:G5", "GENERAL INPUTS");
assumptions.getRange("A6:D6").values = [["Input", "Value", "Units", "Source / rationale"]];
styleTableHeader(assumptions.getRange("A6:D6"));
const generalInputs = [
  ["Kickstarter platform fee", 0.05, "% of DTC product revenue", "Kickstarter official fee"],
  ["Payment processing", 0.04, "% of DTC product revenue", "Midpoint of official 3%–5% range"],
  ["DTC refund/chargeback reserve", 0.02, "% of DTC product revenue", "Planning assumption"],
  ["DTC pick/pack and materials", 3.25, "$ / DTC unit", "Planning assumption; excludes postage"],
  ["Retail-direct net price", 0.50, "% of MSRP", "Public publisher wholesale terms"],
  ["Distributor net price", 0.40, "% of MSRP", "Public publisher distributor terms"],
  ["Wholesale outbound freight", 0.03, "% of wholesale revenue", "Planning assumption"],
  ["Seed liquidity reserve", 75000, "$", "Milestone-gated planning reserve"],
  ["First production target", 1500, "units", "Panda MOQ / practical offset scale"],
  ["Campaign product-pledge goal", 60000, "$", "Management gate; shipping excluded"],
];
assumptions.getRange("A7:D16").values = generalInputs;
styleInput(assumptions.getRange("B7:B16"));
stylePercent(assumptions.getRange("B7:B9"));
styleCurrencyOne(assumptions.getRange("B10"));
stylePercent(assumptions.getRange("B11:B13"));
styleCurrency(assumptions.getRange("B14"));
styleCount(assumptions.getRange("B15"));
styleCurrency(assumptions.getRange("B16"));
assumptions.getRange("A7:D16").format.borders = { bottom: { style: "thin", color: colors.midGray } };
assumptions.getRange("D7:D16").format.wrapText = true;

const assumptionBlocks = { Conservative: 19, Base: 32, Upside: 45 };
const metricRows = [
  ["Units sold", "units", "units"],
  ["MSRP", "$ / unit", "msrp"],
  ["DTC mix", "% units", "dtc"],
  ["Retail-direct mix", "% units", "retail"],
  ["Landed cost", "$ / unit", "landed"],
  ["Development & art", "$", "development"],
  ["Marketing", "$", "marketing"],
  ["Overhead", "$", "overhead"],
  ["Founder compensation", "$", "founderComp"],
];

for (const [caseName, start] of Object.entries(assumptionBlocks)) {
  sectionHeader(assumptions, `A${start}:G${start}`, `${caseName.toUpperCase()} CASE`);
  assumptions.getRange(`A${start + 1}:G${start + 1}`).values = [["Driver", "Units", ...years]];
  styleTableHeader(assumptions.getRange(`A${start + 1}:G${start + 1}`));
  const rows = metricRows.map(([label, unit, key]) => [label, unit, ...cases[caseName][key]]);
  assumptions.getRange(`A${start + 2}:G${start + 10}`).values = rows;
  styleInput(assumptions.getRange(`C${start + 2}:G${start + 10}`));
  styleCount(assumptions.getRange(`C${start + 2}:G${start + 2}`));
  styleCurrencyOne(assumptions.getRange(`C${start + 3}:G${start + 3}`));
  stylePercent(assumptions.getRange(`C${start + 4}:G${start + 5}`));
  styleCurrencyOne(assumptions.getRange(`C${start + 6}:G${start + 6}`));
  styleCurrency(assumptions.getRange(`C${start + 7}:G${start + 10}`));
  assumptions.getRange(`A${start + 2}:G${start + 10}`).format.borders = { bottom: { style: "thin", color: colors.midGray } };
}

addSourceComment(assumptions, "B7", "Source: Kickstarter official fee page | https://help.kickstarter.com/hc/en-us/articles/115005028634-What-are-the-fees | Accessed 2026-07-17");
addSourceComment(assumptions, "B8", "Source: Kickstarter states payment processing is approximately 3%–5%; 4% is the planning midpoint | https://help.kickstarter.com/hc/en-us/articles/115005028634-What-are-the-fees | Accessed 2026-07-17");
addSourceComment(assumptions, "B11", "Source: Public retailer-direct terms commonly show about 50% of MSRP | https://www.level99games.com/pages/retail | Accessed 2026-07-17");
addSourceComment(assumptions, "B12", "Source: Public distributor terms commonly show net 40% of MSRP | https://greenfairystudios.com/about/terms-and-conditions/distributor-terms/ | Accessed 2026-07-17");
addSourceComment(assumptions, "B15", "Source: Panda Game Manufacturing states a 1,500-game MOQ | https://pandagm.com/our-process/ | Accessed 2026-07-17");
addSourceComment(assumptions, "C38", "Source anchor: PrintNinja sample conventional board game was $10.60/unit at 1,000 copies and $14.28 at 500; base landed cost adds duty/receiving/quality contingency | https://printninja.com/sample-pricing/ | Accessed 2026-07-17");

assumptions.getRange("A1:G60").format.font.name = "Aptos";
assumptions.getRange("A1:A60").format.columnWidth = 27;
assumptions.getRange("B1:B60").format.columnWidth = 20;
assumptions.getRange("C1:G60").format.columnWidth = 14;
assumptions.getRange("D7:D16").format.columnWidth = 42;
assumptions.freezePanes.freezeRows(6);

// Operating Model
mergeTitle(operating, "A1:F2", "FIVE-YEAR OPERATING MODEL");
operating.getRange("A3:F3").merge();
operating.getRange("A3").values = [["Revenue excludes sales tax and customer-paid shipping. Green formulas are links; black formulas are calculations."]];
operating.getRange("A3:F3").format = { fill: colors.paleGray, font: { italic: true, color: colors.darkGray } };

const opBlocks = { Base: 4, Conservative: 36, Upside: 68 };
const assumpMetricOffset = {
  units: 2,
  msrp: 3,
  dtc: 4,
  retail: 5,
  landed: 6,
  development: 7,
  marketing: 8,
  overhead: 9,
  founderComp: 10,
};

for (const [caseName, start] of Object.entries(opBlocks)) {
  const aStart = assumptionBlocks[caseName];
  sectionHeader(operating, `A${start}:F${start}`, `${caseName.toUpperCase()} CASE`);
  operating.getRange(`A${start + 1}:F${start + 1}`).values = [["$ unless noted", ...years]];
  styleTableHeader(operating.getRange(`A${start + 1}:F${start + 1}`));
  const labels = [
    "Units Sold", "MSRP / Unit", "DTC Mix", "Retail-Direct Mix", "Distributor Mix",
    "DTC Units", "Retail-Direct Units", "Distributor Units",
    "DTC Net Sales", "Retail-Direct Net Sales", "Distributor Net Sales", "Net Revenue",
    "Landed COGS", "Platform + Payment Fees", "Refund Reserve", "DTC Pick/Pack", "Wholesale Freight",
    "Contribution Profit", "Contribution Margin", "Development & Art", "Marketing", "Overhead", "Founder Compensation",
    "EBITDA before Founder Comp", "EBITDA", "EBITDA Margin", "Cumulative EBITDA"
  ];
  operating.getRange(`A${start + 2}:A${start + 28}`).values = labels.map(x => [x]);
  for (let i = 0; i < 5; i++) {
    const col = String.fromCharCode(66 + i);
    const aCol = String.fromCharCode(67 + i);
    operating.getRange(`${col}${start + 2}`).formulas = [[`='Assumptions'!${aCol}${aStart + assumpMetricOffset.units}`]];
    operating.getRange(`${col}${start + 3}`).formulas = [[`='Assumptions'!${aCol}${aStart + assumpMetricOffset.msrp}`]];
    operating.getRange(`${col}${start + 4}`).formulas = [[`='Assumptions'!${aCol}${aStart + assumpMetricOffset.dtc}`]];
    operating.getRange(`${col}${start + 5}`).formulas = [[`='Assumptions'!${aCol}${aStart + assumpMetricOffset.retail}`]];
    operating.getRange(`${col}${start + 6}`).formulas = [[`=1-${col}${start + 4}-${col}${start + 5}`]];
    operating.getRange(`${col}${start + 7}`).formulas = [[`=${col}${start + 2}*${col}${start + 4}`]];
    operating.getRange(`${col}${start + 8}`).formulas = [[`=${col}${start + 2}*${col}${start + 5}`]];
    operating.getRange(`${col}${start + 9}`).formulas = [[`=${col}${start + 2}*${col}${start + 6}`]];
    operating.getRange(`${col}${start + 10}`).formulas = [[`=${col}${start + 7}*${col}${start + 3}`]];
    operating.getRange(`${col}${start + 11}`).formulas = [[`=${col}${start + 8}*${col}${start + 3}*'Assumptions'!$B$11`]];
    operating.getRange(`${col}${start + 12}`).formulas = [[`=${col}${start + 9}*${col}${start + 3}*'Assumptions'!$B$12`]];
    operating.getRange(`${col}${start + 13}`).formulas = [[`=SUM(${col}${start + 10}:${col}${start + 12})`]];
    operating.getRange(`${col}${start + 14}`).formulas = [[`=${col}${start + 2}*'Assumptions'!${aCol}${aStart + assumpMetricOffset.landed}`]];
    operating.getRange(`${col}${start + 15}`).formulas = [[`=${col}${start + 10}*('Assumptions'!$B$7+'Assumptions'!$B$8)`]];
    operating.getRange(`${col}${start + 16}`).formulas = [[`=${col}${start + 10}*'Assumptions'!$B$9`]];
    operating.getRange(`${col}${start + 17}`).formulas = [[`=${col}${start + 7}*'Assumptions'!$B$10`]];
    operating.getRange(`${col}${start + 18}`).formulas = [[`=(${col}${start + 11}+${col}${start + 12})*'Assumptions'!$B$13`]];
    operating.getRange(`${col}${start + 19}`).formulas = [[`=${col}${start + 13}-SUM(${col}${start + 14}:${col}${start + 18})`]];
    operating.getRange(`${col}${start + 20}`).formulas = [[`=${col}${start + 19}/${col}${start + 13}`]];
    operating.getRange(`${col}${start + 21}`).formulas = [[`='Assumptions'!${aCol}${aStart + assumpMetricOffset.development}`]];
    operating.getRange(`${col}${start + 22}`).formulas = [[`='Assumptions'!${aCol}${aStart + assumpMetricOffset.marketing}`]];
    operating.getRange(`${col}${start + 23}`).formulas = [[`='Assumptions'!${aCol}${aStart + assumpMetricOffset.overhead}`]];
    operating.getRange(`${col}${start + 24}`).formulas = [[`='Assumptions'!${aCol}${aStart + assumpMetricOffset.founderComp}`]];
    operating.getRange(`${col}${start + 25}`).formulas = [[`=${col}${start + 19}-SUM(${col}${start + 21}:${col}${start + 23})`]];
    operating.getRange(`${col}${start + 26}`).formulas = [[`=${col}${start + 25}-${col}${start + 24}`]];
    operating.getRange(`${col}${start + 27}`).formulas = [[`=${col}${start + 26}/${col}${start + 13}`]];
    operating.getRange(`${col}${start + 28}`).formulas = [[i === 0 ? `=${col}${start + 26}` : `=${String.fromCharCode(65 + i)}${start + 28}+${col}${start + 26}`]];
  }
  styleLinkedFormula(operating.getRange(`B${start + 2}:F${start + 5}`));
  styleLinkedFormula(operating.getRange(`B${start + 21}:F${start + 24}`));
  styleCount(operating.getRange(`B${start + 2}:F${start + 2}`));
  styleCurrencyOne(operating.getRange(`B${start + 3}:F${start + 3}`));
  stylePercent(operating.getRange(`B${start + 4}:F${start + 6}`));
  styleCount(operating.getRange(`B${start + 7}:F${start + 9}`));
  styleCurrency(operating.getRange(`B${start + 10}:F${start + 19}`));
  stylePercent(operating.getRange(`B${start + 20}:F${start + 20}`));
  styleCurrency(operating.getRange(`B${start + 21}:F${start + 26}`));
  stylePercent(operating.getRange(`B${start + 27}:F${start + 27}`));
  styleCurrency(operating.getRange(`B${start + 28}:F${start + 28}`));
  styleSubheader(operating.getRange(`A${start + 10}:F${start + 10}`));
  borderTotal(operating.getRange(`A${start + 13}:F${start + 13}`));
  borderTotal(operating.getRange(`A${start + 19}:F${start + 20}`));
  borderTotal(operating.getRange(`A${start + 25}:F${start + 28}`));
  operating.getRange(`A${start + 2}:F${start + 28}`).format.borders = { bottom: { style: "thin", color: colors.midGray } };
}

operating.getRange("A1:F100").format.font.name = "Aptos";
operating.getRange("A1:A100").format.columnWidth = 31;
operating.getRange("B1:F100").format.columnWidth = 15;
operating.freezePanes.freezeRows(5);

// Unit economics
mergeTitle(unitEcon, "A1:G2", "UNIT ECONOMICS & BREAK-EVEN");
unitEcon.getRange("A3:G3").merge();
unitEcon.getRange("A3").values = [["Base-case channel economics. Customer postage and sales tax are excluded."]];
unitEcon.getRange("A3:G3").format = { fill: colors.paleGray, font: { italic: true, color: colors.darkGray } };
for (const [label, row, yearCol] of [["LAUNCH ECONOMICS — 2027", 5, "C"], ["MATURE ECONOMICS — 2031", 14, "G"]]) {
  sectionHeader(unitEcon, `A${row}:G${row}`, label);
  unitEcon.getRange(`A${row + 1}:G${row + 1}`).values = [["Channel", "Revenue / Unit", "Landed COGS", "Fees / Reserve", "Handling / Freight", "Contribution / Unit", "Contribution Margin"]];
  styleTableHeader(unitEcon.getRange(`A${row + 1}:G${row + 1}`));
  unitEcon.getRange(`A${row + 2}:A${row + 5}`).values = [["DTC / Crowdfunding"], ["Retail Direct"], ["Distributor"], ["Weighted Mix"]];
  const baseStart = assumptionBlocks.Base;
  const opStart = opBlocks.Base;
  unitEcon.getRange(`B${row + 2}`).formulas = [[`='Assumptions'!${yearCol}${baseStart + assumpMetricOffset.msrp}`]];
  unitEcon.getRange(`C${row + 2}`).formulas = [[`='Assumptions'!${yearCol}${baseStart + assumpMetricOffset.landed}`]];
  unitEcon.getRange(`D${row + 2}`).formulas = [[`=B${row + 2}*SUM('Assumptions'!$B$7:$B$9)`]];
  unitEcon.getRange(`E${row + 2}`).formulas = [["='Assumptions'!$B$10"]];
  unitEcon.getRange(`F${row + 2}`).formulas = [[`=B${row + 2}-SUM(C${row + 2}:E${row + 2})`]];
  unitEcon.getRange(`G${row + 2}`).formulas = [[`=F${row + 2}/B${row + 2}`]];
  unitEcon.getRange(`B${row + 3}`).formulas = [[`='Assumptions'!${yearCol}${baseStart + assumpMetricOffset.msrp}*'Assumptions'!$B$11`]];
  unitEcon.getRange(`C${row + 3}`).formulas = [[`='Assumptions'!${yearCol}${baseStart + assumpMetricOffset.landed}`]];
  unitEcon.getRange(`D${row + 3}`).values = [[0]];
  unitEcon.getRange(`E${row + 3}`).formulas = [[`=B${row + 3}*'Assumptions'!$B$13`]];
  unitEcon.getRange(`F${row + 3}`).formulas = [[`=B${row + 3}-SUM(C${row + 3}:E${row + 3})`]];
  unitEcon.getRange(`G${row + 3}`).formulas = [[`=F${row + 3}/B${row + 3}`]];
  unitEcon.getRange(`B${row + 4}`).formulas = [[`='Assumptions'!${yearCol}${baseStart + assumpMetricOffset.msrp}*'Assumptions'!$B$12`]];
  unitEcon.getRange(`C${row + 4}`).formulas = [[`='Assumptions'!${yearCol}${baseStart + assumpMetricOffset.landed}`]];
  unitEcon.getRange(`D${row + 4}`).values = [[0]];
  unitEcon.getRange(`E${row + 4}`).formulas = [[`=B${row + 4}*'Assumptions'!$B$13`]];
  unitEcon.getRange(`F${row + 4}`).formulas = [[`=B${row + 4}-SUM(C${row + 4}:E${row + 4})`]];
  unitEcon.getRange(`G${row + 4}`).formulas = [[`=F${row + 4}/B${row + 4}`]];
  const opCol = yearCol === "C" ? "B" : "F";
  unitEcon.getRange(`B${row + 5}`).formulas = [[`='Operating Model'!${opCol}${opStart + 13}/'Operating Model'!${opCol}${opStart + 2}`]];
  unitEcon.getRange(`C${row + 5}`).formulas = [[`='Operating Model'!${opCol}${opStart + 14}/'Operating Model'!${opCol}${opStart + 2}`]];
  unitEcon.getRange(`D${row + 5}`).formulas = [[`=SUM('Operating Model'!${opCol}${opStart + 15}:'Operating Model'!${opCol}${opStart + 16})/'Operating Model'!${opCol}${opStart + 2}`]];
  unitEcon.getRange(`E${row + 5}`).formulas = [[`=SUM('Operating Model'!${opCol}${opStart + 17}:'Operating Model'!${opCol}${opStart + 18})/'Operating Model'!${opCol}${opStart + 2}`]];
  unitEcon.getRange(`F${row + 5}`).formulas = [[`='Operating Model'!${opCol}${opStart + 19}/'Operating Model'!${opCol}${opStart + 2}`]];
  unitEcon.getRange(`G${row + 5}`).formulas = [[`='Operating Model'!${opCol}${opStart + 20}`]];
  styleCurrencyOne(unitEcon.getRange(`B${row + 2}:F${row + 5}`));
  stylePercent(unitEcon.getRange(`G${row + 2}:G${row + 5}`));
  unitEcon.getRange(`A${row + 2}:G${row + 5}`).format.borders = { bottom: { style: "thin", color: colors.midGray } };
  borderTotal(unitEcon.getRange(`A${row + 5}:G${row + 5}`));
}

sectionHeader(unitEcon, "A23:G23", "BASE-CASE BREAK-EVEN");
unitEcon.getRange("A24:D24").values = [["Metric", "2027", "2031", "Interpretation"]];
styleTableHeader(unitEcon.getRange("A24:D24"));
unitEcon.getRange("A25:A28").values = [["Weighted contribution / unit"], ["Fixed cash costs before founder comp"], ["Break-even units before founder comp"], ["Break-even units after founder comp"]];
unitEcon.getRange("B25").formulas = [["=F10"]]; unitEcon.getRange("C25").formulas = [["=F19"]];
unitEcon.getRange("B26").formulas = [["=SUM('Operating Model'!B25:B27)"]]; unitEcon.getRange("C26").formulas = [["=SUM('Operating Model'!F25:F27)"]];
unitEcon.getRange("B27").formulas = [["=B26/B25"]]; unitEcon.getRange("C27").formulas = [["=C26/C25"]];
unitEcon.getRange("B28").formulas = [["=(B26+'Operating Model'!B28)/B25"]]; unitEcon.getRange("C28").formulas = [["=(C26+'Operating Model'!F28)/C25"]];
unitEcon.getRange("D25:D28").values = [["Channel mix matters as much as MSRP"], ["Development, marketing, and overhead"], ["First-year base plan is below this level"], ["Mature base plan exceeds this level"]];
styleCurrencyOne(unitEcon.getRange("B25:C25"));
styleCurrency(unitEcon.getRange("B26:C26"));
styleCount(unitEcon.getRange("B27:C28"));
unitEcon.getRange("A25:D28").format.borders = { bottom: { style: "thin", color: colors.midGray } };
unitEcon.getRange("D25:D28").format.wrapText = true;
unitEcon.getRange("A1:G32").format.font.name = "Aptos";
unitEcon.getRange("A1:A32").format.columnWidth = 38;
unitEcon.getRange("B1:G32").format.columnWidth = 18;
unitEcon.getRange("D1:D32").format.columnWidth = 34;
unitEcon.getRange("A25:A28").format.wrapText = true;
unitEcon.getRange("A25:D28").format.autofitRows();
unitEcon.freezePanes.freezeRows(6);

// Scenario Summary
mergeTitle(summary, "A1:G2", "SCENARIO SUMMARY");
summary.getRange("A3:G3").merge();
summary.getRange("A3").values = [["Scenarios vary unit demand, channel mix, landed cost, and investment cadence. They are not probabilities."]];
summary.getRange("A3:G3").format = { fill: colors.paleGray, font: { italic: true, color: colors.darkGray } };
summary.getRange("A5:G5").values = [["Scenario", "Year-5 Revenue", "Year-5 EBITDA", "5-Year Revenue", "5-Year EBITDA", "Break-even Year", "Year-5 Units"]];
styleTableHeader(summary.getRange("A5:G5"));
summary.getRange("A6:A8").values = [["Conservative"], ["Base"], ["Upside"]];
const caseOrder = ["Conservative", "Base", "Upside"];
for (let i = 0; i < caseOrder.length; i++) {
  const name = caseOrder[i];
  const row = 6 + i;
  const op = opBlocks[name];
  summary.getRange(`B${row}`).formulas = [[`='Operating Model'!F${op + 13}`]];
  summary.getRange(`C${row}`).formulas = [[`='Operating Model'!F${op + 26}`]];
  summary.getRange(`D${row}`).formulas = [[`=SUM('Operating Model'!B${op + 13}:F${op + 13})`]];
  summary.getRange(`E${row}`).formulas = [[`=SUM('Operating Model'!B${op + 26}:F${op + 26})`]];
  summary.getRange(`F${row}`).formulas = [[`=IF('Operating Model'!B${op + 26}>0,2027,IF('Operating Model'!C${op + 26}>0,2028,IF('Operating Model'!D${op + 26}>0,2029,IF('Operating Model'!E${op + 26}>0,2030,IF('Operating Model'!F${op + 26}>0,2031,"Not reached")))))`]];
  summary.getRange(`G${row}`).formulas = [[`='Operating Model'!F${op + 2}`]];
}
styleCurrency(summary.getRange("B6:E8"));
styleCount(summary.getRange("G6:G8"));
summary.getRange("A6:G8").format.borders = { bottom: { style: "thin", color: colors.midGray } };
summary.getRange("A7:G7").format.fill = colors.paleTeal;
summary.getRange("A7:G7").format.font = { bold: true, color: colors.navy };

sectionHeader(summary, "A11:F11", "ANNUAL NET REVENUE");
summary.getRange("A12:F12").values = [["Scenario", ...years]];
styleTableHeader(summary.getRange("A12:F12"));
summary.getRange("A13:A15").values = [["Conservative"], ["Base"], ["Upside"]];
for (let i = 0; i < caseOrder.length; i++) {
  const op = opBlocks[caseOrder[i]];
  summary.getRange(`B${13 + i}:F${13 + i}`).formulas = [[...years.map((_, j) => `='Operating Model'!${String.fromCharCode(66 + j)}${op + 13}`)]];
}
styleCurrency(summary.getRange("B13:F15"));

sectionHeader(summary, "A18:F18", "ANNUAL EBITDA");
summary.getRange("A19:F19").values = [["Scenario", ...years]];
styleTableHeader(summary.getRange("A19:F19"));
summary.getRange("A20:A22").values = [["Conservative"], ["Base"], ["Upside"]];
for (let i = 0; i < caseOrder.length; i++) {
  const op = opBlocks[caseOrder[i]];
  summary.getRange(`B${20 + i}:F${20 + i}`).formulas = [[...years.map((_, j) => `='Operating Model'!${String.fromCharCode(66 + j)}${op + 26}`)]];
}
styleCurrency(summary.getRange("B20:F22"));

const chartColors = ["#2A7F75", "#234E70", "#D6A84B"];
const revenueChart = summary.charts.add("line", {
  chartType: "line",
  title: "Revenue Scales Only After Repeatable Demand ($)",
  hasLegend: true,
});
for (let i = 0; i < caseOrder.length; i++) {
  const series = revenueChart.series.add(caseOrder[i]);
  series.categoryFormula = "'Scenario Summary'!$B$12:$F$12";
  series.formula = `'Scenario Summary'!$B$${13 + i}:$F$${13 + i}`;
  series.fill = chartColors[i];
}
revenueChart.title = "Revenue Scales Only After Repeatable Demand ($)";
revenueChart.hasLegend = true;
revenueChart.xAxis = { axisType: "textAxis" };
revenueChart.yAxis = { numberFormatCode: "$#,##0" };
revenueChart.setPosition("I4", "P17");

const ebitdaChart = summary.charts.add("line", {
  chartType: "line",
  title: "EBITDA Diverges by Scenario ($)",
  hasLegend: true,
});
for (let i = 0; i < caseOrder.length; i++) {
  const series = ebitdaChart.series.add(caseOrder[i]);
  series.categoryFormula = "'Scenario Summary'!$B$19:$F$19";
  series.formula = `'Scenario Summary'!$B$${20 + i}:$F$${20 + i}`;
  series.fill = chartColors[i];
}
ebitdaChart.title = "EBITDA Diverges by Scenario ($)";
ebitdaChart.hasLegend = true;
ebitdaChart.xAxis = { axisType: "textAxis" };
ebitdaChart.yAxis = { numberFormatCode: "$#,##0" };
ebitdaChart.setPosition("I19", "P32");

summary.getRange("A1:P34").format.font.name = "Aptos";
summary.getRange("A1:A34").format.columnWidth = 20;
summary.getRange("B1:G34").format.columnWidth = 18;
summary.getRange("H1:H34").format.columnWidth = 3;
summary.freezePanes.freezeRows(5);

// Cash & Funding
mergeTitle(cash, "A1:F2", "CASH RESERVE & FUNDING GATES");
cash.getRange("A3:F3").merge();
cash.getRange("A3").values = [["Base-case liquidity view. EBITDA is used as a simplified cash proxy; inventory timing, tax, debt, and receivables are not modeled."]];
cash.getRange("A3:F3").format = { fill: colors.paleGray, font: { italic: true, color: colors.darkGray }, wrapText: true };
sectionHeader(cash, "A5:F5", "BASE-CASE CASH ROLL-FORWARD");
cash.getRange("A6:F6").values = [["$", ...years]];
styleTableHeader(cash.getRange("A6:F6"));
cash.getRange("A7:A10").values = [["Beginning Cash"], ["EBITDA"], ["Ending Cash"], ["Minimum Cash Flag"]];
cash.getRange("B7").formulas = [["='Assumptions'!$B$14"]];
cash.getRange("C7:F7").formulas = [["=B9", "=C9", "=D9", "=E9"]];
cash.getRange("B8:F8").formulas = [["='Operating Model'!B30", "='Operating Model'!C30", "='Operating Model'!D30", "='Operating Model'!E30", "='Operating Model'!F30"]];
cash.getRange("B9:F9").formulas = [["=B7+B8", "=C7+C8", "=D7+D8", "=E7+E8", "=F7+F8"]];
cash.getRange("B10:F10").formulas = [["=IF(B9<15000,\"LOW\",\"OK\")", "=IF(C9<15000,\"LOW\",\"OK\")", "=IF(D9<15000,\"LOW\",\"OK\")", "=IF(E9<15000,\"LOW\",\"OK\")", "=IF(F9<15000,\"LOW\",\"OK\")"]];
styleCurrency(cash.getRange("B7:F9"));
cash.getRange("B10:F10").conditionalFormats.add("containsText", { text: "LOW", format: { fill: colors.paleRed, font: { bold: true, color: colors.failRed } } });
cash.getRange("B10:F10").conditionalFormats.add("containsText", { text: "OK", format: { fill: colors.paleTeal, font: { bold: true, color: colors.passGreen } } });
borderTotal(cash.getRange("A9:F9"));

sectionHeader(cash, "A13:D13", "MILESTONE-GATED USE OF FUNDS");
cash.getRange("A14:D14").values = [["Use", "Amount", "Release Gate", "Purpose"]];
styleTableHeader(cash.getRange("A14:D14"));
const uses = [
  ["Prototypes & playtesting", 5000, "Concept gate", "Fast evidence before final art"],
  ["Development, art & graphics", 18000, "External playtest gate", "Staged contractor commitments"],
  ["Campaign creative & prelaunch", 12000, "Blind-test + audience gates", "Preview assets and proven promotion"],
  ["Legal, IP, insurance, compliance", 4000, "Before public brand spend", "Ownership, clearance, risk controls"],
  ["Manufacturing/fulfillment working capital", 20000, "Successful campaign + final quotes", "Deposits and timing buffer"],
  ["General overhead/tools", 4000, "Monthly cap", "Software, samples, administration"],
  ["Contingency", 12000, "Joint founder approval", "Cost, duty, quality, or schedule shock"],
];
cash.getRange("A15:D21").values = uses;
cash.getRange("A22").values = [["Total"]];
cash.getRange("B22").formulas = [["=SUM(B15:B21)"]];
cash.getRange("C22").values = [["Must equal seed reserve"]];
cash.getRange("D22").formulas = [["=IF(B22='Assumptions'!$B$14,\"PASS\",\"FAIL\")"]];
styleCurrency(cash.getRange("B15:B22"));
styleInput(cash.getRange("B15:B21"), '"$"#,##0;[Red]("$"#,##0);-');
borderTotal(cash.getRange("A22:D22"));
cash.getRange("A15:D21").format.borders = { bottom: { style: "thin", color: colors.midGray } };
cash.getRange("C15:D21").format.wrapText = true;

cash.getRange("A24:F27").merge();
cash.getRange("A24").values = [["Funding rule: do not treat the $75,000 reserve as permission to spend. Each tranche is released only after the preceding evidence gate. Crowdfunding product receipts are modeled in revenue and should not be double-counted as financing."]];
cash.getRange("A24:F27").format = { fill: colors.paleGold, font: { bold: true, color: colors.navy }, wrapText: true, verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: colors.gold } };
cash.getRange("A1:F30").format.font.name = "Aptos";
cash.getRange("A1:A30").format.columnWidth = 34;
cash.getRange("B1:B30").format.columnWidth = 16;
cash.getRange("C1:C30").format.columnWidth = 28;
cash.getRange("D1:D30").format.columnWidth = 38;
cash.getRange("E1:F30").format.columnWidth = 15;
cash.freezePanes.freezeRows(6);

// Checks
mergeTitle(checks, "A1:G2", "MODEL CHECKS");
checks.getRange("A3:G3").merge();
checks.getRange("A3").formulas = [["=\"MODEL STATUS: \"&IF(COUNTIF(F7:F13,\"FAIL\")=0,\"PASS\",\"FAIL\")"]];
checks.getRange("A3:G3").format = { fill: colors.paleTeal, font: { bold: true, color: colors.passGreen, size: 14 }, horizontalAlignment: "center" };
checks.getRange("A5:G5").values = [["Check", "Actual", "Expected", "Difference", "Tolerance", "Status", "Where to Fix / Notes"]];
styleTableHeader(checks.getRange("A5:G5"));
const checkLabels = [
  "All channel mixes equal 100%",
  "Use of funds equals seed reserve",
  "Base Year-5 revenue ties to summary",
  "Base ending cash roll-forward ties",
  "Base contribution is positive every year",
  "Scenario ordering: Upside > Base > Conservative",
  "Base Year-5 EBITDA is formula-driven",
];
checks.getRange("A7:A13").values = checkLabels.map(x => [x]);
checks.getRange("B7").formulas = [["=MAX(ABS(SUM('Operating Model'!B8:B10)-1),ABS(SUM('Operating Model'!C8:C10)-1),ABS(SUM('Operating Model'!D8:D10)-1),ABS(SUM('Operating Model'!E8:E10)-1),ABS(SUM('Operating Model'!F8:F10)-1),ABS(SUM('Operating Model'!B40:B42)-1),ABS(SUM('Operating Model'!C40:C42)-1),ABS(SUM('Operating Model'!D40:D42)-1),ABS(SUM('Operating Model'!E40:E42)-1),ABS(SUM('Operating Model'!F40:F42)-1),ABS(SUM('Operating Model'!B72:B74)-1),ABS(SUM('Operating Model'!C72:C74)-1),ABS(SUM('Operating Model'!D72:D74)-1),ABS(SUM('Operating Model'!E72:E74)-1),ABS(SUM('Operating Model'!F72:F74)-1))"]];
checks.getRange("C7").values = [[0]]; checks.getRange("D7").formulas = [["=B7-C7"]]; checks.getRange("E7").values = [[0.0001]]; checks.getRange("F7").formulas = [["=IF(ABS(D7)<=E7,\"PASS\",\"FAIL\")"]]; checks.getRange("G7").values = [["Operating Model channel-mix rows"]];
checks.getRange("B8").formulas = [["='Cash & Funding'!B22"]]; checks.getRange("C8").formulas = [["='Assumptions'!B14"]]; checks.getRange("D8").formulas = [["=B8-C8"]]; checks.getRange("E8").values = [[0]]; checks.getRange("F8").formulas = [["=IF(ABS(D8)<=E8,\"PASS\",\"FAIL\")"]]; checks.getRange("G8").values = [["Cash & Funding use-of-funds table"]];
checks.getRange("B9").formulas = [["='Operating Model'!F17"]]; checks.getRange("C9").formulas = [["='Scenario Summary'!B7"]]; checks.getRange("D9").formulas = [["=B9-C9"]]; checks.getRange("E9").values = [[0.01]]; checks.getRange("F9").formulas = [["=IF(ABS(D9)<=E9,\"PASS\",\"FAIL\")"]]; checks.getRange("G9").values = [["Scenario Summary Base row"]];
checks.getRange("B10").formulas = [["='Cash & Funding'!F9"]]; checks.getRange("C10").formulas = [["='Assumptions'!B14+SUM('Operating Model'!B30:F30)"]]; checks.getRange("D10").formulas = [["=B10-C10"]]; checks.getRange("E10").values = [[0.01]]; checks.getRange("F10").formulas = [["=IF(ABS(D10)<=E10,\"PASS\",\"FAIL\")"]]; checks.getRange("G10").values = [["Cash roll-forward"]];
checks.getRange("B11").formulas = [["=MIN('Operating Model'!B23:F23)"]]; checks.getRange("C11").values = [[0]]; checks.getRange("D11").formulas = [["=B11-C11"]]; checks.getRange("E11").values = [[0]]; checks.getRange("F11").formulas = [["=IF(B11>C11,\"PASS\",\"FAIL\")"]]; checks.getRange("G11").values = [["Base contribution-profit row"]];
checks.getRange("B12").formulas = [["='Scenario Summary'!B8-'Scenario Summary'!B7"]]; checks.getRange("C12").formulas = [["='Scenario Summary'!B7-'Scenario Summary'!B6"]]; checks.getRange("D12").formulas = [["=MIN(B12,C12)"]]; checks.getRange("E12").values = [[0]]; checks.getRange("F12").formulas = [["=IF(D12>0,\"PASS\",\"FAIL\")"]]; checks.getRange("G12").values = [["Year-5 revenue ordering"]];
checks.getRange("B13").formulas = [["='Operating Model'!F30"]]; checks.getRange("C13").formulas = [["='Operating Model'!F29-'Operating Model'!F28"]]; checks.getRange("D13").formulas = [["=B13-C13"]]; checks.getRange("E13").values = [[0.01]]; checks.getRange("F13").formulas = [["=IF(ABS(D13)<=E13,\"PASS\",\"FAIL\")"]]; checks.getRange("G13").values = [["Base EBITDA formula"]];
styleCurrency(checks.getRange("B8:E13"));
stylePercent(checks.getRange("B7:E7"));
checks.getRange("F7:F13").conditionalFormats.add("containsText", { text: "PASS", format: { fill: colors.paleTeal, font: { bold: true, color: colors.passGreen } } });
checks.getRange("F7:F13").conditionalFormats.add("containsText", { text: "FAIL", format: { fill: colors.paleRed, font: { bold: true, color: colors.failRed } } });
checks.getRange("A7:G13").format.borders = { bottom: { style: "thin", color: colors.midGray } };
checks.getRange("G7:G13").format.wrapText = true;
checks.getRange("A1:G18").format.font.name = "Aptos";
checks.getRange("A1:A18").format.columnWidth = 42;
checks.getRange("B1:F18").format.columnWidth = 16;
checks.getRange("G1:G18").format.columnWidth = 36;
checks.freezePanes.freezeRows(5);

// Sources
mergeTitle(sources, "A1:G2", "SOURCES & AUDIT TRAIL");
sources.getRange("A3:G3").merge();
sources.getRange("A3").values = [["Public sources support market context and selected assumptions. Planning assumptions remain estimates until replaced by company-specific evidence."]];
sources.getRange("A3:G3").format = { fill: colors.paleGray, font: { italic: true, color: colors.darkGray }, wrapText: true };
sources.getRange("A5:G5").values = [["ID", "Item", "Value / Claim", "Units / As-of", "Source Type", "Source URL", "Notes"]];
styleTableHeader(sources.getRange("A5:G5"));
const sourceRows = [
  ["S01", "U.S. Games/Puzzles", 4900000000, "$ retail sales / 2025", "Trade association / Circana", "https://www.toyassociation.org/ta/toys/research-and-data/data/us-sales-data.aspx", "Broad category; includes puzzles, TCGs, mass/licensed products"],
  ["S02", "2025 Games/Puzzles growth", 0.37, "% YoY / 2025", "Circana", "https://www.circana.com/post/u-s-toy-industry-returns-to-growth-in-2025-circana-reports", "Driven primarily by Pokémon"],
  ["S03", "U.S./Canada hobby games", 3660000000, "$ / 2025", "ICv2 estimate", "https://icv2.com/articles/markets/view/62035/big-growth-hobby-games-2025", "TCGs roughly two-thirds of dollars"],
  ["S04", "Kickstarter tabletop pledges", 220000000, "$ successful campaigns / 2024", "Kickstarter", "https://updates.kickstarter.com/kickstarter-biggest-platform-for-games/", "6,646 launches; 5,314 funded"],
  ["S05", "Kickstarter tabletop success", 0.80, "% projects / 2024", "Kickstarter", "https://updates.kickstarter.com/kickstarter-biggest-platform-for-games/", "Not a Wallaboo probability forecast"],
  ["S06", "Kickstarter platform fee", 0.05, "% collected funds", "Kickstarter", "https://help.kickstarter.com/hc/en-us/articles/115005028634-What-are-the-fees", "Payment processing separately 3%–5%"],
  ["S07", "Straightforward board-game cost", "8–15", "$ / unit at 1,500+", "Manufacturer guidance", "https://printninja.com/printing-products/board-game-printing/", "Budgetary range, not a quote"],
  ["S08", "Sample board game / 1,000", 10.60, "$ / unit / Apr 2026", "Manufacturer sample", "https://printninja.com/sample-pricing/", "Production example"],
  ["S09", "Sample board game / 500", 14.28, "$ / unit / Apr 2026", "Manufacturer sample", "https://printninja.com/sample-pricing/", "Production example"],
  ["S10", "Panda minimum run", 1500, "units", "Manufacturer", "https://pandagm.com/our-process/", "Other suppliers may differ"],
  ["S11", "PrintNinja minimum run", 500, "units", "Manufacturer", "https://printninja.com/game-quote-guide/", "Configuration dependent"],
  ["S12", "Retail-direct net", 0.50, "% MSRP", "Publisher terms", "https://www.level99games.com/pages/retail", "Public reference; negotiated terms vary"],
  ["S13", "Distributor net", 0.40, "% MSRP", "Publisher terms", "https://greenfairystudios.com/about/terms-and-conditions/distributor-terms/", "Public reference; negotiated terms vary"],
  ["S14", "Children's product classification", "Fact-specific", "Current guidance", "CPSC", "https://www.cpsc.gov/Business--Manufacturing/Business-Education/Childrens-Products", "Board games may be general use or children's products"],
  ["S15", "Toy testing / CPC", "Required for many 12-and-under toys", "Current guidance", "CPSC", "https://www.cpsc.gov/FAQ/Toy-Safety", "Obtain product-specific advice"],
  ["S16", "Trademark base filing fee", 350, "$ per class", "USPTO", "https://www.uspto.gov/trademarks/basics/how-much-does-it-cost", "Legal and additional fees not included"],
  ["S17", "Trademark clearance", "Search before filing", "Current guidance", "USPTO", "https://www.uspto.gov/trademarks/basics/why-search-similar-trademarks", "Internet review is not a legal opinion"],
  ["S18", "WALLABOO third-party brand", "Dutch baby-products brand since 2006", "Observed 2026-07-17", "Company website", "https://www.wallaboo.com/pages/about-us", "Different goods; still requires professional clearance"],
  ["S19", "Temporary import surcharge", 0.10, "% through 2026-07-24 unless changed", "White House proclamation", "https://www.whitehouse.gov/presidential-actions/2026/02/imposing-a-temporary-import-surcharge-to-address-fundamental-international-payments-problems/", "Recheck at every purchase order"],
  ["S20", "Source links", "Founders, brand, and concept provenance", "2026-02-14 / recovered 2026-07-17", "User-provided shared sources", "https://chatgpt.com/share/6a5a60d0-26a0-83ea-8007-3f0a196c1e76", "Additional canvas bodies not exposed"],
];
sources.getRange(`A6:G${5 + sourceRows.length}`).values = sourceRows;
sources.getRange("B6:G25").format.wrapText = true;
sources.getRange("A6:G25").format.borders = { bottom: { style: "thin", color: colors.midGray } };
styleCurrency(sources.getRange("C6"));
stylePercent(sources.getRange("C7"));
styleCurrency(sources.getRange("C8:C9"));
stylePercent(sources.getRange("C10"));
styleCurrencyOne(sources.getRange("C13:C14"));
styleCount(sources.getRange("C15:C16"));
stylePercent(sources.getRange("C17:C18"));
styleCurrency(sources.getRange("C21"));
stylePercent(sources.getRange("C24"));
sources.getRange("A1:G28").format.font.name = "Aptos";
sources.getRange("A1:A28").format.columnWidth = 8;
sources.getRange("B1:B28").format.columnWidth = 29;
sources.getRange("C1:C28").format.columnWidth = 32;
sources.getRange("D1:D28").format.columnWidth = 25;
sources.getRange("E1:E28").format.columnWidth = 24;
sources.getRange("F1:F28").format.columnWidth = 58;
sources.getRange("G1:G28").format.columnWidth = 38;
sources.getRange("A6:G25").format.autofitRows();
sources.freezePanes.freezeRows(5);

// Common print/render cleanup
for (const sheet of workbook.worksheets.items) {
  const used = sheet.getUsedRange();
  if (used) {
    used.format.verticalAlignment = "center";
  }
}

await fs.mkdir(outputDir, { recursive: true });
await fs.rm(qaDir, { recursive: true, force: true });
await fs.mkdir(qaDir, { recursive: true });

const inspectSummary = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 4000 });
console.log(inspectSummary.ndjson);

for (const sheetName of ["Cover", "Assumptions", "Unit Economics", "Operating Model", "Scenario Summary", "Cash & Funding", "Checks", "Sources"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(qaDir, `${sheetName.replaceAll(" ", "_").replaceAll("&", "and")}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(JSON.stringify({ outputPath, qaDir }));
