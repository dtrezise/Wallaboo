import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "/Users/dan/Documents/GitHub/Wallaboo/outputs/Wallaboo_Financial_Model_v0.1.xlsx";
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const checks = [
  ["Cover", "A9:H13"],
  ["Operating Model", "A4:F32"],
  ["Scenario Summary", "A5:G8"],
  ["Cash & Funding", "A6:F10"],
  ["Checks", "A3:G13"],
];

for (const [sheetId, range] of checks) {
  const result = await workbook.inspect({
    kind: "region",
    sheetId,
    range,
    include: "values,formulas",
    tableMaxRows: 40,
    tableMaxCols: 10,
    maxChars: 12000,
  });
  console.log(`\n### ${sheetId}!${range}`);
  console.log(result.ndjson);
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log("\n### Formula error scan");
console.log(errors.ndjson);
