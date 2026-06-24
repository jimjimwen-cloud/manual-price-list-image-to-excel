import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const [xlsxPath, previewPath] = process.argv.slice(2);
if (!xlsxPath || !previewPath) throw new Error("Usage: verify_price_list.mjs <output.xlsx> <preview.png>");

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(xlsxPath));
const sheetName = workbook.worksheets.items[0].name;
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula errors",
});
const preview = await workbook.render({ sheetName, scale: 0.5 });
await fs.writeFile(previewPath, Buffer.from(await preview.arrayBuffer()));
console.log(errors.ndjson);
console.log(JSON.stringify({ xlsxPath, previewPath, sheetName }));

