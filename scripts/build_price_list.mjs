import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const [modelPath, outputPath] = process.argv.slice(2);
if (!modelPath || !outputPath) throw new Error("Usage: build_price_list.mjs <model.json> <output.xlsx>");

const model = JSON.parse(await fs.readFile(modelPath, "utf8"));
const workbook = Workbook.create();
const sheet = workbook.worksheets.add(model.sheetName || "报价单");
const palette = { green: "#A9D18E", modelGreen: "#C6E0B4", blue: "#BDD7EE", red: "#FF0000", black: "#000000", white: "#FFFFFF", ...(model.palette || {}) };
const color = (value, fallback) => palette[value] || value || fallback;
const border = { style: "thin", color: palette.black };
const borders = { top: border, bottom: border, left: border, right: border };
const columns = model.columnWidths?.length || model.sections?.[0]?.headers?.length || 7;
const lastCol = String.fromCharCode(64 + columns);

function style(range, { fill = "white", fontColor = "black", fontSize = 12, bold = true, rowHeight } = {}) {
  range.format.fill = color(fill, palette.white);
  range.format.font = { name: "Arial", size: fontSize, bold, color: color(fontColor, palette.black) };
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  range.format.wrapText = true;
  range.format.borders = borders;
  if (rowHeight) range.format.rowHeight = rowHeight;
}

function mergedBlock(startRow, block) {
  const endRow = startRow + (block.rows || 1) - 1;
  sheet.mergeCells(`A${startRow}:${lastCol}${endRow}`);
  const range = sheet.getRange(`A${startRow}:${lastCol}${endRow}`);
  range.values = [[block.text || ""]];
  style(range, { fill: block.fill || "green", fontColor: block.fontColor || "black", fontSize: block.fontSize || 13, rowHeight: block.rowHeight || 30 });
  return endRow + 1;
}

let row = 1;
for (const block of model.topBlocks || []) row = mergedBlock(row, block);

for (const section of model.sections || []) {
  if (section.headers.length !== columns) throw new Error(`Header width mismatch at row ${row}`);
  const header = sheet.getRange(`A${row}:${lastCol}${row}`);
  header.values = [section.headers];
  style(header, { fill: section.headerFill || "green", fontSize: 12, rowHeight: 52 });
  row++;
  const sectionStart = row;

  for (const group of section.groups || []) {
    const groupStart = row;
    for (const data of group.rows || []) {
      if (data.length !== columns - 2) throw new Error(`Data width mismatch for model ${group.model}`);
      const values = [null, ...data, null];
      const range = sheet.getRange(`A${row}:${lastCol}${row}`);
      range.values = [values];
      style(range, { fill: "white", fontSize: 12, rowHeight: 29 });
      for (let c = 2; c < columns - 1; c++) {
        if (typeof values[c] === "number") sheet.getRange(`${String.fromCharCode(65 + c)}${row}`).format.numberFormat = "0";
      }
      row++;
    }
    const groupEnd = row - 1;
    sheet.mergeCells(`A${groupStart}:A${groupEnd}`);
    const modelRange = sheet.getRange(`A${groupStart}:A${groupEnd}`);
    modelRange.values = [[group.model || ""]];
    style(modelRange, { fill: group.modelFill || "modelGreen", fontColor: group.modelRedText ? "red" : "black", fontSize: 12 });
    for (const [offset, col] of group.blueCells || []) {
      sheet.getRange(`${String.fromCharCode(65 + col)}${groupStart + offset}`).format.fill = palette.blue;
    }
  }

  const sectionEnd = row - 1;
  const noteGroups = section.noteGroups || [{ count: sectionEnd - sectionStart + 1, text: section.note || "", redText: true }];
  let noteRow = sectionStart;
  for (const note of noteGroups) {
    const end = noteRow + note.count - 1;
    if (end > sectionEnd) throw new Error(`Note group exceeds section rows near ${noteRow}`);
    sheet.mergeCells(`${lastCol}${noteRow}:${lastCol}${end}`);
    const noteRange = sheet.getRange(`${lastCol}${noteRow}:${lastCol}${end}`);
    noteRange.values = [[note.text || ""]];
    style(noteRange, { fill: note.fill || "white", fontColor: note.redText === false ? "black" : "red", fontSize: 11 });
    noteRow = end + 1;
  }
  if (noteRow !== sectionEnd + 1) throw new Error(`Note group counts do not cover section ending at row ${sectionEnd}`);
}

for (const block of model.bottomBlocks || []) row = mergedBlock(row, block);
const usedEnd = row - 1;

const widths = model.columnWidths || Array(columns).fill(18);
for (let c = 0; c < columns; c++) sheet.getRange(`${String.fromCharCode(65 + c)}:${String.fromCharCode(65 + c)}`).format.columnWidth = widths[c];
sheet.getRange(`A1:${lastCol}${usedEnd}`).format.borders = borders;
sheet.freezePanes.freezeRows(model.freezeRows || Math.min(5, (model.topBlocks || []).reduce((n, b) => n + (b.rows || 1), 0)));

await fs.mkdir(new URL(".", `file://${outputPath}`).pathname, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, sheetName: model.sheetName || "报价单", rows: usedEnd, columns }));

