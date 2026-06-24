---
name: manual-price-list-image-to-excel
description: Manually transcribe price-list table images into structured, formatted Excel workbooks without using OCR tools, OCR APIs, OCR models, text-recognition libraries, or image-table-ocr skills. Use when the user asks to recognize, parse, convert, or extract a quotation/price-list image into .xlsx while preserving merged cells, numeric price types, text, repeated headers, notes, fills, and basic layout, especially when they explicitly require no OCR.
---

# Manual Price List Image to Excel

Create a verified `.xlsx` from a price-list image using visual inspection and manual transcription only.

## Mandatory Constraints

- Do not use OCR tools, APIs, models, libraries, browser text extraction, or another image-table recognition skill.
- Inspect the source image visually with `view_image`; crop or zoom sections when the image is long or dense.
- Use the Spreadsheets skill and bundled workspace dependencies for workbook authoring and verification.
- Store prices and counts as numbers. Preserve capacities such as `64G`, `512G`, and `1TB` as text. Preserve `-` as text.
- Report exact token usage only when the platform exposes it; otherwise state that it is unavailable and provide a clearly labeled estimate when requested.

## Workflow

1. Confirm the source image and requested output `.xlsx` filename.
2. Inspect the complete image and identify title blocks, column count, repeated headers, model groups, note groups, colored cells, and footer blocks.
3. For tall images, visually inspect top, middle, bottom, and dense sections separately. Manually transcribe every visible cell.
4. Create a JSON workbook model following [references/data-model.md](references/data-model.md). Record merged model/note ranges through grouped rows rather than duplicating text.
5. Compare the completed JSON against the image a second time before generating the workbook.
6. Load workspace dependencies. Create a writable scratch folder under the current workspace, link its `node_modules` to the returned bundled Node packages, and copy `scripts/build_price_list.mjs` plus `scripts/verify_price_list.mjs` into that scratch folder. Run the copied scripts from there so ESM can resolve `@oai/artifact-tool`.
7. Run `<scratch>/build_price_list.mjs <model.json> <output.xlsx>` with the bundled Node runtime.
8. Run `<scratch>/verify_price_list.mjs <output.xlsx> <preview.png>`. Inspect its output and visually review the preview.
9. Correct meaningful content, merge, numeric-type, or layout issues; rebuild and verify once more.
10. Deliver the final `.xlsx` link and briefly mention limitations or uncertain cells.

## Transcription Rules

- Preserve source wording, punctuation, line breaks, and visible adjustment notes. Do not silently rewrite text.
- Model text belongs in the merged first column and may contain adjustment notes on new lines.
- Notes belong in merged note groups spanning the exact affected data rows.
- Use `blueCells` to retain visible blue price highlights. Use `redText` only for source text shown in red.
- When the source has different column schemas by product generation, represent each schema as a separate section with its own repeated header.
- Use one worksheet unless the source clearly contains independent tables or the user requests multiple sheets.

## Verification Contract

- Confirm the first, middle, and last product sections against the image.
- Confirm every model group's row count and capacity sequence.
- Confirm numeric price cells remain numbers and placeholders remain text.
- Confirm model and note merges span the intended rows.
- Scan for formula errors even if no formulas are expected.
- Render the entire used range and fix clipping, unreadable widths, broken fills, or obvious merge problems.
