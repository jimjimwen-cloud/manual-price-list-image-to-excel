# Price List Workbook Model

Use UTF-8 JSON. The builder accepts this shape:

```json
{
  "sheetName": "报价单",
  "palette": {
    "green": "#A9D18E",
    "modelGreen": "#C6E0B4",
    "blue": "#BDD7EE",
    "red": "#FF0000"
  },
  "columnWidths": [18, 13, 20, 20, 20, 22, 24],
  "topBlocks": [
    {"text": "标题", "rows": 1, "fill": "green", "fontColor": "red", "fontSize": 26},
    {"text": "说明文字", "rows": 2, "fill": "green", "fontSize": 13}
  ],
  "sections": [
    {
      "headers": ["型号", "容量", "价格A", "价格B", "靓机", "小花", "备注"],
      "groups": [
        {
          "model": "16pro\n白色+50",
          "modelRedText": true,
          "rows": [
            ["128G", "-", "-", 4200, 4050],
            ["256G", 5750, 5600, 5200, 5000]
          ],
          "blueCells": [[0, 4], [1, 3]]
        }
      ],
      "noteGroups": [
        {"count": 2, "text": "电池低于90%-150", "redText": true}
      ]
    }
  ],
  "bottomBlocks": [
    {"text": "地址与交易说明", "rows": 2, "fill": "green", "fontSize": 14}
  ]
}
```

## Semantics

- `topBlocks` and `bottomBlocks` merge across all columns.
- Each section creates one repeated header row.
- Each group merges the first/model column across all its `rows`.
- Group row arrays contain every column except the model column and final note column. The builder inserts those merged columns.
- `blueCells` entries are `[zeroBasedRowOffset, zeroBasedWorksheetColumn]`. Worksheet column `0` is model, `1` is capacity.
- `noteGroups.count` is the number of consecutive data rows covered by the merged note cell. Counts must sum to the section's total data rows.
- If `note` is supplied instead of `noteGroups`, it spans the whole section.
- Palette names may be used as fill/font values; literal hex colors are also accepted.

## Manual Review Checklist

- Compare model names, capacities, each numeric price, adjustment text, and note text with the source.
- Ensure each section has a consistent column count.
- Ensure note-group counts match the affected rows.
- Ensure highlighted cells use the correct worksheet-column index.
