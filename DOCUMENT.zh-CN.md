# API 文档

bun-spreadsheet 完整 API 参考。

---

## 目录

- [Excel](#excel)
  - [writeExcel](#writeexcelpath-workbook-options)
  - [readExcel](#readexcelpath-options)
  - [buildExcelBuffer](#buildexcelbufferworkbook-options)
- [CSV](#csv)
  - [writeCSV](#writecsvpath-data-options)
  - [readCSV](#readcsvpath-options)
  - [readCSVStream](#readcsvstreampath-options)
  - [createCSVStream](#createcsvstreampath-options)
- [Excel 流式写入](#excel-流式写入)
  - [createExcelStream](#createexcelstreampath-options)
  - [createMultiSheetExcelStream](#createmultisheetexcelstreampath-options)
  - [createChunkedExcelStream](#createchunkedexcelstreampath-options)
- [类型定义](#类型定义)
  - [Workbook](#workbook)
  - [Worksheet](#worksheet)
  - [Row](#row)
  - [Cell](#cell)
  - [CellValue](#cellvalue)
  - [ColumnConfig](#columnconfig)
  - [MergeCell](#mergecell)
  - [Hyperlink](#hyperlink)
- [样式](#样式)
  - [CellStyle](#cellstyle)
  - [FontStyle](#fontstyle)
  - [FillStyle](#fillstyle)
  - [BorderStyle](#borderstyle)
  - [AlignmentStyle](#alignmentstyle)
  - [数字格式](#数字格式)
- [功能](#功能)
  - [公式](#公式)
  - [超链接](#超链接)
  - [合并单元格](#合并单元格)
  - [冻结窗格](#冻结窗格)
- [写入模式对比](#写入模式对比)

---

## Excel

### `writeExcel(path, workbook, options?)`

将 Workbook 写入 `.xlsx` 文件。

**参数：**

| 参数 | 类型 | 必填 | 描述 |
|------|------|------|------|
| `path` | `string` | 是 | 输出文件路径 |
| `workbook` | `Workbook` | 是 | 要写入的工作簿数据 |
| `options` | `ExcelWriteOptions` | 否 | 写入选项 |

**ExcelWriteOptions：**

| 选项 | 类型 | 默认值 | 描述 |
|------|------|--------|------|
| `creator` | `string` | `undefined` | 文件元数据中的作者名称 |
| `compress` | `boolean` | `true` | 启用 ZIP 压缩 |

**返回值：** `Promise<void>`

**示例：**

```typescript
import { writeExcel, type Workbook } from "bun-spreadsheet";

const workbook: Workbook = {
  worksheets: [{
    name: "Sheet1",
    columns: [{ width: 20 }, { width: 15 }],
    rows: [
      {
        cells: [
          { value: "姓名", style: { font: { bold: true } } },
          { value: "分数", style: { font: { bold: true } } },
        ],
      },
      { cells: [{ value: "小明" }, { value: 95 }] },
      { cells: [{ value: "小红" }, { value: 87 }] },
    ],
  }],
};

await writeExcel("report.xlsx", workbook, { creator: "我的应用" });
```

---

### `readExcel(path, options?)`

读取 `.xlsx` 文件并返回 Workbook 对象。

**参数：**

| 参数 | 类型 | 必填 | 描述 |
|------|------|------|------|
| `path` | `string` | 是 | .xlsx 文件路径 |
| `options` | `ExcelReadOptions` | 否 | 读取选项 |

**ExcelReadOptions：**

| 选项 | 类型 | 默认值 | 描述 |
|------|------|--------|------|
| `sheets` | `string[] \| number[]` | 所有工作表 | 指定要读取的工作表（按名称或索引） |
| `includeStyles` | `boolean` | `true` | 是否解析并包含单元格样式 |

**返回值：** `Promise<Workbook>`

**示例：**

```typescript
import { readExcel } from "bun-spreadsheet";

// 读取所有工作表
const workbook = await readExcel("report.xlsx");

// 仅读取指定工作表
const partial = await readExcel("report.xlsx", {
  sheets: ["Sheet1"],
  includeStyles: false,  // 不需要样式时更快
});

// 遍历数据
for (const sheet of workbook.worksheets) {
  console.log(`工作表: ${sheet.name}`);
  for (const row of sheet.rows) {
    console.log(row.cells.map(c => c.value).join(" | "));
  }
}
```

---

### `buildExcelBuffer(workbook, options?)`

将 Excel 文件构建为 `Uint8Array` 缓冲区，不写入磁盘。适用于 HTTP 响应或进一步处理。

**参数：** 与 `writeExcel` 相同，但没有 `path`。

**返回值：** `Uint8Array`

**示例：**

```typescript
import { buildExcelBuffer } from "bun-spreadsheet";

const buffer = buildExcelBuffer(workbook);
// 可用于 HTTP 响应、上传等
```

---

## CSV

### `writeCSV(path, data, options?)`

将数据写入 CSV 文件。

**参数：**

| 参数 | 类型 | 必填 | 描述 |
|------|------|------|------|
| `path` | `string` | 是 | 输出文件路径 |
| `data` | `Workbook \| CellValue[][]` | 是 | 要写入的数据 |
| `options` | `CSVWriteOptions` | 否 | 写入选项 |

**CSVWriteOptions：**

| 选项 | 类型 | 默认值 | 描述 |
|------|------|--------|------|
| `delimiter` | `string` | `","` | 字段分隔符 |
| `quoteChar` | `string` | `"\""` | 引号字符 |
| `lineEnding` | `string` | `"\n"` | 行尾符 |
| `includeHeader` | `boolean` | `false` | 是否包含表头行 |
| `headers` | `string[]` | `undefined` | 自定义表头名称 |
| `bom` | `boolean` | `false` | 添加 UTF-8 BOM（用于 Excel 兼容性） |

**返回值：** `Promise<void>`

**示例：**

```typescript
import { writeCSV } from "bun-spreadsheet";

// 简单数组数据
await writeCSV("data.csv", [
  ["姓名", "年龄", "城市"],
  ["小明", 28, "北京"],
  ["小红", 32, "上海"],
]);

// 带选项
await writeCSV("export.csv", data, {
  delimiter: ";",
  bom: true,           // Excel 兼容 UTF-8
  includeHeader: true,
  headers: ["ID", "名称", "值"],
});
```

---

### `readCSV(path, options?)`

读取 CSV 文件并返回 Workbook 对象（单工作表）。

**参数：**

| 参数 | 类型 | 必填 | 描述 |
|------|------|------|------|
| `path` | `string` | 是 | CSV 文件路径 |
| `options` | `CSVReadOptions` | 否 | 读取选项 |

**CSVReadOptions：**

| 选项 | 类型 | 默认值 | 描述 |
|------|------|--------|------|
| `delimiter` | `string` | `","` | 字段分隔符 |
| `quoteChar` | `string` | `"\""` | 引号字符 |
| `escapeChar` | `string` | `"\""` | 转义字符 |
| `hasHeader` | `boolean` | `false` | 第一行是否为表头 |
| `encoding` | `string` | `"utf-8"` | 文件编码 |
| `skipEmptyLines` | `boolean` | `false` | 跳过空行 |

**返回值：** `Promise<Workbook>`

自动类型检测：
- `"42"` -> `42`（数字）
- `"3.14"` -> `3.14`（数字）
- `"true"` / `"false"` -> `true` / `false`（布尔值）
- `"2024-01-15"` -> `Date` 对象
- 其他 -> `string`

---

### `readCSVStream(path, options?)`

逐行流式读取大型 CSV 文件。返回 `AsyncGenerator`。

**参数：** 与 `readCSV` 相同。

**返回值：** `AsyncGenerator<Row>`

**示例：**

```typescript
import { readCSVStream } from "bun-spreadsheet";

for await (const row of readCSVStream("large.csv")) {
  const values = row.cells.map(c => c.value);
  // 逐行处理，无需将整个文件加载到内存
}
```

---

### `createCSVStream(path, options?)`

创建流式 CSV 写入器。直接写入磁盘。

**参数：** 与 `writeCSV` 相同。

**返回值：** `CSVStreamWriter`

**CSVStreamWriter 方法：**

| 方法 | 描述 |
|------|------|
| `writeRow(values: CellValue[])` | 写入一行 |
| `flush()` | 刷新缓冲区到磁盘 |
| `end(): Promise<void>` | 完成并关闭文件 |

**示例：**

```typescript
import { createCSVStream } from "bun-spreadsheet";

const stream = createCSVStream("output.csv", {
  headers: ["ID", "名称", "值"],
  includeHeader: true,
});

for (let i = 0; i < 100000; i++) {
  stream.writeRow([i + 1, `项目_${i}`, Math.random() * 1000]);
}

await stream.end();
```

---

## Excel 流式写入

三种流式写入模式适用于不同场景：

| 模式 | 内存 | 适用场景 | 共享字符串 |
|------|------|----------|------------|
| `createExcelStream` | 中等 | 大多数场景（< 10 万行） | 是（内存中） |
| `createMultiSheetExcelStream` | 中等 | 多工作表 | 是（内存中） |
| `createChunkedExcelStream` | 恒定（低） | 超大文件（10 万+ 行） | 否（内联字符串） |

---

### `createExcelStream(path, options?)`

创建流式 Excel 写入器。立即将每行序列化为 XML，但共享字符串保留在内存中。

**ExcelStreamOptions：**

| 选项 | 类型 | 默认值 | 描述 |
|------|------|--------|------|
| `sheetName` | `string` | `"Sheet1"` | 工作表名称 |
| `columns` | `ColumnConfig[]` | `undefined` | 列宽配置 |
| `defaultRowHeight` | `number` | `15` | 默认行高 |
| `freezePane` | `{ row, col }` | `undefined` | 冻结窗格位置 |
| `mergeCells` | `MergeCell[]` | `undefined` | 合并单元格区域 |
| `creator` | `string` | `undefined` | 作者名称 |
| `compress` | `boolean` | `true` | 启用 ZIP 压缩 |

**返回值：** `ExcelStreamWriter`

**ExcelStreamWriter 方法：**

| 方法 | 描述 |
|------|------|
| `writeRow(row: Row \| CellValue[])` | 写入一行（Row 对象或纯数组） |
| `flush()` | 刷新缓冲区 |
| `end(): Promise<void>` | 完成 ZIP 并写入磁盘 |

**示例：**

```typescript
import { createExcelStream } from "bun-spreadsheet";

const stream = createExcelStream("report.xlsx", {
  sheetName: "数据",
  columns: [{ width: 10 }, { width: 25 }, { width: 15 }],
  freezePane: { row: 1, col: 0 },
});

// 带样式的表头行（使用 Row 对象）
stream.writeRow({
  cells: [
    { value: "ID", style: { font: { bold: true } } },
    { value: "名称", style: { font: { bold: true } } },
    { value: "价格", style: { font: { bold: true } } },
  ],
});

// 数据行（使用纯数组更方便）
for (let i = 0; i < 50000; i++) {
  stream.writeRow([i + 1, `产品_${i}`, Math.random() * 1000]);
}

await stream.end();
```

---

### `createMultiSheetExcelStream(path, options?)`

创建支持多工作表的流式 Excel 写入器。

**返回值：** `MultiSheetExcelStreamWriter`

**MultiSheetExcelStreamWriter 方法：**

| 方法 | 描述 |
|------|------|
| `addSheet(name, options?)` | 添加新工作表并切换到该表 |
| `writeRow(row: Row \| CellValue[])` | 向当前工作表写入一行 |
| `flush()` | 刷新缓冲区 |
| `end(): Promise<void>` | 完成 |

**示例：**

```typescript
import { createMultiSheetExcelStream } from "bun-spreadsheet";

const stream = createMultiSheetExcelStream("multi.xlsx");

stream.addSheet("收入", {
  columns: [{ width: 15 }, { width: 12 }],
  freezePane: { row: 1, col: 0 },
});
stream.writeRow(["月份", "金额"]);
stream.writeRow(["一月", 50000]);
stream.writeRow(["二月", 62000]);

stream.addSheet("支出", {
  columns: [{ width: 15 }, { width: 12 }],
});
stream.writeRow(["类别", "金额"]);
stream.writeRow(["工资", 30000]);

await stream.end();
```

---

### `createChunkedExcelStream(path, options?)`

创建恒定内存的分块流式 Excel 写入器。行 XML 被写入磁盘上的临时文件，结束时组装为 ZIP。

**ChunkedExcelStreamOptions：** 与 `ExcelStreamOptions` 相同。

**返回值：** `ExcelChunkedStreamWriter`

**ExcelChunkedStreamWriter 方法：**

| 方法 | 描述 |
|------|------|
| `writeRow(row: Row \| CellValue[])` | 写入一行（立即序列化到磁盘） |
| `writeStyledRow(values, styles)` | 写入带逐单元格样式的行 |
| `writeRows(rows)` | 一次写入多行 |
| `flush()` | 刷新临时文件缓冲区 |
| `end(): Promise<void>` | 从临时文件组装 ZIP 并写入输出 |
| `currentRowCount` | 获取当前行数 |

**工作原理：**

1. `writeRow()` — 序列化行 XML 并写入磁盘临时文件（不占用内存）
2. `end()` — 读取临时文件、包装工作表 XML、创建 ZIP、写入输出、删除临时文件

使用**内联字符串**（`<is><t>...</t></is>`）代替共享字符串表，无需在内存中跟踪字符串。

**示例：**

```typescript
import { createChunkedExcelStream } from "bun-spreadsheet";

const stream = createChunkedExcelStream("huge_report.xlsx", {
  sheetName: "报表",
  columns: [{ width: 14 }, { width: 20 }, { width: 12 }],
  freezePane: { row: 1, col: 0 },
});

// 100 万行 — 内存保持恒定
for (let i = 0; i < 1_000_000; i++) {
  stream.writeRow([i, `行 ${i}`, Math.random() * 10000]);
}

await stream.end();
```

---

## 类型定义

### Workbook

```typescript
interface Workbook {
  worksheets: Worksheet[];
  creator?: string;      // 作者名称
  created?: Date;        // 创建时间
  modified?: Date;       // 修改时间
}
```

### Worksheet

```typescript
interface Worksheet {
  name: string;                              // 工作表名称
  rows: Row[];                               // 行数组
  columns?: ColumnConfig[];                  // 列配置
  mergeCells?: MergeCell[];                  // 合并单元格区域
  freezePane?: { row: number; col: number }; // 冻结窗格位置
  defaultRowHeight?: number;                 // 默认行高
  defaultColWidth?: number;                  // 默认列宽
}
```

### Row

```typescript
interface Row {
  cells: Cell[];       // 该行的单元格数组
  height?: number;     // 自定义行高
  style?: CellStyle;   // 该行所有单元格的默认样式
}
```

### Cell

```typescript
interface Cell {
  value: CellValue;                                // 单元格值
  style?: CellStyle;                               // 单元格样式
  type?: "string" | "number" | "boolean" | "date" | "formula";
  formula?: string;                                // 公式（不含 "="）
  formulaResult?: string | number | boolean;       // 公式缓存结果
  hyperlink?: Hyperlink;                           // 单元格超链接
}
```

### CellValue

```typescript
type CellValue = string | number | boolean | Date | null | undefined;
```

### ColumnConfig

```typescript
interface ColumnConfig {
  width?: number;      // 列宽（字符数）
  style?: CellStyle;   // 默认列样式
  header?: string;     // 表头文字
}
```

### MergeCell

```typescript
interface MergeCell {
  startRow: number;    // 起始行（0 索引）
  startCol: number;    // 起始列（0 索引）
  endRow: number;      // 结束行（0 索引）
  endCol: number;      // 结束列（0 索引）
}
```

### Hyperlink

```typescript
interface Hyperlink {
  target: string;      // URL、mailto: 或内部引用如 "Sheet2!A1"
  tooltip?: string;    // 悬停时的提示文字
}
```

---

## 样式

### CellStyle

```typescript
interface CellStyle {
  font?: FontStyle;
  fill?: FillStyle;
  border?: BorderStyle;
  alignment?: AlignmentStyle;
  numberFormat?: string;       // 如 "#,##0.00"、"yyyy-mm-dd"
}
```

### FontStyle

```typescript
interface FontStyle {
  name?: string;       // 字体名称，如 "Arial"、"Calibri"
  size?: number;       // 字体大小（磅）
  bold?: boolean;      // 加粗
  italic?: boolean;    // 斜体
  underline?: boolean; // 下划线
  strike?: boolean;    // 删除线
  color?: string;      // 十六进制颜色（不含 #），如 "FF0000" 表示红色
}
```

**示例：**

```typescript
// 红色加粗表头
{ font: { bold: true, color: "FF0000", size: 14 } }

// 指定字体的斜体
{ font: { italic: true, name: "Times New Roman" } }
```

### FillStyle

```typescript
interface FillStyle {
  type: "pattern" | "gradient";
  pattern?: "solid" | "darkGray" | "mediumGray" | "lightGray" | "none";
  fgColor?: string;    // 前景色（十六进制）
  bgColor?: string;    // 背景色（十六进制）
}
```

**示例：**

```typescript
// 黄色背景
{ fill: { type: "pattern", pattern: "solid", fgColor: "FFFF00" } }

// 浅灰色
{ fill: { type: "pattern", pattern: "solid", fgColor: "D9D9D9" } }
```

### BorderStyle

```typescript
interface BorderStyle {
  top?: BorderEdgeStyle;
  bottom?: BorderEdgeStyle;
  left?: BorderEdgeStyle;
  right?: BorderEdgeStyle;
}

interface BorderEdgeStyle {
  style?: "thin" | "medium" | "thick" | "dotted" | "dashed" | "double"
        | "hair" | "dashDot" | "dashDotDot" | "mediumDashed"
        | "mediumDashDot" | "mediumDashDotDot" | "slantDashDot";
  color?: string;      // 十六进制颜色
}
```

**示例：**

```typescript
// 四边细边框
{
  border: {
    top: { style: "thin", color: "000000" },
    bottom: { style: "thin", color: "000000" },
    left: { style: "thin", color: "000000" },
    right: { style: "thin", color: "000000" },
  }
}
```

### AlignmentStyle

```typescript
interface AlignmentStyle {
  horizontal?: "left" | "center" | "right" | "fill" | "justify";
  vertical?: "top" | "center" | "bottom";
  wrapText?: boolean;       // 自动换行
  textRotation?: number;    // 文字旋转角度（0-180）
  indent?: number;          // 缩进级别
}
```

### 数字格式

设置 CellStyle 上的 `numberFormat` 控制数字和日期的显示方式。

**常用格式：**

| 格式 | 示例输出 | 描述 |
|------|----------|------|
| `"0"` | `1234` | 整数 |
| `"0.00"` | `1234.56` | 2 位小数 |
| `"#,##0"` | `1,234` | 千位分隔符 |
| `"#,##0.00"` | `1,234.56` | 千位分隔符 + 小数 |
| `"0%"` | `12%` | 百分比 |
| `"0.00%"` | `12.34%` | 带小数的百分比 |
| `"$#,##0.00"` | `$1,234.56` | 货币 |
| `"yyyy-mm-dd"` | `2024-01-15` | 日期（ISO） |
| `"dd/mm/yyyy"` | `15/01/2024` | 日期（欧洲） |
| `"hh:mm:ss"` | `14:30:00` | 时间 |

---

## 功能

### 公式

写入带有可选缓存结果的公式。缓存结果在 Excel 重新计算之前显示。

```typescript
{
  cells: [
    { value: 10 },
    { value: 20 },
    { value: 30 },
    {
      value: null,
      formula: "SUM(A1:C1)",        // 公式（不含前导 "="）
      formulaResult: 60,             // 缓存结果
      style: { numberFormat: "#,##0" },
    },
  ],
}
```

支持所有标准 Excel 函数：SUM、AVERAGE、IF、MAX、MIN、COUNT、VLOOKUP 等。

---

### 超链接

三种类型的超链接：

```typescript
// 1. 外部 URL
{
  value: "访问网站",
  hyperlink: { target: "https://example.com", tooltip: "点击打开" },
  style: { font: { color: "0563C1", underline: true } },
}

// 2. 邮件
{
  value: "联系我们",
  hyperlink: { target: "mailto:hello@example.com" },
}

// 3. 内部工作表引用
{
  value: "跳转到汇总",
  hyperlink: { target: "Sheet2!A1" },
}
```

---

### 合并单元格

合并一组单元格。仅左上角单元格的值和样式可见。

```typescript
const worksheet: Worksheet = {
  name: "报表",
  rows: [
    // 第 0 行：合并标题，横跨 A1:F1
    {
      cells: [
        { value: "2024 年度报告", style: { font: { bold: true, size: 16 } } },
        { value: null }, { value: null }, { value: null }, { value: null }, { value: null },
      ],
    },
  ],
  mergeCells: [
    { startRow: 0, startCol: 0, endRow: 0, endCol: 5 },  // A1:F1
  ],
};
```

> 注意：所有行/列索引均为 0 基准。

---

### 冻结窗格

冻结行和/或列，使其在滚动时保持可见。

```typescript
const worksheet: Worksheet = {
  name: "数据",
  rows: [/* ... */],
  freezePane: { row: 1, col: 0 },  // 冻结第一行（表头）
};
```

| 值 | 效果 |
|----|------|
| `{ row: 1, col: 0 }` | 冻结第一行 |
| `{ row: 0, col: 1 }` | 冻结第一列 |
| `{ row: 1, col: 1 }` | 冻结第一行和第一列 |
| `{ row: 2, col: 0 }` | 冻结前 2 行 |

---

## 写入模式对比

| 功能 | `writeExcel` | `createExcelStream` | `createChunkedExcelStream` |
|------|-------------|--------------------|-----------------------------|
| 内存 | 整个工作簿在内存中 | 行 XML 缓冲区在内存中 | 恒定（低） |
| 共享字符串 | 是 | 是 | 否（内联） |
| 多工作表 | 是 | 单工作表 | 单工作表 |
| 多工作表支持 | 通过 Workbook | `createMultiSheetExcelStream` | 不支持 |
| 样式 | 完整支持 | 完整支持 | 完整支持 |
| 公式 | 完整支持 | 完整支持 | 完整支持 |
| 超链接 | 完整支持 | 完整支持 | 完整支持 |
| 合并单元格 | 完整支持 | 完整支持 | 完整支持 |
| 冻结窗格 | 完整支持 | 完整支持 | 完整支持 |
| 适用场景 | 中小型文件 | 中大型文件 | 超大文件（10 万+） |

**如何选择：**

- **`writeExcel`** — 所有数据都已在内存中。最简单的 API。
- **`createExcelStream`** — 数据逐行生成（如来自数据库查询）。功能和内存的良好平衡。
- **`createChunkedExcelStream`** — 需要关注内存的超大文件。以磁盘 I/O 换取恒定内存使用。
