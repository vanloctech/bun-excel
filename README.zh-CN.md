# bun-spreadsheet

[![CI](https://github.com/vanloctech/bun-spreadsheet/actions/workflows/ci.yml/badge.svg)](https://github.com/vanloctech/bun-spreadsheet/actions/workflows/ci.yml)
[![npm version](https://img.shields.io/npm/v/bun-spreadsheet.svg)](https://www.npmjs.com/package/bun-spreadsheet)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Bun](https://img.shields.io/badge/Bun-%E2%89%A51.0-black?logo=bun)](https://bun.sh)
[![TypeScript](https://img.shields.io/badge/TypeScript-%E2%89%A55.0-blue?logo=typescript)](https://www.typescriptlang.org/)

[![English](https://img.shields.io/badge/lang-English-blue)](README.md) [![中文](https://img.shields.io/badge/lang-%E4%B8%AD%E6%96%87-red)](README.zh-CN.md)

一个高性能、针对 Bun 优化的 TypeScript Excel (.xlsx) 和 CSV 库。

> 运行时说明：`bun-spreadsheet` 使用 `Bun.file()`、`Bun.write()` 和 `FileSink` 等 Bun 特有 API。它面向 Bun 运行时，不兼容 Node.js 或 Deno。

## 特性

- **Bun 原生** — 基于 `Bun.file()`、`Bun.write()` 和 `FileSink` 构建，性能最大化
- **Excel (.xlsx)** — 完整读写支持：样式、公式、超链接、合并单元格、冻结窗格
- **CSV** — 读写支持自动类型检测和流式处理
- **3 种写入模式** — 普通写入、流式写入（内存友好）、分块流式写入（恒定内存）
- **丰富样式** — 字体、填充、边框、对齐、数字格式
- **超链接** — 外部 URL、邮件、内部工作表引用
- **公式** — 支持缓存结果的读写（SUM、AVERAGE、IF 等）
- **数据验证** — 下拉列表、数字范围、日期限制、自定义公式
- **安全加固** — XML 炸弹防护、路径遍历保护、输入验证
- **最少依赖** — 仅依赖 [fflate](https://github.com/101arrowz/fflate) 用于 ZIP 压缩

## 安装

```bash
bun add bun-spreadsheet
```

## 快速开始

### 写入 Excel

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

await writeExcel("report.xlsx", workbook);
```

### 读取 Excel

```typescript
import { readExcel } from "bun-spreadsheet";

const workbook = await readExcel("report.xlsx");
for (const sheet of workbook.worksheets) {
  console.log(`工作表: ${sheet.name}`);
  for (const row of sheet.rows) {
    console.log(row.cells.map(c => c.value).join(" | "));
  }
}
```

### CSV

```typescript
import { readCSV, writeCSV } from "bun-spreadsheet";

// 写入
await writeCSV("data.csv", [
  ["姓名", "年龄"],
  ["小明", 28],
]);

// 读取
const csv = await readCSV("data.csv");
```

## 文档

完整 API 参考请查看 [DOCUMENT.zh-CN.md](DOCUMENT.zh-CN.md)，包括：

- 所有函数（`writeExcel`、`readExcel`、`writeCSV`、`readCSV`、流式 API）
- 类型定义（`Workbook`、`Worksheet`、`Cell`、`Row` 等）
- 样式指南（字体、填充、边框、对齐、数字格式）
- 功能说明（公式、超链接、合并单元格、冻结窗格、数据验证）
- 写入模式对比（普通 vs 流式 vs 分块流式）

## 性能测试

30 列 x 30,000 行，包含样式和公式（MacOS, Bun 1.3）：

| 方法 | 总耗时 | 堆内存 | 文件大小 |
|------|--------|--------|----------|
| `writeExcel` | ~1,500ms | ~129MB | 6.85MB |
| `createExcelStream` | ~1,460ms | ~163MB | 6.85MB |
| `createChunkedExcelStream` | ~1,325ms | ~75MB | 6.63MB |

自行运行性能测试：

```bash
bun run benchmark
```

## 示例

```bash
# 运行所有示例
bun run demo

# 大型报表 (30 列 x 30K 行)
bun run large-report

# 性能测试
bun run benchmark
```

## 安全性

本库已进行安全加固：

- **XML 炸弹防护** — 深度限制、节点数量上限、输入大小验证
- **路径遍历保护** — `path.resolve()` + 所有文件路径的空字节检查
- **Zip slip 防护** — 验证 ZIP 归档中的所有路径
- **输入验证** — 最大行数 (1M)、最大列数 (16K)、最大文件大小 (200MB)
- **XML 注入防护** — 所有用户值均经过正确转义
- **原型污染防护** — 动态映射使用 `Object.create(null)`

## 贡献

请查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解开发设置和指南。

## 许可证

[MIT](LICENSE)
