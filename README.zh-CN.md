# bun-spreadsheet

[![CI](https://github.com/vanloctech/bun-spreadsheet/actions/workflows/ci.yml/badge.svg)](https://github.com/vanloctech/bun-spreadsheet/actions/workflows/ci.yml)
[![npm version](https://img.shields.io/npm/v/bun-spreadsheet.svg)](https://www.npmjs.com/package/bun-spreadsheet)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Bun](https://img.shields.io/badge/Bun-%E2%89%A51.0-black?logo=bun)](https://bun.sh)
[![TypeScript](https://img.shields.io/badge/TypeScript-%E2%89%A55.0-blue?logo=typescript)](https://www.typescriptlang.org/)

[![English](https://img.shields.io/badge/lang-English-blue)](README.md) [![中文](https://img.shields.io/badge/lang-%E4%B8%AD%E6%96%87-red)](README.zh-CN.md)

一个高性能、针对 Bun 优化的 TypeScript Excel (.xlsx) 和 CSV 库。

> ⚠️ **Note**: 运行时说明：`bun-spreadsheet` 使用 `Bun.file()`、`Bun.write()` 和 `FileSink` 等 Bun 特有 API。它面向 Bun 运行时，不兼容 Node.js 或 Deno。

## 为什么使用这个包

- **为 Bun 而写，不是从 Node-first 抽象层改出来的** — 核心文件路径直接使用 `Bun.file()`、`Bun.write()`、`FileSink` 和 Bun 原生流式 API。
- **与 Bun 原生文件目标（包括 S3）配合自然** — 支持读取和写入本地路径、`Bun.file(...)` 以及 Bun `S3File` 对象，流式导出也可以直接写入 S3 目标。
- **面向 Bun backend 的生产导出 helper** — 支持进度回调、`AbortSignal`、导出诊断、流式 `Response` helper，以及通过 Bun 原生 writer 选项调优 S3 multipart 上传。
- **TypeScript 优先的表格模型** — `Workbook`、`Worksheet`、`Row`、`Cell` 以及样式对象都清晰、实用，适合 Bun 应用直接使用。
- **聚焦真实报表场景** — 样式、公式、超链接、数据验证、条件格式、自动筛选、冻结/拆分窗格以及工作簿元数据都已覆盖。
- **按工作负载选择写入策略** — 小文件可直接写，大文件可用流式或磁盘落地分块写入来降低内存压力。

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

以下数据是在 Bun `1.3.10` / `darwin arm64` 环境下测得，测试场景为单工作表、压缩 `.xlsx`、`1,000,000` 行 x `10` 列：

| 模式 | 总耗时 | 收尾耗时 | 每秒行数 | Peak RSS | Peak heapUsed | 文件大小 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `createExcelStream()` | `13.1s` | `8.9s` | `76,363` | `110.6MB` | `5.1MB` | `54.33MB` |
| `createChunkedExcelStream()` | `11.9s` | `8.5s` | `84,029` | `120.9MB` | `5.1MB` | `54.33MB` |

当前版本中，单工作表的 `createExcelStream()` 已经走与 chunked writer 相同的磁盘落地低内存路径，所以两者结果接近是正常的。你可以通过下面的命令在自己的机器上重跑这个大数据量 benchmark：

```bash
bun run benchmark:1m
```

如果你想看普通写入、流式写入和分块磁盘写入三种模式在真实 `large-report` 工作负载下的对比，下面这组数据是在 Bun `1.3.10` / `MacOS ARM`、单工作表、压缩 `.xlsx`、`30` 列 x `30,000` 行，并使用与 `examples/large-report.ts` 相同的样式、合并单元格和页脚公式的条件下测得：

| 方法 | 总耗时 | Peak RSS 增量 | Peak heapUsed 增量 | 文件大小 |
| --- | ---: | ---: | ---: | ---: |
| `writeExcel()` | `1.92s` | `518.6MB` | `154.0MB` | `6.20MB` |
| `createExcelStream()` | `1.98s` | `48.0MB` | `39.2MB` | `6.35MB` |
| `createChunkedExcelStream()` | `2.01s` | `2.5MB` | `3.1MB` | `6.35MB` |

这里的内存列表示 benchmark 运行过程中相对基线进程内存的峰值增量，不是写入前后内存差值。

你可以通过下面的命令在自己的机器上重跑这个 benchmark：

```bash
bun run benchmark
```

## 示例

```bash
# 性能测试（普通 vs 流式 vs 分块流式）
bun run benchmark

# 1M 行 Excel 性能测试（流式 vs 分块流式）
bun run benchmark:1m
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
