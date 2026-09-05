# bun-excel

[![CI](https://github.com/vanloctech/bun-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/vanloctech/bun-excel/actions/workflows/ci.yml)
[![npm version](https://img.shields.io/npm/v/bun-excel.svg)](https://www.npmjs.com/package/bun-excel)
[![GitHub stars](https://img.shields.io/github/stars/vanloctech/bun-excel?style=social)](https://github.com/vanloctech/bun-excel/stargazers)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Bun](https://img.shields.io/badge/Bun-%E2%89%A51.4-black?logo=bun)](https://bun.sh)
[![TypeScript](https://img.shields.io/badge/TypeScript-%E2%89%A55.0-blue?logo=typescript)](https://www.typescriptlang.org/)

[![English](https://img.shields.io/badge/lang-English-blue)](README.md) [![中文](https://img.shields.io/badge/lang-%E4%B8%AD%E6%96%87-red)](README.zh-CN.md)

一个高性能、针对 Bun 优化的 TypeScript Excel (.xlsx) 和 CSV 库。

> ⚠️ **Note**: 运行时说明：`bun-excel` 使用 `Bun.file()`、`Bun.write()` 和 `FileSink` 等 Bun 特有 API。它面向 Bun 运行时，不兼容 Node.js 或 Deno。
> 参与项目之前，请先阅读我们的 [Code of Conduct](CODE_OF_CONDUCT.md)。
> 某些安全扫描工具可能会标记包内的 `schemas.openxmlformats.org` 等 URL。它们是 Excel OOXML 格式要求的命名空间和 relationship 标识符，并不是运行时网络请求。

## 为什么使用这个包

- **为 Bun 而写，不是从 Node-first 抽象层改出来的** — 核心文件路径直接使用 `Bun.file()`、`Bun.write()`、`FileSink` 和 Bun 原生流式 API。
- **与 Bun 原生文件目标（包括 S3）配合自然** — 支持读取和写入本地路径、`Bun.file(...)` 以及 Bun `S3File` 对象，流式导出也可以直接写入 S3 目标。
- **面向 Bun backend 的生产导出 helper** — 支持进度回调、`AbortSignal`、导出诊断、流式 `Response` helper，以及通过 Bun 原生 writer 选项调优 S3 multipart 上传。
- **TypeScript 优先的表格模型** — `Workbook`、`Worksheet`、`Row`、`Cell` 以及样式对象都清晰、实用，适合 Bun 应用直接使用。
- **聚焦真实报表场景** — 样式、公式、超链接、数据验证、条件格式、自动筛选、冻结/拆分窗格以及工作簿元数据都已覆盖。
- **按工作负载选择写入策略** — 小文件可直接写，大文件可用流式或磁盘落地分块写入来降低内存压力。

## 安装

```bash
bun add bun-excel
```

要求 Bun 1.4.0 或更新版本。

## 快速开始

### 写入 Excel

```typescript
import { writeExcel, type Workbook } from "bun-excel";

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
import { readExcel } from "bun-excel";

const workbook = await readExcel("report.xlsx");
for (const sheet of workbook.worksheets) {
  console.log(`工作表: ${sheet.name}`);
  for (const row of sheet.rows) {
    console.log(row.cells.map(c => c.value).join(" | "));
  }
}
```

预览和选择性导入可使用 [`readExcelStream()`](DOCUMENT.zh-CN.md#readexcelstreamsource-options)，支持行范围、每个工作表的行数限制和列选择。

[`readExcel()`](DOCUMENT.zh-CN.md#readexcelsource-options) 可通过 `includeImages: false`、`includeComments: false` 和 `includeTables: false` 跳过图片、批注和表格。

### CSV

```typescript
import { readCSV, writeCSV } from "bun-excel";

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

于 2026-09-04 在 Bun `1.4.0` / macOS ARM64 上测量。各项数值为 3 次运行的中位数，每种模式使用独立进程。写入工作负载均导出单工作表的压缩 `.xlsx` 文件。

**1,000,000 行 × 10 列**

| 模式 | 总耗时 | 收尾耗时 | 每秒行数 | Peak RSS | 文件大小 |
| --- | ---: | ---: | ---: | ---: | ---: |
| `createExcelStream()` | `7.43s` | `4.01s` | `134,505` | `88.2 MiB` | `54.31 MiB` |
| `createChunkedExcelStream()` | `6.80s` | `3.28s` | `147,004` | `94.4 MiB` | `54.31 MiB` |

```bash
bun run benchmark:1m
```

**大型报表：30,000 行数据 × 30 列**，包含 `examples/large-report.ts` 中的样式、合并单元格和页脚公式。

| 方法 | 总耗时 | Peak RSS | 采样峰值 heapUsed | 文件大小 |
| --- | ---: | ---: | ---: | ---: |
| `writeExcel()` | `0.86s` | `337.0 MiB` | `121.8 MiB` | `5.90 MiB` |
| `createExcelStream()` | `0.94s` | `81.7 MiB` | `6.3 MiB` | `6.03 MiB` |
| `createChunkedExcelStream()` | `0.96s` | `82.5 MiB` | `8.5 MiB` | `6.03 MiB` |

```bash
bun run benchmark
```

**流式读取：`readExcelStream()`**

在同一 Bun `1.4.0` 运行时下，对比旧版读取器（自定义 XML 解析器）与当前 Bun XML 原生读取器。每个文件包含 30,009 行，含报表表头和页脚。每种实现使用独立进程交替运行 3 次，取中位数；耗时包含读取全部行并对 JSON 输出计算哈希。输出校验和一致。

| XLSX 文件 | 旧版耗时 | 原生版耗时 | 旧版 Peak RSS | 原生版 Peak RSS |
| --- | ---: | ---: | ---: | ---: |
| 共享字符串（`bench-normal.xlsx`） | `2.107s` | `1.709s` | `184.7 MiB` | `174.6 MiB` |
| 内联字符串（`bench-stream.xlsx`） | `2.214s` | `1.709s` | `153.5 MiB` | `149.4 MiB` |

该对比涵盖完整读取流程，包括 XML 分批处理和 ZIP 解压，并非仅比较 XML 解析器。

Peak RSS 为操作系统记录的进程内存峰值，包含运行时开销；heapUsed 通过采样测量，可能遗漏短暂峰值。内存和文件大小使用 MiB。结果受机器和系统负载影响；RSS 不可与之前同一进程内的内存增量直接比较。

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
