import { createHash } from 'node:crypto';
import { type CellValue, readExcelStream, readExcelValuesStream } from '../src';

const [, , file = 'output/bench-normal.xlsx', mode] = process.argv;
if (mode) {
  if (mode !== 'rows' && mode !== 'values') throw new Error('Unknown mode');
  const hash = createHash('sha256');
  let count = 0;
  const start = performance.now();
  function consume(sheetIndex: number, rowIndex: number, values: CellValue[]) {
    hash.update(JSON.stringify([sheetIndex, rowIndex, values]));
    count++;
  }
  if (mode === 'rows') {
    for await (const item of readExcelStream(file)) {
      consume(
        item.sheetIndex,
        item.rowIndex,
        Array.from(item.row.cells, (cell) => cell?.value ?? null),
      );
    }
  } else {
    for await (const batch of readExcelValuesStream(file)) {
      for (let i = 0; i < batch.rows.length; i++)
        consume(batch.sheetIndex, batch.rowIndices[i], batch.rows[i]);
    }
  }
  console.log(
    JSON.stringify({
      mode,
      milliseconds: performance.now() - start,
      peakRssMiB: process.resourceUsage().maxRSS / 1024,
      rows: count,
      checksum: hash.digest('hex'),
    }),
  );
} else {
  const results: Record<
    string,
    { milliseconds: number; peakRssMiB: number }[]
  > = { rows: [], values: [] };
  let checksum: string | undefined;
  console.log(
    `Bun ${Bun.version}, ${process.platform}/${process.arch}, ${file}`,
  );
  for (let run = 0; run < 3; run++) {
    for (const childMode of run % 2 ? ['values', 'rows'] : ['rows', 'values']) {
      const child = Bun.spawn(
        [process.execPath, import.meta.path, file, childMode],
        { stdout: 'pipe', stderr: 'inherit' },
      );
      const output = await new Response(child.stdout).text();
      if (await child.exited)
        throw new Error(
          'Benchmark failed; generate fixtures with bun run benchmark first',
        );
      const result = JSON.parse(output);
      checksum ??= result.checksum;
      if (checksum !== result.checksum) throw new Error('Values do not match');
      results[childMode].push(result);
      console.log(output.trim());
    }
  }
  console.log('| Reader | Median time | Median peak RSS |');
  console.log('| --- | ---: | ---: |');
  for (const [reader, runs] of Object.entries(results)) {
    const times = runs.map((run) => run.milliseconds).sort((a, b) => a - b);
    const memory = runs.map((run) => run.peakRssMiB).sort((a, b) => a - b);
    console.log(
      `| ${reader} | ${times[1].toFixed(1)} ms | ${memory[1].toFixed(1)} MiB |`,
    );
  }
}
