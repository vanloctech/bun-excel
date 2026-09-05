import { createNativeInflaters } from './native-inflate';
// ============================================
// XLSX Reader — Bun-optimized Excel file reader
// ============================================

import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Unzip, unzipSync } from 'fflate';
import {
  describeFileSource,
  getRuntimeFileSize,
  toReadableFile,
} from '../runtime-io';
import type {
  AlignmentStyle,
  BorderEdgeStyle,
  BorderStyle,
  Cell,
  CellStyle,
  CellValue,
  ColumnConfig,
  DefinedName,
  ExcelReadOptions,
  ExcelReadStreamOptions,
  ExcelReadStreamRow,
  FileSource,
  FillStyle,
  FontStyle,
  HeaderFooter,
  PageMargins,
  PageSetup,
  RichTextRun,
  Row,
  Workbook,
  WorkbookView,
  Worksheet,
  WorksheetProtection,
} from '../types';
import { parseAutoFilter } from './auto-filter';
import { parseCommentsXML } from './comments';
import { parseConditionalFormattings } from './conditional-formatting';
import { parseDataValidations } from './data-validation';
import { parseDrawingImages } from './images';
import {
  elementChildren,
  findChild,
  findChildren,
  getTextContent,
  MAX_XML_SIZE,
  parseXML,
  type XMLNode,
} from './native-xml';
import { createTempRuntimeId } from './runtime-utils';
import { parseTableXML } from './tables';
import { excelSerialToDate } from './xlsx-writer';
import { letterToColIndex, parseCellRef } from './xml-builder';
import { streamNativeElements } from './xml-elements';

// Top-level regex for performance (biome: useTopLevelRegex)
const CELL_REF_REGEX = /^([A-Z]+)(\d+)$/;
const RGB_PREFIX_REGEX = /^FF/;
const COL_LETTER_REGEX = /^([A-Z]+)/;
const WORKSHEET_ENTRY_REGEX = /^xl\/worksheets\/[^/]+\.xml$/;
const LEADING_EQUALS_REGEX = /^=+/;
const LEADING_SINGLE_QUOTE_REGEX = /^'/;
const TRAILING_SINGLE_QUOTE_REGEX = /'$/;
const HEADER_FOOTER_LEFT_REGEX = /&L([^&]*)/;
const HEADER_FOOTER_CENTER_REGEX = /&C([^&]*)/;
const HEADER_FOOTER_RIGHT_REGEX = /&R([^&]*)/;
const BUILTIN_NUMFMTS: Record<number, string> = {
  14: 'yyyy-mm-dd',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
};

/** Security limits */
const MAX_FILE_SIZE = 200 * 1024 * 1024; // 200MB max file size
const MAX_DECOMPRESSED_SIZE = 1024 * 1024 * 1024; // 1GB max decompressed size
const MAX_ZIP_ENTRIES = 10_000; // 10K max entries in ZIP
const MAX_ROWS = 1_048_576; // Excel max rows
const MAX_COLS = 16_384; // Excel max columns (XFD)
const MAX_SHARED_STRINGS = 10_000_000; // 10M max shared strings

interface SheetRelationships {
  hyperlinks: Map<string, string>;
  commentsPath?: string;
  drawingPath?: string;
  tablePaths: string[];
}

interface StreamSheetFile {
  entryPath: string;
  tempPath: string;
}

interface StreamSheetTempResource extends StreamSheetFile {
  writer: Bun.FileSink;
}

interface WorkbookSheetDescriptor {
  index: number;
  name: string;
  path: string;
}

interface StreamRowSelection {
  startRow: number;
  endRow: number;
  maxRows: number;
  columns?: ReadonlySet<number>;
}

function validateReadIndex(
  value: number,
  name: string,
  maximum: number,
): number {
  if (!Number.isInteger(value) || value < 0 || value > maximum)
    throw new RangeError(`${name} must be an integer between 0 and ${maximum}`);
  return value;
}

function streamRowSelection(
  options?: ExcelReadStreamOptions,
): StreamRowSelection {
  const startRow = validateReadIndex(
    options?.startRow === undefined ? 0 : options.startRow,
    'startRow',
    MAX_ROWS - 1,
  );
  const endRow = validateReadIndex(
    options?.endRow === undefined ? MAX_ROWS - 1 : options.endRow,
    'endRow',
    MAX_ROWS - 1,
  );
  if (endRow < startRow)
    throw new RangeError('endRow must be greater than or equal to startRow');
  const maxRows =
    options?.maxRows === undefined
      ? Number.POSITIVE_INFINITY
      : validateReadIndex(options.maxRows, 'maxRows', MAX_ROWS);
  let columns: Set<number> | undefined;
  if (options?.columns !== undefined) {
    if (!Array.isArray(options.columns))
      throw new TypeError('columns must be an array of column indices');
    columns = new Set();
    for (const column of options.columns)
      columns.add(validateReadIndex(column, 'columns entry', MAX_COLS - 1));
  }
  return { startRow, endRow, maxRows, columns };
}

function parseRangeRef(rangeRef: string) {
  const [startRef, endRef = startRef] = rangeRef.split(':');
  const start = parseCellRef(startRef.replace(/\$/g, ''));
  const end = parseCellRef(endRef.replace(/\$/g, ''));
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

function resolveSheetPartPath(baseDir: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1);
  }

  const segments = baseDir.split('/').filter(Boolean);
  for (const part of target.split('/')) {
    if (part === '..') {
      segments.pop();
    } else if (part !== '.' && part !== '') {
      segments.push(part);
    }
  }
  return segments.join('/');
}

function concatUint8Arrays(chunks: Uint8Array[]): Uint8Array {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const combined = new Uint8Array(totalLength);
  let offset = 0;
  for (const chunk of chunks) {
    combined.set(chunk, offset);
    offset += chunk.length;
  }
  return combined;
}

function shouldBufferStreamZipEntry(
  path: string,
  includeStyles: boolean,
): boolean {
  return (
    path === 'xl/workbook.xml' ||
    path === 'xl/_rels/workbook.xml.rels' ||
    (includeStyles && path === 'xl/styles.xml')
  );
}

function shouldSpoolWorksheetEntry(path: string): boolean {
  return WORKSHEET_ENTRY_REGEX.test(path);
}

async function cleanupStreamSheetTempResources(
  resources: StreamSheetTempResource[],
): Promise<void> {
  await Promise.allSettled(
    resources.map(async ({ writer }) => {
      try {
        const result = writer.end();
        if (result instanceof Promise) {
          await result;
        }
      } catch {
        // Ignore temp writer close errors during cleanup.
      }
    }),
  );

  await Promise.allSettled(
    resources.map(async ({ tempPath }) => {
      try {
        await Bun.file(tempPath).delete();
      } catch {
        // Ignore temp file deletion errors during cleanup.
      }
    }),
  );
}

async function unzipXlsxForStreaming(
  source: FileSource,
  {
    includeStyles = true,
    selectedPaths,
    metadataOnly = false,
    selection,
  }: {
    includeStyles?: boolean;
    selectedPaths?: ReadonlySet<string>;
    metadataOnly?: boolean;
    selection?: ExcelReadOptions;
  } = {},
): Promise<{
  bufferedEntries: Record<string, Uint8Array>;
  sheetFiles: StreamSheetFile[];
}> {
  const inflaters = createNativeInflaters();
  const bufferedEntries: Record<string, Uint8Array> = {};
  const sheetFiles: StreamSheetFile[] = [];
  const resources: StreamSheetTempResource[] = [];
  const activeWriters = new Set<Bun.FileSink>();
  let endings: Promise<unknown>[] = [];
  let totalDeclared = 0;
  let totalActual = 0;
  let entries = 0;
  let incomplete = 0;
  let metadataChecked = false;
  const unzip = new Unzip((entry) => {
    if (++entries > MAX_ZIP_ENTRIES)
      throw new Error('ZIP has too many entries');
    if (
      entry.name.startsWith('/') ||
      entry.name.startsWith('\\') ||
      entry.name.includes('..')
    ) {
      throw new Error(
        `Malicious zip entry detected: "${entry.name}" — potential Zip Slip attack`,
      );
    }
    totalDeclared += entry.originalSize ?? 0;
    if (totalDeclared > MAX_DECOMPRESSED_SIZE)
      throw new Error('Declared decompressed size exceeds limit');
    const buffered = shouldBufferStreamZipEntry(
      entry.name,
      includeStyles && !metadataOnly,
    );
    const spool =
      !metadataOnly &&
      ((shouldSpoolWorksheetEntry(entry.name) &&
        (!selectedPaths || selectedPaths.has(entry.name))) ||
        entry.name === 'xl/sharedStrings.xml');
    if (!buffered && !spool) return;
    incomplete++;
    const chunks: Uint8Array[] = [];
    let entrySize = 0;
    let writer: Bun.FileSink | undefined;
    if (spool) {
      const tempPath = join(
        tmpdir(),
        `bun-excel-stream-${createTempRuntimeId()}.xml`,
      );
      writer = Bun.file(tempPath).writer({ highWaterMark: 64 * 1024 });
      const resource = { entryPath: entry.name, tempPath, writer };
      resources.push(resource);
      sheetFiles.push(resource);
      activeWriters.add(writer);
    }
    entry.ondata = (error, chunk, final) => {
      if (error) throw error;
      totalActual += chunk.length;
      entrySize += chunk.length;
      if (totalActual > MAX_DECOMPRESSED_SIZE)
        throw new Error('Actual decompressed size exceeds limit');
      if (buffered && entrySize > MAX_XML_SIZE)
        throw new Error('XML metadata exceeds size limit');
      if (writer) writer.write(chunk);
      else if (chunk.length) chunks.push(chunk);
      if (final) {
        incomplete--;
        if (writer) {
          activeWriters.delete(writer);
          endings.push(Promise.resolve(writer.end()));
        } else {
          bufferedEntries[entry.name] = concatUint8Arrays(chunks);
          chunks.length = 0;
        }
      }
    };
    entry.start();
  });
  unzip.register(inflaters.decoder);
  const reader = toReadableFile(source).stream().getReader();
  async function drain(): Promise<void> {
    await inflaters.drain();
    const pending = endings;
    endings = [];
    for (const writer of activeWriters)
      pending.push(Promise.resolve(writer.flush()));
    await Promise.all(pending);
  }
  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      // Bound both decompression bursts and queued sink data even if the source
      // returns a large chunk. No Promise closure retains every output chunk.
      for (let offset = 0; offset < value.length; offset += 64 * 1024) {
        unzip.push(value.subarray(offset, offset + 64 * 1024));
        await drain();
        if (
          metadataOnly &&
          !metadataChecked &&
          !incomplete &&
          bufferedEntries['xl/workbook.xml'] &&
          bufferedEntries['xl/_rels/workbook.xml.rels']
        ) {
          metadataChecked = true;
          // The extraction pass validates the full archive. Stop this lookup
          // early when there is a selected sheet; empty selections still scan
          // to the end so excluded entries receive the same ZIP checks.
          if (getStreamDescriptors(bufferedEntries, selection).length)
            return { bufferedEntries, sheetFiles };
        }
      }
    }
    unzip.push(new Uint8Array(0), true);
    await drain();
    if (incomplete) throw new Error('Truncated ZIP entry');
    return { bufferedEntries, sheetFiles };
  } catch (error) {
    await inflaters.abort();
    await Promise.allSettled(endings);
    await cleanupStreamSheetTempResources(resources);
    throw error;
  } finally {
    await reader.cancel();
    reader.releaseLock();
  }
}

function parseWorkbookSheetDescriptors(
  workbookXml: string,
  relsXml: string,
  options?: ExcelReadOptions,
): WorkbookSheetDescriptor[] {
  const workbookRoot = parseXML(workbookXml);
  const sheetsNode = findChild(workbookRoot, 'sheets');
  const sheetNodes = sheetsNode ? findChildren(sheetsNode, 'sheet') : [];

  const relsDoc = parseXML(relsXml);
  const relMap = new Map<string, string>();
  for (const rel of elementChildren(relsDoc)) {
    relMap.set(rel.attributes.Id, rel.attributes.Target);
  }

  const selectedSheets = options?.sheets;
  const descriptors: WorkbookSheetDescriptor[] = [];
  for (let index = 0; index < sheetNodes.length; index++) {
    const sheetNode = sheetNodes[index];
    const name = sheetNode.attributes.name;
    const relationshipId = sheetNode.attributes['r:id'];
    const target = relationshipId ? relMap.get(relationshipId) : undefined;
    if (!name || !target) continue;

    if (selectedSheets) {
      const streamSelectedSheets = selectedSheets as readonly (
        | string
        | number
      )[];
      const matchesName = streamSelectedSheets.includes(name);
      const matchesIndex = streamSelectedSheets.includes(index);
      if (!matchesName && !matchesIndex) {
        continue;
      }
    }

    const path = resolveSheetPartPath('xl', target);
    if (path.includes('..') || path.startsWith('/')) continue;

    descriptors.push({ index, name, path });
  }

  return descriptors;
}

function getStreamDescriptors(
  entries: Record<string, Uint8Array>,
  options?: ExcelReadOptions,
): WorkbookSheetDescriptor[] {
  if (!entries['xl/workbook.xml'] || !entries['xl/_rels/workbook.xml.rels'])
    throw new Error('Invalid XLSX file: workbook metadata is missing');
  const decoder = new TextDecoder();
  return parseWorkbookSheetDescriptors(
    decoder.decode(entries['xl/workbook.xml']),
    decoder.decode(entries['xl/_rels/workbook.xml.rels']),
    options,
  );
}

async function* streamWorksheetRows(
  tempPath: string,
  sharedStrings: string[],
  styles: CellStyle[],
  selection: StreamRowSelection,
): AsyncGenerator<{ rowIndex: number; row: Row }> {
  let emitted = 0;
  const rowSelection =
    selection.startRow === 0 &&
    selection.endRow === MAX_ROWS - 1 &&
    !selection.columns
      ? undefined
      : selection;
  for await (const node of streamNativeElements(
    Bun.file(tempPath).stream(),
    'row',
  )) {
    // Only the current bounded native batch is retained while a consumer waits.
    const row = parseWorksheetRow(node, sharedStrings, styles, rowSelection);
    if (row) {
      yield row;
      if (++emitted >= selection.maxRows) return;
    }
  }
}

/**
 * Read an Excel file as a row-by-row async stream.
 * Uses Bun-native streams for local files and S3 files, while avoiding
 * materializing the full worksheet XML tree in memory.
 */
export async function* readExcelStream(
  source: FileSource,
  options?: ExcelReadStreamOptions,
): AsyncGenerator<ExcelReadStreamRow> {
  const selection = streamRowSelection(options);
  if (selection.maxRows === 0) return;
  const file = toReadableFile(source);
  const exists = await file.exists();
  if (!exists) {
    throw new Error(`File not found: ${describeFileSource(source)}`);
  }

  const fileSize = await getRuntimeFileSize(file);
  if (fileSize > MAX_FILE_SIZE) {
    throw new Error(
      `File too large: ${fileSize} bytes (max: ${MAX_FILE_SIZE})`,
    );
  }

  const decoder = new TextDecoder('utf-8');
  let selectedPaths: Set<string> | undefined;
  let selectedDescriptors: WorkbookSheetDescriptor[] | undefined;
  if (options?.sheets) {
    // Resolve workbook relationships before extracting worksheets, regardless
    // of ZIP entry order. This pass never inflates worksheet/style data.
    const metadata = await unzipXlsxForStreaming(source, {
      metadataOnly: true,
      selection: options,
    });
    const descriptors = getStreamDescriptors(metadata.bufferedEntries, options);
    if (!descriptors.length) return;
    selectedDescriptors = descriptors;
    selectedPaths = new Set(descriptors.map(({ path }) => path));
  }
  const { bufferedEntries, sheetFiles } = await unzipXlsxForStreaming(source, {
    includeStyles: options?.includeStyles !== false,
    selectedPaths,
  });
  const cleanupPaths = sheetFiles.map((sheet) => sheet.tempPath);

  try {
    if (
      !bufferedEntries['xl/workbook.xml'] ||
      !bufferedEntries['xl/_rels/workbook.xml.rels']
    )
      throw new Error('Invalid XLSX file: workbook metadata is missing');

    const descriptors = getStreamDescriptors(bufferedEntries, options);
    if (
      selectedDescriptors &&
      (selectedDescriptors.length !== descriptors.length ||
        selectedDescriptors.some((selected, index) => {
          const current = descriptors[index];
          return (
            selected.index !== current.index ||
            selected.name !== current.name ||
            selected.path !== current.path
          );
        }))
    ) {
      throw new Error('XLSX sheet selection changed between read passes');
    }

    const stringFile = sheetFiles.find(
      (file) => file.entryPath === 'xl/sharedStrings.xml',
    );
    const sharedStrings = stringFile
      ? await readSharedStringElements(Bun.file(stringFile.tempPath).stream())
      : [];
    const styles =
      options?.includeStyles !== false
        ? parseStyles(bufferedEntries, decoder)
        : { cellStyles: [], differentialStyles: [] };
    const sheetPathToTempPath = new Map(
      sheetFiles.map((sheetFile) => [sheetFile.entryPath, sheetFile.tempPath]),
    );

    for (const path of Object.keys(bufferedEntries))
      delete bufferedEntries[path];

    for (const descriptor of descriptors) {
      const tempPath = sheetPathToTempPath.get(descriptor.path);
      if (!tempPath) continue;

      for await (const { rowIndex, row } of streamWorksheetRows(
        tempPath,
        sharedStrings,
        styles.cellStyles,
        selection,
      )) {
        yield {
          sheetIndex: descriptor.index,
          sheetName: descriptor.name,
          rowIndex,
          row,
        };
      }
    }
  } finally {
    await Promise.all(cleanupPaths.map((path) => Bun.file(path).delete()));
  }
}

/** Validate every directory entry, including entries excluded from extraction. */
function extractBufferedEntries(
  buffer: Uint8Array,
  include: (name: string) => boolean,
): Record<string, Uint8Array> {
  let entries = 0;
  let declaredSize = 0;
  return unzipSync(buffer, {
    filter(file) {
      if (++entries > MAX_ZIP_ENTRIES)
        throw new Error('ZIP has too many entries');
      declaredSize += file.originalSize;
      if (declaredSize > MAX_DECOMPRESSED_SIZE)
        throw new Error(
          'Declared decompressed size exceeds limit — potential zip bomb',
        );
      if (
        file.name.startsWith('/') ||
        file.name.startsWith('\\') ||
        file.name.includes('..')
      )
        throw new Error(
          `Malicious zip entry detected: "${file.name}" — potential Zip Slip attack`,
        );
      return include(file.name);
    },
  });
}

function relationshipPartPath(path: string): string {
  const slash = path.lastIndexOf('/');
  return `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
}

/** Resolve only resources consumed by the selected worksheets. */
function extractSelectedResources(
  buffer: Uint8Array,
  selectedPaths: ReadonlySet<string>,
  metadata: Record<string, Uint8Array>,
  availablePaths: ReadonlySet<string>,
  includeStyles: boolean,
): Record<string, Uint8Array> {
  const required = new Set([...selectedPaths, 'docProps/core.xml']);
  if (selectedPaths.size) required.add('xl/sharedStrings.xml');
  if (includeStyles) required.add('xl/styles.xml');
  const sheetRels = new Set(
    [...selectedPaths]
      .map(relationshipPartPath)
      .filter((path) => availablePaths.has(path)),
  );
  const relationships = sheetRels.size
    ? extractBufferedEntries(buffer, (name) => sheetRels.has(name))
    : {};
  const drawingRels = new Map<string, string>();
  const decoder = new TextDecoder();
  for (const data of Object.values(relationships)) {
    const parsed = parseSheetRelationships(decoder.decode(data));
    for (const target of [parsed.commentsPath, ...parsed.tablePaths]) {
      if (target) required.add(resolveSheetPartPath('xl/worksheets', target));
    }
    if (parsed.drawingPath) {
      const drawing = resolveSheetPartPath('xl/worksheets', parsed.drawingPath);
      required.add(drawing);
      drawingRels.set(relationshipPartPath(drawing), drawing);
    }
  }
  if (drawingRels.size) {
    const entries = extractBufferedEntries(buffer, (name) =>
      drawingRels.has(name),
    );
    for (const [path, data] of Object.entries(entries)) {
      const drawing = drawingRels.get(path);
      if (!drawing) continue;
      const base = drawing.slice(0, drawing.lastIndexOf('/'));
      for (const rel of elementChildren(parseXML(decoder.decode(data)))) {
        if (rel.attributes.TargetMode === 'External') continue;
        const target = rel.attributes.Target;
        if (target && rel.attributes.Type?.endsWith('/image'))
          required.add(resolveSheetPartPath(base, target));
      }
    }
    Object.assign(relationships, entries);
  }
  return Object.assign(
    extractBufferedEntries(buffer, (name) => required.has(name)),
    metadata,
    relationships,
  );
}

/**
 * Read an Excel file and return a Workbook
 * Uses Bun.file().arrayBuffer() for optimized binary reading
 */
export async function readExcel(
  source: FileSource,
  options?: ExcelReadOptions,
): Promise<Workbook> {
  const opts = options || {};
  const file = toReadableFile(source);
  const exists = await file.exists();
  if (!exists) {
    throw new Error(`File not found: ${describeFileSource(source)}`);
  }

  const fileSize = await getRuntimeFileSize(file);
  if (fileSize > MAX_FILE_SIZE) {
    throw new Error(
      `File too large: ${fileSize} bytes (max: ${MAX_FILE_SIZE})`,
    );
  }
  // Read bytes directly as Uint8Array for unzipSync()
  const buffer = await file.bytes();

  let zip: Record<string, Uint8Array>;
  if (opts.sheets) {
    const availablePaths = new Set<string>();
    const metadata = extractBufferedEntries(buffer, (name) => {
      availablePaths.add(name);
      return (
        name === 'xl/workbook.xml' || name === 'xl/_rels/workbook.xml.rels'
      );
    });
    const selectedPaths = new Set(
      getStreamDescriptors(metadata, opts).map(({ path }) => path),
    );
    zip = extractSelectedResources(
      buffer,
      selectedPaths,
      metadata,
      availablePaths,
      opts.includeStyles !== false,
    );
  } else {
    zip = extractBufferedEntries(
      buffer,
      (name) => opts.includeStyles !== false || name !== 'xl/styles.xml',
    );
  }

  const decoder = new TextDecoder('utf-8');

  // Parse shared strings
  const sharedStrings = await parseSharedStrings(zip);
  delete zip['xl/sharedStrings.xml'];

  // Parse styles
  const styles =
    opts.includeStyles !== false
      ? parseStyles(zip, decoder)
      : { cellStyles: [], differentialStyles: [] };
  delete zip['xl/styles.xml'];

  // Parse workbook to get sheet info
  const workbookXML = decoder.decode(zip['xl/workbook.xml']);
  const workbookRoot = parseXML(workbookXML);
  delete zip['xl/workbook.xml'];
  const sheetsNode = findChild(workbookRoot, 'sheets');
  const sheetNodes = sheetsNode ? findChildren(sheetsNode, 'sheet') : [];
  const workbookView = parseWorkbookView(workbookRoot);
  const definedNames = parseDefinedNames(workbookRoot);

  // Parse workbook rels to get sheet paths
  const relsXML = decoder.decode(zip['xl/_rels/workbook.xml.rels']);
  const relsDoc = parseXML(relsXML);
  delete zip['xl/_rels/workbook.xml.rels'];
  const relMap = new Map<string, string>();
  for (const rel of elementChildren(relsDoc)) {
    relMap.set(rel.attributes.Id, rel.attributes.Target);
  }

  const workbookProps = parseWorkbookProperties(zip, decoder);
  delete zip['docProps/core.xml'];

  // A worksheet part can be referenced more than once. Release its bytes only
  // after the last reference, while retaining shared feature/image resources.
  const remainingSheetReferences = new Map<string, number>();
  for (const sheet of sheetNodes) {
    const target = relMap.get(sheet.attributes['r:id']);
    if (target) {
      const path = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
      remainingSheetReferences.set(
        path,
        (remainingSheetReferences.get(path) ?? 0) + 1,
      );
    }
  }

  // Parse worksheets
  const worksheets: Worksheet[] = [];
  const worksheetsByOriginalIndex = new Map<number, Worksheet>();

  for (let i = 0; i < sheetNodes.length; i++) {
    const sheetNode = sheetNodes[i];
    const sheetName = sheetNode.attributes.name;
    const rId = sheetNode.attributes['r:id'];

    // Filter sheets if specified
    if (opts.sheets) {
      const matchesName = (
        opts.sheets as readonly (string | number)[]
      ).includes(sheetName);
      const matchesIndex = (
        opts.sheets as readonly (string | number)[]
      ).includes(i);
      if (!matchesName && !matchesIndex) continue;
    }

    const target = relMap.get(rId);
    if (!target) continue;

    const sheetPath = target.startsWith('/') ? target.slice(1) : `xl/${target}`;

    // Zip Slip: validate resolved sheet path
    if (sheetPath.includes('..') || sheetPath.startsWith('/')) {
      continue; // skip suspicious paths
    }

    if (!zip[sheetPath]) continue;

    const sheetXML = decoder.decode(zip[sheetPath]);

    // Parse per-sheet rels for hyperlinks
    const sheetRelsPath = `xl/worksheets/_rels/${
      target.includes('/') ? target.split('/').pop() : target
    }.rels`;
    const sheetRelationships =
      sheetRelsPath in zip
        ? parseSheetRelationships(decoder.decode(zip[sheetRelsPath]))
        : { hyperlinks: new Map<string, string>(), tablePaths: [] };

    const remaining = (remainingSheetReferences.get(sheetPath) ?? 1) - 1;
    remainingSheetReferences.set(sheetPath, remaining);
    if (!remaining) {
      delete zip[sheetPath];
      delete zip[sheetRelsPath];
    }

    const worksheet = parseWorksheet(
      sheetXML,
      sheetName,
      sharedStrings,
      styles.cellStyles,
      styles.differentialStyles,
      sheetRelationships,
      zip,
      decoder,
    );
    const sheetState = sheetNode.attributes.state as
      | Worksheet['state']
      | undefined;
    if (sheetState) {
      worksheet.state = sheetState;
    }
    worksheets.push(worksheet);
    worksheetsByOriginalIndex.set(i, worksheet);
  }

  for (const definedName of definedNames) {
    if (definedName.name !== '_xlnm.Print_Area') continue;
    let worksheet =
      definedName.localSheetId === undefined
        ? undefined
        : worksheetsByOriginalIndex.get(definedName.localSheetId);
    let ref = definedName.refersTo;

    const bangIndex = ref.lastIndexOf('!');
    if (bangIndex !== -1) {
      const sheetNameRef = ref
        .slice(0, bangIndex)
        .replace(LEADING_EQUALS_REGEX, '');
      ref = ref.slice(bangIndex + 1);
      if (definedName.localSheetId === undefined) {
        const normalizedSheetName = sheetNameRef
          .replace(LEADING_SINGLE_QUOTE_REGEX, '')
          .replace(TRAILING_SINGLE_QUOTE_REGEX, '');
        worksheet = worksheets.find(
          (worksheet) => worksheet.name === normalizedSheetName,
        );
      }
    }

    if (worksheet) worksheet.printArea = parseRangeRef(ref);
  }

  return {
    worksheets,
    ...workbookProps,
    definedNames: definedNames.length > 0 ? definedNames : undefined,
    views: workbookView,
  };
}

function parseWorkbookProperties(
  zip: Record<string, Uint8Array>,
  decoder: TextDecoder,
): Pick<Workbook, 'creator' | 'created' | 'modified'> {
  const coreProps = zip['docProps/core.xml'];
  if (!coreProps) return {};

  const xml = decoder.decode(coreProps);
  const root = parseXML(xml);
  if (!root) return {};

  const props: Pick<Workbook, 'creator' | 'created' | 'modified'> = {};
  const creator = findChild(root, 'creator');
  const created = findChild(root, 'created');
  const modified = findChild(root, 'modified');

  if (creator) {
    props.creator = getTextContent(creator);
  }

  const createdValue = created ? new Date(getTextContent(created)) : undefined;
  if (createdValue && !Number.isNaN(createdValue.getTime())) {
    props.created = createdValue;
  }

  const modifiedValue = modified
    ? new Date(getTextContent(modified))
    : undefined;
  if (modifiedValue && !Number.isNaN(modifiedValue.getTime())) {
    props.modified = modifiedValue;
  }

  return props;
}

function parseSheetRelationships(relsXml: string): SheetRelationships {
  const relsDoc = parseXML(relsXml);
  const relationships: SheetRelationships = {
    hyperlinks: new Map<string, string>(),
    tablePaths: [],
  };

  for (const rel of elementChildren(relsDoc)) {
    const target = rel.attributes.Target;
    const type = rel.attributes.Type || '';
    if (!target) continue;

    if (type.includes('hyperlink')) {
      relationships.hyperlinks.set(rel.attributes.Id, target);
      continue;
    }
    if (type.includes('/comments')) {
      relationships.commentsPath = target;
      continue;
    }
    if (type.includes('/drawing')) {
      relationships.drawingPath = target;
      continue;
    }
    if (type.includes('/table')) {
      relationships.tablePaths.push(target);
    }
  }

  return relationships;
}

/**
 * Parse shared strings from XLSX
 */
async function readSharedStringElements(
  stream: ReadableStream<Uint8Array>,
): Promise<string[]> {
  const strings: string[] = [];
  for await (const si of streamNativeElements(stream, 'si')) {
    if (strings.length >= MAX_SHARED_STRINGS)
      throw new Error('Too many shared strings');
    const text = findChild(si, 't');
    strings.push(
      text
        ? getTextContent(text)
        : findChildren(si, 'r')
            .map((run) => getTextContent(findChild(run, 't')))
            .join(''),
    );
  }
  return strings;
}

async function parseSharedStrings(
  zip: Record<string, Uint8Array>,
): Promise<string[]> {
  const data = zip['xl/sharedStrings.xml'];
  if (!data) return [];
  let offset = 0;
  return readSharedStringElements(
    new ReadableStream<Uint8Array>({
      pull(controller) {
        if (offset >= data.length) {
          controller.close();
          return;
        }
        controller.enqueue(data.subarray(offset, offset + 64 * 1024));
        offset += 64 * 1024;
      },
    }),
  );
}

function parseStyles(
  zip: Record<string, Uint8Array>,
  decoder: TextDecoder,
): { cellStyles: CellStyle[]; differentialStyles: CellStyle[] } {
  const data = zip['xl/styles.xml'];
  if (!data) return { cellStyles: [], differentialStyles: [] };

  const xml = decoder.decode(data);
  const root = parseXML(xml);
  if (!root) return { cellStyles: [], differentialStyles: [] };

  // Parse fonts
  const fonts: FontStyle[] = [];
  const fontsNode = findChild(root, 'fonts');
  if (fontsNode) {
    for (const fontNode of findChildren(fontsNode, 'font')) {
      fonts.push(parseFontNode(fontNode));
    }
  }

  // Parse fills
  const fills: FillStyle[] = [];
  const fillsNode = findChild(root, 'fills');
  if (fillsNode) {
    for (const fillNode of findChildren(fillsNode, 'fill')) {
      fills.push(parseFillNode(fillNode));
    }
  }

  // Parse borders
  const bordersList: BorderStyle[] = [];
  const bordersNode = findChild(root, 'borders');
  if (bordersNode) {
    for (const borderNode of findChildren(bordersNode, 'border')) {
      bordersList.push(parseBorderNode(borderNode));
    }
  }

  // Parse number formats
  const numFmtMap = new Map<number, string>();
  const numFmtsNode = findChild(root, 'numFmts');
  if (numFmtsNode) {
    for (const nf of findChildren(numFmtsNode, 'numFmt')) {
      numFmtMap.set(
        Number.parseInt(nf.attributes.numFmtId, 10),
        nf.attributes.formatCode,
      );
    }
  }

  // Parse cell xfs
  const cellStyles: CellStyle[] = [];
  const cellXfsNode = findChild(root, 'cellXfs');
  if (cellXfsNode) {
    for (const xf of findChildren(cellXfsNode, 'xf')) {
      cellStyles.push(
        parseCellXfStyle(xf, fonts, fills, bordersList, numFmtMap),
      );
    }
  }

  const differentialStyles: CellStyle[] = [];
  const dxfsNode = findChild(root, 'dxfs');
  if (dxfsNode) {
    for (const dxf of findChildren(dxfsNode, 'dxf')) {
      const style: CellStyle = {};
      const font = findChild(dxf, 'font');
      const fill = findChild(dxf, 'fill');
      const border = findChild(dxf, 'border');
      const numFmt = findChild(dxf, 'numFmt');
      const alignment = findChild(dxf, 'alignment');

      if (font) style.font = parseFontNode(font);
      if (fill) style.fill = parseFillNode(fill);
      if (border) style.border = parseBorderNode(border);
      if (numFmt?.attributes.formatCode) {
        style.numberFormat = numFmt.attributes.formatCode;
      }
      if (alignment) {
        style.alignment = parseAlignmentNode(alignment);
      }

      differentialStyles.push(style);
    }
  }

  return { cellStyles, differentialStyles };
}

function parseFontNode(fontNode: XMLNode): FontStyle {
  const font: FontStyle = {};
  if (findChild(fontNode, 'b')) font.bold = true;
  if (findChild(fontNode, 'i')) font.italic = true;
  if (findChild(fontNode, 'u')) font.underline = true;
  if (findChild(fontNode, 'strike')) font.strike = true;

  const sz = findChild(fontNode, 'sz');
  if (sz?.attributes.val) {
    font.size = Number.parseFloat(sz.attributes.val);
  }

  const name = findChild(fontNode, 'name');
  if (name?.attributes.val) {
    font.name = name.attributes.val;
  }

  const color = findChild(fontNode, 'color');
  if (color?.attributes.rgb) {
    font.color = color.attributes.rgb.replace(RGB_PREFIX_REGEX, '');
  }

  return font;
}

function parseFillNode(fillNode: XMLNode): FillStyle {
  const gradientFill =
    fillNode.name === 'gradientFill'
      ? fillNode
      : findChild(fillNode, 'gradientFill');
  if (gradientFill) {
    const stops = findChildren(gradientFill, 'stop');
    const firstColor = stops[0]
      ? findChild(stops[0], 'color')?.attributes.rgb
      : undefined;
    const lastColor =
      stops.length > 0
        ? findChild(stops[stops.length - 1], 'color')?.attributes.rgb
        : undefined;
    const fill: FillStyle = { type: 'gradient' };
    if (firstColor) {
      fill.fgColor = firstColor.replace(RGB_PREFIX_REGEX, '');
    }
    if (lastColor) {
      fill.bgColor = lastColor.replace(RGB_PREFIX_REGEX, '');
    }
    return fill;
  }

  const patternFill =
    fillNode.name === 'patternFill'
      ? fillNode
      : findChild(fillNode, 'patternFill');

  if (patternFill) {
    const fill: FillStyle = {
      type: 'pattern',
      pattern:
        (patternFill.attributes.patternType as FillStyle['pattern']) || 'none',
    };
    const fgColor = findChild(patternFill, 'fgColor');
    if (fgColor?.attributes.rgb) {
      fill.fgColor = fgColor.attributes.rgb.replace(RGB_PREFIX_REGEX, '');
    }
    const bgColor = findChild(patternFill, 'bgColor');
    if (bgColor?.attributes.rgb) {
      fill.bgColor = bgColor.attributes.rgb.replace(RGB_PREFIX_REGEX, '');
    }
    return fill;
  }

  return { type: 'pattern', pattern: 'none' };
}

function parseBorderNode(borderNode: XMLNode): BorderStyle {
  const border: BorderStyle = {};
  for (const side of ['left', 'right', 'top', 'bottom'] as const) {
    const sideNode = findChild(borderNode, side);
    if (sideNode?.attributes.style) {
      const edge: BorderEdgeStyle = {
        style: sideNode.attributes.style as BorderEdgeStyle['style'],
      };
      const color = findChild(sideNode, 'color');
      if (color?.attributes.rgb) {
        edge.color = color.attributes.rgb.replace(RGB_PREFIX_REGEX, '');
      }
      border[side] = edge;
    }
  }
  return border;
}

function parseAlignmentNode(alignment: XMLNode): AlignmentStyle {
  const style: AlignmentStyle = {};
  if (alignment.attributes.horizontal) {
    style.horizontal = alignment.attributes
      .horizontal as AlignmentStyle['horizontal'];
  }
  if (alignment.attributes.vertical) {
    style.vertical = alignment.attributes
      .vertical as AlignmentStyle['vertical'];
  }
  if (alignment.attributes.wrapText === '1') {
    style.wrapText = true;
  }
  if (alignment.attributes.textRotation) {
    style.textRotation = Number.parseInt(alignment.attributes.textRotation, 10);
  }
  if (alignment.attributes.indent) {
    style.indent = Number.parseInt(alignment.attributes.indent, 10);
  }
  return style;
}

function parseProtectionNode(protection: XMLNode): CellStyle['protection'] {
  const style: NonNullable<CellStyle['protection']> = {};
  if (protection.attributes.locked !== undefined) {
    style.locked = protection.attributes.locked === '1';
  }
  if (protection.attributes.hidden !== undefined) {
    style.hidden = protection.attributes.hidden === '1';
  }
  return Object.keys(style).length > 0 ? style : undefined;
}

function isDateNumberFormat(numberFormat: string | undefined): boolean {
  if (!numberFormat) return false;

  const normalized = numberFormat
    .toLowerCase()
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*]/g, '')
    .replace(/\\./g, '');

  const hasDateToken =
    normalized.includes('y') ||
    normalized.includes('d') ||
    normalized.includes('h') ||
    normalized.includes('s');

  return hasDateToken || (normalized.includes('m') && normalized.includes('/'));
}

function parseCellXfStyle(
  xf: XMLNode,
  fonts: FontStyle[],
  fills: FillStyle[],
  bordersList: BorderStyle[],
  numFmtMap: Map<number, string>,
): CellStyle {
  const style: CellStyle = {};
  const fontId = Number.parseInt(xf.attributes.fontId || '0', 10);
  const fillId = Number.parseInt(xf.attributes.fillId || '0', 10);
  const borderId = Number.parseInt(xf.attributes.borderId || '0', 10);
  const numFmtId = Number.parseInt(xf.attributes.numFmtId || '0', 10);

  if (xf.attributes.applyFont === '1' && fonts[fontId]) {
    style.font = fonts[fontId];
  }
  if (xf.attributes.applyFill === '1' && fills[fillId]) {
    style.fill = fills[fillId];
  }
  if (xf.attributes.applyBorder === '1' && bordersList[borderId]) {
    style.border = bordersList[borderId];
  }
  if (xf.attributes.applyNumberFormat === '1' && numFmtId > 0) {
    style.numberFormat =
      numFmtMap.get(numFmtId) || BUILTIN_NUMFMTS[numFmtId] || '';
  }

  const alignment = findChild(xf, 'alignment');
  if (alignment) {
    style.alignment = parseAlignmentNode(alignment);
  }

  const protection = findChild(xf, 'protection');
  if (protection) {
    style.protection = parseProtectionNode(protection);
  }

  return style;
}

function parseWorkbookView(root: XMLNode): WorkbookView | undefined {
  const bookViews = findChild(root, 'bookViews');
  const workbookView = bookViews
    ? findChild(bookViews, 'workbookView')
    : undefined;
  if (!workbookView) return undefined;

  const view: WorkbookView = {};
  if (workbookView.attributes.activeTab) {
    view.activeTab = Number.parseInt(workbookView.attributes.activeTab, 10);
  }
  if (workbookView.attributes.firstSheet) {
    view.firstSheet = Number.parseInt(workbookView.attributes.firstSheet, 10);
  }
  if (workbookView.attributes.visibility) {
    view.visibility = workbookView.attributes
      .visibility as WorkbookView['visibility'];
  }
  if (workbookView.attributes.xWindow) {
    view.xWindow = Number.parseInt(workbookView.attributes.xWindow, 10);
  }
  if (workbookView.attributes.yWindow) {
    view.yWindow = Number.parseInt(workbookView.attributes.yWindow, 10);
  }
  if (workbookView.attributes.windowWidth) {
    view.windowWidth = Number.parseInt(workbookView.attributes.windowWidth, 10);
  }
  if (workbookView.attributes.windowHeight) {
    view.windowHeight = Number.parseInt(
      workbookView.attributes.windowHeight,
      10,
    );
  }
  return Object.keys(view).length > 0 ? view : undefined;
}

function parseDefinedNames(root: XMLNode): DefinedName[] {
  const definedNamesNode = findChild(root, 'definedNames');
  if (!definedNamesNode) return [];

  const definedNames: DefinedName[] = [];
  for (const node of findChildren(definedNamesNode, 'definedName')) {
    const name = node.attributes.name;
    if (!name) continue;
    definedNames.push({
      name,
      refersTo: getTextContent(node),
      comment: node.attributes.comment,
      hidden: node.attributes.hidden === '1',
      localSheetId: node.attributes.localSheetId
        ? Number.parseInt(node.attributes.localSheetId, 10)
        : undefined,
    });
  }
  return definedNames;
}

function parseRichTextRuns(isNode: XMLNode): RichTextRun[] {
  const runs: RichTextRun[] = [];
  for (const rNode of findChildren(isNode, 'r')) {
    const tNode = findChild(rNode, 't');
    if (!tNode) continue;
    const run: RichTextRun = { text: getTextContent(tNode) };
    const rPr = findChild(rNode, 'rPr');
    if (rPr) {
      const font: FontStyle = {};
      if (findChild(rPr, 'b')) font.bold = true;
      if (findChild(rPr, 'i')) font.italic = true;
      if (findChild(rPr, 'u')) font.underline = true;
      if (findChild(rPr, 'strike')) font.strike = true;
      const sz = findChild(rPr, 'sz');
      if (sz?.attributes.val) {
        font.size = Number.parseFloat(sz.attributes.val);
      }
      const color = findChild(rPr, 'color');
      if (color?.attributes.rgb) {
        font.color = color.attributes.rgb.replace(RGB_PREFIX_REGEX, '');
      }
      const rFont = findChild(rPr, 'rFont');
      if (rFont?.attributes.val) {
        font.name = rFont.attributes.val;
      }
      if (Object.keys(font).length > 0) {
        run.font = font;
      }
    }
    runs.push(run);
  }
  return runs;
}

function parseHeaderFooterSection(value: string | undefined) {
  if (!value) return undefined;
  const section: NonNullable<HeaderFooter['oddHeader']> = {};
  const leftMatch = value.match(HEADER_FOOTER_LEFT_REGEX);
  const centerMatch = value.match(HEADER_FOOTER_CENTER_REGEX);
  const rightMatch = value.match(HEADER_FOOTER_RIGHT_REGEX);
  if (leftMatch) section.left = leftMatch[1];
  if (centerMatch) section.center = centerMatch[1];
  if (rightMatch) section.right = rightMatch[1];
  return Object.keys(section).length > 0 ? section : undefined;
}

function parseHeaderFooter(root: XMLNode): HeaderFooter | undefined {
  const headerFooterNode = findChild(root, 'headerFooter');
  if (!headerFooterNode) return undefined;

  const headerFooter: HeaderFooter = {
    differentFirst: headerFooterNode.attributes.differentFirst === '1',
    differentOddEven: headerFooterNode.attributes.differentOddEven === '1',
    oddHeader: parseHeaderFooterSection(
      getTextContent(
        findChild(headerFooterNode, 'oddHeader') ?? {
          name: '',
          attributes: Object.create(null),
          children: [],
        },
      ),
    ),
    oddFooter: parseHeaderFooterSection(
      getTextContent(
        findChild(headerFooterNode, 'oddFooter') ?? {
          name: '',
          attributes: Object.create(null),
          children: [],
        },
      ),
    ),
    evenHeader: parseHeaderFooterSection(
      getTextContent(
        findChild(headerFooterNode, 'evenHeader') ?? {
          name: '',
          attributes: Object.create(null),
          children: [],
        },
      ),
    ),
    evenFooter: parseHeaderFooterSection(
      getTextContent(
        findChild(headerFooterNode, 'evenFooter') ?? {
          name: '',
          attributes: Object.create(null),
          children: [],
        },
      ),
    ),
    firstHeader: parseHeaderFooterSection(
      getTextContent(
        findChild(headerFooterNode, 'firstHeader') ?? {
          name: '',
          attributes: Object.create(null),
          children: [],
        },
      ),
    ),
    firstFooter: parseHeaderFooterSection(
      getTextContent(
        findChild(headerFooterNode, 'firstFooter') ?? {
          name: '',
          attributes: Object.create(null),
          children: [],
        },
      ),
    ),
  };

  return Object.values(headerFooter).some(
    (value) => value !== undefined && value !== false,
  )
    ? headerFooter
    : undefined;
}

function parsePageMargins(root: XMLNode): PageMargins | undefined {
  const node = findChild(root, 'pageMargins');
  if (!node) return undefined;
  const margins: PageMargins = {};
  for (const key of [
    'left',
    'right',
    'top',
    'bottom',
    'header',
    'footer',
  ] as const) {
    if (node.attributes[key]) {
      margins[key] = Number.parseFloat(node.attributes[key]);
    }
  }
  return Object.keys(margins).length > 0 ? margins : undefined;
}

function parsePageSetup(root: XMLNode): PageSetup | undefined {
  const node = findChild(root, 'pageSetup');
  if (!node) return undefined;
  const pageSetup: PageSetup = {};
  if (node.attributes.orientation) {
    pageSetup.orientation = node.attributes
      .orientation as PageSetup['orientation'];
  }
  if (node.attributes.paperSize) {
    pageSetup.paperSize = Number.parseInt(node.attributes.paperSize, 10);
  }
  if (node.attributes.scale) {
    pageSetup.scale = Number.parseInt(node.attributes.scale, 10);
  }
  if (node.attributes.fitToWidth) {
    pageSetup.fitToWidth = Number.parseInt(node.attributes.fitToWidth, 10);
  }
  if (node.attributes.fitToHeight) {
    pageSetup.fitToHeight = Number.parseInt(node.attributes.fitToHeight, 10);
  }
  if (node.attributes.firstPageNumber) {
    pageSetup.firstPageNumber = Number.parseInt(
      node.attributes.firstPageNumber,
      10,
    );
  }
  if (node.attributes.useFirstPageNumber) {
    pageSetup.useFirstPageNumber = node.attributes.useFirstPageNumber === '1';
  }
  return Object.keys(pageSetup).length > 0 ? pageSetup : undefined;
}

function parseWorksheetProtection(
  root: XMLNode,
): WorksheetProtection | undefined {
  const node = findChild(root, 'sheetProtection');
  if (!node) return undefined;
  const protection: WorksheetProtection = {};
  if (node.attributes.password) protection.password = node.attributes.password;
  for (const key of [
    'sheet',
    'objects',
    'scenarios',
    'formatCells',
    'formatColumns',
    'formatRows',
    'insertColumns',
    'insertRows',
    'insertHyperlinks',
    'deleteColumns',
    'deleteRows',
    'selectLockedCells',
    'sort',
    'autoFilter',
    'pivotTables',
    'selectUnlockedCells',
  ] as const) {
    if (node.attributes[key] !== undefined) {
      protection[key] = node.attributes[key] === '1';
    }
  }
  return protection;
}

function parseWorksheetColumns(root: XMLNode): ColumnConfig[] | undefined {
  const colsNode = findChild(root, 'cols');
  if (!colsNode) return undefined;

  const colConfigs: ColumnConfig[] = [];
  for (const col of findChildren(colsNode, 'col')) {
    const min = Number.parseInt(col.attributes.min, 10) - 1;
    const max = Number.parseInt(col.attributes.max, 10) - 1;
    const width = col.attributes.width
      ? Number.parseFloat(col.attributes.width)
      : undefined;
    for (let c = min; c <= max; c++) {
      if (c >= MAX_COLS) break;
      while (colConfigs.length <= c) colConfigs.push({});
      colConfigs[c] = {
        width,
        hidden: col.attributes.hidden === '1',
        collapsed: col.attributes.collapsed === '1',
        outlineLevel: col.attributes.outlineLevel
          ? Number.parseInt(col.attributes.outlineLevel, 10)
          : undefined,
      };
    }
  }

  return colConfigs;
}

function parseWorksheetMergeCells(
  root: XMLNode,
): Worksheet['mergeCells'] | undefined {
  const mergeCellsNode = findChild(root, 'mergeCells');
  if (!mergeCellsNode) return undefined;

  const mergeCells: NonNullable<Worksheet['mergeCells']> = [];
  for (const mergeCellNode of findChildren(mergeCellsNode, 'mergeCell')) {
    const ref = mergeCellNode.attributes.ref;
    const [start, end] = ref.split(':');
    const startMatch = start.match(CELL_REF_REGEX);
    const endMatch = end.match(CELL_REF_REGEX);
    if (!startMatch || !endMatch) continue;
    mergeCells.push({
      startRow: Number.parseInt(startMatch[2], 10) - 1,
      startCol: letterToColIndex(startMatch[1]),
      endRow: Number.parseInt(endMatch[2], 10) - 1,
      endCol: letterToColIndex(endMatch[1]),
    });
  }

  return mergeCells;
}

function parseCellValueNode(
  cellNode: XMLNode,
  cellType: string | undefined,
  sharedStrings: string[],
): { value: CellValue; type: Cell['type'] } {
  const valueNode = findChild(cellNode, 'v');
  if (!valueNode) {
    return { value: null, type: 'string' };
  }

  const rawValue = getTextContent(valueNode);
  if (cellType === 's') {
    const index = Number.parseInt(rawValue, 10);
    return {
      value:
        index >= 0 && index < sharedStrings.length
          ? sharedStrings[index]
          : rawValue,
      type: 'string',
    };
  }

  if (cellType === 'b') {
    return { value: rawValue === '1', type: 'boolean' };
  }

  if (cellType === 'str' || cellType === 'inlineStr') {
    return { value: rawValue, type: 'string' };
  }

  const numberValue = Number.parseFloat(rawValue);
  return {
    value: Number.isNaN(numberValue) ? rawValue : numberValue,
    type: 'number',
  };
}

function parseInlineStringValue(
  inlineStringNode: XMLNode | undefined,
): { value: string; richText?: RichTextRun[] } | undefined {
  if (!inlineStringNode) return undefined;

  const textNode = findChild(inlineStringNode, 't');
  if (textNode) {
    return { value: getTextContent(textNode) };
  }

  const richTextRuns = parseRichTextRuns(inlineStringNode);
  if (richTextRuns.length === 0) return undefined;

  return {
    value: richTextRuns.map((run) => run.text).join(''),
    richText: richTextRuns,
  };
}

function parseWorksheetCell(
  cellNode: XMLNode,
  sharedStrings: string[],
  styles: CellStyle[],
  columns?: ReadonlySet<number>,
): { colIndex: number; cell: Cell } | undefined {
  const ref = cellNode.attributes.r;
  if (!ref) return undefined;

  const match = ref.match(COL_LETTER_REGEX);
  if (!match) return undefined;

  const colIndex = letterToColIndex(match[1]);
  if (colIndex < 0 || colIndex >= MAX_COLS) return undefined;
  if (columns && !columns.has(colIndex)) return undefined;

  const cellType = cellNode.attributes.t;
  const styleIndex = Number.parseInt(cellNode.attributes.s || '0', 10);
  const parsedValue = parseCellValueNode(cellNode, cellType, sharedStrings);
  const inlineString = parseInlineStringValue(findChild(cellNode, 'is'));
  const cell: Cell = {
    value: inlineString?.value ?? parsedValue.value,
    type: inlineString ? 'string' : parsedValue.type,
  };

  if (inlineString?.richText) {
    cell.richText = inlineString.richText;
  }

  const formulaNode = findChild(cellNode, 'f');
  if (formulaNode) {
    cell.formula = getTextContent(formulaNode);
    cell.type = 'formula';
  }

  if (styleIndex > 0 && styles[styleIndex]) {
    cell.style = styles[styleIndex];
  }

  const numberFormat = styles[styleIndex]?.numberFormat;
  if (typeof cell.value === 'number' && isDateNumberFormat(numberFormat)) {
    cell.value = excelSerialToDate(cell.value);
    if (!cell.formula) {
      cell.type = 'date';
    }
  }

  return { colIndex, cell };
}

function parseWorksheetRow(
  rowNode: XMLNode,
  sharedStrings: string[],
  styles: CellStyle[],
  selection?: StreamRowSelection,
): { rowIndex: number; row: Row } | undefined {
  const rowIndex = Number.parseInt(rowNode.attributes.r, 10) - 1;
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= MAX_ROWS)
    return undefined;
  if (
    selection &&
    (rowIndex < selection.startRow || rowIndex > selection.endRow)
  )
    return undefined;

  const row: Row = { cells: [] };
  if (rowNode.attributes.ht) {
    row.height = Number.parseFloat(rowNode.attributes.ht);
  }
  row.hidden = rowNode.attributes.hidden === '1';
  row.collapsed = rowNode.attributes.collapsed === '1';
  if (rowNode.attributes.outlineLevel) {
    row.outlineLevel = Number.parseInt(rowNode.attributes.outlineLevel, 10);
  }
  if (selection?.columns?.size === 0) return { rowIndex, row };

  const cells: Cell[] = [];
  for (const cellNode of findChildren(rowNode, 'c')) {
    const parsedCell = parseWorksheetCell(
      cellNode,
      sharedStrings,
      styles,
      selection?.columns,
    );
    if (!parsedCell) continue;
    if (!selection?.columns)
      while (cells.length <= parsedCell.colIndex) {
        cells.push({ value: null });
      }
    cells[parsedCell.colIndex] = parsedCell.cell;
  }

  row.cells = cells;
  return { rowIndex, row };
}

function parseWorksheetRows(
  root: XMLNode,
  sharedStrings: string[],
  styles: CellStyle[],
): Row[] {
  const sheetDataNode = findChild(root, 'sheetData');
  if (!sheetDataNode) return [];

  const rows: Row[] = [];
  let maxRow = -1;

  for (const rowNode of findChildren(sheetDataNode, 'row')) {
    const parsedRow = parseWorksheetRow(rowNode, sharedStrings, styles);
    if (!parsedRow) continue;

    while (rows.length <= parsedRow.rowIndex) {
      rows.push({ cells: [] });
    }
    rows[parsedRow.rowIndex] = parsedRow.row;
    maxRow = Math.max(maxRow, parsedRow.rowIndex);
  }

  return maxRow >= 0 ? rows.slice(0, maxRow + 1) : [];
}

function parseWorksheetHyperlinks(
  root: XMLNode,
  worksheet: Worksheet,
  sheetRelationships: SheetRelationships,
): void {
  const hyperlinksNode = findChild(root, 'hyperlinks');
  if (!hyperlinksNode) return;

  for (const hyperlinkNode of findChildren(hyperlinksNode, 'hyperlink')) {
    const ref = hyperlinkNode.attributes.ref;
    if (!ref) continue;

    const refMatch = ref.match(CELL_REF_REGEX);
    if (!refMatch) continue;

    const colIndex = letterToColIndex(refMatch[1]);
    const rowIndex = Number.parseInt(refMatch[2], 10) - 1;
    const row = worksheet.rows[rowIndex];
    const cell = row?.cells[colIndex];
    if (!cell) continue;

    const relationshipId = hyperlinkNode.attributes['r:id'];
    const location = hyperlinkNode.attributes.location;
    const tooltip = hyperlinkNode.attributes.tooltip;

    let target = '';
    if (relationshipId && sheetRelationships.hyperlinks.has(relationshipId)) {
      target = sheetRelationships.hyperlinks.get(relationshipId) ?? '';
    } else if (location) {
      target = location;
    }

    if (target) {
      cell.hyperlink = { target, tooltip };
    }
  }
}

function parseWorksheetViews(root: XMLNode, worksheet: Worksheet): void {
  const sheetViews = findChild(root, 'sheetViews');
  const sheetView = sheetViews ? findChild(sheetViews, 'sheetView') : undefined;
  const pane = sheetView ? findChild(sheetView, 'pane') : undefined;
  if (!pane) return;

  const state = pane.attributes.state;
  const xSplit = Number.parseFloat(pane.attributes.xSplit || '0');
  const ySplit = Number.parseFloat(pane.attributes.ySplit || '0');

  if (state === 'frozen' && (xSplit > 0 || ySplit > 0)) {
    worksheet.freezePane = { row: ySplit, col: xSplit };
    return;
  }

  if (xSplit > 0 || ySplit > 0) {
    const topLeftRef = pane.attributes.topLeftCell;
    worksheet.splitPane = {
      x: xSplit,
      y: ySplit,
      topLeftCell: topLeftRef ? parseCellRef(topLeftRef) : undefined,
    };
  }
}

function ensureWorksheetCell(
  worksheet: Worksheet,
  rowIndex: number,
  colIndex: number,
): Cell {
  while (worksheet.rows.length <= rowIndex) {
    worksheet.rows.push({ cells: [] });
  }

  const row = worksheet.rows[rowIndex];
  if (!row) {
    worksheet.rows[rowIndex] = { cells: [] };
    return ensureWorksheetCell(worksheet, rowIndex, colIndex);
  }

  while (row.cells.length <= colIndex) {
    row.cells.push({ value: null });
  }

  const existingCell = row.cells[colIndex];
  if (existingCell) return existingCell;

  const newCell: Cell = { value: null };
  row.cells[colIndex] = newCell;
  return newCell;
}

function applyWorksheetComments(
  worksheet: Worksheet,
  sheetRelationships: SheetRelationships,
  zip: Record<string, Uint8Array>,
  decoder: TextDecoder,
): void {
  if (!sheetRelationships.commentsPath) return;

  const commentsPath = resolveSheetPartPath(
    'xl/worksheets',
    sheetRelationships.commentsPath,
  );
  const commentsData = zip[commentsPath];
  if (!commentsData) return;

  for (const entry of parseCommentsXML(decoder.decode(commentsData))) {
    ensureWorksheetCell(worksheet, entry.row, entry.col).comment =
      entry.comment;
  }
}

function applyWorksheetImages(
  worksheet: Worksheet,
  sheetRelationships: SheetRelationships,
  zip: Record<string, Uint8Array>,
  decoder: TextDecoder,
): void {
  if (!sheetRelationships.drawingPath) return;

  const drawingPath = resolveSheetPartPath(
    'xl/worksheets',
    sheetRelationships.drawingPath,
  );
  const drawingData = zip[drawingPath];
  const drawingRelsPath = `${drawingPath.replace(
    'xl/drawings/',
    'xl/drawings/_rels/',
  )}.rels`;
  const drawingRelsData = zip[drawingRelsPath];
  if (!drawingData || !drawingRelsData) return;

  const images = parseDrawingImages(
    decoder.decode(drawingData),
    decoder.decode(drawingRelsData),
    zip,
  );
  if (images.length > 0) {
    worksheet.images = images;
  }
}

function applyWorksheetTables(
  worksheet: Worksheet,
  sheetRelationships: SheetRelationships,
  zip: Record<string, Uint8Array>,
  decoder: TextDecoder,
): void {
  if (sheetRelationships.tablePaths.length === 0) return;

  const tables = sheetRelationships.tablePaths
    .map((tablePath) => {
      const resolvedPath = resolveSheetPartPath('xl/worksheets', tablePath);
      const tableData = zip[resolvedPath];
      if (!tableData) return undefined;
      return parseTableXML(decoder.decode(tableData));
    })
    .filter((table): table is NonNullable<typeof table> => !!table);

  if (tables.length > 0) {
    worksheet.tables = tables;
  }
}

/**
 * Parse a single worksheet
 */
function parseWorksheet(
  xml: string,
  name: string,
  sharedStrings: string[],
  styles: CellStyle[],
  differentialStyles: CellStyle[],
  sheetRelationships: SheetRelationships,
  zip: Record<string, Uint8Array>,
  decoder: TextDecoder,
): Worksheet {
  const root = parseXML(xml);

  const worksheet: Worksheet = { name, rows: [] };

  const columns = parseWorksheetColumns(root);
  if (columns) {
    worksheet.columns = columns;
  }

  const mergeCells = parseWorksheetMergeCells(root);
  if (mergeCells) {
    worksheet.mergeCells = mergeCells;
  }

  const autoFilter = parseAutoFilter(root);
  if (autoFilter) {
    worksheet.autoFilter = autoFilter;
  }

  const conditionalFormattings = parseConditionalFormattings(
    root,
    differentialStyles,
  );
  if (conditionalFormattings.length > 0) {
    worksheet.conditionalFormattings = conditionalFormattings;
  }

  worksheet.rows = parseWorksheetRows(root, sharedStrings, styles);

  const dataValidations = parseDataValidations(root);
  if (dataValidations.length > 0) {
    worksheet.dataValidations = dataValidations;
  }

  parseWorksheetHyperlinks(root, worksheet, sheetRelationships);
  parseWorksheetViews(root, worksheet);

  const headerFooter = parseHeaderFooter(root);
  if (headerFooter) {
    worksheet.headerFooter = headerFooter;
  }

  const pageMargins = parsePageMargins(root);
  if (pageMargins) {
    worksheet.pageMargins = pageMargins;
  }

  const pageSetup = parsePageSetup(root);
  if (pageSetup) {
    worksheet.pageSetup = pageSetup;
  }

  const protection = parseWorksheetProtection(root);
  if (protection) {
    worksheet.protection = protection;
  }

  applyWorksheetComments(worksheet, sheetRelationships, zip, decoder);
  applyWorksheetImages(worksheet, sheetRelationships, zip, decoder);
  applyWorksheetTables(worksheet, sheetRelationships, zip, decoder);

  return worksheet;
}
