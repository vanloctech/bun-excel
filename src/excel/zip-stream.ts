import { Zip, ZipDeflate, ZipPassThrough } from 'fflate';
import type { FileTarget, S3WriterOptions } from '../types';
import { ManagedFileSink } from './file-sink';

const encoder = new TextEncoder();

export type ZipFilePart = string | Uint8Array | Blob;

function isBlobPart(part: ZipFilePart): part is Blob {
  return typeof part === 'object' && 'stream' in part;
}

type ZipEntry = ZipDeflate | ZipPassThrough;

export interface StreamingZipWriterOptions {
  compress?: boolean;
  level?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
  highWaterMark?: number;
  flushThreshold?: number;
  s3WriterOptions?: S3WriterOptions;
}

export class StreamingZipWriter {
  private readonly output: ManagedFileSink;
  private readonly compress: boolean;
  private readonly level: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
  private readonly zip: Zip;
  private zipError: Error | null = null;

  constructor(target: FileTarget, options: StreamingZipWriterOptions = {}) {
    this.output = new ManagedFileSink(target, {
      highWaterMark: options.highWaterMark,
      flushThreshold: options.flushThreshold,
      s3WriterOptions: options.s3WriterOptions,
    });
    this.compress = options.compress !== false;
    this.level = options.level ?? 6;
    this.zip = new Zip((err, data) => {
      if (err) {
        this.zipError = err;
        return;
      }
      this.output.write(data);
    });
  }

  async addFile(
    filename: string,
    parts: readonly ZipFilePart[],
  ): Promise<void> {
    this.throwIfErrored();

    const entry = this.createEntry(filename);
    this.zip.add(entry);

    for (const part of parts) {
      if (isBlobPart(part)) {
        await this.pipeBlob(entry, part);
      } else if (typeof part === 'string') {
        if (part.length === 0) {
          continue;
        }
        entry.push(encoder.encode(part), false);
      } else if (part.length > 0) {
        entry.push(part, false);
      }

      await this.output.drain();
      this.throwIfErrored();
    }

    entry.push(new Uint8Array(0), true);
    await this.output.flush();
    this.throwIfErrored();
  }

  async close(): Promise<void> {
    this.throwIfErrored();
    this.zip.end();
    await this.output.end();
    this.throwIfErrored();
  }

  private createEntry(filename: string): ZipEntry {
    if (this.compress) {
      return new ZipDeflate(filename, { level: this.level });
    }
    return new ZipPassThrough(filename);
  }

  private async pipeBlob(entry: ZipEntry, blob: Blob): Promise<void> {
    const reader = blob.stream().getReader();

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          break;
        }

        entry.push(value, false);
        await this.output.drain();
        this.throwIfErrored();
      }
    } finally {
      reader.releaseLock();
    }
  }

  private throwIfErrored(): void {
    if (this.zipError) {
      throw this.zipError;
    }
  }
}
