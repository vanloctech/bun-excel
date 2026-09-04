import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { createDeflateRaw } from 'node:zlib';
import { Zip, type ZipInputFile, ZipPassThrough } from 'fflate';
import type { FileTarget, S3WriterOptions } from '../types';
import { ManagedFileSink } from './file-sink';

const encoder = new TextEncoder();

export type ZipFilePart = string | Uint8Array | Blob;

function isBlobPart(part: ZipFilePart): part is Blob {
  return typeof part === 'object' && 'stream' in part;
}

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

    if (this.compress) {
      await this.addCompressedFile(filename, parts);
      return;
    }
    const entry = new ZipPassThrough(filename);
    this.zip.add(entry);
    this.throwIfErrored();

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

  private async addCompressedFile(
    filename: string,
    parts: readonly ZipFilePart[],
  ): Promise<void> {
    const entry: ZipInputFile = { filename, compression: 8, crc: 0, size: 0 };
    this.zip.add(entry);
    this.throwIfErrored();
    const controller = new AbortController();
    const source = Readable.from(
      this.readParts(parts, entry, controller.signal),
      {
        objectMode: false,
        highWaterMark: 64 * 1024,
      },
    );
    const compressor = createDeflateRaw({
      level: this.level,
      chunkSize: 64 * 1024,
    });
    // Attach rejection handling before consuming output. Either input or the
    // destination can fail while native compression is in flight.
    const completion = pipeline(source, compressor).then(
      () => ({ error: undefined }),
      (error: unknown) => ({ error }),
    );
    try {
      for await (const chunk of compressor) {
        entry.ondata?.(null, chunk, false);
        await this.output.drain();
        this.throwIfErrored();
      }
      const { error } = await completion;
      if (error) throw error;
      entry.ondata?.(null, new Uint8Array(0), true);
      await this.output.flush();
      this.throwIfErrored();
    } catch (error) {
      this.zipError = error instanceof Error ? error : new Error(String(error));
      throw error;
    } finally {
      controller.abort();
      source.destroy();
      compressor.destroy();
      await completion;
    }
  }

  private async *readParts(
    parts: readonly ZipFilePart[],
    entry: ZipInputFile,
    signal: AbortSignal,
  ): AsyncGenerator<Uint8Array> {
    for (const part of parts) {
      signal.throwIfAborted();
      if (isBlobPart(part)) {
        const reader = part.stream().getReader();
        const cancel = () => {
          void reader.cancel().catch(() => {});
        };
        signal.addEventListener('abort', cancel, { once: true });
        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            yield* this.checksumChunks(value, entry);
          }
        } finally {
          signal.removeEventListener('abort', cancel);
          try {
            await reader.cancel();
          } finally {
            reader.releaseLock();
          }
        }
      } else {
        yield* this.checksumChunks(
          typeof part === 'string' ? encoder.encode(part) : part,
          entry,
        );
      }
    }
  }

  private *checksumChunks(
    bytes: Uint8Array,
    entry: ZipInputFile,
  ): Generator<Uint8Array> {
    for (let offset = 0; offset < bytes.length; offset += 64 * 1024) {
      const chunk = bytes.subarray(offset, offset + 64 * 1024);
      entry.crc = Bun.hash.crc32(chunk, entry.crc);
      entry.size += chunk.length;
      yield chunk;
    }
  }

  private async pipeBlob(entry: ZipPassThrough, blob: Blob): Promise<void> {
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
