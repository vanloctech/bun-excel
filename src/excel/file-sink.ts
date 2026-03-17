import { Buffer } from 'node:buffer';
import { isS3File, toWriteTarget } from '../runtime-io';
import type { FileTarget, S3WriterOptions } from '../types';

type SinkChunk = Parameters<Bun.FileSink['write']>[0];

function getChunkSize(chunk: SinkChunk): number {
  if (typeof chunk === 'string') {
    return Buffer.byteLength(chunk);
  }

  if (chunk instanceof ArrayBuffer || chunk instanceof SharedArrayBuffer) {
    return chunk.byteLength;
  }

  if (ArrayBuffer.isView(chunk)) {
    return chunk.byteLength;
  }

  return 0;
}

export interface ManagedFileSinkOptions {
  highWaterMark?: number;
  flushThreshold?: number;
  s3WriterOptions?: S3WriterOptions;
}

export class ManagedFileSink {
  private readonly sink: Bun.FileSink | Bun.NetworkSink;
  private readonly flushThreshold: number;
  private bufferedBytes = 0;
  private flushQueued = false;
  private flushPromise: Promise<void> = Promise.resolve();
  private closed = false;

  constructor(target: FileTarget, options: ManagedFileSinkOptions = {}) {
    const resolvedTarget = toWriteTarget(target);
    if (typeof resolvedTarget === 'string') {
      this.sink = Bun.file(resolvedTarget).writer({
        highWaterMark: options.highWaterMark ?? 256 * 1024,
      });
    } else if (isS3File(resolvedTarget)) {
      this.sink = resolvedTarget.writer(options.s3WriterOptions);
    } else {
      this.sink = resolvedTarget.writer({
        highWaterMark: options.highWaterMark ?? 256 * 1024,
      });
    }
    this.flushThreshold = options.flushThreshold ?? 512 * 1024;
  }

  write(chunk: SinkChunk): void {
    if (this.closed) {
      throw new Error('Cannot write to a closed sink');
    }

    this.sink.write(chunk);
    this.bufferedBytes += getChunkSize(chunk);

    if (this.bufferedBytes >= this.flushThreshold) {
      this.queueFlush();
    }
  }

  drain(): Promise<void> {
    return this.flushPromise;
  }

  async flush(): Promise<void> {
    if (this.bufferedBytes > 0) {
      this.queueFlush();
    }
    await this.flushPromise;
  }

  async end(): Promise<void> {
    if (this.closed) {
      return;
    }

    await this.flush();
    this.closed = true;

    const result = this.sink.end();
    if (result instanceof Promise) {
      await result;
    }
  }

  private queueFlush(): void {
    if (this.flushQueued || this.closed) {
      return;
    }

    this.flushQueued = true;
    this.bufferedBytes = 0;

    this.flushPromise = this.flushPromise
      .then(async () => {
        const result = this.sink.flush();
        if (result instanceof Promise) {
          await result;
        }
      })
      .finally(() => {
        this.flushQueued = false;
        if (this.bufferedBytes >= this.flushThreshold && !this.closed) {
          this.queueFlush();
        }
      });
  }
}
