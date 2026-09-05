import type { ExcelReadProgress, ExcelReadStreamOptions } from '../types';

/** Observe cancellation without leaving an unhandled rejection from the work. */
export function awaitRead<T>(
  work: Promise<T>,
  signal?: AbortSignal,
): Promise<T> {
  if (!signal) return work;
  return new Promise<T>((resolve, reject) => {
    const abort = () => {
      signal.removeEventListener('abort', abort);
      reject(signal.reason);
    };
    signal.addEventListener('abort', abort, { once: true });
    work
      .then(resolve, reject)
      .finally(() => signal.removeEventListener('abort', abort));
    if (signal.aborted) abort();
  });
}

/** Cancel pending stream reads as well as checking between synchronous batches. */
export function cancelReadOnAbort(
  reader: ReadableStreamDefaultReader<Uint8Array>,
  signal?: AbortSignal,
): () => void {
  const abort = () => {
    void reader.cancel(signal?.reason).catch(() => {});
  };
  signal?.addEventListener('abort', abort, { once: true });
  if (signal?.aborted) abort();
  return () => signal?.removeEventListener('abort', abort);
}

export class ReadControl {
  readonly enabled: boolean;
  readonly signal?: AbortSignal;
  readonly interval: number;
  readonly progress: Omit<ExcelReadProgress, 'elapsedMs'> = {
    stage: 'metadata',
    bytesRead: 0,
    rowsRead: 0,
    sharedStringsRead: 0,
  };
  private readonly callback?: ExcelReadStreamOptions['onProgress'];
  private readonly started = performance.now();
  private lastReport = Number.NEGATIVE_INFINITY;

  constructor(options?: ExcelReadStreamOptions) {
    this.enabled = !!(options?.signal || options?.onProgress);
    this.signal = options?.signal;
    this.callback = options?.onProgress;
    this.interval =
      options?.progressIntervalRows === undefined
        ? 1000
        : options.progressIntervalRows;
    if (!Number.isSafeInteger(this.interval) || this.interval < 1)
      throw new RangeError(
        'progressIntervalRows must be a positive safe integer',
      );
    this.signal?.throwIfAborted();
  }

  async report(
    stage: ExcelReadProgress['stage'],
    force = false,
  ): Promise<void> {
    this.signal?.throwIfAborted();
    const changed = this.progress.stage !== stage;
    this.progress.stage = stage;
    if (!this.callback) return;
    const now = performance.now();
    if (!force && !changed && now - this.lastReport < 100) return;
    this.lastReport = now;
    await awaitRead(
      Promise.resolve(
        this.callback({ ...this.progress, elapsedMs: now - this.started }),
      ),
      this.signal,
    );
    this.signal?.throwIfAborted();
  }
}
