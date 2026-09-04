import type { AsyncFlateStreamHandler, UnzipDecoder } from 'fflate';

/** Native raw-DEFLATE decoders with input backpressure for fflate ZIP framing. */
export function createNativeInflaters(): {
  decoder: { new (): UnzipDecoder; compression: number };
  drain(): Promise<void>;
  abort(): Promise<void>;
} {
  const active = new Set<NativeInflate>();
  let pending: Promise<void>[] = [];
  let failure: unknown;
  function track(promise: Promise<void>) {
    // Attach the handler immediately; ZIP framing itself can throw before drain.
    const guarded = promise.catch((error: unknown) => {
      failure ??= error;
    });
    pending.push(guarded);
  }
  class NativeInflate {
    static compression = 8;
    ondata!: AsyncFlateStreamHandler;
    private readonly stream = new DecompressionStream('deflate-raw');
    private readonly input = this.stream.writable.getWriter();
    private readonly output = this.stream.readable.getReader();
    private readonly completion: Promise<void>;
    private aborted = false;
    constructor() {
      active.add(this);
      this.completion = this.read();
      void this.completion.catch((error: unknown) => {
        failure ??= error;
      });
    }
    private async read(): Promise<void> {
      try {
        while (!this.aborted) {
          const { done, value } = await this.output.read();
          if (this.aborted) return;
          if (done) {
            this.ondata(null, new Uint8Array(0), true);
            return;
          }
          this.ondata(null, value, false);
        }
      } catch (error) {
        failure ??= error;
        this.aborted = true;
        await Promise.allSettled([
          this.output.cancel(error),
          this.input.abort(error),
        ]);
        throw error;
      } finally {
        active.delete(this);
      }
    }
    push(data: Uint8Array, final: boolean): void {
      const bytes =
        data.buffer instanceof ArrayBuffer
          ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
          : new Uint8Array(data);
      track(
        this.input.write(bytes).then(async () => {
          if (final) {
            await this.input.close();
            await this.completion;
          }
        }),
      );
    }
    async terminate(): Promise<void> {
      this.aborted = true;
      await Promise.allSettled([
        this.output.cancel(),
        this.input.abort(),
        this.completion,
      ]);
    }
  }
  return {
    decoder: NativeInflate,
    async drain() {
      const writes = pending;
      pending = [];
      await Promise.all(writes);
      if (failure) throw failure;
    },
    async abort() {
      await Promise.allSettled(
        [...active].map((decoder) => decoder.terminate()),
      );
      await Promise.allSettled(pending);
      pending = [];
    },
  };
}
