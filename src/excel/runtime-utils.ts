export function createTempRuntimeId(): string {
  if ('randomUUIDv7' in Bun && typeof Bun.randomUUIDv7 === 'function') {
    return Bun.randomUUIDv7();
  }

  return crypto.randomUUID();
}
