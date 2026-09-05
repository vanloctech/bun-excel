import { closeSync, mkdtempSync, openSync, rmSync } from 'node:fs';
import { rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';

/** Internal temporary files only: each file owns its private directory. */
export function createPrivateTempFile(
  prefix: string,
  parent = tmpdir(),
): string {
  const directory = mkdtempSync(join(parent, `${prefix}-`));
  const path = join(directory, 'data.tmp');
  try {
    closeSync(openSync(path, 'wx', 0o600));
    return path;
  } catch (error) {
    rmSync(directory, { recursive: true, force: true });
    throw error;
  }
}

/** Only accepts paths returned by createPrivateTempFile. */
export async function removePrivateTempFile(path: string): Promise<void> {
  await rm(dirname(path), { recursive: true, force: true });
}
