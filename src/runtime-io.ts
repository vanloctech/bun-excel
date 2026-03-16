import { resolve } from 'node:path';
import type { FileSource, FileTarget } from './types';

type RuntimeFile = Bun.BunFile | Bun.S3File;

export function validatePath(filePath: string): string {
  if (filePath.includes('\0')) {
    throw new Error('Invalid file path: contains null bytes');
  }
  return resolve(filePath);
}

export function toReadableFile(source: FileSource): RuntimeFile {
  if (typeof source === 'string') {
    return Bun.file(validatePath(source));
  }
  return source;
}

export function toWriteTarget(target: FileTarget): string | RuntimeFile {
  if (typeof target === 'string') {
    return validatePath(target);
  }
  return target;
}

export function isS3File(target: RuntimeFile): target is Bun.S3File {
  return 'presign' in target;
}

export async function getRuntimeFileSize(file: RuntimeFile): Promise<number> {
  const stat = await file.stat();
  return stat.size;
}

export function describeFileSource(source: FileSource): string {
  if (typeof source === 'string') {
    return validatePath(source);
  }

  return source.name || '[blob-source]';
}
