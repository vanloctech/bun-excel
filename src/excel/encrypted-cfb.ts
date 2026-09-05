// MS-CFB v3 layout for large encrypted packages. The payload is already on disk
// at offset 512; only the header and trailing allocation metadata are buffered.
// https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/05060311-bfce-4b12-874d-71fd4ce63aea
const SECTOR = 512;
const FREE = 0xffffffff;
const END = 0xfffffffe;
const FAT = 0xfffffffd;
const DIFAT = 0xfffffffc;

export function encryptedCfbMetadata(payloadSize: number, info: Uint8Array) {
  if (
    !Number.isSafeInteger(payloadSize) ||
    payloadSize < 4096 ||
    payloadSize >= 0x80000000
  )
    throw new Error('Encrypted package exceeds the supported CFB size range');
  if (info.length > 4096)
    throw new Error('EncryptionInfo exceeds its reserved space');
  const payloadSectors = Math.ceil(payloadSize / SECTOR);
  let fatCount = 0;
  let difatCount = 0;
  for (;;) {
    const next = Math.ceil((payloadSectors + 9 + fatCount + difatCount) / 128);
    if (next === fatCount) break;
    fatCount = next;
    difatCount = Math.ceil(Math.max(0, fatCount - 109) / 127);
  }
  const fatStart = payloadSectors + 9;
  const difatStart = fatStart + fatCount;
  if ((1 + difatStart + difatCount) * SECTOR > 0x80000000)
    throw new Error('Encrypted compound file exceeds the 2 GiB CFB v3 limit');
  const header = Buffer.alloc(SECTOR);
  header.set(Buffer.from('d0cf11e0a1b11ae1', 'hex'));
  header.writeUInt16LE(0x003e, 24);
  header.writeUInt16LE(3, 26);
  header.writeUInt16LE(0xfffe, 28);
  header.writeUInt16LE(9, 30);
  header.writeUInt16LE(6, 32);
  header.writeUInt32LE(fatCount, 44);
  header.writeUInt32LE(payloadSectors + 8, 48);
  header.writeUInt32LE(4096, 56);
  header.writeUInt32LE(END, 60);
  header.writeUInt32LE(difatCount ? difatStart : END, 68);
  header.writeUInt32LE(difatCount, 72);
  for (let i = 0; i < 109; i++)
    header.writeUInt32LE(i < fatCount ? fatStart + i : FREE, 76 + i * 4);

  const footer = Buffer.alloc((9 + fatCount + difatCount) * SECTOR);
  // A 4096-byte XML stream uses the normal FAT, avoiding a separate mini stream.
  footer.fill(0x20, 0, 4096);
  footer.set(info);
  const directory = footer.subarray(4096, 4608);
  function entry(
    index: number,
    name: string,
    type: number,
    start: number,
    size: number,
  ) {
    const record = directory.subarray(index * 128, (index + 1) * 128);
    record.write(name, 0, 64, 'utf16le');
    record.writeUInt16LE((name.length + 1) * 2, 64);
    record[66] = type;
    record[67] = 1;
    for (const offset of [68, 72, 76]) record.writeUInt32LE(FREE, offset);
    record.writeUInt32LE(start, 116);
    record.writeBigUInt64LE(BigInt(size), 120);
    return record;
  }
  entry(0, 'Root Entry', 5, END, 0).writeUInt32LE(1, 76);
  entry(1, 'EncryptionInfo', 2, payloadSectors, 4096).writeUInt32LE(2, 72);
  entry(2, 'EncryptedPackage', 2, 0, payloadSize)[67] = 0;
  // Remaining entry is unallocated.
  for (const offset of [68, 72, 76])
    directory.writeUInt32LE(FREE, 384 + offset);

  const fat = footer.subarray(4608, (9 + fatCount) * SECTOR);
  fat.fill(0xff);
  for (let i = 0; i < payloadSectors; i++)
    fat.writeUInt32LE(i === payloadSectors - 1 ? END : i + 1, i * 4);
  for (let i = payloadSectors; i < payloadSectors + 8; i++)
    fat.writeUInt32LE(i === payloadSectors + 7 ? END : i + 1, i * 4);
  fat.writeUInt32LE(END, (payloadSectors + 8) * 4);
  for (let i = fatStart; i < difatStart; i++) fat.writeUInt32LE(FAT, i * 4);
  for (let i = 0; i < difatCount; i++) {
    fat.writeUInt32LE(DIFAT, (difatStart + i) * 4);
    const sector = footer.subarray(
      (9 + fatCount + i) * SECTOR,
      (10 + fatCount + i) * SECTOR,
    );
    sector.fill(0xff);
    for (let j = 0; j < 127; j++) {
      const index = 109 + i * 127 + j;
      if (index < fatCount) sector.writeUInt32LE(fatStart + index, j * 4);
    }
    sector.writeUInt32LE(i === difatCount - 1 ? END : difatStart + i + 1, 508);
  }
  return { header, footer, footerOffset: (1 + payloadSectors) * SECTOR };
}
