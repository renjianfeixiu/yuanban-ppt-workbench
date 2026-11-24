import { createWriteStream } from 'node:fs';
import { stat } from 'node:fs/promises';
import { pipeline } from 'node:stream/promises';
import { buffer as readBuffer } from 'node:stream/consumers';
import yauzl from 'yauzl';
import yazl from 'yazl';

export const MAX_TEMPLATE_BYTES = 160 * 1024 * 1024;
const MAX_EXPANDED_BYTES = 512 * 1024 * 1024;

export async function openArchive(file) {
  const { size } = await stat(file);
  if (size > MAX_TEMPLATE_BYTES) throw new Error('模板不能超过 160 MiB。');
  const zip = await new Promise((resolve, reject) => {
    yauzl.open(file, { lazyEntries: true, autoClose: false, strictFileNames: true }, (error, value) => {
      if (error) reject(new Error('文件不是可读取的 PPTX 压缩包。', { cause: error }));
      else resolve(value);
    });
  });
  const entries = new Map();
  let expanded = 0;
  try {
    if (zip.entryCount > 20000) throw new Error('模板包含的文件部件过多。');
    await new Promise((resolve, reject) => {
      zip.once('error', reject);
      zip.once('end', resolve);
      zip.on('entry', entry => {
        expanded += entry.uncompressedSize;
        if (expanded > MAX_EXPANDED_BYTES) return reject(new Error('模板展开后超过 512 MiB。'));
        if (entry.isEncrypted()) return reject(new Error('请先将加密模板另存为普通 PPTX。'));
        if (!entry.fileName.endsWith('/')) {
          if (entries.has(entry.fileName)) return reject(new Error('模板含有重复的文件部件。'));
          entries.set(entry.fileName, entry);
        }
        zip.readEntry();
      });
      zip.readEntry();
    });
  } catch (error) {
    zip.close();
    throw error;
  }
  const stream = (entry, callback) => zip.openReadStream(entry, callback);
  const bytes = async (part, limit = 16 * 1024 * 1024) => {
    const entry = entries.get(part);
    if (!entry) throw new Error(`模板缺少文件部件：${part}`);
    if (entry.uncompressedSize > limit) throw new Error(`文件部件过大：${part}`);
    const readable = await new Promise((resolve, reject) => {
      stream(entry, (error, value) => error ? reject(error) : resolve(value));
    });
    return readBuffer(readable);
  };
  return {
    entries,
    stream,
    bytes,
    xml: async part => new TextDecoder('utf-8', { fatal: true }).decode(await bytes(part)),
    close: () => zip.close(),
  };
}

// Repack the file without reserializing untouched XML or media.
export async function writeArchive(archive, patches, destination) {
  const zip = new yazl.ZipFile();
  const written = pipeline(zip.outputStream, createWriteStream(destination, { flags: 'wx' }));
  zip.on('error', error => zip.outputStream.destroy(error));
  try {
    const remaining = new Map(patches);
    for (const [part, entry] of archive.entries) {
      if (remaining.has(part)) {
        zip.addBuffer(Buffer.from(remaining.get(part)), part);
        remaining.delete(part);
      } else {
        zip.addReadStreamLazy(part, { size: entry.uncompressedSize, compress: entry.compressionMethod !== 0 }, callback => {
          archive.stream(entry, callback);
        });
      }
    }
    for (const [part, data] of remaining) zip.addBuffer(Buffer.from(data), part);
    zip.end();
  } catch (error) {
    zip.outputStream.destroy(error);
  }
  await written;
}
