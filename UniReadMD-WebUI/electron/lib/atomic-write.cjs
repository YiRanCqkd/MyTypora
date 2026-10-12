const { randomUUID } = require('node:crypto');
const path = require('node:path');

async function atomicWrite(fs, target, content) {
  const temporary = path.join(path.dirname(target),
    `.${path.basename(target)}.${randomUUID()}.tmp`);
  try {
    const handle = await fs.open(temporary, 'wx');
    try {
      await handle.writeFile(content, 'utf8');
      await handle.sync();
    } finally {
      await handle.close();
    }
    // 同目录重命名保证旧文件不会在写入中断时被清空。
    await fs.rename(temporary, target);
  } finally {
    await fs.unlink(temporary).catch(() => {});
  }
}

module.exports = { atomicWrite };
