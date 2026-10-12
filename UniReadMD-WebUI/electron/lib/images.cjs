const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { constants } = require('node:fs');
const { IMAGE_TYPES } = require('./assets.cjs');
const { createContractError, requireString } = require('./errors.cjs');
const { isWithin } = require('./workspace-tools.cjs');
const { atomicWrite } = require('./atomic-write.cjs');
const { collectDocumentImages } = require('./image-collection.cjs');

async function readClipboardImage(clipboard, nativeImage) {
  // 兼容旧同步接口和当前基于 ClipboardItem 的异步接口。
  if (typeof clipboard.readImage === 'function') return clipboard.readImage();
  for (const item of await clipboard.read()) {
    for (const type of ['image/png', 'image/jpeg', 'image/bmp', 'image/tiff']) {
      if (!item.types.includes(type)) continue;
      const blob = await item.getType(type);
      if (blob.size > 20 * 1024 * 1024) throw new Error('截图超过 20 MiB');
      const image = nativeImage.createFromBuffer(Buffer.from(await blob.arrayBuffer()));
      if (!image.isEmpty()) return image;
    }
  }
  return null;
}

function createImageService({ fs, assets, files, authorization, dialog, clipboard,
  nativeImage, getEventWindow, approvalsPath }) {
  let approvals = null;
  let loading = null;
  let pending = Promise.resolve();
  const documentKey = (filePath) => path.resolve(filePath).toLowerCase();
  async function readApprovals() {
    if (approvals) return approvals;
    if (!loading) loading = (async () => {
      let saved = {};
      try {
        if (approvalsPath && (await fs.stat(approvalsPath)).size <= 256 * 1024) {
          const value = JSON.parse(await fs.readFile(approvalsPath, 'utf8'));
          if (value && typeof value === 'object' && !Array.isArray(value)) saved = value;
        }
      } catch { /* 损坏授权记录不能扩大权限。 */ }
      approvals = saved;
      return saved;
    })();
    return loading;
  }
  async function destination(event, payload) {
    const source = requireString(payload.sourceFilePath, 'sourceFilePath');
    authorization.assertAuthorized(event, source);
    const root = await fs.realpath(files.getWorkspaceRoot(event) || path.dirname(source));
    const copyDirectory = requireString(payload.copyDirectory || 'assets', 'copyDirectory', 1024);
    const directory = path.resolve(path.dirname(source), copyDirectory);
    // 验证已存在的父目录后再创建，拒绝通过目录联接写入工作区之外。
    let ancestor = directory;
    while (true) {
      try { ancestor = await fs.realpath(ancestor); break; }
      catch (error) {
        if (error.code !== 'ENOENT' || path.dirname(ancestor) === ancestor) throw error;
        ancestor = path.dirname(ancestor);
      }
    }
    if (!isWithin(root, directory) || !isWithin(root, ancestor)) {
      throw createContractError('ERR_RESOURCE_NOT_AUTHORIZED', '图片保存目录超出工作区');
    }
    await fs.mkdir(directory, { recursive: true });
    return { source, directory };
  }

  async function select(event, payload) {
    const result = await dialog.showOpenDialog(getEventWindow(event), {
      title: '插入本地图片', properties: ['openFile', 'multiSelections'],
      filters: [{ name: '图片', extensions: Object.keys(IMAGE_TYPES).map((item) => item.slice(1)) }],
    });
    if (result.canceled) return { images: [] };
    const { source, directory } = await destination(event, payload);
    const images = [];
    for (const selected of result.filePaths) {
      const extension = path.extname(selected).toLowerCase();
      const stat = await fs.stat(selected);
      if (!IMAGE_TYPES[extension] || !stat.isFile() || stat.size > 20 * 1024 * 1024) {
        throw createContractError('ERR_INVALID_IMAGE', '图片格式不支持或超过 20 MiB');
      }
      const target = path.join(directory, `${randomUUID()}${extension}`);
      await fs.copyFile(selected, target, constants.COPYFILE_EXCL);
      images.push({ reference: path.relative(path.dirname(source), target).replace(/\\/g, '/'),
        name: path.basename(selected) });
    }
    return { images };
  }

  async function drop(event, payload) {
    const selected = await fs.realpath(requireString(payload.filePath, 'filePath'));
    const extension = path.extname(selected).toLowerCase();
    const stat = await fs.stat(selected);
    if (!IMAGE_TYPES[extension] || !stat.isFile() || stat.size > 20 * 1024 * 1024) {
      throw createContractError('ERR_INVALID_IMAGE', '图片格式不支持或超过 20 MiB');
    }
    const { source, directory } = await destination(event, payload);
    const target = path.join(directory, `${randomUUID()}${extension}`);
    await fs.copyFile(selected, target, constants.COPYFILE_EXCL);
    return { images: [{ name: path.basename(selected),
      reference: path.relative(path.dirname(source), target).replace(/\\/g, '/') }] };
  }

  async function applyRoot(event, payload) {
    const source = requireString(payload.sourceFilePath, 'sourceFilePath');
    authorization.assertAuthorized(event, source);
    const content = requireString(payload.frontMatter || ' ', 'frontMatter', 16384);
    const value = /^typora-root-url:\s*(.+)$/m.exec(content)?.[1]?.trim()
      .replace(/^['"]|['"]$/g, '');
    if (!value) { assets.setImageRoot(event, source, undefined); return { ok: true }; }
    const requested = await fs.realpath(path.resolve(path.dirname(source), value));
    const root = await fs.realpath(files.getWorkspaceRoot(event) || path.dirname(source));
    const remembered = (await readApprovals())[documentKey(source)];
    if (!isWithin(root, requested) && !authorization.isAuthorized(event, requested)
      && remembered !== requested) {
      throw createContractError('ERR_RESOURCE_NOT_AUTHORIZED', '图片根目录未授权，请通过目录选择设置');
    }
    assets.setImageRoot(event, source, requested);
    return { ok: true };
  }

  async function paste(event, payload) {
    const image = await readClipboardImage(clipboard, nativeImage);
    if (!image || image.isEmpty()) return { images: [] };
    const bytes = image.toPNG();
    if (bytes.length > 20 * 1024 * 1024) throw new Error('截图超过 20 MiB');
    const { source, directory } = await destination(event, payload);
    const target = path.join(directory, `${randomUUID()}.png`);
    await fs.writeFile(target, bytes, { flag: 'wx' });
    return { images: [{ name: '截图',
      reference: path.relative(path.dirname(source), target).replace(/\\/g, '/') }] };
  }

  async function setRoot(event, payload) {
    const source = requireString(payload.sourceFilePath, 'sourceFilePath');
    authorization.assertAuthorized(event, source);
    const result = await dialog.showOpenDialog(getEventWindow(event), {
      title: '选择本文图片根目录', properties: ['openDirectory'],
    });
    if (result.canceled || !result.filePaths[0]) return null;
    const directoryPath = await fs.realpath(result.filePaths[0]);
    authorization.grantDirectory(event, directoryPath);
    const records = await readApprovals();
    records[documentKey(source)] = directoryPath;
    if (approvalsPath) {
      pending = pending.catch(() => {}).then(() => {
        return atomicWrite(fs, approvalsPath, JSON.stringify(records));
      });
      await pending;
    }
    assets.issue(event, source, files.getWorkspaceRoot(event));
    assets.setImageRoot(event, source, directoryPath);
    return { directoryPath };
  }
  async function collect(event, payload) {
    const source = await fs.realpath(requireString(payload.sourceFilePath, 'sourceFilePath'));
    authorization.assertAuthorized(event, source);
    if (payload.content === '') return { content: '', count: 0 };
    const content = requireString(payload.content, 'content', 20 * 1024 * 1024);
    const root = files.getWorkspaceRoot(event) || path.dirname(source);
    const result = await dialog.showOpenDialog(getEventWindow(event), {
      title: '选择工作区内的图片整理目录', defaultPath: root, properties: ['openDirectory'],
    });
    if (result.canceled || !result.filePaths[0]) return null;
    return collectDocumentImages({ fs, assets, event, source, content, root,
      destination: result.filePaths[0] });
  }
  return { select, paste, setRoot, drop, applyRoot, collect };
}

module.exports = { createImageService, readClipboardImage };
