const { randomUUID } = require('node:crypto');
const path = require('node:path');
const { createContractError } = require('./errors.cjs');

const IMAGE_TYPES = {
  '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
  '.gif': 'image/gif', '.webp': 'image/webp', '.svg': 'image/svg+xml',
  '.bmp': 'image/bmp', '.ico': 'image/x-icon', '.avif': 'image/avif',
};

function createAssetService({ fs, getServerUrl }) {
  const tokens = new Map();
  const documents = new Map();
  function issue(owner, filePath, rootDirectory) {
    const ownerId = owner.sender?.id ?? owner.webContents?.id;
    const key = `${ownerId}:${filePath}`;
    let token = documents.get(key);
    // 导出复用已有能力，不能把工作区授权缩小为文档所在目录。
    if (token && rootDirectory === undefined) {
      return `${getServerUrl()}__unireadmd_asset__/${token}?path=`;
    }
    rootDirectory ||= path.dirname(filePath);
    const relative = path.relative(rootDirectory, filePath);
    if (relative === '..' || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
      rootDirectory = path.dirname(filePath);
    }
    if (!token) {
      token = randomUUID();
      documents.set(key, token);
    }
    tokens.set(token, { ...tokens.get(token), ownerId, filePath, rootDirectory });
    return `${getServerUrl()}__unireadmd_asset__/${token}?path=`;
  }
  async function read(token, reference) {
    const document = tokens.get(token);
    if (!document || typeof reference !== 'string' || reference.includes('\0')) {
      throw createContractError('ERR_RESOURCE_NOT_AUTHORIZED', '资源未授权');
    }
    const base = document.imageRootDirectory && reference.startsWith('/')
      ? document.imageRootDirectory : path.dirname(document.filePath);
    const requested = document.imageRootDirectory && reference.startsWith('/')
      ? reference.replace(/^\/+/, '') : reference;
    const target = await fs.realpath(path.resolve(base, requested));
    const roots = await Promise.all([document.rootDirectory, document.imageRootDirectory]
      .filter(Boolean).map((root) => fs.realpath(root)));
    const allowed = roots.some((root) => {
      const relative = path.relative(root, target);
      return relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
    });
    const contentType = IMAGE_TYPES[path.extname(target).toLowerCase()];
    if (!contentType || !allowed) {
      throw createContractError('ERR_RESOURCE_NOT_AUTHORIZED', '图片超出文档授权目录');
    }
    const stat = await fs.stat(target);
    if (!stat.isFile() || stat.size > 20 * 1024 * 1024) {
      throw createContractError('ERR_RESOURCE_TOO_LARGE', '图片超过大小限制');
    }
    return { contentType, content: await fs.readFile(target) };
  }
  function revoke(owner) {
    const ownerId = typeof owner === 'number' ? owner : owner.sender?.id ?? owner.webContents?.id;
    for (const [token, document] of tokens) {
      if (document.ownerId === ownerId) tokens.delete(token);
    }
    for (const [key] of documents) {
      if (key.startsWith(`${ownerId}:`)) documents.delete(key);
    }
  }
  async function readForOwner(owner, token, reference) {
    const ownerId = owner.sender?.id ?? owner.webContents?.id;
    if (tokens.get(token)?.ownerId !== ownerId) {
      throw createContractError('ERR_RESOURCE_NOT_AUTHORIZED', '资源不属于当前窗口');
    }
    return read(token, reference);
  }
  function setImageRoot(owner, filePath, imageRootDirectory) {
    const ownerId = owner.sender?.id ?? owner.webContents?.id;
    const token = documents.get(`${ownerId}:${filePath}`);
    const document = tokens.get(token);
    if (document) document.imageRootDirectory = imageRootDirectory;
  }
  return { issue, read, readForOwner, revoke, setImageRoot };
}

module.exports = { createAssetService, IMAGE_TYPES };
