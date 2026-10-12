const path = require('node:path');
const fs = require('node:fs');

const { createContractError, requireString } = require('./errors.cjs');

function toAuthorizationKey(filePath) {
  let resolvedPath = path.resolve(requireString(filePath, 'filePath'));
  try {
    resolvedPath = fs.realpathSync.native(resolvedPath);
  } catch {
    try {
      resolvedPath = path.join(fs.realpathSync.native(path.dirname(resolvedPath)),
        path.basename(resolvedPath));
    } catch {
      // 不存在的路径仍保留规范化名称，实际读写时报告文件错误。
    }
  }
  return process.platform === 'win32' ? resolvedPath.toLowerCase() : resolvedPath;
}

function resolveOwnerId(owner) {
  const ownerId = owner?.sender?.id ?? owner?.webContents?.id ?? owner;
  if ((typeof ownerId !== 'number' && typeof ownerId !== 'string') || ownerId === '') {
    throw createContractError('ERR_AUTHORIZATION_OWNER', '缺少授权所有者');
  }
  return String(ownerId);
}

function createAuthorizationService() {
  const authorizedPaths = new Map();
  const authorizedDirectories = new Map();

  function grant(ownerId, filePath) {
    const key = resolveOwnerId(ownerId);
    const paths = authorizedPaths.get(key) || new Set();
    paths.add(toAuthorizationKey(filePath));
    authorizedPaths.set(key, paths);
    return path.resolve(filePath);
  }

  function isAuthorized(ownerId, filePath) {
    const ownerKey = resolveOwnerId(ownerId);
    const fileKey = toAuthorizationKey(filePath);
    if (authorizedPaths.get(ownerKey)?.has(fileKey)) return true;
    return [...(authorizedDirectories.get(ownerKey) || [])].some((root) => {
      const relative = path.relative(root, fileKey);
      return !relative.startsWith(`..${path.sep}`)
        && relative !== '..' && !path.isAbsolute(relative);
    });
  }

  function assertAuthorized(ownerId, filePath) {
    if (!isAuthorized(ownerId, filePath)) {
      throw createContractError('ERR_PATH_NOT_AUTHORIZED', '路径未获得当前窗口授权');
    }
    return path.resolve(filePath);
  }

  function revoke(ownerId) {
    authorizedPaths.delete(resolveOwnerId(ownerId));
    authorizedDirectories.delete(resolveOwnerId(ownerId));
  }

  function grantMany(ownerId, filePaths) {
    return filePaths.map((filePath) => grant(ownerId, filePath));
  }

  function grantDirectory(owner, directory) {
    const key = resolveOwnerId(owner);
    const roots = authorizedDirectories.get(key) || new Set();
    roots.add(toAuthorizationKey(directory));
    authorizedDirectories.set(key, roots);
    return path.resolve(directory);
  }

  return {
    assertAuthorized,
    grant,
    grantMany,
    grantDirectory,
    isAuthorized,
    revoke,
  };
}

module.exports = {
  createAuthorizationService,
  resolveOwnerId,
  toAuthorizationKey,
};
