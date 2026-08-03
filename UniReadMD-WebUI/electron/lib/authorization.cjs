const path = require('node:path');

const { createContractError, requireString } = require('./errors.cjs');

function toAuthorizationKey(filePath) {
  const resolvedPath = path.resolve(requireString(filePath, 'filePath'));
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

  function grant(ownerId, filePath) {
    const key = resolveOwnerId(ownerId);
    const paths = authorizedPaths.get(key) || new Set();
    paths.add(toAuthorizationKey(filePath));
    authorizedPaths.set(key, paths);
    return path.resolve(filePath);
  }

  function isAuthorized(ownerId, filePath) {
    const paths = authorizedPaths.get(resolveOwnerId(ownerId));
    return Boolean(paths?.has(toAuthorizationKey(filePath)));
  }

  function assertAuthorized(ownerId, filePath) {
    if (!isAuthorized(ownerId, filePath)) {
      throw createContractError('ERR_PATH_NOT_AUTHORIZED', '路径未获得当前窗口授权');
    }
    return path.resolve(filePath);
  }

  function revoke(ownerId) {
    authorizedPaths.delete(resolveOwnerId(ownerId));
  }

  function grantMany(ownerId, filePaths) {
    return filePaths.map((filePath) => grant(ownerId, filePath));
  }

  return {
    assertAuthorized,
    grant,
    grantMany,
    isAuthorized,
    revoke,
  };
}

module.exports = {
  createAuthorizationService,
  resolveOwnerId,
  toAuthorizationKey,
};
