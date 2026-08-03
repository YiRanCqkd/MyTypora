const path = require('node:path');

const retryDelaysMs = Object.freeze([
  250,
  500,
  750,
  1000,
  1500,
]);

const waitFor = (milliseconds) => new Promise((resolve) => {
  setTimeout(resolve, milliseconds);
});

const isTargetRename = ({
  source,
  destination,
  platform,
}) => platform === 'win32'
  && path.basename(source) === 'win-unpacked.tmp'
  && path.basename(destination) === 'win-unpacked'
  && path.dirname(source) === path.dirname(destination);

const isTransientRenameError = (error) => error instanceof Error
  && error.code === 'EPERM'
  && error.syscall === 'rename';

// Electron 解压完成后，Windows 文件过滤器可能短暂占用目录中的可执行文件。
// 仅对已确认的 win-unpacked 目录 EPERM 使用有限预算，其他错误立即保持原样抛出。
const renameWithTransientRetry = async ({
  source,
  destination,
  rename,
  wait = waitFor,
  platform = process.platform,
  onRetry = () => {},
}) => {
  const targetRename = isTargetRename({
    source,
    destination,
    platform,
  });

  for (let attempt = 0; ; attempt += 1) {
    try {
      await rename(source, destination);
      return;
    } catch (error) {
      if (!targetRename
        || !isTransientRenameError(error)
        || attempt >= retryDelaysMs.length) {
        throw error;
      }

      const delayMs = retryDelaysMs[attempt];
      onRetry({
        attempt: attempt + 1,
        delayMs,
      });
      await wait(delayMs);
    }
  }
};

module.exports = {
  renameWithTransientRetry,
  retryDelaysMs,
};
