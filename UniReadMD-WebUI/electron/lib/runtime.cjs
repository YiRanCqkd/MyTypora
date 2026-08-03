const crypto = require('node:crypto');

const { safeErrorMessage } = require('./errors.cjs');

function createRuntime(options) {
  const { app, fs, path, spawn } = options;
  let logFilePath = '';
  let settingsFilePath = '';

  async function initializePaths() {
    const userDataPath = app.getPath('userData');
    const logsDirectory = path.join(userDataPath, 'logs');
    await fs.mkdir(logsDirectory, { recursive: true });
    logFilePath = path.join(logsDirectory, 'runtime.log');
    settingsFilePath = path.join(userDataPath, 'settings.json');
  }

  function createTraceId(prefix = 'srv') {
    return `${prefix}-${Date.now()}-${crypto.randomBytes(4).toString('hex')}`;
  }

  function normalizeTraceId(value) {
    if (typeof value !== 'string') return createTraceId();
    const trimmed = value.trim();
    return /^[A-Za-z0-9:._-]{4,64}$/.test(trimmed) ? trimmed : createTraceId();
  }

  async function appendLog(level, message, traceId, detail) {
    if (!logFilePath) return;
    const line = JSON.stringify({
      time: new Date().toISOString(),
      level,
      traceId,
      message,
      detail: detail ?? null,
    });
    await fs.appendFile(logFilePath, `${line}\n`, 'utf8');
  }

  function log(level, message, traceId, detail) {
    const line = `[${level}] [${traceId}] ${message}`;
    if (level === 'error') console.error(line, detail ?? '');
    else console.log(line, detail ?? '');
    void appendLog(level, message, traceId, detail).catch((error) => {
      console.error('[logger] append failed', safeErrorMessage(error));
    });
  }

  function runCommand(command, arguments_, timeoutMs = 5000) {
    return new Promise((resolve, reject) => {
      const child = spawn(command, arguments_, {
        windowsHide: true,
        stdio: ['ignore', 'pipe', 'pipe'],
      });
      let stdout = '';
      let stderr = '';
      let settled = false;
      const finish = (callback) => {
        if (settled) return false;
        settled = true;
        callback();
        return true;
      };
      const timer = setTimeout(() => finish(() => {
        child.kill('SIGKILL');
        reject(new Error(`${command} timeout after ${timeoutMs}ms`));
      }), timeoutMs);
      child.stdout.on('data', (chunk) => { stdout += String(chunk || ''); });
      child.stderr.on('data', (chunk) => { stderr += String(chunk || ''); });
      child.once('error', (error) => finish(() => {
        clearTimeout(timer);
        reject(error);
      }));
      child.once('close', (code) => finish(() => {
        clearTimeout(timer);
        if (code === 0) resolve({ stdout, stderr });
        else reject(new Error(stderr.trim() || `${command} exit code ${code}`));
      }));
    });
  }

  return {
    createTraceId,
    getSettingsFilePath: () => settingsFilePath,
    initializePaths,
    log,
    normalizeTraceId,
    runCommand,
  };
}

module.exports = { createRuntime };
