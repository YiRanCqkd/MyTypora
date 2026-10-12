const { createContractError } = require('./errors.cjs');

const PRODUCTION_CSP = [
  "default-src 'none'", "script-src 'self'", "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob: https: http:", "font-src 'self' data:",
  "connect-src 'self'", "object-src 'none'", "frame-src 'none'",
  "base-uri 'none'", "form-action 'none'",
].join('; ');

function validateIpcSender(event, BrowserWindow, getRendererUrl) {
  const win = BrowserWindow.fromWebContents(event.sender);
  const frame = event.senderFrame;
  let valid = Boolean(win && !win.isDestroyed() && win.webContents.id === event.sender.id);
  try {
    valid = valid && Boolean(frame && !frame.parent)
      && frame === event.sender.mainFrame
      && new URL(frame.url).origin === new URL(getRendererUrl()).origin;
  } catch {
    valid = false;
  }
  if (!valid) throw createContractError('ERR_IPC_SENDER', 'IPC 来源未授权');
}

module.exports = { PRODUCTION_CSP, validateIpcSender };
