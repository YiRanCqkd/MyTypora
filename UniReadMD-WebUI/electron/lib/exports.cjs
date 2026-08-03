const { createContractError, requireString } = require('./errors.cjs');
const {
  BUILTIN_EXPORT_FORMATS,
  PANDOC_EXPORT_FORMATS,
} = require('./export-engines.cjs');

const MAX_EXPORT_BYTES = 20 * 1024 * 1024;
const SUPPORTED_FORMATS = new Set([...BUILTIN_EXPORT_FORMATS, ...PANDOC_EXPORT_FORMATS]);

function optionalString(value, fieldName) {
  return value === undefined || value === null ? null : requireString(value, fieldName);
}

function assertExportSize(value, fieldName) {
  if (typeof value !== 'string') {
    throw createContractError('ERR_INVALID_ARGUMENT', `${fieldName} 必须是字符串`);
  }
  if (Buffer.byteLength(value, 'utf8') > MAX_EXPORT_BYTES) {
    throw createContractError('ERR_EXPORT_SOURCE_TOO_LARGE', '导出源超过 20 MiB 限制');
  }
  return value;
}

function validateExportDocumentPayload(payload) {
  const format = requireString(payload?.format, 'format', 32).toLowerCase();
  if (!SUPPORTED_FORMATS.has(format)) {
    throw createContractError('ERR_UNSUPPORTED_EXPORT_FORMAT', `不支持 ${format} 导出`);
  }
  return {
    format,
    html: assertExportSize(payload?.html, 'html'),
    markdown: assertExportSize(payload?.markdown, 'markdown'),
    sourceFilePath: optionalString(payload?.sourceFilePath, 'sourceFilePath'),
    targetPath: optionalString(payload?.targetPath, 'targetPath'),
  };
}

function validateExportHtmlPayload(payload) {
  return {
    html: assertExportSize(payload?.html, 'html'),
    sourceFilePath: optionalString(payload?.sourceFilePath, 'sourceFilePath'),
    targetPath: optionalString(payload?.targetPath, 'targetPath'),
  };
}

function deriveExportPath(sourceFilePath, format, path) {
  const extensions = {
    docx: 'docx',
    html: 'html',
    jpeg: 'jpg',
    odt: 'odt',
    pdf: 'pdf',
    png: 'png',
    rtf: 'rtf',
  };
  const extension = extensions[format] || 'html';
  if (!sourceFilePath) return `document.${extension}`;
  const baseName = path.basename(sourceFilePath, path.extname(sourceFilePath));
  return `${baseName || 'document'}.${extension}`;
}

function getExportDialogFilter(format) {
  const filters = {
    docx: { name: 'Word', extensions: ['docx'] },
    html: { name: 'HTML', extensions: ['html'] },
    jpeg: { name: 'JPEG Image', extensions: ['jpg', 'jpeg'] },
    odt: { name: 'OpenDocument', extensions: ['odt'] },
    pdf: { name: 'PDF', extensions: ['pdf'] },
    png: { name: 'PNG Image', extensions: ['png'] },
    rtf: { name: 'RTF', extensions: ['rtf'] },
  };
  return filters[format] || filters.html;
}

function createExportsService(options) {
  const {
    authorization,
    dialog,
    enforceAuthorization = false,
    engines,
    fs,
    getEventWindow,
    log = () => {},
    path,
  } = options;

  function focusForDialog(win) {
    if (!win || win.isDestroyed()) return;
    if (win.isMinimized()) win.restore();
    win.focus();
  }

  async function resolveTarget(event, payload, traceId) {
    if (payload.targetPath) {
      if (enforceAuthorization) authorization.assertAuthorized(event, payload.targetPath);
      return payload.targetPath;
    }
    const win = getEventWindow(event);
    focusForDialog(win);
    const dialogOptions = {
      defaultPath: deriveExportPath(payload.sourceFilePath, payload.format, path),
      filters: [getExportDialogFilter(payload.format)],
    };
    const result = win && !win.isDestroyed()
      ? await dialog.showSaveDialog(win, dialogOptions)
      : await dialog.showSaveDialog(dialogOptions);
    if (result.canceled || !result.filePath) {
      log('info', 'export canceled', traceId, { format: payload.format });
      return '';
    }
    authorization.grant(event, result.filePath);
    return result.filePath;
  }

  async function exportDocument(event, rawPayload, traceId) {
    const payload = validateExportDocumentPayload(rawPayload);
    const targetPath = await resolveTarget(event, payload, traceId);
    if (!targetPath) return null;
    let engine = 'builtin';
    log('info', 'export document start', traceId, { format: payload.format, targetPath });
    if (payload.format === 'html') {
      await fs.writeFile(targetPath, payload.html, 'utf8');
    } else if (['pdf', 'png', 'jpeg'].includes(payload.format)) {
      await engines.exportChromium(payload.html, payload.format, targetPath);
    } else {
      engine = 'pandoc';
      await engines.exportPandoc(payload.markdown, payload.format, targetPath, traceId);
    }
    const result = { filePath: targetPath, format: payload.format, engine };
    log('info', 'export document success', traceId, result);
    return result;
  }

  async function exportHtml(event, rawPayload, traceId) {
    const payload = validateExportHtmlPayload(rawPayload);
    return exportDocument(event, {
      format: 'html',
      html: payload.html,
      markdown: '',
      sourceFilePath: payload.sourceFilePath || undefined,
      targetPath: payload.targetPath || undefined,
    }, traceId);
  }

  return {
    exportDocument,
    exportHtml,
    getCapabilities: engines.getCapabilities,
  };
}

module.exports = {
  MAX_EXPORT_BYTES,
  createExportsService,
  deriveExportPath,
  getExportDialogFilter,
  validateExportDocumentPayload,
  validateExportHtmlPayload,
};
