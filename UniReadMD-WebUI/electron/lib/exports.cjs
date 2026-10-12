const { createContractError, requireString } = require('./errors.cjs');
const { atomicWrite } = require('./atomic-write.cjs');
const { prepareExportHtml } = require('./export-html.cjs');
const MarkdownIt = require('markdown-it');
const { validatePdfOptions } = require('./pdf-options.cjs');
const { validateExportOptions } = require('./export-options.cjs');
const { IMAGE_TYPES } = require('./assets.cjs');
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
    ...(payload?.pdfOptions ? { pdfOptions: validatePdfOptions(payload.pdfOptions) } : {}),
    ...(payload?.options ? { options: validateExportOptions(payload.options) } : {}),
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
    epub: 'epub',
    latex: 'tex',
    revealjs: 'html', mediawiki: 'wiki', rst: 'rst', textile: 'textile', opml: 'opml',
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
    epub: { name: 'EPUB', extensions: ['epub'] },
    latex: { name: 'LaTeX', extensions: ['tex'] },
    revealjs: { name: 'RevealJS 幻灯片', extensions: ['html'] },
    mediawiki: { name: 'MediaWiki', extensions: ['wiki'] },
    rst: { name: 'reStructuredText', extensions: ['rst'] },
    textile: { name: 'Textile', extensions: ['textile'] },
    opml: { name: 'OPML 大纲', extensions: ['opml'] },
  };
  return filters[format] || filters.html;
}

function createExportsService(options) {
  const {
    authorization,
    assets,
    dialog,
    enforceAuthorization = true,
    engines,
    fs,
    getEventWindow,
    log = () => {},
    path,
    files,
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
    if (payload.sourceFilePath && enforceAuthorization) {
      authorization.assertAuthorized(event, payload.sourceFilePath);
    }
    for (const key of ['referenceDoc', 'epubCover']) {
      const filePath = payload.options?.[key];
      if (!filePath) continue;
      if (enforceAuthorization) authorization.assertAuthorized(event, filePath);
      const real = await fs.realpath(filePath);
      if (enforceAuthorization) authorization.assertAuthorized(event, real);
      const stat = await fs.stat(real);
      const extension = path.extname(real).toLowerCase();
      if (!stat.isFile() || stat.size > MAX_EXPORT_BYTES
        || (key === 'referenceDoc' ? extension !== '.docx' : !IMAGE_TYPES[extension])) {
        throw createContractError('ERR_INVALID_EXPORT_RESOURCE', '参考文档或封面格式、大小无效');
      }
      payload.options[key] = real;
    }
    const targetPath = await resolveTarget(event, payload, traceId);
    if (!targetPath) return null;
    const extensions = getExportDialogFilter(payload.format).extensions;
    if (!extensions.includes(path.extname(targetPath).slice(1).toLowerCase())) {
      throw createContractError('ERR_EXPORT_TARGET_EXTENSION', '导出文件扩展名与格式不一致');
    }
    let engine = 'builtin';
    log('info', 'export document start', traceId, { format: payload.format, targetPath });
    if (payload.format === 'html') {
      await atomicWrite(fs, targetPath, await prepareExportHtml(payload.html, event, assets));
    } else if (['pdf', 'png', 'jpeg'].includes(payload.format)) {
      const html = await prepareExportHtml(payload.html, event, assets);
      await engines.exportChromium(html, payload.format, targetPath,
        payload.pdfOptions, payload.options);
    } else {
      engine = 'pandoc';
      const localImages = new Map();
      let resourceBytes = 0;
      if (assets) {
        const token = payload.sourceFilePath
          ? new URL(assets.issue(event, payload.sourceFilePath)).pathname.split('/').at(-1) : '';
        const visit = async (tokens) => {
          for (const item of tokens) {
            if (item.type === 'image') {
              const source = item.attrGet('src') || '';
              if (!/^(https?:|data:image\/)/i.test(source)) {
                if (!token) throw new Error('请先保存文档，再导出包含本地图片的文档');
                const reference = decodeURIComponent(source.split(/[?#]/)[0]);
                if (!localImages.has(reference)) {
                  const resource = await assets.readForOwner(event, token, reference);
                  resourceBytes += resource.content.length;
                  if (resourceBytes > 64 * 1024 * 1024) throw new Error('导出图片超过 64 MiB');
                  localImages.set(reference, resource);
                }
              }
            }
            if (item.children) await visit(item.children);
          }
        };
        await visit(new MarkdownIt({ html: true }).parse(payload.markdown, {}));
      }
      await engines.exportPandoc(
        payload.markdown, payload.format, targetPath, traceId, payload.sourceFilePath, localImages,
        payload.options,
      );
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

  async function importDocument(event) {
    const capabilities = await engines.getCapabilities();
    if (!capabilities.pandocAvailable) throw new Error('导入需要安装 Pandoc 3.10');
    const win = getEventWindow(event);
    const selected = await dialog.showOpenDialog(win, {
      title: '导入文档', properties: ['openFile'],
      filters: [{ name: '文档', extensions: ['docx', 'odt', 'rtf', 'epub', 'html', 'tex', 'rst'] }],
    });
    if (selected.canceled || !selected.filePaths[0]) return null;
    const source = await fs.realpath(selected.filePaths[0]);
    const stat = await fs.stat(source);
    if (!stat.isFile() || stat.size > MAX_EXPORT_BYTES) throw new Error('导入文档超过 20 MiB');
    const saved = await dialog.showSaveDialog(win, {
      title: '保存导入的 Markdown', defaultPath: deriveExportPath(source, 'html', path)
        .replace(/\.html$/, '.md'), filters: [{ name: 'Markdown', extensions: ['md'] }],
    });
    if (saved.canceled || !saved.filePath) return null;
    if (!/\.md$/i.test(saved.filePath)) throw new Error('请选择 .md 目标文件');
    authorization.grant(event, saved.filePath);
    await engines.importPandoc(source, saved.filePath);
    return files.openFromPath(event, { filePath: saved.filePath });
  }

  async function chooseResource(event, payload) {
    const kind = payload.kind;
    if (!['referenceDoc', 'epubCover'].includes(kind)) {
      throw createContractError('ERR_INVALID_EXPORT_RESOURCE', '导出资源类型无效');
    }
    const result = await dialog.showOpenDialog(getEventWindow(event), {
      title: kind === 'referenceDoc' ? '选择 Word 参考文档' : '选择 EPUB 封面',
      properties: ['openFile'], filters: [{ name: '导出资源',
        extensions: kind === 'referenceDoc' ? ['docx'] : ['png', 'jpg', 'jpeg', 'webp'] }],
    });
    if (result.canceled || !result.filePaths[0]) return null;
    const filePath = await fs.realpath(result.filePaths[0]);
    authorization.grant(event, filePath);
    return { filePath };
  }

  return {
    exportDocument,
    exportHtml,
    getCapabilities: engines.getCapabilities,
    importDocument,
    chooseResource,
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
