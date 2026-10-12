const { createContractError } = require('./errors.cjs');

function validateExportOptions(input = {}) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) {
    throw createContractError('ERR_INVALID_EXPORT_OPTIONS', '导出配置无效');
  }
  const result = {
    imageWidth: input.imageWidth ?? 1280, jpegQuality: input.jpegQuality ?? 92,
    includeStyles: input.includeStyles ?? true, headingPageBreak: input.headingPageBreak ?? false,
    theme: input.theme ?? 'current', title: input.title ?? '', author: input.author ?? '',
    language: input.language ?? 'zh-CN', referenceDoc: input.referenceDoc ?? '',
    epubCover: input.epubCover ?? '',
    tableOfContents: input.tableOfContents ?? false, tocDepth: input.tocDepth ?? 3,
    epubChapterLevel: input.epubChapterLevel ?? 1,
  };
  if (!Number.isInteger(result.imageWidth) || result.imageWidth < 256 || result.imageWidth > 4096
    || !Number.isInteger(result.jpegQuality) || result.jpegQuality < 1 || result.jpegQuality > 100
    || typeof result.includeStyles !== 'boolean' || typeof result.headingPageBreak !== 'boolean'
    || typeof result.tableOfContents !== 'boolean'
    || !Number.isInteger(result.tocDepth) || result.tocDepth < 1 || result.tocDepth > 6
    || !Number.isInteger(result.epubChapterLevel)
    || result.epubChapterLevel < 1 || result.epubChapterLevel > 6
    || !['current', 'light', 'dark'].includes(result.theme)
    || ['title', 'author'].some((key) => typeof result[key] !== 'string'
      || result[key].length > 200 || /[\x00-\x1f]/.test(result[key]))
    || typeof result.language !== 'string' || !/^[a-z]{2,3}(?:-[a-zA-Z0-9]{2,8})*$/.test(result.language)
    || ['referenceDoc', 'epubCover'].some((key) => typeof result[key] !== 'string'
      || result[key].length > 4096 || result[key].includes('\0'))) {
    throw createContractError('ERR_INVALID_EXPORT_OPTIONS', '导出配置超出允许范围');
  }
  return result;
}

function pandocOptionArguments(settings, format) {
  const args = [];
  for (const key of ['title', 'author', 'language']) {
    if (settings[key]) args.push('--metadata', `${key === 'language' ? 'lang' : key}=${settings[key]}`);
  }
  if (format === 'docx' && settings.referenceDoc) args.push('--reference-doc', settings.referenceDoc);
  if (format === 'epub' && settings.epubCover) args.push('--epub-cover-image', settings.epubCover);
  if (settings.tableOfContents) args.push('--toc', '--toc-depth', String(settings.tocDepth));
  if (format === 'epub') args.push('--split-level', String(settings.epubChapterLevel));
  if (['revealjs', 'opml'].includes(format)) args.push('--standalone');
  return args;
}

module.exports = { validateExportOptions, pandocOptionArguments };
