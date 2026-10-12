const { createContractError } = require('./errors.cjs');

function validatePdfOptions(input = {}) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) {
    throw createContractError('ERR_INVALID_PDF_OPTIONS', 'PDF 设置格式无效');
  }
  const result = {
    pageSize: input.pageSize ?? 'A4', marginMm: input.marginMm ?? 15,
    landscape: input.landscape ?? false, bookmarks: input.bookmarks ?? true,
    header: input.header ?? '', footer: input.footer ?? '',
  };
  if (!['A4', 'A5', 'Letter', 'Legal', 'Custom'].includes(result.pageSize)
    || typeof result.marginMm !== 'number' || !Number.isFinite(result.marginMm)
    || result.marginMm < 0 || result.marginMm > 50
    || typeof result.landscape !== 'boolean' || typeof result.bookmarks !== 'boolean'
    || ['header', 'footer'].some((key) => typeof result[key] !== 'string'
      || result[key].length > 200)) {
    throw createContractError('ERR_INVALID_PDF_OPTIONS', 'PDF 设置超出允许范围');
  }
  const bounded = (value, min, max) => typeof value === 'number' && Number.isFinite(value)
    && value >= min && value <= max;
  if (result.pageSize === 'Custom') {
    if (!bounded(input.widthMm, 50, 1000) || !bounded(input.heightMm, 50, 1000)) {
      throw createContractError('ERR_INVALID_PDF_OPTIONS', '自定义纸张范围为 50 至 1000 毫米');
    }
    result.widthMm = input.widthMm;
    result.heightMm = input.heightMm;
  }
  if (input.marginsMm !== undefined) {
    if (!input.marginsMm || ['top', 'bottom', 'left', 'right'].some((side) => {
      return !bounded(input.marginsMm[side], 0, 50);
    })) throw createContractError('ERR_INVALID_PDF_OPTIONS', '各侧页边距范围为 0 至 50 毫米');
    result.marginsMm = Object.fromEntries(['top', 'bottom', 'left', 'right']
      .map((side) => [side, input.marginsMm[side]]));
  }
  return result;
}

function headerTemplate(text) {
  const escaped = text.replace(/[&<>"']/g, (char) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  })[char]);
  const content = escaped.replace(/\{\{page\}\}|\$\{pageNo\}/g,
    '<span class="pageNumber"></span>').replace(/\{\{pages\}\}|\$\{pageCount\}/g,
    '<span class="totalPages"></span>');
  return `<div style="font-size:9px;width:100%;text-align:center">${content}</div>`;
}

function createPrintOptions(input) {
  const settings = validatePdfOptions(input);
  const margin = settings.marginMm / 25.4;
  return {
    printBackground: true,
    pageSize: settings.pageSize === 'Custom'
      ? { width: settings.widthMm / 25.4, height: settings.heightMm / 25.4 } : settings.pageSize,
    landscape: settings.landscape,
    margins: settings.marginsMm ? Object.fromEntries(Object.entries(settings.marginsMm)
      .map(([side, value]) => [side, value / 25.4]))
      : { top: margin, bottom: margin, left: margin, right: margin },
    generateDocumentOutline: settings.bookmarks, generateTaggedPDF: true,
    displayHeaderFooter: Boolean(settings.header || settings.footer),
    headerTemplate: headerTemplate(settings.header), footerTemplate: headerTemplate(settings.footer),
  };
}

module.exports = { validatePdfOptions, createPrintOptions };
