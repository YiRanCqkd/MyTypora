import type { ExportFormat } from '../../types/native';

export const BUILTIN_EXPORT_FORMATS: readonly ExportFormat[] = [
  'html',
  'pdf',
  'png',
  'jpeg',
];

export const EXPORT_LABELS: Readonly<Record<ExportFormat, string>> = {
  html: 'HTML',
  pdf: 'PDF',
  png: 'PNG',
  jpeg: 'JPEG',
  docx: 'Word(docx)',
  odt: 'ODT',
  rtf: 'RTF',
  epub: 'EPUB',
  latex: 'LaTeX',
  revealjs: 'RevealJS 幻灯片', mediawiki: 'MediaWiki', rst: 'reStructuredText',
  textile: 'Textile', opml: 'OPML 大纲',
};
