import type { ExportOptions } from '../../types/native';
let resources = { referenceDoc: '', epubCover: '' };

export function setExportResources(options: ExportOptions) {
  resources = { referenceDoc: options.referenceDoc, epubCover: options.epubCover };
}

export function defaultExportOptions(): ExportOptions {
  return {
    imageWidth: 1280, jpegQuality: 92, includeStyles: true, headingPageBreak: false,
    theme: 'current', title: '', author: '', language: 'zh-CN', referenceDoc: '', epubCover: '',
    tableOfContents: false, tocDepth: 3, epubChapterLevel: 1,
  };
}

export function readExportOptions(): ExportOptions {
  const defaults = defaultExportOptions();
  try {
    const saved = JSON.parse(localStorage.getItem('unireadmd.export-options') || 'null');
    if (!saved || typeof saved !== 'object' || Array.isArray(saved)) return defaults;
    for (const key of Object.keys(defaults) as Array<keyof ExportOptions>) {
      if (typeof saved[key] === typeof defaults[key]) {
        Object.assign(defaults, { [key]: saved[key] });
      }
    }
    // 文件选择授权属于当前窗口，重启后需要重新选择，不能复用路径扩大权限。
    Object.assign(defaults, resources);
  } catch { /* 损坏设置恢复默认，不阻断导出。 */ }
  return defaults;
}
