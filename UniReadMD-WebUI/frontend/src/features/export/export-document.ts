import type { ExportFormat, PdfOptions } from '../../types/native';
import { bridgeService } from '../../services/bridge';
import { createTraceId, logBridge } from '../../utils/trace';
import { EXPORT_LABELS } from './export-formats';
import { nextTick } from 'vue';
import { readExportOptions } from './export-options';
import { readWritingPreferences } from '../settings/writing-preferences';
import type { useWorkbenchFoundation } from '../editor/workbench-foundation';
import type { useDiagramFeature } from '../markdown/diagram-feature';
import type { useAvailableExportFormats } from './available-formats';
import type { useApplicationStatus } from '../editor/application-status';
import type { useExportCapabilitiesState } from './export-capabilities-state';
import type { usePreviewEditingState } from '../editor/preview-editing-state';
import type { useTaskRunner } from '../workspace/task-runner';
import type { useWorkbenchStore } from '../editor/editor-store';
type UseDocumentExportDependencies =
  ReturnType<typeof useWorkbenchFoundation> &
  ReturnType<typeof useDiagramFeature> &
  ReturnType<typeof useAvailableExportFormats> &
  ReturnType<typeof useApplicationStatus> &
  ReturnType<typeof useExportCapabilitiesState> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof useTaskRunner> &
  ReturnType<typeof useWorkbenchStore>;

export function useDocumentExport(input: UseDocumentExportDependencies) {
  const { availableExportFormats,errorMessage,exportCapabilities,
    renderedHtml,runTask,store,
  } = input;

  async function buildExportHtmlDocument() {
    const options = readExportOptions();
    input.runtime.flushEdits?.();
    await input.ensureMathPluginLoaded();
    await nextTick();
    await input.renderMermaidBlocks(true);
    const documentCopy = document.implementation.createHTMLDocument();
    documentCopy.documentElement.lang = 'zh-CN';
    documentCopy.documentElement.dataset.theme = options.theme === 'current'
      ? document.documentElement.dataset.theme || 'light' : options.theme;
    const writing = readWritingPreferences();
    documentCopy.documentElement.dataset.writingTheme = options.theme === 'current'
      ? document.documentElement.dataset.writingTheme || 'paper'
      : options.theme === 'dark' ? writing.darkTheme : writing.lightTheme;
    documentCopy.title = options.title || store.filePath || 'UniReadMD 导出';
    documentCopy.documentElement.lang = options.language;
    const styles: string[] = [];
    for (const sheet of options.includeStyles ? document.styleSheets : []) {
      try {
        let css = Array.from(sheet.cssRules, (rule) => rule.cssText).join('\n');
        const references = [...css.matchAll(/url\(\s*['"]?([^'"\s)]+)['"]?\s*\)/g)];
        for (const match of references) {
          const url = new URL(match[1], sheet.href || document.baseURI);
          if (url.origin !== location.origin) continue;
          const response = await fetch(url);
          if (!response.ok) throw new Error('导出字体读取失败');
          const blob = await response.blob();
          if (blob.size > 10 * 1024 * 1024) throw new Error('导出字体超过大小限制');
          const data = await new Promise<string>((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(String(reader.result));
            reader.onerror = () => reject(new Error('导出字体读取失败'));
            reader.readAsDataURL(blob);
          });
          css = css.replaceAll(match[0], `url('${data}')`);
        }
        styles.push(css);
      } catch (error) {
        if (error instanceof DOMException && error.name === 'SecurityError') continue;
        throw error;
      }
    }
    styles.push('.preview { display: block; height: auto; overflow: visible; margin: 0 auto; }'
      + ' .preview-editing-block { visibility: visible; }'
      + ' body { display: block; padding: 32px; height: auto; overflow: visible; }');
    const style = documentCopy.createElement('style');
    style.textContent = styles.join('\n');
    if (options.headingPageBreak) {
      style.textContent += '\n@media print { '
        + '.preview h1:not(:first-child) { break-before: page; } }';
    }
    documentCopy.head.append(style);
    const article = documentCopy.createElement('article');
    article.className = 'preview';
    article.id = 'write';
    article.innerHTML = input.previewHost.value?.innerHTML || renderedHtml.value;
    article.querySelectorAll(
      '.preview-inline-editor, .preview-table-cell-editor, .preview-empty-line',
    )
      .forEach((element) => element.remove());
    article.querySelectorAll('.preview-editing-fragment, .preview-editing-block')
      .forEach((element) => {
        element.classList.remove('preview-editing-fragment', 'preview-editing-block');
        (element as HTMLElement).style.removeProperty('--preview-edit-width');
        (element as HTMLElement).style.removeProperty('--preview-edit-height');
      });
    documentCopy.body.append(article);
    return `<!doctype html>\n${documentCopy.documentElement.outerHTML}`;
  }

  function getExportFallbackHint(format: ExportFormat) {
    if (format === 'docx' || format === 'odt' || format === 'rtf') {
      return '请确认已安装 Pandoc，或先导出 HTML/PDF/Image。';
    }

    if (format === 'pdf' || format === 'png' || format === 'jpeg') {
      return '可先导出 HTML，再通过系统打印或截图作为临时回退。';
    }

    return '可检查目标文件路径权限，或改用“另存为”后重试。';
  }

  function buildExportMarkdownDocument() {
    return store.content.replace(/\r\n/g, '\n');
  }

  async function exportDocumentByFormat(format: ExportFormat) {
    if (!availableExportFormats.value.has(format)) {
      const taskName = `export-${format}`;
      const traceId = createTraceId(taskName);
      const hint = getExportFallbackHint(format);
      errorMessage.value = `当前环境暂不支持导出 ${EXPORT_LABELS[format]}（追踪ID: ${traceId}）；建议：${hint}`;
      await logBridge('error', `${taskName}:unsupported`, traceId, {
        format,
        capability: exportCapabilities.value,
        hint,
      });
      return;
    }

    let pdfOptions: PdfOptions | undefined;
    if (format === 'pdf') {
      try {
        pdfOptions = JSON.parse(localStorage.getItem('unireadmd.pdf-options') || 'null');
      } catch { /* 损坏设置使用默认纸张，不能阻断导出。 */ }
    }
    await runTask(`export-${format}`, async (traceId) => {
      const options = readExportOptions();
      return bridgeService.exportDocument({
        format,
        html: await buildExportHtmlDocument(),
        markdown: buildExportMarkdownDocument(),
        sourceFilePath: store.filePath || undefined,
        ...(format === 'pdf' ? { pdfOptions: pdfOptions || undefined } : {}),
        traceId,
        options,
      });
    }, {
      fallbackHint: getExportFallbackHint(format),
    });
  }

  async function exportHtml() {
    await exportDocumentByFormat('html');
  }

  return { buildExportHtmlDocument,getExportFallbackHint,buildExportMarkdownDocument,
    exportDocumentByFormat,exportHtml,
  };
}
