import type { ExportFormat } from '../../types/native';
import { bridgeService } from '../../services/bridge';
import { createTraceId, logBridge } from '../../utils/trace';
import { EXPORT_LABELS } from './export-formats';
import type { useAvailableExportFormats } from './available-formats';
import type { useApplicationStatus } from '../editor/application-status';
import type { useExportCapabilitiesState } from './export-capabilities-state';
import type { usePreviewEditingState } from '../editor/preview-editing-state';
import type { useTaskRunner } from '../workspace/task-runner';
import type { useWorkbenchStore } from '../editor/editor-store';
type UseDocumentExportDependencies =
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

  function buildExportHtmlDocument() {
    const titleText = store.filePath || 'UniReadMD Export';

    return `<!doctype html>
  <html lang='zh-CN'>
    <head>
      <meta charset='UTF-8' />
      <meta name='viewport' content='width=device-width, initial-scale=1.0' />
      <title>${titleText}</title>
    </head>
    <body>
      ${renderedHtml.value}
    </body>
  </html>`;
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

    await runTask(`export-${format}`, (traceId) => {
      return bridgeService.exportDocument({
        format,
        html: buildExportHtmlDocument(),
        markdown: buildExportMarkdownDocument(),
        sourceFilePath: store.filePath || undefined,
        traceId,
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
