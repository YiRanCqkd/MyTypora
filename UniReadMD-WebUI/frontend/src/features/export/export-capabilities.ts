import { bridgeService } from '../../services/bridge';
import { BUILTIN_EXPORT_FORMATS } from './export-formats';
import type { useExportCapabilitiesState } from './export-capabilities-state';
import type { useDocumentExport } from './export-document';
import type { useTaskRunner } from '../workspace/task-runner';
type UseExportCapabilitiesDependencies =
  ReturnType<typeof useExportCapabilitiesState> &
  ReturnType<typeof useDocumentExport> &
  ReturnType<typeof useTaskRunner>;

export function useExportCapabilities(input: UseExportCapabilitiesDependencies) {
  const { exportCapabilities,exportDocumentByFormat,runTask,
  } = input;

  async function refreshExportCapabilities() {
    const result = await runTask('export-capabilities', (traceId) => {
      return bridgeService.getExportCapabilities(traceId);
    }, {
      fallbackHint: '导出能力探测失败时将仅保留 HTML 导出。',
    });

    if (result) {
      exportCapabilities.value = result;
      return;
    }

    exportCapabilities.value = {
      builtinFormats: [...BUILTIN_EXPORT_FORMATS],
      pandocAvailable: false,
      pandocVersion: '',
      pandocFormats: [],
    };
  }

  async function exportPdf() {
    await exportDocumentByFormat('pdf');
  }

  async function exportPng() {
    await exportDocumentByFormat('png');
  }

  async function exportJpeg() {
    await exportDocumentByFormat('jpeg');
  }

  async function exportDocx() {
    await exportDocumentByFormat('docx');
  }

  return { refreshExportCapabilities,exportPdf,exportPng,exportJpeg,exportDocx,
  };
}
