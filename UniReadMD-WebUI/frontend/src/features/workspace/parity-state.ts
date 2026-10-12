import { computed, ref, watch } from 'vue';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import type { FileTreeRow } from './workspace-contracts';
import type { DraftItem, ExportFormat, PdfOptions } from '../../types/native';
import { readExportOptions } from '../export/export-options';

export function createParityState(domains: WorkbenchDomains) {
  const { store, errorMessage } = domains.editor.foundation;
  const workspace = domains.workspace.foundation();
  const dialog = ref<'' | 'quick' | 'drafts' | 'name' | 'pdf' | 'export'>('');
  const exportOptions = ref(readExportOptions());
  const quickQuery = ref('');
  const drafts = ref<DraftItem[]>([]);
  const entryName = ref('');
  const entryAction = ref('');
  const selectedRow = ref<FileTreeRow | null>(null);
  const fileMenu = ref(false);
  const fileMenuPosition = ref<Record<string, string>>({ left: '0px', top: '0px' });
  const notice = ref('');
  const externalChanged = ref(false);
  const busy = ref(false);
  let rootRevision = 0;
  watch(() => [store.filePath, /^---\r?\n([\s\S]*?)\r?\n---/.exec(store.content)?.[1] || ''],
    async ([sourceFilePath, frontMatter]) => {
      if (!sourceFilePath || !window.nativeBridge.images?.applyRoot) return;
      const revision = ++rootRevision;
      try {
        await window.nativeBridge.images.applyRoot({ sourceFilePath, frontMatter });
        if (revision === rootRevision) domains.editor.presentation().mathPluginRevision.value += 1;
      } catch (error) { errorMessage.value = String(error); }
    });
  const selectedExport = ref<ExportFormat>('pdf');
  const pdfOptions = ref<PdfOptions>({
    pageSize: 'A4', widthMm: 210, heightMm: 297, marginMm: 15,
    landscape: false, bookmarks: true, header: '', footer: '',
  });
  try {
    const saved = JSON.parse(localStorage.getItem('unireadmd.pdf-options') || 'null');
    if (saved) pdfOptions.value = { ...pdfOptions.value, ...saved };
  } catch { /* 损坏设置恢复默认，不影响文档。 */ }

  const quickFiles = computed(() => {
    const paths = [...new Set([...workspace.currentDirectoryMarkdownFiles.value,
      ...workspace.recentFiles.value])];
    const query = quickQuery.value.toLowerCase();
    return paths.filter((filePath) => {
      let position = -1;
      return [...query].every((character) => {
        position = filePath.toLowerCase().indexOf(character, position + 1);
        return position >= 0;
      });
    }).slice(0, 50);
  });
  const pandocAvailable = computed(() => {
    return Boolean(domains.export.foundation().exportCapabilities.value?.pandocAvailable);
  });
  const canExportSelected = computed(() => {
    return domains.export.foundation().availableExportFormats.value.has(selectedExport.value);
  });


  return { dialog, quickQuery, drafts, entryName, entryAction, selectedRow,
    fileMenu, fileMenuPosition, notice, externalChanged, busy, selectedExport,
    pdfOptions, exportOptions, quickFiles, pandocAvailable, canExportSelected };
}
