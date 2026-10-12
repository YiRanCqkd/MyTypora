import type { Ref } from 'vue';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import type { FileTreeRow } from './workspace-contracts';
import type { WorkflowRunner } from './parity-images';
import { bridgeService } from '../../services/bridge';

export function createParityFileOperations(input: {
  domains: WorkbenchDomains; execute: WorkflowRunner;
  selectedRow: Ref<FileTreeRow | null>; fileMenu: Ref<boolean>;
  fileMenuPosition: Ref<{ left: string; top: string }>;
  entryName: Ref<string>; entryAction: Ref<string>; dialog: Ref<string>;
}) {
  const { domains, execute, selectedRow, fileMenu, fileMenuPosition,
    entryName, entryAction, dialog } = input;
  const editor = domains.editor.foundation;
  const { store, runtime } = editor;
  const actions = domains.workspace.actions();
  function openFileMenu(event: MouseEvent, row: FileTreeRow) {
    if (row.path.startsWith('dir:')) return;
    selectedRow.value = row;
    fileMenuPosition.value = { left: `${Math.min(event.clientX, innerWidth - 220)}px`,
      top: `${Math.min(event.clientY, innerHeight - 290)}px` };
    fileMenu.value = true;
  }

  function requestEntry(action: string) {
    entryAction.value = action;
    fileMenu.value = false;
    const row = selectedRow.value;
    if (!row) return;
    entryName.value = action === 'mkdir' ? '新建文件夹' : action === 'create' ? '新文档.md'
      : action === 'duplicate' ? row.label.replace(/(\.[^.]+)$/, ' 副本$1') : row.label;
    dialog.value = 'name';
  }

  async function applyFileOperation(action: string, directoryPath?: string) {
    const row = selectedRow.value;
    const operate = window.nativeBridge.file.operate;
    if (!row || !operate) return;
    runtime.flushEdits?.();
    const current = store.filePath.replace(/\\/g, '/').toLowerCase();
    const source = row.path.replace(/\\/g, '/').toLowerCase();
    const affectsDocument = current === source || current.startsWith(`${source}/`);
    if ((action === 'trash' && affectsDocument || action === 'create')
      && !await editor.ensureCanReplaceDocument()) return;
    // 先保存活动文档再迁移，避免用磁盘旧版本覆盖未保存的图片引用。
    if (action === 'move' && affectsDocument && store.isDirty
      && !await editor.saveCurrentDocument()) return;
    const revision = store.documentRevision;
    const originalContent = store.content;
    await execute('文件操作', async () => {
      const result = await operate({ action, filePath: row.path,
        name: ['move', 'trash'].includes(action) ? undefined : entryName.value, directoryPath });
      if (action === 'trash' && affectsDocument) store.setFile('', '');
      if (['rename', 'move'].includes(action) && affectsDocument) {
        const target = result.filePath + store.filePath.slice(row.path.length);
        const disk = await bridgeService.openFromPath(target);
        if (disk && store.documentRevision === revision) {
          if (action === 'move' && store.content === originalContent) {
            store.setContent(disk.content);
          }
          store.markSaved(target, disk.content, disk);
          domains.settings.session().pushRecentFile(target);
        }
      }
      if (action === 'create') {
        const document = await bridgeService.openFromPath(result.filePath);
        if (document) {
          store.setFile(document.filePath, document.content, document);
          domains.settings.session().pushRecentFile(document.filePath);
        }
      }
      await actions.refreshCurrentDirectoryFiles(store.filePath);
      dialog.value = '';
      fileMenu.value = false;
    });
  }

  async function trashEntry() {
    fileMenu.value = false;
    if (!window.confirm(`将“${selectedRow.value?.label}”移入回收站？`)) return;
    await applyFileOperation('trash');
  }

  async function moveEntry() {
    fileMenu.value = false;
    const result = await window.nativeBridge.file.chooseMoveDirectory?.();
    if (result) await applyFileOperation('move', result.directoryPath);
  }

  function startFileDrag(event: DragEvent, row: FileTreeRow) {
    if (row.path.startsWith('dir:')) return;
    event.dataTransfer?.setData('application/x-unireadmd-file', row.path);
  }

  async function dropOnDirectory(event: DragEvent, row: FileTreeRow) {
    const source = event.dataTransfer?.getData('application/x-unireadmd-file');
    if (!source || row.type !== 'dir') return;
    event.preventDefault();
    event.stopPropagation();
    selectedRow.value = { ...row, path: source };
    await applyFileOperation('move', row.path);
  }


  return { openFileMenu, requestEntry, applyFileOperation, trashEntry, moveEntry,
    startFileDrag, dropOnDirectory };
}
