import type { Ref } from 'vue';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import type { ImageInsertResult } from '../../types/native';

export type WorkflowRunner = (
  name: string, action: () => Promise<unknown>,
) => Promise<unknown>;

export function createParityImages(input: {
  domains: WorkbenchDomains; execute: WorkflowRunner; insertText: (text: string) => void;
  dialog: Ref<string>; busy: Ref<boolean>; notice: Ref<string>;
}) {
  const { domains, execute, insertText, dialog, busy, notice } = input;
  const editor = domains.editor.foundation;
  const { store, runtime, errorMessage } = editor;
  function imagePayload() {
    const frontMatter = /^---\r?\n([\s\S]*?)\r?\n---/.exec(store.content)?.[1] || '';
    const directory = /^typora-copy-images-to:\s*(.+)$/m.exec(frontMatter)?.[1]
      ?.trim().replace(/^['"]|['"]$/g, '');
    const filename = store.filePath.replace(/\\/g, '/').split('/').at(-1)
      ?.replace(/\.[^.]+$/, '') || 'document';
    return { sourceFilePath: store.filePath,
      copyDirectory: (directory || 'assets').replaceAll('${filename}', filename) };
  }

  function insertImages(result: ImageInsertResult) {
    if (!result.images.length) return;
    const markdown = result.images.map((image) => {
      const name = image.name.replace(/[[\]\\]/g, '\\$&');
      return `![${name}](${encodeURI(image.reference).replace(/[()]/g, (char) => {
        return `%${char.charCodeAt(0).toString(16)}`;
      })})`;
    }).join('\n\n');
    insertText(`\n\n${markdown}\n\n`);
  }

  async function ensureImageDocument() {
    runtime.flushEdits?.();
    if (!store.filePath && !await editor.saveCurrentDocument()) return false;
    if (!window.nativeBridge.images) throw new Error('图片资源操作仅在桌面版提供');
    return true;
  }

  async function insertLocalImages(paste = false) {
    await execute(paste ? '粘贴截图' : '插入图片', async () => {
      if (!await ensureImageDocument()) return;
      const result = await window.nativeBridge.images![paste ? 'paste' : 'select'](imagePayload());
      insertImages(result);
    });
  }

  async function setImageRoot() {
    await execute('设置本文图片根目录', async () => {
      if (!await ensureImageDocument()) return;
      const result = await window.nativeBridge.images!.setRoot(imagePayload());
      if (!result) return;
      const value = JSON.stringify(result.directoryPath.replace(/\\/g, '/'));
      let content = store.content;
      if (/^---\r?\n/.test(content)) {
        const match = /^---\r?\n([\s\S]*?)\r?\n---/.exec(content);
        if (!match) throw new Error('文档头部元数据未闭合，请先补全');
        const metadata = /^typora-root-url:/m.test(match[1])
          ? match[1].replace(/^typora-root-url:.*$/m, `typora-root-url: ${value}`)
          : `${match[1]}\ntypora-root-url: ${value}`;
        content = content.replace(match[0], `---\n${metadata}\n---`);
      } else content = `---\ntypora-root-url: ${value}\n---\n\n${content}`;
      store.setContent(content);
      notice.value = '已设置本文图片根目录；以 / 开头的图片路径将从该目录读取。';
    });
  }

  function pasteImage(event: ClipboardEvent) {
    const target = event.target instanceof Element ? event.target : null;
    if (!target?.closest('.preview-inline-editor, .editor-host, #typora-source, .preview')) return;
    if (!window.nativeBridge.images || !event.clipboardData?.files.length
      || ![...event.clipboardData.files].some((file) => file.type.startsWith('image/'))) return;
    event.preventDefault();
    void insertLocalImages(true);
  }

  function dropImage(event: DragEvent) {
    const dropped = [...(event.dataTransfer?.files || [])].filter((item) => {
      return /^image\//.test(item.type) || /\.(png|jpg|jpeg|gif|svg|webp|bmp)$/i.test(item.name);
    });
    if (!dropped.length || !window.nativeBridge.images?.drop || dialog.value || busy.value) return;
    event.preventDefault();
    event.stopPropagation();
    runtime.fileDragDepth = 0;
    domains.workspace.foundation().isDragOverFile.value = false;
    if (!store.filePath) {
      errorMessage.value = '请先保存文档，再拖入图片。';
      return;
    }
    const dropping = dropped.slice(0, 16).map((file) => {
      const result = window.nativeBridge.images!.drop!(file, imagePayload());
      void result.catch(() => {});
      return result;
    });
    void execute('拖入图片', async () => {
      const results = await Promise.allSettled(dropping);
      insertImages({ images: results.flatMap((result) => {
        return result.status === 'fulfilled' ? result.value.images : [];
      }) });
      const failed = results.find((result) => result.status === 'rejected');
      if (failed?.status === 'rejected') throw failed.reason;
    });
  }


  return { insertLocalImages, setImageRoot, pasteImage, dropImage };
}
