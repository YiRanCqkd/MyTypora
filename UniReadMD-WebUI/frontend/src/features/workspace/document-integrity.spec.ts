import { beforeEach, describe, expect, it, vi } from 'vitest';
import { createPinia, setActivePinia } from 'pinia';
import { ref } from 'vue';
import { useEditorStore } from '../../stores/editor';
import { bridgeService } from '../../services/bridge';
import { createWorkbenchRuntime } from '../editor/workbench-runtime';
import { useDocumentIntegrity } from './document-integrity';

describe('文档替换与保存完整性', () => {
  it('另存为图片文档应用主进程迁移正文，保存期间新增编辑不会丢失', async () => {
    const input = create('save');
    const original = '![图片](old.png)';
    const relocated = '![图片](document-assets/image.png)';
    input.store.setContent(original);
    const save = vi.spyOn(bridgeService, 'saveMarkdown').mockResolvedValueOnce({
      filePath: 'I:/new/document.md', content: relocated,
    });
    expect(await input.saveCurrentDocument(true)).toBe(true);
    expect(save).toHaveBeenCalledWith(expect.objectContaining({
      sourceFilePath: 'I:/test.md', content: original,
    }));
    expect(input.store.content).toBe(relocated);
    expect(input.store.savedContent).toBe(relocated);
    input.store.setContent(original);
    save.mockImplementationOnce(async () => {
      input.store.setContent('保存期间新写的内容');
      return { filePath: 'I:/new/second.md', content: relocated };
    });
    expect(await input.saveCurrentDocument(true)).toBe(false);
    expect(input.store.content).toBe('保存期间新写的内容');
    expect(input.store.savedContent).toBe(relocated);
    expect(input.store.isDirty).toBe(true);
  });

  it('首次保存图片草稿没有原始文件路径，迁移结果相同不重写编辑状态', async () => {
    const input = create('save');
    input.store.setFile('', '');
    input.store.setContent('![图片](https://example.com/a.png)');
    const save = vi.spyOn(bridgeService, 'saveMarkdown').mockImplementationOnce(async (payload) => {
      expect(payload.sourceFilePath).toBeUndefined();
      return { filePath: 'I:/first.md', content: payload.content };
    });
    expect(await input.saveCurrentDocument(true)).toBe(true);
    expect(save).toHaveBeenCalledOnce();
    expect(input.store.isDirty).toBe(false);
  });
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.restoreAllMocks();
  });

  function create(action: 'save' | 'discard' | 'cancel') {
    const store = useEditorStore();
    store.setFile('I:/test.md', '原文');
    store.setContent('编辑内容');
    window.nativeBridge = { session: {
      confirmDiscard: vi.fn().mockResolvedValue({ action }),
    } } as unknown as Window['nativeBridge'];
    const runtime = createWorkbenchRuntime().runtime;
    const errorMessage = ref('');
    return { store, runtime, errorMessage,
      ...useDocumentIntegrity({ store, runtime, errorMessage }),
    };
  }

  it('取消保留正文和脏状态，放弃才允许替换', async () => {
    const cancelled = create('cancel');
    expect(await cancelled.ensureCanReplaceDocument()).toBe(false);
    expect(cancelled.store.content).toBe('编辑内容');
    expect(cancelled.store.isDirty).toBe(true);
    expect(await create('discard').ensureCanReplaceDocument()).toBe(true);
  });

  it('保存失败或取消另存为时禁止替换，显示错误', async () => {
    const input = create('save');
    vi.spyOn(bridgeService, 'saveMarkdown').mockRejectedValueOnce(new Error('磁盘已满'))
      .mockResolvedValueOnce(null);
    expect(await input.ensureCanReplaceDocument()).toBe(false);
    expect(input.errorMessage.value).toBe('磁盘已满');
    expect(await input.ensureCanReplaceDocument()).toBe(false);
    expect(input.store.isDirty).toBe(true);
  });

  it('先提交浮层编辑，保存期间新增修改仍保持未保存', async () => {
    const input = create('save');
    input.runtime.flushEdits = () => input.store.setContent('已提交浮层');
    vi.spyOn(bridgeService, 'saveMarkdown').mockImplementation(async (payload) => {
      expect(payload.content).toBe('已提交浮层');
      input.store.setContent('保存期间的修改');
      return { filePath: 'I:/test.md', version: '新版本' };
    });
    expect(await input.saveCurrentDocument()).toBe(false);
    expect(input.store.savedContent).toBe('已提交浮层');
    expect(input.store.content).toBe('保存期间的修改');
    expect(input.store.isDirty).toBe(true);
  });

  it('保存成功允许替换，且不会把旧文档保存结果套到新文档', async () => {
    const input = create('save');
    vi.spyOn(bridgeService, 'saveMarkdown').mockResolvedValueOnce({ filePath: 'I:/test.md' })
      .mockImplementationOnce(async () => {
        input.store.setFile('I:/new.md', '新文档');
        return { filePath: 'I:/test.md' };
      });
    expect(await input.ensureCanReplaceDocument()).toBe(true);
    input.store.setContent('再次编辑');
    expect(await input.saveCurrentDocument()).toBe(false);
    expect(input.store.filePath).toBe('I:/new.md');
    expect(input.store.content).toBe('新文档');
  });

  it('重复确认共用一个对话框，正在保存时不会发出第二次写入', async () => {
    const input = create('save');
    let finish!: (value: { filePath: string }) => void;
    const save = vi.spyOn(bridgeService, 'saveMarkdown').mockImplementation(() => {
      return new Promise((resolve) => { finish = resolve; });
    });
    const first = input.ensureCanReplaceDocument();
    expect(input.ensureCanReplaceDocument()).toBe(first);
    await vi.waitFor(() => expect(save).toHaveBeenCalledOnce());
    expect(await input.saveCurrentDocument()).toBe(false);
    finish({ filePath: 'I:/test.md' });
    expect(await first).toBe(true);
    expect(window.nativeBridge.session?.confirmDiscard).toHaveBeenCalledOnce();
    expect(save).toHaveBeenCalledOnce();
  });

  it('确认或保存返回非 Error 异常时仍保留草稿并显示提示', async () => {
    const input = create('save');
    vi.mocked(window.nativeBridge.session!.confirmDiscard).mockRejectedValue('确认异常');
    expect(await input.ensureCanReplaceDocument()).toBe(false);
    expect(input.errorMessage.value).toBe('无法确认未保存修改');
    vi.spyOn(bridgeService, 'saveMarkdown').mockRejectedValue('保存异常');
    expect(await input.saveCurrentDocument()).toBe(false);
    expect(input.errorMessage.value).toContain('保存失败');
    expect(input.store.content).toBe('编辑内容');
    expect(input.store.isDirty).toBe(true);
  });

  it('浏览器取消阻止替换，干净文档无需确认，另存为不携带原文件版本', async () => {
    const input = create('cancel');
    window.nativeBridge = {} as Window['nativeBridge'];
    const confirmation = vi.spyOn(window, 'confirm').mockReturnValueOnce(false)
      .mockReturnValueOnce(true);
    expect(await input.ensureCanReplaceDocument()).toBe(false);
    expect(await input.ensureCanReplaceDocument()).toBe(true);
    input.store.version = '旧版本';
    const save = vi.spyOn(bridgeService, 'saveMarkdown')
      .mockResolvedValue({ filePath: 'I:/另存为.md' });
    expect(await input.saveCurrentDocument(true)).toBe(true);
    expect(save).toHaveBeenCalledWith({ filePath: undefined, version: undefined,
      content: '编辑内容',
    });
    expect(await input.ensureCanReplaceDocument()).toBe(true);
    expect(confirmation).toHaveBeenCalledTimes(2);
    expect(input.store.filePath).toBe('I:/另存为.md');
  });
});
