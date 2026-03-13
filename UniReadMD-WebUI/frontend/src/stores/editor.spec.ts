import { setActivePinia, createPinia } from 'pinia';
import { describe, expect, it, beforeEach } from 'vitest';
import { useEditorStore } from './editor';

describe('editor store', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
  });

  it('setFile 应更新内容并清理脏状态', () => {
    const store = useEditorStore();

    store.setDirty(true);
    store.setFile('I:/a.md', '# title');

    expect(store.filePath).toBe('I:/a.md');
    expect(store.content).toBe('# title');
    expect(store.isDirty).toBe(false);
  });

  it('setFile 应移除 UTF-8 BOM，保证首行标题可正常渲染', () => {
    const store = useEditorStore();

    store.setFile('I:/bom.md', '\uFEFF# 标题');

    expect(store.content).toBe('# 标题');
  });

  it('setCursorHint 应更新跳转提示行', () => {
    const store = useEditorStore();

    store.setCursorHint(42);

    expect(store.cursorHint).toBe(42);
  });

  it('setContent 应更新内容并置为未保存', () => {
    const store = useEditorStore();

    store.setFile('I:/a.md', '# old');
    store.setContent('# new');

    expect(store.content).toBe('# new');
    expect(store.isDirty).toBe(true);
  });

  it('restoreSession 应恢复会话状态', () => {
    const store = useEditorStore();

    store.restoreSession({
      filePath: 'I:/restored.md',
      content: '# restored',
      isDirty: true,
      cursorHint: 8,
    });

    expect(store.filePath).toBe('I:/restored.md');
    expect(store.content).toBe('# restored');
    expect(store.isDirty).toBe(true);
    expect(store.cursorHint).toBe(8);
  });

  it('restoreSession 应移除 UTF-8 BOM', () => {
    const store = useEditorStore();

    store.restoreSession({
      filePath: 'I:/restored-bom.md',
      content: '\uFEFF# restored',
      isDirty: true,
      cursorHint: 2,
    });

    expect(store.content).toBe('# restored');
  });
});
