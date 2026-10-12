import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createPinia } from 'pinia';
import { mount, flushPromises, type VueWrapper } from '@vue/test-utils';
import App from './App.vue';
import { useEditorStore } from './stores/editor';
import { installNativeBridgeFallback } from './services/native-bridge-fallback';

let wrapper: VueWrapper | undefined;
let pinia: ReturnType<typeof createPinia>;
const rect = new DOMRect(0, 0, 700, 400);

beforeEach(async () => {
  localStorage.clear();
  document.body.innerHTML = '';
  pinia = createPinia();
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    return window.setTimeout(() => callback(performance.now()), 16);
  });
  vi.stubGlobal('cancelAnimationFrame', (id: number) => window.clearTimeout(id));
  vi.stubGlobal('ResizeObserver', class {
    observe() {} unobserve() {} disconnect() {}
  });
  vi.spyOn(Element.prototype, 'getBoundingClientRect').mockReturnValue(rect);
  Object.defineProperty(Range.prototype, 'getBoundingClientRect', {
    configurable: true, value: () => rect,
  });
  Object.defineProperty(Range.prototype, 'getClientRects', {
    configurable: true, value: () => [rect],
  });
  Object.defineProperty(Element.prototype, 'scrollIntoView', {
    configurable: true, value: vi.fn(),
  });
  Object.defineProperty(Element.prototype, 'scrollTo', {
    configurable: true, value: vi.fn(),
  });
  Reflect.deleteProperty(window, 'nativeBridge');
  installNativeBridgeFallback();
  await window.nativeBridge.settings.set('documentToolsVisible', true);
  await window.nativeBridge.settings.set('preferencesVisible', true);
});

afterEach(async () => {
  wrapper?.unmount();
  wrapper = undefined;
  await flushPromises();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

async function start(content = '# 集成测试\n\n正文', filePath = '') {
  localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
    filePath, content, isDirty: false, cursorHint: 1,
  }));
  wrapper = mount(App, { attachTo: document.body, global: { plugins: [pinia] } });
  await flushPromises();
  return { app: wrapper, store: useEditorStore(pinia) };
}

function shortcut(key: string, options: KeyboardEventInit = {}) {
  window.dispatchEvent(new KeyboardEvent('keydown', { key, ctrlKey: true, ...options }));
}

describe('实际工作台集成', () => {
  it('插图插入当前选区，根目录写入元数据，取消选图保留内容', async () => {
    const { app, store } = await start('正文', '/workspace/a.md');
    const select = vi.fn(async () => ({ images: [{ name: '图片.png', reference: 'assets/a.png' }] }));
    window.nativeBridge.images = {
      select, paste: async () => ({ images: [] }),
      setRoot: async () => ({ directoryPath: '/workspace/images' }),
      applyRoot: async () => ({ ok: true }),
    };
    const button = (label: string) => app.findAll('.document-tools button')
      .find((item) => item.text() === label)!;
    await button('插入图片').trigger('click');
    await flushPromises();
    expect(store.content).toContain('![图片.png](assets/a.png)');
    expect(select).toHaveBeenCalledWith({
      sourceFilePath: '/workspace/a.md', copyDirectory: 'assets',
    });
    await button('图片根目录').trigger('click');
    await flushPromises();
    expect(store.content).toContain('typora-root-url: "/workspace/images"');
    const before = store.content;
    select.mockResolvedValue({ images: [] });
    await button('插入图片').trigger('click');
    await flushPromises();
    expect(store.content).toBe(before);
  });

  it('文件重命名更新当前路径但保留未保存修改和撤销历史', async () => {
    const { app, store } = await start('原始内容', '/workspace/a.md');
    window.nativeBridge.file.openDirectory = async () => ({
      directoryPath: '/workspace', filePaths: ['/workspace/a.md'],
    });
    window.nativeBridge.file.openFromPath = async () => ({
      filePath: '/workspace/改名.md', content: '原始内容', version: 'disk-version',
    });
    const operate = vi.fn(async () => ({ filePath: '/workspace/改名.md', action: 'rename' }));
    window.nativeBridge.file.operate = operate;
    await app.find('[data-testid=tab-files]').trigger('click');
    await app.find('[aria-label=打开文件夹]').trigger('click');
    await flushPromises();
    store.setContent('未保存修改');
    await flushPromises();
    await app.find('.files-tree-entry.file').trigger('contextmenu');
    await flushPromises();
    await app.findAll('.workspace-context-menu button').find((item) => item.text() === '重命名')!
      .trigger('click');
    await flushPromises();
    await app.find('.workbench-dialog input').setValue('改名.md');
    await app.find('.workbench-dialog form').trigger('submit');
    await flushPromises();
    expect(operate).toHaveBeenCalledWith(expect.objectContaining({
      action: 'rename', filePath: '/workspace/a.md', name: '改名.md',
    }));
    expect(store.filePath).toBe('/workspace/改名.md');
    expect(store.content).toBe('未保存修改');
    expect(store.savedContent).toBe('原始内容');
    expect(store.isDirty).toBe(true);
    expect(store.version).toBe('disk-version');
  });

  it('表格单元格编辑和新增行只更新对应表格，预览与源码保持一致', async () => {
    const { app, store } = await start('前文\n\n| 标题 |\n| --- |\n| 原值 |\n\n后文');
    await app.find('.preview tbody td').trigger('click');
    await flushPromises();
    const cell = app.find('.preview td[contenteditable=true]');
    expect(app.find('.preview-table-cell-editor-input').exists()).toBe(false);
    cell.element.textContent = '新值';
    await cell.trigger('input');
    await cell.trigger('keydown', { key: 'Enter' });
    await flushPromises();
    expect(store.content).toContain('新值');
    expect(store.content.startsWith('前文\n\n')).toBe(true);
    expect(store.content.endsWith('\n\n后文')).toBe(true);
    // 最后一行回车会提交单元格并新建下一行。
    expect(app.findAll('.preview tbody tr')).toHaveLength(2);
    await app.find('.preview tbody td').trigger('contextmenu');
    await flushPromises();
    await app.findAll('.preview-table-menu button')
      .find((button) => button.text().includes('下方插入行'))!.trigger('click');
    await flushPromises();
    expect(app.findAll('.preview tbody tr')).toHaveLength(3);
    expect(store.content.endsWith('\n\n后文')).toBe(true);
  });

  it('预览中的正则空匹配不会拆坏表情符号或陷入循环', async () => {
    const { app } = await start('😀');
    shortcut('f');
    await flushPromises();
    await app.find('[aria-label=正则表达式]').trigger('click');
    await app.find('[data-testid=doc-search-query]').setValue('(?=)');
    await flushPromises();
    expect(app.find('[data-testid=doc-search-summary]').text()).toBe('1/2');
    expect(app.find('.preview').text()).toBe('😀');
    expect(app.findAll('.doc-search-highlight')).toHaveLength(2);
  });
  it('挂载真实编辑器并完成搜索、替换、配色、专注和源码切换', async () => {
    const { app, store } = await start('foo foo\n\n# 标题');
    expect(app.find('.preview h1').text()).toBe('标题');
    await app.find('#toggle-sourceview-btn').trigger('click');
    shortcut('h');
    await flushPromises();
    await app.find('[data-testid=doc-search-query]').setValue('foo');
    await app.find('[data-testid=doc-replace-query]').setValue('$&');
    await app.findAll('.doc-search-btn').find((button) => button.text() === 'Replace All')!
      .trigger('click');
    await flushPromises();
    expect(store.content).toBe('$& $&\n\n# 标题');
    expect(store.isDirty).toBe(true);
    shortcut('f', { altKey: true });
    await flushPromises();
    expect(app.find('[data-testid=app-layout]').classes()).toContain('on-focus-mode');
    shortcut('t', { altKey: true });
    await app.find('.preferences-panel summary').trigger('click');
    await app.find('.preferences-panel select').setValue('dark');
    await flushPromises();
    expect(document.documentElement.dataset.theme).toBe('dark');
    await app.find('#toggle-sourceview-btn').trigger('click');
    expect(app.find('.preview h1').text()).toBe('标题');
  });

  it('任务切换、预览编辑、格式快捷键和取消不改动其他源块', async () => {
    const { app, store } = await start('- [ ] 任务\n\n普通段落');
    await app.find('.task-list-item-checkbox').trigger('click');
    await flushPromises();
    expect(store.content).toBe('- [x] 任务\n\n普通段落');
    await app.find('.preview > p').trigger('click', { altKey: true });
    await flushPromises();
    const textarea = app.find<HTMLTextAreaElement>('.preview-inline-editor-input');
    expect(textarea.exists()).toBe(true);
    await textarea.setValue('新段落');
    await textarea.trigger('compositionstart');
    await textarea.trigger('compositionend');
    await flushPromises();
    expect(store.content).toBe('- [x] 任务\n\n新段落');
    textarea.element.setSelectionRange(0, 3);
    await textarea.trigger('keydown', { key: 'b', ctrlKey: true });
    await textarea.trigger('compositionend');
    await flushPromises();
    expect(store.content).toContain('**新段落**');
    await textarea.trigger('keydown', { key: 'Escape' });
    await flushPromises();
    expect(store.content).toBe('- [x] 任务\n\n普通段落');
  });

  it('文件夹操作、正文结果定位和草稿恢复走真实组件入口', async () => {
    const { app, store } = await start();
    window.nativeBridge.file.openDirectory = vi.fn(async () => ({
      directoryPath: '/workspace', filePaths: ['/workspace/a.md'],
      directories: ['/workspace/空目录'],
    }));
    window.nativeBridge.file.openFromPath = vi.fn(async () => ({
      filePath: '/workspace/a.md', content: '# 文件\n\n目标正文',
    }));
    window.nativeBridge.file.searchDirectory = vi.fn(async () => ({
      matches: [{ filePath: '/workspace/a.md', line: 3, text: '目标正文' }],
    }));
    window.nativeBridge.session = {
      load: async () => null, save: async () => ({ ok: true }),
      confirmDiscard: async () => ({ action: 'discard' }),
      listDrafts: async () => ({
        drafts: [{ id: 'draft', filePath: '', savedAt: 1, title: '历史' }],
      }),
      restoreDraft: async () => ({
        filePath: '', originalFilePath: '', content: '历史草稿', isDirty: true,
      }),
    };
    await app.find('[data-testid=tab-files]').trigger('click');
    await app.find('[aria-label=打开文件夹]').trigger('click');
    await flushPromises();
    expect(app.find('.files-tree').text()).toContain('a.md');
    expect(app.find('.files-tree').text()).toContain('空目录');
    await app.find('[aria-label=搜索文件夹正文]').setValue('目标');
    await app.find('.workspace-search').trigger('submit');
    await flushPromises();
    await app.find('.workspace-search-results button').trigger('click');
    await flushPromises();
    expect(store.filePath).toBe('/workspace/a.md');
    expect(store.cursorHint).toBe(3);
    await app.findAll('.document-tools button').find((button) => button.text() === '历史草稿')!
      .trigger('click');
    await flushPromises();
    await app.find('.dialog-results button').trigger('click');
    await flushPromises();
    expect(store.filePath).toBe('');
    expect(store.content).toBe('历史草稿');
    expect(store.isDirty).toBe(true);
  });
});
