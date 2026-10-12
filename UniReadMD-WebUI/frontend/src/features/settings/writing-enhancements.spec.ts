import { afterEach, describe, expect, it, vi } from 'vitest';
import { mount, flushPromises } from '@vue/test-utils';
import { defineComponent, h, reactive, ref } from 'vue';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import { useWritingEnhancements } from './writing-enhancements';
import { WRITING_KEY } from './writing-preferences';

const originalBridge = Object.getOwnPropertyDescriptor(window, 'nativeBridge');
const cleanups: Array<() => void> = [];
const saved = { statUnit: 'characters', readingSpeed: 300, equationNumbering: true,
  lightTheme: 'sepia', darkTheme: 'midnight' };

afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup());
  if (originalBridge) Object.defineProperty(window, 'nativeBridge', originalBridge);
  else Reflect.deleteProperty(window, 'nativeBridge');
  localStorage.clear();
  delete document.documentElement.dataset.writingTheme;
  delete document.documentElement.dataset.theme;
});

function setup(getAll: () => Promise<unknown>, set = vi.fn(async (...args: unknown[]) => {
  void args;
  return {};
})) {
  Object.defineProperty(window, 'nativeBridge', {
    configurable: true, value: { settings: { getAll, set } },
  });
  const errorMessage = ref('');
  const mathPluginRevision = ref(0);
  const foundation = { store: reactive({ documentRevision: 0 }), runtime: { editorView: null },
    previewHost: ref(null), errorMessage };
  const presentation = { contentCharCount: ref(20), contentWordCount: ref(10), mathPluginRevision };
  const domains = { editor: { foundation, presentation: () => presentation } };
  let writing!: ReturnType<typeof useWritingEnhancements>;
  const wrapper = mount(defineComponent({ setup() {
    writing = useWritingEnhancements(domains as unknown as WorkbenchDomains);
    return () => h('div');
  } }));
  cleanups.push(() => wrapper.unmount());
  return { writing, set, errorMessage, mathPluginRevision };
}

describe('原生写作设置持久化', () => {
  it('空浏览器存储从原生配置恢复，并同步公式和实际主题', async () => {
    document.documentElement.dataset.theme = 'dark';
    const state = setup(async () => ({ settings: {
      writingPreferencesJson: JSON.stringify(saved),
    } }));
    await flushPromises();
    expect(state.writing.preferences).toEqual(saved);
    expect(document.documentElement.dataset.writingTheme).toBe('midnight');
    expect(JSON.parse(localStorage.getItem(WRITING_KEY)!)).toEqual(saved);
    expect(state.mathPluginRevision.value).toBe(1);
    state.writing.preferences.readingSpeed = 400;
    await flushPromises();
    expect(state.set).toHaveBeenLastCalledWith('writingPreferencesJson',
      JSON.stringify({ ...saved, readingSpeed: 400 }), undefined);
  });

  it('启动读取期间的用户修改优先于旧设置，IPC 只传可克隆字符串', async () => {
    let resolve!: (value: unknown) => void;
    const state = setup(() => new Promise((done) => { resolve = done; }));
    state.writing.preferences.readingSpeed = 500;
    resolve({ settings: { writingPreferencesJson: JSON.stringify(saved) } });
    await flushPromises();
    expect(state.writing.preferences.readingSpeed).toBe(500);
    expect(typeof state.set.mock.calls[0][1]).toBe('string');
    expect(JSON.parse(String(state.set.mock.calls[0][1])).readingSpeed).toBe(500);
  });

  it('原生读取失败提示错误并保留本地配置，不写回默认值', async () => {
    const state = setup(async () => { throw new Error('磁盘读取失败'); });
    await flushPromises();
    expect(state.errorMessage.value).toContain('写作设置读取失败');
    expect(state.set).not.toHaveBeenCalled();
  });
});
