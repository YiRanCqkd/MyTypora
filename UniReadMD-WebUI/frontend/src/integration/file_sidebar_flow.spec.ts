import { beforeEach, describe, expect, it, vi } from 'vitest';
import { createPinia, setActivePinia } from 'pinia';
import { useEditorStore } from '../stores/editor';
import { bridgeService } from '../services/bridge';
import {
  extractHeadings,
  findActiveHeadingLine,
  findMatchedLines,
} from '../utils/markdown';

function installMockBridge() {
  const bridge = {
    app: {
      getVersion: vi.fn(),
      logEvent: vi.fn(),
    },
    file: {
      openMarkdown: vi.fn(),
      openFromPath: vi.fn(),
      saveMarkdown: vi.fn(),
    },
    exporter: {
      exportHtml: vi.fn(),
      exportDocument: vi.fn(),
      getCapabilities: vi.fn(),
    },
    window: {
      minimize: vi.fn(),
      toggleMaximize: vi.fn(),
      isMaximized: vi.fn(),
      close: vi.fn(),
      toggleDevTools: vi.fn(),
    },
    settings: {
      getAll: vi.fn(),
      get: vi.fn(),
      set: vi.fn(),
      selectUserCssFile: vi.fn(),
    },
    spell: {
      getState: vi.fn(),
      setEnabled: vi.fn(),
      setLanguage: vi.fn(),
    },
    getAppVersion: vi.fn(),
    openMarkdownFile: vi.fn(),
    saveMarkdownFile: vi.fn(),
  };

  Object.defineProperty(window, 'nativeBridge', {
    configurable: true,
    value: bridge,
  });

  return bridge;
}

describe('integration: file + sidebar flow', () => {
  beforeEach(() => {
    setActivePinia(createPinia());
    vi.restoreAllMocks();
  });

  it('应完成打开-编辑-保存主链路', async () => {
    const bridge = installMockBridge();
    const store = useEditorStore();

    bridge.file.openMarkdown.mockResolvedValue({
      filePath: 'I:/docs/sample.md',
      content: '# H1\n\nold content',
    });
    bridge.file.saveMarkdown.mockResolvedValue({
      filePath: 'I:/docs/sample.md',
    });

    const opened = await bridgeService.openMarkdown('trace-open');
    expect(opened).not.toBeNull();

    store.setFile(opened!.filePath, opened!.content);
    expect(store.filePath).toBe('I:/docs/sample.md');
    expect(store.isDirty).toBe(false);

    store.setContent('# H1\n\nnew content');
    expect(store.isDirty).toBe(true);

    const saved = await bridgeService.saveMarkdown({
      filePath: store.filePath,
      content: store.content,
      traceId: 'trace-save',
    });
    expect(saved?.filePath).toBe('I:/docs/sample.md');

    store.markSaved(saved!.filePath);
    expect(store.isDirty).toBe(false);
  });

  it('应保持 Outline 与 Search 的状态联动', () => {
    const store = useEditorStore();
    store.setFile(
      'I:/docs/sidebar.md',
      '# Root\n\ntext\n## Child A\nhello world\n## Child B\nsearch target',
    );

    const headings = extractHeadings(store.content);
    const matched = findMatchedLines(store.content, 'search');
    const activeHeadingLine = findActiveHeadingLine(headings, 7);

    expect(headings.map((item) => item.text)).toEqual([
      'Root',
      'Child A',
      'Child B',
    ]);
    expect(matched).toEqual([
      {
        lineNo: 7,
        line: 'search target',
      },
    ]);
    expect(activeHeadingLine).toBe(6);

    store.setCursorHint(activeHeadingLine);
    expect(store.cursorHint).toBe(6);
  });
});
