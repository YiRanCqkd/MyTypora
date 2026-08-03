import { afterEach, describe, expect, it, vi } from 'vitest';
import { bridgeService } from './bridge';
import type { SettingsState } from '../types/native';

function mockBridge() {
  const bridge = {
    app: {
      getVersion: vi.fn(),
      logEvent: vi.fn(),
      consumeLaunchFilePath: vi.fn(),
      onLaunchFile: vi.fn(),
    },
    file: {
      openMarkdown: vi.fn(),
      openFromPath: vi.fn(),
      saveMarkdown: vi.fn(),
      getPathForFile: vi.fn(),
      listMarkdownInDirectory: vi.fn(),
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
      getRecentFiles: vi.fn(),
      get: vi.fn(),
      set: vi.fn(),
      setRecentFiles: vi.fn(),
      selectUserCssFile: vi.fn(),
    },
    context: {
      showFileTreeMenu: vi.fn(),
      showSourceEditorMenu: vi.fn(),
    },
    spell: {
      getState: vi.fn(),
      setEnabled: vi.fn(),
      setLanguage: vi.fn(),
    },
    menu: {
      onCommand: vi.fn(),
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

describe('bridgeService', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('getVersion 应透传 traceId 并校验返回结构', async () => {
    const bridge = mockBridge();
    bridge.app.getVersion.mockResolvedValue({
      version: '9.9.9',
    });

    const result = await bridgeService.getVersion('trace-1');

    expect(bridge.app.getVersion).toHaveBeenCalledWith('trace-1');
    expect(result.version).toBe('9.9.9');
  });

  it('openMarkdown 在用户取消时应返回 null', async () => {
    const bridge = mockBridge();
    bridge.file.openMarkdown.mockResolvedValue(null);

    const result = await bridgeService.openMarkdown('trace-open');

    expect(bridge.file.openMarkdown).toHaveBeenCalledWith({
      traceId: 'trace-open',
    });
    expect(result).toBeNull();
  });

  it('consumeLaunchFilePath 应返回并校验启动文件路径', async () => {
    const bridge = mockBridge();
    bridge.app.consumeLaunchFilePath.mockResolvedValue({
      filePath: 'I:/docs/startup.md',
    });

    const result = await bridgeService.consumeLaunchFilePath('trace-launch');

    expect(bridge.app.consumeLaunchFilePath).toHaveBeenCalledWith('trace-launch');
    expect(result?.filePath).toBe('I:/docs/startup.md');
  });

  it('openFromPath 应透传路径与 traceId', async () => {
    const bridge = mockBridge();
    bridge.file.openFromPath.mockResolvedValue({
      filePath: 'I:/doc.md',
      content: '# title',
    });

    const result = await bridgeService.openFromPath('I:/doc.md', 'trace-path');

    expect(bridge.file.openFromPath).toHaveBeenCalledWith({
      filePath: 'I:/doc.md',
      traceId: 'trace-path',
    });
    expect(result?.content).toBe('# title');
  });

  it('saveMarkdown 返回非法结构时应抛错', async () => {
    const bridge = mockBridge();
    bridge.file.saveMarkdown.mockResolvedValue({
      filePath: 123,
    });

    await expect(
      bridgeService.saveMarkdown({
        content: '# md',
      }),
    ).rejects.toThrow('file:saveMarkdown 返回非法字段: filePath');
  });

  it('settings getAll 应校验 settings 对象', async () => {
    const bridge = mockBridge();
    const settings: SettingsState = {
      theme: 'light',
      fontSize: 14,
      contentZoom: 100,
      viewMode: 'split',
      autoSave: false,
      userCssEnabled: false,
      userCssPath: '',
      spellCheckEnabled: true,
      spellCheckLanguage: 'en-US',
    };

    bridge.settings.getAll.mockResolvedValue({
      settings,
    });

    const result = await bridgeService.getAllSettings('trace-settings');

    expect(bridge.settings.getAll).toHaveBeenCalledWith('trace-settings');
    expect(result.settings.theme).toBe('light');
  });

  it('setSetting 应透传 key/value/traceId', async () => {
    const bridge = mockBridge();
    const settings: SettingsState = {
      theme: 'dark',
      fontSize: 14,
      contentZoom: 100,
      viewMode: 'split',
      autoSave: false,
      userCssEnabled: false,
      userCssPath: '',
      spellCheckEnabled: true,
      spellCheckLanguage: 'en-US',
    };

    bridge.settings.set.mockResolvedValue({
      key: 'theme',
      value: 'dark',
      settings,
    });

    const result = await bridgeService.setSetting('theme', 'dark', 'trace-set');

    expect(bridge.settings.set).toHaveBeenCalledWith('theme', 'dark', 'trace-set');
    expect(result.value).toBe('dark');
  });

  it('recent files get/set 应透传并校验返回结构', async () => {
    const bridge = mockBridge();
    bridge.settings.getRecentFiles.mockResolvedValue({
      filePaths: ['I:/docs/a.md', 'I:/docs/b.md'],
    });
    bridge.settings.setRecentFiles.mockResolvedValue({
      filePaths: ['I:/docs/c.md', 'I:/docs/a.md'],
    });

    const current = await bridgeService.getRecentFiles('trace-recent-get');
    expect(bridge.settings.getRecentFiles).toHaveBeenCalledWith('trace-recent-get');
    expect(current.filePaths).toEqual(['I:/docs/a.md', 'I:/docs/b.md']);

    const updated = await bridgeService.setRecentFiles(
      ['I:/docs/c.md', 'I:/docs/a.md'],
      'trace-recent-set',
    );
    expect(bridge.settings.setRecentFiles).toHaveBeenCalledWith(
      ['I:/docs/c.md', 'I:/docs/a.md'],
      'trace-recent-set',
    );
    expect(updated.filePaths).toEqual(['I:/docs/c.md', 'I:/docs/a.md']);
  });

  it('window action 应透传 traceId', async () => {
    const bridge = mockBridge();
    bridge.window.minimize.mockResolvedValue({ ok: true });

    const result = await bridgeService.minimizeWindow('trace-min');

    expect(bridge.window.minimize).toHaveBeenCalledWith('trace-min');
    expect(result.ok).toBe(true);
  });

  it('导出编排能力应支持 capability 查询与导出调用', async () => {
    const bridge = mockBridge();
    bridge.exporter.getCapabilities.mockResolvedValue({
      builtinFormats: ['html', 'pdf', 'png', 'jpeg'],
      pandocAvailable: true,
      pandocVersion: 'pandoc 3.1',
      pandocFormats: ['docx'],
    });
    bridge.exporter.exportDocument.mockResolvedValue({
      filePath: 'I:/export/output.docx',
      format: 'docx',
      engine: 'pandoc',
    });

    const capabilityResult = await bridgeService.getExportCapabilities('trace-cap');
    expect(bridge.exporter.getCapabilities).toHaveBeenCalledWith('trace-cap');
    expect(capabilityResult.pandocAvailable).toBe(true);
    expect(capabilityResult.pandocFormats).toContain('docx');

    const exportResult = await bridgeService.exportDocument({
      format: 'docx',
      html: '<h1>doc</h1>',
      markdown: '# doc',
      traceId: 'trace-export',
    });
    expect(bridge.exporter.exportDocument).toHaveBeenCalledWith({
      format: 'docx',
      html: '<h1>doc</h1>',
      markdown: '# doc',
      traceId: 'trace-export',
    });
    expect(exportResult?.engine).toBe('pandoc');
  });

  it('showFileTreeContextMenu 应透传文件路径与坐标', async () => {
    const bridge = mockBridge();
    bridge.context.showFileTreeMenu.mockResolvedValue({
      ok: true,
    });

    const result = await bridgeService.showFileTreeContextMenu({
      filePath: 'I:/docs/sample.md',
      x: 120,
      y: 88,
      traceId: 'trace-context-menu',
    });

    expect(bridge.context.showFileTreeMenu).toHaveBeenCalledWith({
      filePath: 'I:/docs/sample.md',
      x: 120,
      y: 88,
      traceId: 'trace-context-menu',
    });
    expect(result.ok).toBe(true);
  });

  it('showSourceEditorContextMenu 在原生通道存在时应返回 handled=true', async () => {
    const bridge = mockBridge();
    bridge.context.showSourceEditorMenu.mockResolvedValue({
      ok: true,
    });

    const result = await bridgeService.showSourceEditorContextMenu({
      x: 300,
      y: 190,
      traceId: 'trace-source-context',
    });

    expect(bridge.context.showSourceEditorMenu).toHaveBeenCalledWith({
      x: 300,
      y: 190,
      traceId: 'trace-source-context',
    });
    expect(result.ok).toBe(true);
    expect(result.handled).toBe(true);
  });

  it('showSourceEditorContextMenu 在原生通道缺失时应回退 handled=false', async () => {
    const bridge = mockBridge();
    const contextApi = bridge.context as unknown as Record<string, unknown>;
    delete contextApi.showSourceEditorMenu;

    const result = await bridgeService.showSourceEditorContextMenu({
      x: 12,
      y: 34,
      traceId: 'trace-source-fallback',
    });

    expect(result.ok).toBe(true);
    expect(result.handled).toBe(false);
  });

  it('spell get/set 应校验返回结构', async () => {
    const bridge = mockBridge();
    bridge.spell.getState.mockResolvedValue({
      enabled: true,
      language: 'en-US',
      availableLanguages: ['en-US', 'zh-CN'],
      dictionaryStatus: 'ready',
    });
    bridge.spell.setLanguage.mockResolvedValue({
      enabled: true,
      language: 'zh-CN',
      availableLanguages: ['en-US', 'zh-CN'],
      dictionaryStatus: 'ready',
    });

    const current = await bridgeService.getSpellCheckState('trace-spell');
    expect(current.language).toBe('en-US');
    expect(bridge.spell.getState).toHaveBeenCalledWith('trace-spell');

    const updated = await bridgeService.setSpellCheckLanguage('zh-CN', 'trace-spell-2');
    expect(updated.language).toBe('zh-CN');
    expect(bridge.spell.setLanguage).toHaveBeenCalledWith('zh-CN', 'trace-spell-2');
  });

  it('缺少 nativeBridge 时应返回清晰错误', async () => {
    Object.defineProperty(window, 'nativeBridge', {
      configurable: true,
      value: undefined,
    });

    await expect(bridgeService.getVersion()).rejects.toThrow('nativeBridge 不可用');
  });

  it('launch listener 应校验事件并保留取消订阅函数', () => {
    const bridge = mockBridge();
    const unsubscribe = vi.fn();
    bridge.app.onLaunchFile.mockImplementation((listener) => {
      listener({ filePath: 'I:/docs/launch.md' });
      listener(null);
      return unsubscribe;
    });
    const listener = vi.fn();

    const result = bridgeService.onLaunchFile(listener);

    expect(listener).toHaveBeenNthCalledWith(1, { filePath: 'I:/docs/launch.md' });
    expect(listener).toHaveBeenNthCalledWith(2, null);
    expect(result).toBe(unsubscribe);
  });

  it('file adapters 应覆盖有效、取消、目录及路径回退合同', async () => {
    const bridge = mockBridge();
    bridge.file.openMarkdown.mockResolvedValue({
      content: '# doc',
      filePath: 'I:/docs/a.md',
    });
    bridge.file.openFromPath.mockResolvedValue(null);
    bridge.file.saveMarkdown.mockResolvedValue(null);
    bridge.file.getPathForFile.mockReturnValue('I:/docs/file.md');
    bridge.file.listMarkdownInDirectory.mockResolvedValue({
      directoryPath: 'I:/docs',
      filePaths: ['I:/docs/a.md', 7, 'I:/docs/b.md'],
    });

    await expect(bridgeService.openMarkdown()).resolves.toEqual({
      content: '# doc',
      filePath: 'I:/docs/a.md',
    });
    await expect(bridgeService.openFromPath('I:/none.md')).resolves.toBeNull();
    await expect(bridgeService.saveMarkdown({ content: 'draft' })).resolves.toBeNull();
    expect(bridgeService.getPathForFile(new File([], 'file.md'))).toBe('I:/docs/file.md');
    await expect(bridgeService.listMarkdownInDirectory('I:/docs/a.md')).resolves.toEqual({
      directoryPath: 'I:/docs',
      filePaths: ['I:/docs/a.md', 'I:/docs/b.md'],
    });

    bridge.file.getPathForFile.mockImplementation(() => {
      throw new Error('path failed');
    });
    expect(bridgeService.getPathForFile(new File([], 'file.md'))).toBe('');
    Object.defineProperty(bridge.file, 'getPathForFile', {
      configurable: true,
      value: undefined,
    });
    expect(bridgeService.getPathForFile(new File([], 'file.md'))).toBe('');
  });

  it('export adapters 应校验空结果、格式列表和 engine', async () => {
    const bridge = mockBridge();
    bridge.exporter.exportHtml.mockResolvedValue({ filePath: 'I:/out/doc.html' });
    bridge.exporter.exportDocument.mockResolvedValue(null);
    bridge.exporter.getCapabilities.mockResolvedValue({
      builtinFormats: ['html', 1, 'pdf'],
      pandocAvailable: 1,
      pandocVersion: 3,
      pandocFormats: null,
    });

    await expect(bridgeService.exportHtml({ html: '<p>doc</p>' })).resolves.toEqual({
      filePath: 'I:/out/doc.html',
    });
    await expect(bridgeService.exportDocument({
      format: 'pdf',
      html: '<p>doc</p>',
      markdown: 'doc',
    })).resolves.toBeNull();
    await expect(bridgeService.getExportCapabilities()).resolves.toEqual({
      builtinFormats: ['html', 'pdf'],
      pandocAvailable: true,
      pandocVersion: '',
      pandocFormats: [],
    });

    bridge.exporter.exportDocument.mockResolvedValue({
      engine: 'invalid',
      filePath: 'I:/out/doc.pdf',
      format: 'pdf',
    });
    await expect(bridgeService.exportDocument({
      format: 'pdf',
      html: '<p>doc</p>',
      markdown: 'doc',
    })).rejects.toThrow('export:document 返回非法字段: engine');
  });

  it('settings adapters 应覆盖非法结构、过滤及 CSS 选择', async () => {
    const bridge = mockBridge();
    bridge.settings.getAll.mockResolvedValue(null);
    await expect(bridgeService.getAllSettings()).rejects.toThrow(
      'settings:getAll 返回非法数据',
    );

    bridge.settings.getRecentFiles.mockResolvedValue({ filePaths: null });
    await expect(bridgeService.getRecentFiles()).resolves.toEqual({ filePaths: [] });
    bridge.settings.setRecentFiles.mockResolvedValue({ filePaths: ['a.md', 2] });
    await expect(bridgeService.setRecentFiles(['a.md'])).resolves.toEqual({
      filePaths: ['a.md'],
    });
    bridge.settings.selectUserCssFile.mockResolvedValue({ filePath: 'I:/theme.css' });
    await expect(bridgeService.selectUserCssFile()).resolves.toEqual({
      filePath: 'I:/theme.css',
    });
    bridge.settings.selectUserCssFile.mockResolvedValue(null);
    await expect(bridgeService.selectUserCssFile()).resolves.toBeNull();
  });

  it('window、context 与 spell adapters 应覆盖剩余合同和回退路径', async () => {
    const bridge = mockBridge();
    bridge.window.toggleMaximize.mockResolvedValue({ ok: true });
    bridge.window.close.mockResolvedValue({ ok: true });
    bridge.context.showFileTreeMenu.mockResolvedValue({ ok: true });
    bridge.spell.setEnabled.mockResolvedValue({
      availableLanguages: ['en-US', 1],
      dictionaryStatus: 'missing',
      enabled: 1,
      language: 'en-US',
    });

    await expect(bridgeService.toggleMaximizeWindow()).resolves.toEqual({ ok: true });
    await expect(bridgeService.closeWindow()).resolves.toEqual({ ok: true });
    await expect(bridgeService.showFileTreeContextMenu({
      filePath: 'I:/docs/a.md',
      x: 1,
      y: 2,
    })).resolves.toEqual({ ok: true });
    await expect(bridgeService.setSpellCheckEnabled(true)).resolves.toEqual({
      availableLanguages: ['en-US'],
      dictionaryStatus: 'missing',
      enabled: true,
      language: 'en-US',
    });

    Object.defineProperty(bridge, 'context', {
      configurable: true,
      value: undefined,
    });
    await expect(bridgeService.showFileTreeContextMenu({
      filePath: 'I:/docs/a.md',
      x: 1,
      y: 2,
    })).resolves.toEqual({ ok: true });
  });
});
