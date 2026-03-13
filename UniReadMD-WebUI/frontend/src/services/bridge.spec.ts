import { afterEach, describe, expect, it, vi } from 'vitest';
import { bridgeService } from './bridge';
import type { SettingsState } from '../types/native';

function mockBridge() {
  const bridge = {
    app: {
      getVersion: vi.fn(),
      logEvent: vi.fn(),
      consumeLaunchFilePath: vi.fn(),
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
});
