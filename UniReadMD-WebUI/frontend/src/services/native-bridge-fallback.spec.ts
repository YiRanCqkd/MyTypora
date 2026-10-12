import { afterEach, describe, expect, it, vi } from 'vitest';

afterEach(() => {
  Reflect.deleteProperty(window, 'nativeBridge');
  localStorage.removeItem('unireadmd.web-theme.v1');
  localStorage.removeItem('unireadmd.view-visibility.v1');
  vi.restoreAllMocks();
  vi.resetModules();
});

async function install() {
  Reflect.deleteProperty(window, 'nativeBridge');
  const module = await import('./native-bridge-fallback');
  module.installNativeBridgeFallback();
  return window.nativeBridge;
}

describe('浏览器主题持久化', () => {
  it('视图区域默认隐藏，开关独立且重载后保留', async () => {
    const first = await install();
    expect((await first.settings.getAll()).settings).toMatchObject({
      documentToolsVisible: false, preferencesVisible: false,
    });
    await first.settings.set('documentToolsVisible', true);
    vi.resetModules();
    const second = await install();
    expect((await second.settings.getAll()).settings).toMatchObject({
      documentToolsVisible: true, preferencesVisible: false,
    });
    localStorage.setItem('unireadmd.view-visibility.v1', '损坏');
    vi.resetModules();
    expect((await (await install()).settings.getAll()).settings.documentToolsVisible).toBe(false);
  });
  it('刷新后仍读取系统主题选项', async () => {
    Reflect.deleteProperty(window, 'nativeBridge');
    const first = await import('./native-bridge-fallback');
    first.installNativeBridgeFallback();
    await window.nativeBridge.settings.set('theme', 'system');
    Reflect.deleteProperty(window, 'nativeBridge');
    vi.resetModules();
    const second = await import('./native-bridge-fallback');
    second.installNativeBridgeFallback();
    expect((await window.nativeBridge.settings.getAll()).settings.theme).toBe('system');
  });

  it('损坏主题回退到浅色', async () => {
    localStorage.setItem('unireadmd.web-theme.v1', 'invalid');
    Reflect.deleteProperty(window, 'nativeBridge');
    const bridge = await import('./native-bridge-fallback');
    bridge.installNativeBridgeFallback();
    expect((await window.nativeBridge.settings.getAll()).settings.theme).toBe('light');
  });

  it('虚拟文档可保存、重新打开，同目录列表过滤非 Markdown 和其他目录', async () => {
    const bridge = await install();
    const first = await bridge.file.saveMarkdown({ content: '# 新文档' });
    expect(first?.filePath).toMatch(/^\/virtual\/document-\d+\.md$/);
    await bridge.file.saveMarkdown({ filePath: '/virtual/空白.markdown', content: '' });
    await bridge.file.saveMarkdown({ filePath: '/other/其他.md', content: '其他目录' });
    await bridge.file.saveMarkdown({ filePath: '/virtual/图片.txt', content: '非 Markdown' });
    expect(await bridge.file.openFromPath({ filePath: first!.filePath }))
      .toMatchObject({ content: '# 新文档' });
    expect(await bridge.file.openFromPath({ filePath: '/virtual/空白.markdown' }))
      .toMatchObject({ content: '' });
    expect(await bridge.file.openFromPath({ filePath: '/missing.md' })).toBeNull();
    const listed = await bridge.file.listMarkdownInDirectory({ filePath: first!.filePath });
    expect(listed.directoryPath).toBe('/virtual');
    expect(listed.filePaths).toEqual(expect.arrayContaining([
      first!.filePath, '/virtual/空白.markdown',
    ]));
    expect(listed.filePaths).not.toContain('/other/其他.md');
    expect(listed.filePaths).not.toContain('/virtual/图片.txt');
    expect((await bridge.file.listMarkdownInDirectory({ filePath: 'root.md' })).directoryPath)
      .toBe('/');
  });

  it('浏览器导出保持虚拟内容，缺少 Pandoc 时拒绝对应格式', async () => {
    const bridge = await install();
    expect((await bridge.exporter.getCapabilities()).pandocAvailable).toBe(false);
    const html = await bridge.exporter.exportHtml({ html: '<p>正文</p>' });
    expect(await bridge.file.openFromPath({ filePath: html!.filePath }))
      .toMatchObject({ content: '<p>正文</p>' });
    await bridge.exporter.exportHtml({ html: '指定输出', targetPath: '/virtual/指定.html' });
    expect(await bridge.exporter.exportDocument({ format: 'html', html: '内容', markdown: '',
      sourceFilePath: 'C:\\docs\\中文.md' }))
      .toMatchObject({ filePath: '/virtual/中文.html', engine: 'builtin' });
    const plain = await bridge.exporter.exportDocument({ format: 'html', html: '',
      markdown: '# 正文', targetPath: '/virtual/plain.html' });
    expect(await bridge.file.openFromPath({ filePath: plain!.filePath }))
      .toMatchObject({ content: '# 正文' });
    const unnamed = await bridge.exporter.exportDocument({
      format: 'html', html: '', markdown: '',
    });
    expect(unnamed?.filePath)
      .toBe('/virtual/document.html');
    await expect(bridge.exporter.exportDocument({ format: 'docx', html: '正文', markdown: '' }))
      .rejects.toThrow('未检测到 Pandoc');
  });

  it('最近文档去重并限制数量，拼写语言与设置读取保持一致', async () => {
    const bridge = await install();
    const files = ['C:\\docs\\a.md', 'c:/docs/A.md', '', 'bad.txt', 123 as never,
      ...Array.from({ length: 20 }, (_, index) => `/docs/${index}.markdown`)];
    await bridge.settings.setRecentFiles(files);
    const recent = (await bridge.settings.getRecentFiles()).filePaths;
    expect(recent).toHaveLength(12);
    expect(recent.filter((file) => /a\.md$/i.test(file))).toHaveLength(1);
    expect(recent).not.toContain('bad.txt');
    await bridge.settings.set('fontSize', 18);
    expect((await bridge.settings.get('fontSize')).value).toBe(18);
    expect(await bridge.settings.selectUserCssFile()).toBeNull();
    expect((await bridge.spell.getState()).dictionaryStatus).toBe('ready');
    expect((await bridge.spell.setEnabled(false)).enabled).toBe(false);
    expect((await bridge.spell.setLanguage('unknown')).dictionaryStatus).toBe('missing');
    expect((await bridge.spell.setLanguage('')).language).toBe('en-US');
  });

  it('浏览器宿主兼容接口可调用，已有原生桥接始终保留', async () => {
    const bridge = await install();
    vi.spyOn(console, 'log').mockImplementation(() => {});
    expect((await bridge.app.getVersion()).version).toBe('web-fallback');
    expect(await bridge.app.consumeLaunchFilePath()).toBeNull();
    bridge.app.onLaunchFile(() => {})();
    bridge.menu!.onCommand(() => {})();
    await bridge.app.logEvent({ level: 'info', message: '信息' });
    await bridge.app.logEvent({ level: 'error', message: '错误', traceId: 'trace' });
    expect(await bridge.file.openMarkdown()).toBeNull();
    expect(bridge.file.getPathForFile(new File([], '文档.md'))).toBe('');
    expect((await bridge.getAppVersion())).toBe('web-fallback');
    expect(await bridge.openMarkdownFile()).toBeNull();
    expect(await bridge.saveMarkdownFile({ content: '兼容保存' })).toHaveProperty('filePath');
    expect(await bridge.window.isMaximized()).toEqual({ isMaximized: false });
    for (const action of ['minimize', 'toggleMaximize', 'close', 'toggleDevTools'] as const) {
      expect(await bridge.window[action]()).toEqual({ ok: true });
    }
    expect(await bridge.context!.showFileTreeMenu({ filePath: '/virtual/a.md' }))
      .toEqual({ ok: true });
    (await import('./native-bridge-fallback')).installNativeBridgeFallback();
    expect(window.nativeBridge).toBe(bridge);
  });

  it('禁用存储仍可切换当前主题，Electron 缺少 preload 时不伪造原生能力', async () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new Error('拒绝'); });
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('拒绝'); });
    const bridge = await install();
    expect((await bridge.settings.set('theme', 'dark')).value).toBe('dark');
    Reflect.deleteProperty(window, 'nativeBridge');
    vi.spyOn(navigator, 'userAgent', 'get').mockReturnValue('Electron');
    const error = vi.spyOn(console, 'error').mockImplementation(() => {});
    (await import('./native-bridge-fallback')).installNativeBridgeFallback();
    expect(window.nativeBridge).toBeUndefined();
    expect(error).toHaveBeenCalledWith(expect.stringContaining('nativeBridge missing'));
  });
});
