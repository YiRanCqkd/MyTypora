import { beforeEach, describe, expect, it, vi } from 'vitest';

import { installNativeBridgeFallback } from './native-bridge-fallback';

function requireBridge() {
  const bridge = window.nativeBridge;
  if (!bridge) {
    throw new Error('测试要求 fallback bridge 已安装');
  }
  return bridge;
}

describe('native bridge fallback', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    Object.defineProperty(window, 'nativeBridge', {
      configurable: true,
      value: undefined,
      writable: true,
    });
    Object.defineProperty(navigator, 'userAgent', {
      configurable: true,
      value: 'Mozilla/5.0',
    });
  });

  it(
    'Given web runtime, when saving and opening files, then the virtual store is consistent',
    async () => {
    installNativeBridgeFallback();
    const bridge = requireBridge();
    const first = await bridge.file.saveMarkdown({ content: '# first' });
    const second = await bridge.file.saveMarkdown({
      content: '# second',
      filePath: '/virtual/folder/second.markdown',
    });

    expect(first?.filePath).toMatch(/^\/virtual\/document-\d+\.md$/);
    expect(second?.filePath).toBe('/virtual/folder/second.markdown');
    expect(await bridge.file.openFromPath({
      filePath: second?.filePath || '',
    })).toEqual({
      content: '# second',
      filePath: '/virtual/folder/second.markdown',
    });
    expect(await bridge.file.openFromPath({ filePath: '/virtual/missing.md' })).toBeNull();
    expect(await bridge.file.openMarkdown()).toBeNull();
    expect(bridge.file.getPathForFile(new File([], 'sample.md'))).toBe('');
    },
  );

  it(
    'Given virtual files, when listing a directory, then only Markdown siblings are returned',
    async () => {
    installNativeBridgeFallback();
    const bridge = requireBridge();
    await bridge.file.saveMarkdown({ content: 'a', filePath: '/virtual/docs/b.md' });
    await bridge.file.saveMarkdown({ content: 'b', filePath: '/virtual/docs/a.markdown' });
    await bridge.file.saveMarkdown({ content: 'c', filePath: '/virtual/other/c.md' });
    await bridge.file.saveMarkdown({ content: 'x', filePath: '/virtual/docs/ignored.txt' });
    await bridge.file.saveMarkdown({ content: 'root', filePath: 'root.md' });

    await expect(bridge.file.listMarkdownInDirectory({
      filePath: '/virtual/docs/current.md',
    })).resolves.toEqual({
      directoryPath: '/virtual/docs',
      filePaths: [
        '/virtual/docs/a.markdown',
        '/virtual/docs/b.md',
      ],
    });
    await expect(bridge.file.listMarkdownInDirectory({ filePath: 'root.md' })).resolves.toEqual({
      directoryPath: '/',
      filePaths: ['root.md'],
    });
    },
  );

  it(
    'Given settings input, when reading and updating state, then normalization is applied',
    async () => {
    installNativeBridgeFallback();
    const bridge = requireBridge();
    const initial = await bridge.settings.getAll();
    expect(initial.settings.theme).toBe('light');

    const update = await bridge.settings.set('theme', 'dark');
    expect(update.value).toBe('dark');
    await expect(bridge.settings.get('theme')).resolves.toEqual({
      key: 'theme',
      value: 'dark',
    });

    const candidates = [
      ' /virtual/A.md ',
      '/virtual/a.MD',
      '/virtual/ignored.txt',
      '',
      ...Array.from({ length: 14 }, (_, index) => `/virtual/${index}.md`),
    ];
    const recent = await bridge.settings.setRecentFiles(candidates);
    expect(recent.filePaths).toHaveLength(12);
    expect(recent.filePaths[0]).toBe('/virtual/A.md');
    expect((await bridge.settings.getRecentFiles()).filePaths).toEqual(recent.filePaths);
    await expect(bridge.settings.selectUserCssFile()).resolves.toBeNull();
    },
  );

  it(
    'Given export payloads, when using builtin formats, then targets and capabilities are stable',
    async () => {
    installNativeBridgeFallback();
    const bridge = requireBridge();
    await expect(bridge.exporter.exportHtml({ html: '<h1>doc</h1>' })).resolves.toEqual({
      filePath: '/virtual/export.html',
    });
    await expect(bridge.exporter.exportDocument({
      format: 'pdf',
      html: '<h1>doc</h1>',
      markdown: '# doc',
      sourceFilePath: 'I:\\docs\\guide.md',
    })).resolves.toEqual({
      engine: 'builtin',
      filePath: '/virtual/guide.pdf',
      format: 'pdf',
    });
    await expect(bridge.exporter.exportDocument({
      format: 'png',
      html: '<h1>doc</h1>',
      markdown: '# doc',
    })).resolves.toMatchObject({
      filePath: '/virtual/document.png',
    });
    await expect(bridge.exporter.getCapabilities()).resolves.toMatchObject({
      builtinFormats: ['html', 'pdf', 'png', 'jpeg'],
      pandocAvailable: false,
    });
    },
  );

  it(
    'Given an unsupported export, when exporting, then a visible Pandoc error is returned',
    async () => {
    installNativeBridgeFallback();
    const bridge = requireBridge();

    await expect(bridge.exporter.exportDocument({
      format: 'docx',
      html: '<h1>doc</h1>',
      markdown: '# doc',
    })).rejects.toThrow('当前环境未检测到 Pandoc');
    },
  );

  it(
    'Given fallback services, when invoking app, window and context APIs, then contracts hold',
    async () => {
    installNativeBridgeFallback();
    const bridge = requireBridge();
    const info = vi.spyOn(console, 'log').mockImplementation(() => {});
    const context = bridge.context;
    const menu = bridge.menu;
    if (!context || !menu) {
      throw new Error('测试要求 context 和 menu fallback 可用');
    }

    await expect(bridge.app.getVersion()).resolves.toEqual({ version: 'web-fallback' });
    await expect(bridge.app.consumeLaunchFilePath()).resolves.toBeNull();
    await expect(bridge.app.logEvent({ level: 'info', message: 'ready' })).resolves.toEqual({
      ok: true,
    });
    expect(info).toHaveBeenCalledWith('[bridge:info]', 'ready', '');
    expect(typeof bridge.app.onLaunchFile(() => {})).toBe('function');
    await expect(bridge.window.minimize()).resolves.toEqual({ ok: true });
    await expect(bridge.window.toggleMaximize()).resolves.toEqual({ ok: true });
    await expect(bridge.window.isMaximized()).resolves.toEqual({ isMaximized: false });
    await expect(bridge.window.close()).resolves.toEqual({ ok: true });
    await expect(bridge.window.toggleDevTools()).resolves.toEqual({ ok: true });
    await expect(context.showFileTreeMenu({
      filePath: '/virtual/a.md',
      x: 1,
      y: 2,
    })).resolves.toEqual({ ok: true });
    expect(typeof menu.onCommand(() => {})).toBe('function');
    },
  );

  it(
    'Given spell settings, when language and enabled state change, then status follows',
    async () => {
    installNativeBridgeFallback();
    const bridge = requireBridge();
    await expect(bridge.spell.setEnabled(false)).resolves.toMatchObject({ enabled: false });
    await expect(bridge.spell.setLanguage('unknown')).resolves.toMatchObject({
      dictionaryStatus: 'missing',
      language: 'unknown',
    });
    await expect(bridge.spell.setLanguage('zh-CN')).resolves.toMatchObject({
      dictionaryStatus: 'ready',
      language: 'zh-CN',
    });
    await expect(bridge.spell.getState()).resolves.toMatchObject({ language: 'zh-CN' });
    },
  );

  it('Given legacy methods, when called, then they preserve fallback compatibility', async () => {
    installNativeBridgeFallback();
    const bridge = requireBridge();
    await expect(bridge.getAppVersion()).resolves.toBe('web-fallback');
    await expect(bridge.openMarkdownFile()).resolves.toBeNull();
    const saved = await bridge.saveMarkdownFile({ content: 'legacy' });
    expect(saved?.filePath).toMatch(/^\/virtual\/document-\d+\.md$/);
  });

  it('Given an existing bridge, when installing fallback, then it is not replaced', () => {
    const existing = { marker: 'existing' };
    Object.defineProperty(window, 'nativeBridge', {
      configurable: true,
      value: existing,
      writable: true,
    });

    installNativeBridgeFallback();

    expect(window.nativeBridge).toBe(existing);
  });

  it(
    'Given Electron without preload, when installing fallback, then the security boundary holds',
    () => {
    const error = vi.spyOn(console, 'error').mockImplementation(() => {});
    Object.defineProperty(navigator, 'userAgent', {
      configurable: true,
      value: 'Electron/37.2.0',
    });

    installNativeBridgeFallback();

    expect(window.nativeBridge).toBeUndefined();
    expect(error).toHaveBeenCalledWith(
      '[bridge:fallback] nativeBridge missing in Electron runtime',
    );
    },
  );
});
