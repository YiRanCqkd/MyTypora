const BUILTIN_EXPORT_FORMATS = Object.freeze(['html', 'pdf', 'png', 'jpeg']);
const PANDOC_EXPORT_FORMATS = Object.freeze(['docx', 'odt', 'rtf']);

const { safeErrorMessage } = require('./errors.cjs');

function createExportEngines(options) {
  const { app, BrowserWindow, fs, log, path, runCommand } = options;
  let pandocProbe = null;

  async function probePandoc() {
    try {
      const versionResult = await runCommand('pandoc', ['--version']);
      const version = String(versionResult.stdout || '')
        .split(/\r?\n/)
        .map((item) => item.trim())
        .find(Boolean) || '';
      let formats = '';
      try {
        formats = (await runCommand('pandoc', ['--list-output-formats'])).stdout;
      } catch {
        formats = '';
      }
      return {
        available: true,
        version,
        outputFormats: String(formats)
          .split(/\s+/)
          .map((item) => item.toLowerCase())
          .filter(Boolean),
      };
    } catch {
      return { available: false, version: '', outputFormats: [] };
    }
  }

  async function getPandocProbe() {
    if (!pandocProbe) pandocProbe = probePandoc();
    return pandocProbe;
  }

  async function getCapabilities(_event, _payload, traceId) {
    const pandoc = await getPandocProbe();
    const result = {
      builtinFormats: [...BUILTIN_EXPORT_FORMATS],
      pandocAvailable: pandoc.available,
      pandocVersion: pandoc.version,
      pandocFormats: pandoc.available
        ? PANDOC_EXPORT_FORMATS.filter((format) => pandoc.outputFormats.includes(format))
        : [],
    };
    log('info', 'export capabilities', traceId, result);
    return result;
  }

  async function exportChromium(html, format, targetPath) {
    const win = new BrowserWindow({
      show: false,
      width: 1280,
      height: 960,
      webPreferences: { contextIsolation: true, sandbox: true, nodeIntegration: false },
    });
    try {
      await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
      await win.webContents.executeJavaScript(
        'Promise.resolve(document.fonts?.ready).catch(() => null).then(() => true);',
        true,
      );
      if (format === 'pdf') {
        await fs.writeFile(targetPath, await win.webContents.printToPDF({ printBackground: true }));
      } else {
        const image = await win.webContents.capturePage();
        await fs.writeFile(targetPath, format === 'jpeg' ? image.toJPEG(92) : image.toPNG());
      }
    } finally {
      if (!win.isDestroyed()) win.destroy();
    }
  }

  async function exportPandoc(markdown, format, targetPath, traceId) {
    const pandoc = await getPandocProbe();
    if (!pandoc.available) {
      throw new Error(
        '未检测到 Pandoc。请先安装 Pandoc（https://pandoc.org），或改用 HTML/PDF/Image 导出。',
      );
    }
    if (!pandoc.outputFormats.includes(format)) {
      throw new Error(
        `Pandoc 当前不支持 ${format} 导出。请检查 pandoc --list-output-formats 输出。`,
      );
    }
    const directory = await fs.mkdtemp(path.join(app.getPath('temp'), 'unireadmd-export-'));
    const source = path.join(directory, 'source.md');
    try {
      await fs.writeFile(source, markdown, 'utf8');
      await runCommand(
        'pandoc',
        [source, '--from', 'markdown', '--to', format, '--output', targetPath],
        30_000,
      );
      log('info', 'pandoc export success', traceId, { format, targetPath });
    } catch (error) {
      throw new Error(`Pandoc 导出失败：${safeErrorMessage(error)}`);
    } finally {
      await fs.rm(directory, { recursive: true, force: true });
    }
  }

  return {
    exportChromium,
    exportPandoc,
    getCapabilities,
  };
}

module.exports = {
  BUILTIN_EXPORT_FORMATS,
  PANDOC_EXPORT_FORMATS,
  createExportEngines,
};
