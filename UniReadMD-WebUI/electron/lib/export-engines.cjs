const BUILTIN_EXPORT_FORMATS = Object.freeze(['html', 'pdf', 'png', 'jpeg']);
const PANDOC_EXPORT_FORMATS = Object.freeze([
  'docx', 'odt', 'rtf', 'epub', 'latex', 'revealjs', 'mediawiki', 'rst', 'textile', 'opml',
]);

const { createContractError, safeErrorMessage } = require('./errors.cjs');
const { atomicWrite } = require('./atomic-write.cjs');
const { createPrintOptions } = require('./pdf-options.cjs');
const { validateExportOptions, pandocOptionArguments } = require('./export-options.cjs');

function createExportEngines(options) {
  const { app, BrowserWindow, fs, log = () => {}, path, runCommand } = options;
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
        available: /^pandoc 3\.10(?:\s|$)/.test(version),
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

  async function exportChromium(html, format, targetPath, pdfOptions, exportOptions) {
    const settings = validateExportOptions(exportOptions);
    const win = new BrowserWindow({
      show: false,
      width: settings.imageWidth,
      height: 960,
      webPreferences: { contextIsolation: true, sandbox: true, nodeIntegration: false,
        backgroundThrottling: false,
      },
    });
    try {
      await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
      await win.webContents.executeJavaScript(
        `Promise.race([
          Promise.all([
            document.fonts.ready,
            ...Array.from(document.images, (image) => image.decode()),
          ]),
          new Promise((_, reject) => setTimeout(
            () => reject(new Error('导出资源加载超时')), 15000,
          )),
        ]).then(() => true);`,
        true,
      );
      if (format === 'pdf') {
        await atomicWrite(fs, targetPath,
          await win.webContents.printToPDF(createPrintOptions(pdfOptions)));
      } else {
        const size = await win.webContents.executeJavaScript(
          `({ width: ${settings.imageWidth},
            height: Math.ceil(document.documentElement.scrollHeight) })`,
        );
        if (size.height > 32767 || size.width * size.height > 64_000_000) {
          throw new Error('文档超过图片尺寸限制，请使用 PDF 导出');
        }
        win.setContentSize(size.width, Math.max(1, size.height));
        await win.webContents.executeJavaScript(
          'new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))',
        );
        const image = await win.webContents.capturePage({ x: 0, y: 0, ...size });
        await atomicWrite(fs, targetPath,
          format === 'jpeg' ? image.toJPEG(settings.jpegQuality) : image.toPNG());
      }
    } finally {
      if (!win.isDestroyed()) win.destroy();
    }
  }

  async function exportPandoc(markdown, format, targetPath, traceId, sourceFilePath, localImages,
    exportOptions) {
    const settings = validateExportOptions(exportOptions);
    const pandoc = await getPandocProbe();
    if (!pandoc.available) {
      throw createContractError(
        'ERR_PANDOC_UNAVAILABLE', '未检测到 Pandoc 3.10，请安装该版本或使用 HTML/PDF/图片导出。',
      );
    }
    if (!pandoc.outputFormats.includes(format)) {
      throw createContractError(
        'ERR_PANDOC_FORMAT_UNAVAILABLE', `Pandoc 当前不支持 ${format} 导出。`,
      );
    }
    const directory = await fs.mkdtemp(path.join(app.getPath('temp'), 'unireadmd-export-'));
    const source = path.join(directory, 'source.md');
    const output = path.join(directory, `result.${format}`);
    let persistentMedia = '';
    let completed = false;
    try {
      await fs.writeFile(source, markdown, 'utf8');
      let input = source;
      let inputFormat = 'markdown-raw_html-raw_tex';
      if (localImages?.size) {
        const astResult = await runCommand('pandoc', [source, '--from', inputFormat,
          '--to', 'json'], 30_000);
        const ast = JSON.parse(astResult.stdout);
        const resolved = new Map();
        if (['latex', 'revealjs', 'mediawiki', 'rst', 'textile', 'opml'].includes(format)) {
          // 这些格式引用外部图片，不能引用导出结束后会清理的临时目录。
          persistentMedia = await fs.mkdtemp(path.join(path.dirname(targetPath),
            `${path.parse(targetPath).name}-assets-`));
        }
        let index = 0;
        for (const [reference, resource] of localImages) {
          const extension = resource.contentType.split('/')[1]
            .replace('svg+xml', 'svg').replace('x-icon', 'ico');
          const imagePath = path.join(persistentMedia || directory,
            `image-${index++}.${extension}`);
          await fs.writeFile(imagePath, resource.content, { flag: 'wx' });
          const referencePath = persistentMedia
            ? path.relative(path.dirname(targetPath), imagePath) : imagePath;
          resolved.set(reference, referencePath.replace(/\\/g, '/'));
        }
        const visit = (node) => {
          if (!node || typeof node !== 'object') return;
          if (node.t === 'Image') {
            const target = node.c?.[2];
            const uri = target?.[0];
            if (typeof uri !== 'string') throw new Error('图片转换结构无效');
            if (!/^(https?:|data:image\/)/i.test(uri)) {
              const replacement = resolved.get(decodeURIComponent(uri.split(/[?#]/)[0]));
              if (!replacement) throw new Error('图片未通过文档授权检查');
              target[0] = replacement;
            }
          }
          for (const value of Object.values(node)) {
            if (Array.isArray(value)) value.forEach(visit);
            else if (value && typeof value === 'object') visit(value);
          }
        };
        visit(ast);
        input = path.join(directory, 'source.json');
        inputFormat = 'json';
        await fs.writeFile(input, JSON.stringify(ast), 'utf8');
      }
      await runCommand(
        'pandoc',
        [input, '--from', inputFormat, '--to', format, '--output', output,
          ...pandocOptionArguments(settings, format),
          ...(sourceFilePath ? ['--resource-path', path.dirname(sourceFilePath)] : [])],
        30_000,
      );
      await atomicWrite(fs, targetPath, await fs.readFile(output));
      completed = true;
      log('info', 'pandoc export success', traceId, { format, targetPath });
    } catch (error) {
      throw new Error(`Pandoc 导出失败：${safeErrorMessage(error)}`);
    } finally {
      await fs.rm(directory, { recursive: true, force: true });
      if (persistentMedia && !completed) {
        await fs.rm(persistentMedia, { recursive: true, force: true });
      }
    }
  }

  async function importPandoc(sourceFilePath, targetPath) {
    const pandoc = await getPandocProbe();
    if (!pandoc.available) {
      throw createContractError('ERR_PANDOC_UNAVAILABLE', '导入需要安装 Pandoc 3.10');
    }
    const directory = await fs.mkdtemp(path.join(path.dirname(targetPath), 'import-assets-'));
    try {
      const result = await runCommand('pandoc', [sourceFilePath, '--to', 'gfm',
        '--wrap', 'none', '--extract-media', directory], 30_000);
      let markdown = String(result.stdout || '');
      if (Buffer.byteLength(markdown) > 20 * 1024 * 1024) throw new Error('导入文档超过 20 MiB');
      markdown = markdown.split(directory.replace(/\\/g, '/'))
        .join(path.basename(directory)).split(directory).join(path.basename(directory));
      await atomicWrite(fs, targetPath, markdown);
    } catch (error) {
      await fs.rm(directory, { recursive: true, force: true });
      throw error;
    }
  }

  return {
    exportChromium,
    exportPandoc,
    importPandoc,
    getCapabilities,
  };
}

module.exports = {
  BUILTIN_EXPORT_FORMATS,
  PANDOC_EXPORT_FORMATS,
  createExportEngines,
};
