const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const { execFile } = require('node:child_process');
const { promisify } = require('node:util');
const { JSDOM } = require('jsdom');
const { createExportEngines } = require('../electron/lib/export-engines.cjs');

// 验证实际转换产物，不能仅以子进程成功退出判断参考样式和封面生效。
async function main() {
  const executable = path.resolve(process.argv[2]);
  const root = path.resolve(process.argv[3]);
  const archiveTool = path.resolve(process.argv[4]);
  await fs.mkdir(root, { recursive: true });
  const execute = promisify(execFile);
  const run = (command, args, extra = {}) => execute(command, args, {
    timeout: 30_000, windowsHide: true, maxBuffer: 64 * 1024 * 1024, ...extra,
  });
  const readEntry = async (archive, entry) => (await run(archiveTool,
    ['x', '-so', archive, entry], { encoding: 'buffer' })).stdout;
  const engines = createExportEngines({ fs, path,
    app: { getPath: () => root },
    runCommand: (_command, args) => run(executable, args),
  });
  const results = [];
  const reference = path.join(root, '中文参考样式.docx');
  const defaultReference = await run(executable,
    ['--print-default-data-file', 'reference.docx'], { encoding: 'buffer' });
  await fs.writeFile(reference, defaultReference.stdout);
  const stylesDirectory = path.join(root, '参考样式内容');
  await run(archiveTool, ['x', '-y', `-o${stylesDirectory}`, reference]);
  const stylesFile = path.join(stylesDirectory, 'word', 'styles.xml');
  const originalStyles = await fs.readFile(stylesFile, 'utf8');
  const marker = '<w:color w:val="BE12EF"/>';
  const styles = originalStyles.replace(
    /(<w:style\b[^>]*w:styleId="Normal"[^>]*>)([\s\S]*?)(<\/w:style>)/,
    (_match, start, body, end) => {
      const updated = body.includes('<w:rPr>')
        ? body.replace('<w:rPr>', `<w:rPr>${marker}`) : `${body}<w:rPr>${marker}</w:rPr>`;
      return start + updated + end;
    },
  );
  assert.ok(styles.includes(marker), '默认 Word 参考文件缺少 Normal 样式');
  await fs.writeFile(stylesFile, styles, 'utf8');
  await execute(archiveTool, ['u', reference, 'word/styles.xml'], {
    cwd: stylesDirectory, windowsHide: true, timeout: 30_000,
  });
  const markdown = '# 高级导出回归\n\n中文正文。';
  const options = { title: '自定义中文标题', author: '回归作者', language: 'zh-CN' };
  const docx = path.join(root, '参考样式导出.docx');
  await engines.exportPandoc(markdown, 'docx', docx, 'options-smoke', '', undefined,
    { ...options, referenceDoc: reference });
  assert.ok((await readEntry(docx, 'word/styles.xml')).toString().includes('BE12EF'));
  const core = (await readEntry(docx, 'docProps/core.xml')).toString();
  assert.ok(core.includes(options.title) && core.includes(options.author));
  results.push({ check: 'Word 参考样式与中文标题作者写入实际产物', passed: true });

  // 使用前轮真实 Chromium 导出的 PNG，避免损坏图片使封面测试失真。
  const cover = path.join(root, '中文封面.png');
  const png = await fs.readFile(path.resolve(process.argv[5]));
  assert.equal(png.subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
  await fs.writeFile(cover, png);
  const epub = path.join(root, '含封面导出.epub');
  await engines.exportPandoc(markdown, 'epub', epub, 'options-smoke', '', undefined,
    { ...options, epubCover: cover });
  const packageXml = (await readEntry(epub, 'EPUB/content.opf')).toString();
  assert.ok(packageXml.includes('cover-image') && packageXml.includes(options.title));
  assert.ok(packageXml.includes(options.author) && packageXml.includes('zh-CN'));
  const packageDocument = new JSDOM(packageXml, { contentType: 'application/xml' }).window.document;
  const coverEntry = packageDocument.querySelector('item[properties~="cover-image"]')
    ?.getAttribute('href');
  assert.ok(coverEntry, 'EPUB 未声明封面图片');
  assert.deepEqual(await readEntry(epub, `EPUB/${coverEntry}`), png);
  results.push({ check: 'EPUB 封面字节与中文元数据写入实际产物', passed: true });

  const preserved = path.join(root, '失败时保留.docx');
  await fs.writeFile(preserved, '原文件必须保留', 'utf8');
  await assert.rejects(engines.exportPandoc(markdown, 'docx', preserved, 'options-smoke', '',
    undefined, { referenceDoc: path.join(root, '不存在.docx') }), /Pandoc 导出失败/);
  assert.equal(await fs.readFile(preserved, 'utf8'), '原文件必须保留');
  assert.ok(!(await fs.readdir(root)).some((name) => name.startsWith('unireadmd-export-')));
  results.push({ check: '转换失败保留原输出文件并清理临时目录', passed: true });
  await fs.writeFile(path.join(root, '验证结果.json'),
    JSON.stringify({ results }, null, 2), 'utf8');
  console.log(JSON.stringify({ results }, null, 2));
}

main().catch((error) => { console.error(error); process.exitCode = 1; });
