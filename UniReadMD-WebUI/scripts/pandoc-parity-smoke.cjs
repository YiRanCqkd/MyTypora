const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const { execFile } = require('node:child_process');
const { promisify } = require('node:util');
const { createExportEngines } = require('../electron/lib/export-engines.cjs');

// 使用官方便携程序验证实际转换；不修改系统 PATH 或安装外部依赖。
async function main() {
  const executable = path.resolve(process.argv[2]);
  const root = path.resolve(process.argv[3]);
  await fs.mkdir(root, { recursive: true });
  const execute = promisify(execFile);
  const versionResult = await execute(executable, ['--version'], {
    timeout: 30_000, windowsHide: true,
  });
  console.log('实际 Pandoc 版本', JSON.stringify(versionResult.stdout));
  const engines = createExportEngines({ fs, path,
    app: { getPath: () => root },
    runCommand: (_command, args, timeout = 30_000) => execute(executable, args, {
      timeout, windowsHide: true, maxBuffer: 64 * 1024 * 1024,
    }).catch((error) => { console.error('实际 Pandoc 调用失败', args, error); throw error; }),
  });
  const png = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aR9sAAAAASUVORK5CYII=',
    'base64',
  );
  const markdown = '# 实际转换测试\n\n中文正文 **加粗**。\n\n![图片](../images/a.png)';
  const resources = new Map([['../images/a.png', { content: png, contentType: 'image/png' }]]);
  const formats = ['docx', 'odt', 'rtf', 'epub', 'latex'];
  const results = [];
  console.log('实际转换能力', await engines.getCapabilities(null, {}, 'smoke'));
  for (const format of formats) {
    const target = path.join(root, `实际转换.${format === 'latex' ? 'tex' : format}`);
    await engines.exportPandoc(markdown, format, target, 'smoke', '', resources);
    assert.ok((await fs.stat(target)).size > 0);
    results.push({ operation: 'export', format, ok: true });
    const restored = path.join(root, `导入-${format}.md`);
    await engines.importPandoc(target, restored);
    assert.match(await fs.readFile(restored, 'utf8'), /实际转换测试/);
    results.push({ operation: 'import', format, ok: true });
  }
  for (const [format, source] of [
    ['html', '<h1>实际转换测试</h1><p>中文正文</p>'],
    ['rst', '实际转换测试\n============\n\n中文正文'],
  ]) {
    const input = path.join(root, `原始.${format}`);
    await fs.writeFile(input, source, 'utf8');
    const target = path.join(root, `导入-${format}.md`);
    await engines.importPandoc(input, target);
    assert.match(await fs.readFile(target, 'utf8'), /实际转换测试/);
    results.push({ operation: 'import', format, ok: true });
  }
  const capabilities = await engines.getCapabilities(null, {}, 'smoke');
  await fs.writeFile(path.join(root, '验证结果.json'), JSON.stringify({
    capabilities, results,
  }, null, 2), 'utf8');
  console.log(JSON.stringify({ capabilities, results }, null, 2));
}

main().catch((error) => { console.error(error); process.exitCode = 1; });
