const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const { execFile } = require('node:child_process');
const { promisify } = require('node:util');
const { JSDOM } = require('jsdom');
const { createExportEngines } = require('../electron/lib/export-engines.cjs');

async function main() {
  const executable = path.resolve(process.argv[2]);
  const root = path.resolve(process.argv[3]);
  const png = await fs.readFile(path.resolve(process.argv[4]));
  await fs.mkdir(root, { recursive: true });
  const execute = promisify(execFile);
  const engines = createExportEngines({ fs, path, app: { getPath: () => root },
    runCommand: (_command, args) => execute(executable, args, {
      windowsHide: true, timeout: 30_000, maxBuffer: 64 * 1024 * 1024,
    }),
  });
  const source = path.join(root, '源文档.md');
  const markdown = '# 回归标题\n\n中文正文。\n\n![中文图片](图.png)\n\n## 子标题\n\n正文';
  await fs.writeFile(source, markdown);
  await fs.writeFile(path.join(root, '图.png'), png);
  const resources = new Map([['图.png', { content: png, contentType: 'image/png' }]]);
  const results = [];
  for (const [format, extension] of [
    ['revealjs', 'html'], ['mediawiki', 'wiki'], ['rst', 'rst'],
    ['textile', 'textile'], ['opml', 'opml'],
  ]) {
    const target = path.join(root, `实际导出-${format}.${extension}`);
    await engines.exportPandoc(markdown, format, target, 'writing-smoke', source, resources,
      { title: '写作回归', tableOfContents: true, tocDepth: 3 });
    const output = await fs.readFile(target, 'utf8');
    assert.ok(output.includes('回归标题'), `${format} 缺少正文标题`);
    if (format === 'revealjs') {
      const html = new JSDOM(output).window.document;
      assert.ok(html.querySelector('.reveal .slides'), '未生成幻灯片结构');
      const image = html.querySelector('img[src], img[data-src]');
      assert.ok(image, '幻灯片缺少图片');
      assert.deepEqual(await fs.readFile(path.resolve(root,
        decodeURIComponent(image.getAttribute('src') || image.getAttribute('data-src')))), png);
    } else if (format === 'opml') {
      const xml = new JSDOM(output, { contentType: 'application/xml' }).window.document;
      assert.ok(xml.querySelector('opml body outline outline'), 'OPML 缺少嵌套大纲');
      const note = xml.querySelector('outline').getAttribute('_note');
      const reference = /!\[中文图片\]\(([^)]+)\)/.exec(note);
      assert.ok(reference, 'OPML 缺少图片引用');
      assert.deepEqual(await fs.readFile(path.resolve(root, decodeURI(reference[1]))), png);
    } else {
      const directories = (await fs.readdir(root)).filter((name) =>
        name.startsWith(`${path.parse(target).name}-assets-`));
      assert.equal(directories.length, 1, `${format} 缺少持久化资源目录`);
      assert.deepEqual(await fs.readFile(path.join(root, directories[0], 'image-0.png')), png);
      assert.ok(output.includes(directories[0]), `${format} 图片引用未指向持久化副本`);
    }
    results.push({ format, passed: true });
  }
  const epub = path.join(root, '章节配置.epub');
  await engines.exportPandoc(markdown, 'epub', epub, 'writing-smoke', source, resources,
    { tableOfContents: true, tocDepth: 4, epubChapterLevel: 2 });
  assert.ok((await fs.stat(epub)).size > 0);
  results.push({ format: 'epub', check: '目录与章节拆分参数', passed: true });
  assert.deepEqual(await fs.readFile(path.join(root, '图.png')), png);
  assert.ok(!(await fs.readdir(root)).some((name) => name.startsWith('unireadmd-export-')));
  await fs.writeFile(path.join(root, '结果.json'), JSON.stringify({ results }, null, 2));
  console.log(JSON.stringify({ results }, null, 2));
}

main().catch((error) => { console.error(error); process.exitCode = 1; });
