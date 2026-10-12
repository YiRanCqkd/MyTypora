import { expect, test, type Page } from './fixtures';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';

async function openContent(page: Page, content: string, sourceMode = false) {
  await page.addInitScript(({ text, source }) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      filePath: '', isDirty: false, content: text, cursorHint: 1, sourceMode: source,
    }));
  }, { text: content, source: sourceMode });
  await page.goto('/');
  await expect(page.getByTestId('preview-host')).toBeAttached();
}

test('任务点击只修改对应源行，不影响代码示例和其他引用任务', async ({ page }) => {
  const source = '```md\n- [ ] 示例\n```\n\n> - [ ] 引用任务\n\n- [ ] 普通任务';
  await openContent(page, source);
  await expect(page.locator('.task-list-item-checkbox')).toHaveCount(2);
  await page.locator('.task-list-item-checkbox').first().click();
  await page.locator('#toggle-sourceview-btn').click();
  await expect.poll(async () => (await page.locator('.cm-content .cm-line').allTextContents())
    .join('\n')).toBe(source.replace('> - [ ]', '> - [x]'));
});

test('Ctrl 点击内部标题链接跳转到本文标题并保留当前页面', async ({ page }) => {
  await openContent(page, '[跳转](#目标标题)\n\n' + '正文\n\n'.repeat(100) + '# 目标标题');
  const pagesBefore = page.context().pages().length;
  await page.locator('.preview a', { hasText: '跳转' }).click({ modifiers: ['Control'] });
  await expect(page.locator('#md-目标标题')).toBeInViewport();
  expect(page.context().pages()).toHaveLength(pagesBefore);
  await expect(page.locator('.error-banner')).toHaveCount(0);
});

test('文档搜索支持中文全词和超过四百个匹配的完整导航', async ({ page }) => {
  await openContent(page, ('测试\n').repeat(450) + '测试文档', true);
  await page.keyboard.press('Control+f');
  await page.getByTestId('doc-search-query').fill('测试');
  await page.getByRole('button', { name: '全词匹配', exact: true }).click();
  await expect(page.locator('.doc-search-summary')).toHaveText('1/450');
  await page.getByRole('button', { name: '上一个匹配' }).click();
  await expect(page.locator('.doc-search-summary')).toHaveText('450/450');
});

test('批量普通替换保留美元字面量', async ({ page }) => {
  await openContent(page, 'foo foo', true);
  await page.keyboard.press('Control+h');
  const tools = await page.getByRole('toolbar', { name: '文档工具' }).boundingBox();
  const search = await page.getByTestId('doc-search-panel').boundingBox();
  expect(tools && search && tools.y + tools.height <= search.y).toBe(true);
  await page.getByTestId('doc-search-query').fill('foo');
  await page.getByTestId('doc-replace-query').fill('$&');
  await page.getByRole('button', { name: 'Replace All', exact: true }).click();
  await expect(page.locator('.cm-content').first()).toHaveText('$& $&');
});

test('正则单次替换保留上下文，空匹配在表情符号旁不会死循环', async ({ page }) => {
  await openContent(page, 'abc abc 😀', true);
  await page.keyboard.press('Control+h');
  await page.getByTestId('doc-search-query').fill('(?<=a)(b)(?=c)');
  await page.getByRole('button', { name: '正则表达式', exact: true }).click();
  await page.getByTestId('doc-replace-query').fill('$1-$0');
  await page.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect(page.locator('.cm-content').first()).toHaveText('ab-bc abc 😀');
  await page.getByTestId('doc-search-query').fill('(?=)');
  await expect(page.locator('.doc-search-summary')).toHaveText(/\/12$/);
});

test('中文字数逐字统计，表情符号按一个字符统计', async ({ page }) => {
  await openContent(page, '中文内容'.repeat(500) + '😀');
  await expect(page.getByText('2000 Words', { exact: true })).toBeVisible();
  await expect(page.getByText('2001 chars', { exact: true })).toBeVisible();
});

test('预览代码使用语法颜色且保留安全转义', async ({ page }) => {
  await openContent(page, '```js\nconst value = "<script>";\n```');
  await expect(page.locator('.preview .hljs-keyword')).toHaveText('const');
  await expect(page.locator('.preview script')).toHaveCount(0);
});

test('预览配对、正文续写和相邻块切换无需进入源码模式', async ({ page }) => {
  await openContent(page, '第一段\n\n第二段');
  await page.locator('.preview p').first().click({ modifiers: ['Alt'] });
  const textarea = page.locator('.preview-inline-editor-input');
  await expect(textarea).toBeFocused();
  await textarea.press('Control+End');
  await textarea.press('(');
  await expect(textarea).toHaveValue('第一段()');
  await textarea.press(')');
  await expect(textarea).toHaveValue('第一段()');
  await textarea.press('Enter');
  await textarea.pressSequentially('new');
  await expect(textarea).toHaveValue('第一段()\n\nnew');
  await textarea.press('Control+End');
  await textarea.press('ArrowDown');
  await expect(textarea).toHaveValue('第二段');
  await expect(textarea).toBeFocused();
});

test('专注模式淡化其他段落，快速打开和 PDF 设置可键盘关闭', async ({ page }) => {
  await openContent(page, '第一段\n\n第二段');
  await page.keyboard.press('Control+Alt+f');
  await expect(page.locator('.preview p').last()).toHaveCSS('opacity', '0.35');
  await page.keyboard.press('Control+p');
  await expect(page.getByRole('dialog')).toBeVisible();
  await expect(page.getByRole('textbox', { name: '查找文件名' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await page.getByRole('button', { name: 'PDF 设置' }).click();
  await page.getByRole('combobox', { name: '纸张大小' }).selectOption('Letter');
  await page.getByRole('spinbutton', { name: '页边距（毫米）' }).fill('20');
  await page.getByRole('button', { name: '保存设置' }).click();
  expect(await page.evaluate(() => JSON.parse(localStorage.getItem('unireadmd.pdf-options')!)))
    .toMatchObject({ pageSize: 'Letter', marginMm: 20 });
});

test('连续输入五个字符仅产生一次防抖恢复写入，载荷可结构化克隆', async ({ page }) => {
  await openContent(page, '正文', true);
  await page.evaluate(() => {
    Reflect.set(window, '__snapshotWrites', []);
    window.nativeBridge.session = {
      save: async (payload) => {
        const writes = Reflect.get(window, '__snapshotWrites');
        writes.push(structuredClone(payload));
        return { ok: true };
      },
      load: async () => null,
      confirmDiscard: async () => ({ action: 'cancel' }),
    };
  });
  await page.waitForTimeout(1200);
  await page.evaluate(() => { Reflect.set(window, '__snapshotWrites', []); });
  await page.locator('.cm-content').first().click();
  await page.keyboard.press('Control+End');
  await page.keyboard.type('abcde', { delay: 40 });
  await page.waitForTimeout(1300);
  const writes = await page.evaluate(() => Reflect.get(window, '__snapshotWrites'));
  expect(writes).toHaveLength(1);
  expect(writes[0].content).toBe('正文abcde');
});

test('预览打字机模式把编辑光标保持在视口中部', async ({ page }) => {
  await openContent(page, Array.from({ length: 40 }, (_, index) => `段落 ${index}`).join('\n\n'));
  await page.keyboard.press('Control+Alt+t');
  await page.locator('.preview p').nth(20).click({ modifiers: ['Alt'] });
  const textarea = page.locator('.preview-inline-editor-input');
  await expect(textarea).toBeFocused();
  await textarea.press('End');
  await textarea.pressSequentially('abc');
  await expect.poll(async () => page.evaluate(() => {
    const host = document.querySelector<HTMLElement>('.preview')!;
    const input = document.querySelector<HTMLTextAreaElement>('.preview-inline-editor-input')!;
    const viewport = host.getBoundingClientRect();
    const rect = input.getBoundingClientRect();
    const lineHeight = parseFloat(getComputedStyle(input).lineHeight) || 22;
    return Math.abs(rect.top + lineHeight / 2 - (viewport.top + viewport.height / 2));
  })).toBeLessThan(35);
});

test('真实页面五 MiB 文档可打开、搜索和滚动，并记录交互耗时', async ({ page }, testInfo) => {
  test.setTimeout(90_000);
  await openContent(page, '# 性能测试', true);
  const content = '# 性能测试\n\n' + ('正文内容 '.repeat(100) + 'search-target\n\n').repeat(5000);
  expect(Buffer.byteLength(content)).toBeGreaterThan(5 * 1024 * 1024);
  await page.evaluate((markdown) => {
    window.nativeBridge.file.openMarkdown = async () => ({
      filePath: '/virtual/large.md', content: markdown,
    });
    Reflect.set(window, '__openStarted', performance.now());
  }, content);
  await page.keyboard.press('Control+o');
  await expect(page.locator('.preview p')).toHaveCount(5000);
  const openMs = await page.evaluate(() => performance.now() - Reflect.get(window, '__openStarted'));
  const searchStarted = Date.now();
  await page.keyboard.press('Control+f');
  await page.getByTestId('doc-search-query').fill('search-target');
  await expect(page.locator('.doc-search-summary')).toHaveText('1/5000');
  const searchMs = Date.now() - searchStarted;
  await page.keyboard.press('Escape');
  await page.locator('.cm-content').first().focus();
  await page.keyboard.press('Control+End');
  const editStarted = Date.now();
  await page.keyboard.insertText('尾部编辑验证');
  await expect(page.locator('.preview p').last()).toHaveText('尾部编辑验证');
  const editMs = Date.now() - editStarted;
  await page.locator('#toggle-sourceview-btn').click();
  await page.getByTestId('preview-host').evaluate((node) => { node.scrollTop = node.scrollHeight; });
  await expect(page.locator('.preview p').last()).toBeInViewport();
  const metrics = { bytes: Buffer.byteLength(content), openMs, searchMs, editMs };
  const reportDirectory = path.resolve(testInfo.config.rootDir,
    '../../release/verification-20261009');
  await mkdir(reportDirectory, { recursive: true });
  await writeFile(path.join(reportDirectory, 'browser-performance.json'),
    JSON.stringify(metrics, null, 2), 'utf8');
  console.log('真实页面性能', metrics);
  await testInfo.attach('真实页面性能.json', {
    body: JSON.stringify(metrics),
    contentType: 'application/json',
  });
  expect(openMs).toBeLessThan(30_000);
  expect(searchMs).toBeLessThan(10_000);
  expect(editMs).toBeLessThan(10_000);
});
