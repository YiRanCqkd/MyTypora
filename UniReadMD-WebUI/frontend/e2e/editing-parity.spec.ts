import { expect, test, type Page } from './fixtures';

async function openContent(page: Page, text: string, source = false) {
  await page.addInitScript(({ text, source }) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      filePath: '', content: text, isDirty: false, cursorHint: 1, sourceMode: source,
    }));
  }, { text, source });
  await page.goto('/');
  await expect(page.getByTestId('preview-host')).toBeAttached();
}

test('表格编辑保持对齐及转义单元格，列对齐菜单可用', async ({ page }) => {
  await openContent(page, '| 甲 | 乙 |\n| :---: | ---: |\n| a\\|b | c |');
  await expect(page.locator('.preview tbody td')).toHaveCount(2);
  await page.locator('.preview tbody td').first().click();
  const input = page.locator('.preview-table-cell-editor input');
  await expect(input).toHaveValue('a\\|b');
  await input.fill('new|value');
  await input.press('Enter');
  await expect(page.locator('.preview tbody td').first()).toHaveText('new|value');
  await expect(page.locator('.preview tbody td').last()).toHaveCSS('text-align', 'right');
  await page.locator('.preview tbody td').first().click({ button: 'right' });
  await page.getByRole('button', { name: '列左对齐', exact: true }).click();
  await expect(page.locator('.preview tbody td').first()).toHaveCSS('text-align', 'left');
  await page.locator('#toggle-sourceview-btn').click();
  await expect(page.locator('.cm-content')).toContainText('| :--- | ---: |');
  await expect(page.locator('.cm-content')).toContainText('new\\|value');
});

test('源码富文本粘贴保留网页格式并过滤危险链接', async ({ page }) => {
  await openContent(page, '', true);
  await page.locator('.cm-content').focus();
  await page.locator('.cm-content').evaluate((element) => {
    const data = new DataTransfer();
    data.setData('text/html', '<h2>网页标题</h2><p><b>加粗</b> '
      + '<a href="https://example.com/">链接</a></p><pre><code>  code\n\nend</code></pre>'
      + '<a href="javascript:alert(1)">安全文字</a>');
    element.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data,
      bubbles: true, cancelable: true }));
  });
  await expect(page.locator('.cm-content')).toContainText('## 网页标题');
  await expect(page.locator('.cm-content')).toContainText('**加粗**');
  await expect(page.locator('.cm-content')).not.toContainText('javascript:');
  await page.locator('#toggle-sourceview-btn').click();
  await expect(page.locator('.preview h2')).toHaveText('网页标题');
  await expect(page.locator('.preview strong')).toHaveText('加粗');
});

test('预览编辑框接受富文本，复制同时提供 Markdown 和 HTML', async ({ page }) => {
  await openContent(page, '原文');
  await page.locator('.preview p').click({ modifiers: ['Alt'] });
  const textarea = page.locator('.preview-inline-editor-input');
  await textarea.selectText();
  await textarea.evaluate((element) => {
    const data = new DataTransfer();
    data.setData('text/html', '<p><b>格式内容</b></p>');
    element.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data,
      bubbles: true, cancelable: true }));
  });
  await expect(textarea).toHaveValue('**格式内容**');
  await textarea.selectText();
  const copied = await textarea.evaluate((element) => {
    const data = new DataTransfer();
    element.dispatchEvent(new ClipboardEvent('copy', { clipboardData: data,
      bubbles: true, cancelable: true }));
    return { plain: data.getData('text/plain'), html: data.getData('text/html') };
  });
  expect(copied.plain).toBe('**格式内容**');
  expect(copied.html).toContain('<strong>格式内容</strong>');
});

test('点击加粗内部时光标落在对应源码字符，输入后不跳到块末尾', async ({ page }) => {
  const source = '开头 **加粗内容** 结尾 [链接](https://example.com)';
  await openContent(page, source);
  const point = await page.locator('.preview strong').evaluate((element) => {
    const range = document.createRange();
    range.setStart(element.firstChild!, 2);
    range.setEnd(element.firstChild!, 3);
    const rect = range.getBoundingClientRect();
    return { x: rect.left + 1, y: rect.top + rect.height / 2 };
  });
  await page.keyboard.down('Alt');
  await page.mouse.click(point.x, point.y);
  await page.keyboard.up('Alt');
  const textarea = page.locator('.preview-inline-editor-input');
  await expect(textarea).toBeFocused();
  const offset = await textarea.evaluate((element: HTMLTextAreaElement) => element.selectionStart);
  await expect(textarea).toHaveValue('**加粗内容**');
  expect(offset).toBeGreaterThanOrEqual(2);
  expect(offset).toBeLessThanOrEqual(6);
  await expect(page.locator('.preview a')).toBeVisible();
  await textarea.pressSequentially('X');
  await page.waitForTimeout(250);
  expect(await textarea.evaluate((element: HTMLTextAreaElement) => element.selectionStart))
    .toBe(offset + 1);
});

test('跨块鼠标选区不会被 click 清除，可以复制带格式正文', async ({ page }) => {
  await openContent(page, '**第一段**\n\n第二段');
  const result = await page.getByTestId('preview-host').evaluate((host) => {
    const paragraphs = host.querySelectorAll('p');
    const range = document.createRange();
    range.setStart(paragraphs[0].querySelector('strong')!.firstChild!, 0);
    range.setEnd(paragraphs[1].firstChild!, 3);
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    paragraphs[1].dispatchEvent(new MouseEvent('click', { bubbles: true }));
    const data = new DataTransfer();
    host.dispatchEvent(new ClipboardEvent('copy', { clipboardData: data,
      bubbles: true, cancelable: true }));
    return { selected: selection.toString(), copied: data.getData('text/plain') };
  });
  expect(result.selected).toContain('第一段');
  expect(result.selected).toContain('第二段');
  expect(result.copied).toContain('**第一段**');
  await expect(page.locator('.preview-inline-editor-input')).toHaveCount(0);
});

test('高级导出与独立 PDF 边距保存后重新打开仍保持配置', async ({ page }) => {
  await openContent(page, '# 导出正文');
  await page.getByRole('button', { name: '导出设置', exact: true }).click();
  await page.getByLabel('标题', { exact: true }).fill('导出标题');
  await page.getByLabel('图片宽度（像素）').fill('2048');
  await page.getByLabel('JPEG 质量（1–100）').fill('75');
  await page.getByLabel('一级标题前分页').check();
  await page.getByRole('button', { name: '保存设置', exact: true }).click();
  await page.getByRole('button', { name: '导出设置', exact: true }).click();
  await expect(page.getByLabel('标题', { exact: true })).toHaveValue('导出标题');
  await expect(page.getByLabel('图片宽度（像素）')).toHaveValue('2048');
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'PDF 设置', exact: true }).click();
  await page.getByLabel('纸张大小').selectOption('Custom');
  await page.getByLabel('纸张宽度（毫米）').fill('180');
  await page.getByLabel('分别设置四侧页边距').check();
  await page.getByLabel('上侧（毫米）').fill('20');
  await page.getByRole('button', { name: '保存设置', exact: true }).click();
  const settings = await page.evaluate(() => {
    return JSON.parse(localStorage.getItem('unireadmd.pdf-options')!);
  });
  expect(settings.widthMm).toBe(180);
  expect(settings.marginsMm.top).toBe(20);
});

test('跨段落加粗按段应用，保持正常渲染且不把空行夹入格式标记', async ({ page }) => {
  await openContent(page, '第一段\n\n第二段');
  await page.getByTestId('preview-host').evaluate((host) => {
    const paragraphs = host.querySelectorAll('p');
    const range = document.createRange();
    range.setStart(paragraphs[0].firstChild!, 0);
    range.setEnd(paragraphs[1].firstChild!, 3);
    window.getSelection()!.removeAllRanges();
    window.getSelection()!.addRange(range);
  });
  await page.keyboard.press('Control+b');
  await expect(page.locator('.preview strong')).toHaveCount(2);
  await expect(page.locator('.preview strong').first()).toHaveText('第一段');
  await expect(page.locator('.preview strong').last()).toHaveText('第二段');
});

test('Ctrl Shift V 保留纯文本粘贴通道，不强制转换 HTML', async ({ page }) => {
  await openContent(page, '原文', true);
  await page.locator('.cm-content').focus();
  await page.locator('.cm-content').evaluate((element) => {
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'v', ctrlKey: true, shiftKey: true }));
    const data = new DataTransfer();
    data.setData('text/plain', '纯文本');
    data.setData('text/html', '<b>格式文本</b>');
    const event = new ClipboardEvent('paste', { clipboardData: data,
      bubbles: true, cancelable: true });
    element.dispatchEvent(event);
  });
  await expect(page.locator('.cm-content')).toContainText('纯文本');
  await expect(page.locator('.cm-content')).not.toContainText('格式文本');
});
