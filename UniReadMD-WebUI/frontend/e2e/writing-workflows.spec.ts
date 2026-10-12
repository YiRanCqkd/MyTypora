import { expect, test } from './fixtures';

async function openContent(page: Parameters<typeof test>[0]['page'], content: string,
  sourceMode = false, numbering = false) {
  await page.addInitScript(({ content, sourceMode, numbering }) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content, sourceMode, filePath: '', isDirty: true, cursorHint: 1,
    }));
    if (!localStorage.getItem('unireadmd.writing.v1')) {
      localStorage.setItem('unireadmd.writing.v1', JSON.stringify({ equationNumbering: numbering }));
    }
  }, { content, sourceMode, numbering });
  await page.goto('/');
  await expect(page.getByTestId('app-layout')).toBeVisible();
}

test('格式片段单独展开，取消与撤销不改变相邻正文', async ({ page }) => {
  await openContent(page, '前文 **加粗** 与 *斜体* [链接](https://example.com)');
  await page.locator('.preview strong').click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await expect(input).toHaveValue('**加粗**');
  await expect(page.locator('.preview em')).toBeVisible();
  await input.fill('**修改**');
  await page.waitForTimeout(200);
  await input.press('Escape');
  await expect(page.locator('.preview strong')).toHaveText('加粗');
  await page.locator('.preview strong').click({ modifiers: ['Alt'] });
  await input.fill('**提交**');
  await expect(input).toHaveValue('**提交**');
  await input.press('Control+Enter');
  await expect(page.locator('.preview strong')).toHaveText('提交');
  await page.locator('#toggle-sourceview-btn').click();
  await page.locator('.cm-content').focus();
  await page.keyboard.press('Control+z');
  await expect(page.locator('.cm-content')).toContainText('**加粗** 与 *斜体*');
});

test('源码和预览选区显示字数，写作单位与阅读速度可持久化', async ({ page }) => {
  await openContent(page, '**中文** alpha', true);
  const source = page.locator('.cm-content');
  await source.focus();
  await page.keyboard.press('Control+a');
  await expect(page.locator('#footer-word-count')).toHaveText('选区 3 Words');
  await page.locator('#toggle-sourceview-btn').click();
  await page.locator('.preview p').evaluate((element) => {
    const range = document.createRange();
    range.selectNodeContents(element);
    window.getSelection()!.removeAllRanges();
    window.getSelection()!.addRange(range);
  });
  await expect(page.locator('#footer-word-count')).toHaveText('选区 3 Words');
  await page.locator('#footer-word-count').click();
  const dialog = page.getByRole('dialog', { name: '写作设置' });
  await dialog.getByLabel('状态栏统计').selectOption('characters');
  await dialog.getByLabel('阅读速度（字／分钟）').fill('300');
  await dialog.getByRole('button', { name: '完成' }).click();
  await expect(page.locator('#footer-word-count')).toContainText('chars');
  await page.reload();
  await page.locator('#footer-word-count').click();
  await expect(dialog.getByLabel('状态栏统计')).toHaveValue('characters');
  await expect(dialog.getByLabel('阅读速度（字／分钟）')).toHaveValue('300');
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
  await expect(page.locator('#footer-word-count')).toBeFocused();
});

test('移除片段格式后结束编辑，同段其他格式仍可见且内容保留', async ({ page }) => {
  await openContent(page, '**第一项** 与 **第二项**');
  await page.locator('.preview strong').first().click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await input.fill('普通文字');
  await expect(input).toHaveCount(0);
  await expect(page.locator('.preview strong')).toHaveText('第二项');
  await expect(page.locator('.preview strong')).toBeVisible();
  await expect(page.locator('.preview p')).toHaveText('普通文字 与 第二项');
});

test('片段中文组合输入不提前写回，组合期间 Escape 不取消文档编辑', async ({ page }) => {
  await openContent(page, '前文 **原文** 后文');
  await page.locator('.preview strong').click({ modifiers: ['Alt'] });
  const input = page.locator('.preview-inline-editor-input');
  await input.evaluate((element: HTMLTextAreaElement) => {
    element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    element.value = '**中文输入**';
    element.dispatchEvent(new InputEvent('input', { bubbles: true, isComposing: true }));
    element.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Escape', bubbles: true, isComposing: true,
    }));
  });
  await expect(input).toBeVisible();
  await expect(page.locator('.preview strong')).toHaveText('原文');
  await input.evaluate((element) => {
    element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
  });
  await expect(page.locator('.preview strong')).toHaveText('中文输入');
  await input.press('Control+Enter');
  await expect(page.locator('.preview p')).toHaveText('前文 中文输入 后文');
});

test('公式前向引用、自动编号和化学表达式实际渲染', async ({ page }) => {
  await openContent(page, '参见 $\\eqref{mass}$\n\n$$\nE=mc^2\\label{mass}\n$$\n\n'
    + '$\\ce{H2O}$', false, true);
  await expect(page.locator('.math-reference')).toHaveText('(1)');
  await expect(page.locator('.preview .katex-error')).toHaveCount(0);
  await expect(page.locator('.preview .katex')).toHaveCount(2);
  await page.locator('.math-reference').click({ modifiers: ['Control'] });
  await expect(page.locator('#equation-mass')).toBeInViewport();
});

test('明暗写作主题分别生效，系统变化与重载保持配置', async ({ page }) => {
  await openContent(page, '# 写作主题');
  await page.locator('#footer-word-count').click();
  const dialog = page.getByRole('dialog', { name: '写作设置' });
  await dialog.getByLabel('浅色写作主题').selectOption('sepia');
  await dialog.getByLabel('深色写作主题').selectOption('midnight');
  await dialog.getByRole('button', { name: '完成' }).click();
  await page.getByText('偏好设置', { exact: true }).click();
  await page.getByRole('combobox', { name: '主题', exact: true }).selectOption('system');
  await page.emulateMedia({ colorScheme: 'light' });
  await expect(page.locator('html')).toHaveAttribute('data-writing-theme', 'sepia');
  await expect(page.locator('.preview')).toHaveCSS('background-color', 'rgb(250, 244, 232)');
  await page.emulateMedia({ colorScheme: 'dark' });
  await expect(page.locator('html')).toHaveAttribute('data-writing-theme', 'midnight');
  await expect(page.locator('.preview')).toHaveCSS('background-color', 'rgb(20, 30, 46)');
  await expect(page.locator('.outline-item.active')).toHaveCSS('color', 'rgb(218, 229, 244)');
  await page.locator('#footer-word-count').click();
  await expect(dialog.getByRole('button', { name: '完成', exact: true }))
    .toHaveCSS('color', 'rgb(218, 229, 244)');
  await dialog.getByRole('button', { name: '完成', exact: true }).click();
  await page.reload();
  await expect(page.locator('html')).toHaveAttribute('data-writing-theme', 'midnight');
  await expect(page.locator('.preview')).toHaveCSS('background-color', 'rgb(20, 30, 46)');
});

test('屏幕外图表延后渲染，滚动可显示，导出包含全部图表', async ({ page }) => {
  test.setTimeout(90_000);
  const content = Array.from({ length: 10 }, (_, index) =>
    `## 图表 ${index}\n\n\`\`\`mermaid\ngraph TD\nA${index}-->B${index}\n\`\`\`\n\n`
      + '间隔正文\n'.repeat(25) + '\n').join('');
  await openContent(page, content);
  await expect(page.locator('.mermaid-output svg').first()).toBeVisible();
  expect(await page.locator('.mermaid-output svg').count()).toBeLessThan(10);
  await page.locator('.mermaid-block').last().scrollIntoViewIfNeeded();
  await expect(page.locator('.mermaid-block').last().locator('svg')).toBeVisible();
  await page.evaluate(() => {
    window.nativeBridge.exporter.exportDocument = async (payload) => {
      Reflect.set(window, '__writingExport', payload.html);
      return { filePath: '/virtual/diagrams.html', format: 'html', engine: 'builtin' };
    };
  });
  await page.getByRole('combobox', { name: '导出格式' }).selectOption('html');
  await page.getByRole('button', { name: '导出', exact: true }).click();
  await expect.poll(() => page.evaluate(() => {
    const html = String(Reflect.get(window, '__writingExport') || '');
    return new DOMParser().parseFromString(html, 'text/html').querySelectorAll('.mermaid-output svg')
      .length;
  }), { timeout: 40_000 }).toBe(10);
});
