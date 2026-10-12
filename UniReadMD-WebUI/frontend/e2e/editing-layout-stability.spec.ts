import { expect, test, type Page } from './fixtures';

async function openDocument(page: Page, content: string) {
  await page.addInitScript((content) => {
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content, sourceMode: false, filePath: '', isDirty: true,
    }));
  }, content);
  await page.goto('/');
  await expect(page.locator('.preview')).toBeVisible();
}

test('活动正文只保留插入光标，不绘制段落左侧装饰竖线', async ({ page }) => {
  await openDocument(page, '前一段\n\n正在编辑的正文\n\n后一段');
  const paragraph = page.locator('.preview p').nth(1);
  await paragraph.click();
  await expect(paragraph).toHaveAttribute('contenteditable', 'true');
  await paragraph.press('Home');
  await expect(paragraph).toBeFocused();
  await expect(paragraph).toHaveCSS('box-shadow', 'none');
  await expect(paragraph).toHaveCSS('border-left-width', '0px');
});

test('普通正文单击和输入不会插入上下空行，保存会话与重新渲染保持原有段距', async ({ page }) => {
  await openDocument(page, '前一段\n\n正在编辑的正文\n\n后一段');
  const paragraphs = page.locator('.preview > p');
  const before = await paragraphs.evaluateAll((nodes) => nodes.map((node) => ({
    top: node.getBoundingClientRect().top, height: node.getBoundingClientRect().height,
    marginTop: getComputedStyle(node).marginTop, marginBottom: getComputedStyle(node).marginBottom,
  })));
  const paragraph = paragraphs.nth(1);
  await paragraph.click();
  await expect(paragraph).toHaveAttribute('contenteditable', 'true');
  const assertGeometry = async () => {
    const after = await paragraphs.evaluateAll((nodes) => nodes.map((node) => ({
      top: node.getBoundingClientRect().top, height: node.getBoundingClientRect().height,
      marginTop: getComputedStyle(node).marginTop,
      marginBottom: getComputedStyle(node).marginBottom,
    })));
    expect(after).toEqual(before);
  };
  await assertGeometry();
  await paragraph.press('End');
  await page.keyboard.insertText('新增');
  await page.waitForTimeout(150);
  await assertGeometry();
  await page.screenshot({ path: test.info().outputPath('plain-paragraph-editing.png') });
  await paragraph.press('Control+Enter');
  await expect(paragraphs).toHaveCount(3);
  await assertGeometry();
  await page.locator('#toggle-sourceview-btn').check();
  await expect(page.locator('.cm-content')).toHaveText('前一段正在编辑的正文新增后一段');
  const source = await page.locator('.cm-line').allTextContents();
  expect(source.map((line) => line.replace(/\u200b/g, '')))
    .toEqual(['前一段', '', '正在编辑的正文新增', '', '后一段']);
});

test('正文已有换行在进入和退出可视编辑时不增加行距或移动后续段落', async ({ page }) => {
  await openDocument(page, '第一行 **加粗**\n第二行  \n第三行 `代码`\n\n后续段落');
  const paragraph = page.locator('.preview p').first();
  const following = page.locator('.preview p').last();
  const before = (await paragraph.boundingBox())!;
  const nextBefore = (await following.boundingBox())!;
  await paragraph.click();
  await expect(paragraph).toHaveAttribute('contenteditable', 'true');
  const editing = (await paragraph.boundingBox())!;
  expect(Math.abs(editing.height - before.height)).toBeLessThanOrEqual(1);
  expect(Math.abs((await following.boundingBox())!.y - nextBefore.y)).toBeLessThanOrEqual(1);
  await paragraph.press('End');
  await page.keyboard.insertText('新增');
  const edited = (await paragraph.boundingBox())!;
  await paragraph.press('Control+Enter');
  await expect(paragraph).toContainText('新增');
  expect(Math.abs((await paragraph.boundingBox())!.height - edited.height))
    .toBeLessThanOrEqual(1);
  await page.locator('#toggle-sourceview-btn').check();
  await expect(page.locator('.cm-content')).toContainText('**加粗**');
  await expect(page.locator('.cm-content')).toContainText('`代码`');
});

test('标题、列表和引用进入可视编辑时保持原有尺寸', async ({ page }) => {
  await openDocument(page, '# 一级标题\n\n## 二级标题\n\n###### 六级标题'
    + '\n\n- 第一项\n- 第二项\n\n> 引用首行\n> 引用次行');
  for (const selector of ['h1', 'h2', 'h6', 'li', 'blockquote p']) {
    const block = page.locator(`.preview ${selector}`).first();
    const before = (await block.boundingBox())!;
    await block.click();
    await expect(block).toHaveAttribute('contenteditable', 'true');
    const editing = (await block.boundingBox())!;
    expect(Math.abs(editing.height - before.height), selector).toBeLessThanOrEqual(1);
    expect(Math.abs(editing.width - before.width), selector).toBeLessThanOrEqual(1);
    await block.press('Control+Enter');
  }
});

for (const width of [1280, 960]) {
  test(`源码模式切换保持顶部和底部控件位置，视口宽度 ${width}`, async ({ page }) => {
    await page.setViewportSize({ width, height: 800 });
    await openDocument(page, '# 布局验证\n\n正文');
    const source = page.locator('#toggle-sourceview-btn');
    const sidebar = page.locator('#outline-btn');
    const selectors = ['.status-bar', '#toggle-sourceview-btn', '#outline-btn', '.top-titlebar'];
    for (const sidebarVisible of [true, false]) {
      await expect(page.getByTestId('sidebar'))[sidebarVisible ? 'toBeVisible' : 'toBeHidden']();
      // 侧栏操作自身有过渡动画；采样时先等待布局稳定。
      await page.waitForTimeout(250);
      const before = await Promise.all(selectors.map((selector) =>
        page.locator(selector).boundingBox()));
      await source.check();
      await expect(page.locator('.editor-host')).toBeVisible();
      await page.waitForTimeout(250);
      for (const [index, selector] of selectors.entries()) {
        const after = (await page.locator(selector).boundingBox())!;
        expect(Math.abs(after.x - before[index]!.x), selector).toBeLessThanOrEqual(1);
        expect(Math.abs(after.y - before[index]!.y), selector).toBeLessThanOrEqual(1);
        expect(Math.abs(after.width - before[index]!.width), selector).toBeLessThanOrEqual(1);
      }
      await source.uncheck();
      await expect(page.locator('.preview')).toBeVisible();
      if (sidebarVisible) await sidebar.click();
    }
  });
}
