import { expect, test } from '@playwright/test';

const SESSION_STORAGE_KEY = 'unireadmd.phase2.session.v1';

const FIXED_SESSION = {
  filePath: 'I:/fixtures/visual-baseline.md',
  content: '# Visual Baseline\n\nTypora parity visual snapshot.\n\n- item 1\n- item 2',
  isDirty: false,
  cursorHint: 1,
  recentFiles: [
    'I:/fixtures/visual-baseline.md',
    'I:/fixtures/alpha.md',
    'I:/fixtures/beta.md',
  ],
  pinOutline: true,
  sourceMode: false,
  focusMode: false,
  typewriterMode: false,
  sidebarWidth: 320,
  fileViewMode: 'list',
  fileSortMode: 'natural',
  filesFilter: '',
  searchPanelVisible: false,
  replacePanelVisible: false,
  searchQuery: '',
  replaceQuery: '',
  searchCaseSensitive: false,
  searchWholeWord: false,
  searchMatchIndex: 0,
  fileTreeExpanded: {},
};

async function bootWithFixedSession(page: Parameters<typeof test>[0]['page']) {
  await page.addInitScript(
    ({ key, snapshot }) => {
      localStorage.setItem(key, JSON.stringify(snapshot));
    },
    {
      key: SESSION_STORAGE_KEY,
      snapshot: FIXED_SESSION,
    },
  );

  await page.goto('/');
  await expect(page.getByTestId('app-layout')).toBeVisible();
  await expect(page.getByTestId('preview-host')).toContainText('Visual Baseline');
}

async function ensureSourceEditorVisible(page: Parameters<typeof test>[0]['page']) {
  const layout = page.getByTestId('app-layout');
  const sourceButton = page.locator('#toggle-sourceview-btn');
  if (!(await layout.getAttribute('class'))?.includes('typora-sourceview-on')) {
    await sourceButton.click();
  }
  await expect(layout).toHaveClass(/typora-sourceview-on/);
  await expect(page.getByTestId('editor-host')).toBeVisible();
}

test.describe('visual baseline', () => {
  test.use({
    viewport: {
      width: 1440,
      height: 900,
    },
  });

  test.beforeEach(async ({ page }) => {
    await bootWithFixedSession(page);
  });

  test('布局基线: pin-outline', async ({ page }) => {
    await expect(page.getByTestId('app-layout')).toHaveScreenshot('layout-pin-outline.png', {
      animations: 'disabled',
      caret: 'hide',
      scale: 'css',
    });
  });

  test('布局基线: 侧栏隐藏', async ({ page }) => {
    await page.locator('#outline-btn').click();
    await expect(page.getByTestId('app-layout')).toHaveScreenshot('layout-sidebar-hidden.png', {
      animations: 'disabled',
      caret: 'hide',
      scale: 'css',
    });
  });

  test('布局基线: 搜索替换面板', async ({ page }) => {
    await page.keyboard.press('Control+F');
    await page.getByTestId('doc-search-query').fill('Visual');
    await page.getByTestId('doc-open-replace').click();
    await page.getByTestId('doc-replace-query').fill('Typora');

    await expect(page.locator('.main')).toHaveScreenshot('main-search-replace.png', {
      animations: 'disabled',
      caret: 'hide',
      scale: 'css',
    });
  });

  test('布局基线: 渲染态块编辑浮层', async ({ page }) => {
    await ensureSourceEditorVisible(page);
    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText([
      '# Visual Inline',
      '',
      'inline content',
    ].join('\n'));
    await page.locator('#toggle-sourceview-btn').click();

    await page.locator('.preview p').first().dblclick();
    await expect(page.getByTestId('preview-inline-editor')).toBeVisible();

    await expect(page.locator('.main')).toHaveScreenshot('main-preview-inline-editor.png', {
      animations: 'disabled',
      caret: 'hide',
      scale: 'css',
    });
  });

  test('布局基线: 表格右键菜单', async ({ page }) => {
    await ensureSourceEditorVisible(page);
    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText(
      [
        '| Name | Value |',
        '| --- | --- |',
        '| A | 1 |',
        '| B | 2 |',
      ].join('\n'),
    );
    await page.locator('#toggle-sourceview-btn').click();

    const tableCell = page.locator('.preview table tbody tr').first().locator('td').first();
    await tableCell.click({
      button: 'right',
    });
    await expect(page.getByTestId('preview-table-menu')).toBeVisible();

    await expect(page.locator('.main')).toHaveScreenshot('main-preview-table-menu.png', {
      animations: 'disabled',
      caret: 'hide',
      scale: 'css',
    });
  });
});
