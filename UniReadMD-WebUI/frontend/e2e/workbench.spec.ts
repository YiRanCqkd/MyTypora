import { expect, test } from '@playwright/test';
import { validateMarkdownSyntax } from '../src/utils/markdown';

test.describe('UniReadMD workbench', () => {
  async function ensureSourceEditorVisible(page: Parameters<typeof test>[0]['page']) {
    const layout = page.getByTestId('app-layout');
    const sourceButton = page.locator('#toggle-sourceview-btn');
    if (!(await layout.getAttribute('class'))?.includes('typora-sourceview-on')) {
      await sourceButton.click();
    }
    await expect(layout).toHaveClass(/typora-sourceview-on/);
    await expect(page.getByTestId('editor-host')).toBeVisible();
  }

  test('应支持在 Outline 中按标题搜索并保留父标题', async ({ page }) => {
    await page.goto('/');
    await expect(page.getByTestId('app-layout')).toBeVisible();
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.type('# Root\n\n## Child\n\ncontent line');

    await page.getByTestId('tab-outline').click();
    await expect(page.getByTestId('outline-item')).toHaveCount(2);
    await expect(page.getByTestId('outline-item').first()).toContainText('Root');

    await expect(page.getByTestId('sidebar-search-input')).toHaveAttribute(
      'placeholder',
      '匹配当前文档标题',
    );
    await page.getByTestId('sidebar-search-input').fill('Child');
    await expect(page.getByTestId('outline-item')).toHaveCount(2);
    await expect(page.getByTestId('outline-item').first()).toContainText('Root');
    await expect(page.getByTestId('outline-item').nth(1)).toContainText('Child');
  });

  test('Markdown 链接应渲染为新标签打开属性', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('[OpenAI](https://openai.com)');

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const link = page.locator('.preview a').first();
    await expect(link).toHaveAttribute('href', 'https://openai.com');
    await expect(link).toHaveAttribute('target', '_blank');
    await expect(link).toHaveAttribute('rel', 'noopener noreferrer');
  });

  test('渲染态链接应默认进入编辑，Ctrl/Cmd+点击不拦截默认行为', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('[OpenAI](https://openai.com)');

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const link = page.locator('.preview a').first();
    const plainClickCancelled = await link.evaluate((node) => {
      const event = new MouseEvent('click', {
        bubbles: true,
        cancelable: true,
      });
      return !node.dispatchEvent(event);
    });
    expect(plainClickCancelled).toBe(true);
    await expect(page.getByTestId('preview-inline-editor')).toBeVisible();

    await page.keyboard.press('Escape');
    await expect(page.getByTestId('preview-inline-editor')).toHaveCount(0);

    const modifier = process.platform === 'darwin' ? 'metaKey' : 'ctrlKey';
    const modifierClickCancelled = await link.evaluate((node, key) => {
      const options: MouseEventInit = {
        bubbles: true,
        cancelable: true,
      };
      if (key === 'metaKey') {
        options.metaKey = true;
      } else {
        options.ctrlKey = true;
      }
      const event = new MouseEvent('click', options);
      return !node.dispatchEvent(event);
    }, modifier);
    expect(modifierClickCancelled).toBe(false);
  });

  test('预览模式点击大纲标题应滚动到对应位置', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');

    const sections = Array.from(
      {
        length: 36,
      },
      (_, index) => {
        const no = index + 1;
        return [
          `## Section ${no}`,
          `paragraph ${no}`,
          `paragraph ${no} extended line for scrolling`,
          '',
        ].join('\n');
      },
    ).join('\n');

    await page.keyboard.insertText([
      '# Root',
      '',
      sections,
    ].join('\n'));

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    await page.getByTestId('tab-outline').click();
    await page.getByTestId('outline-item').filter({
      hasText: 'Section 30',
    }).click();

    const preview = page.getByTestId('preview-host');
    await expect.poll(async () => {
      return await preview.evaluate((node) => Math.round((node as HTMLElement).scrollTop));
    }).toBeGreaterThan(500);

    const targetHeading = page.locator('.preview h2', {
      hasText: 'Section 30',
    });
    await expect(targetHeading).toBeVisible();
    await expect.poll(async () => {
      return await targetHeading.evaluate((node) => {
        const host = node.closest('.preview') as HTMLElement | null;
        if (!host) {
          return -1;
        }

        const rect = node.getBoundingClientRect();
        const hostRect = host.getBoundingClientRect();
        return Math.round(rect.top - hostRect.top);
      });
    }).toBeLessThan(240);
  });

  test('预览滚动时大纲高亮应与可视标题同步', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');

    const sections = Array.from(
      {
        length: 34,
      },
      (_, index) => {
        const no = index + 1;
        return [
          `## Sync ${no}`,
          `line ${no}`,
          `line ${no} more text for viewport sync`,
          '',
        ].join('\n');
      },
    ).join('\n');

    await page.keyboard.insertText([
      '# Root',
      '',
      sections,
    ].join('\n'));

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();
    await page.getByTestId('tab-outline').click();

    await page.getByRole('button', {
      name: 'Sync 1',
      exact: true,
    }).click();

    const preview = page.getByTestId('preview-host');
    await preview.evaluate((node) => {
      const host = node as HTMLElement;
      host.scrollTop = host.scrollHeight * 0.7;
      host.dispatchEvent(new Event('scroll'));
    });

    const expectedActiveText = await preview.evaluate((node) => {
      const host = node as HTMLElement;
      const anchorTop = host.scrollTop + 18;
      const headings = Array.from(
        host.querySelectorAll<HTMLElement>("[data-source-token='heading_open'][data-source-start]"),
      );
      let activeText = '';
      headings.forEach((heading) => {
        if (heading.offsetTop <= anchorTop) {
          activeText = (heading.textContent || '').trim();
        }
      });
      return activeText;
    });

    expect(expectedActiveText.length).toBeGreaterThan(0);
    const activeOutlineItem = page.locator('.outline-item.active').first();
    await expect(activeOutlineItem).toContainText(expectedActiveText);
  });

  test('选中内容区后鼠标滚轮应可正常滚动', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');

    const sections = Array.from(
      {
        length: 48,
      },
      (_, index) => {
        const no = index + 1;
        return [
          `## Wheel ${no}`,
          `wheel line ${no}`,
          `wheel line ${no} extended`,
          '',
        ].join('\n');
      },
    ).join('\n');

    await page.keyboard.insertText([
      '# Root',
      '',
      sections,
    ].join('\n'));

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const preview = page.getByTestId('preview-host');
    await preview.click();
    await page.mouse.wheel(0, 900);

    await expect.poll(async () => {
      return await preview.evaluate((node) => Math.round((node as HTMLElement).scrollTop));
    }).toBeGreaterThan(240);
  });

  test('Source 与 Render 切换应保持光标与滚动位置一致', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');

    const sections = Array.from(
      {
        length: 72,
      },
      (_, index) => {
        const no = index + 1;
        return [
          `## Anchor ${no}`,
          `line ${no}`,
          `line ${no} extra`,
          '',
        ].join('\n');
      },
    ).join('\n');

    await page.keyboard.insertText([
      '# Root',
      '',
      sections,
    ].join('\n'));

    await page.keyboard.press('Control+End');
    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const preview = page.getByTestId('preview-host');
    await expect.poll(async () => {
      return await preview.evaluate((node) => Math.round((node as HTMLElement).scrollTop));
    }).toBeGreaterThan(1200);
    const previewScrollBefore = await preview.evaluate((node) => Math.round((node as HTMLElement).scrollTop));

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();

    await expect.poll(async () => {
      return await page.locator('.editor-host .cm-scroller').evaluate((node) => {
        return Math.round((node as HTMLElement).scrollTop);
      });
    }).toBeGreaterThan(600);
    const sourceScrollTop = await page.locator('.editor-host .cm-scroller').evaluate((node) => {
      return Math.round((node as HTMLElement).scrollTop);
    });

    expect(previewScrollBefore).toBeGreaterThan(0);
    expect(sourceScrollTop).toBeGreaterThan(0);

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    await expect.poll(async () => {
      return await preview.evaluate((node) => Math.round((node as HTMLElement).scrollTop));
    }).toBeGreaterThan(1100);
  });

  test('应支持保存并更新文件状态', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.type('# Save Flow\n\ncontent');

    await expect(page.locator('.dirty-dot')).toBeVisible();
    await page.getByTestId('tab-files').click();
    await expect(page.getByTestId('files-tree-view')).toContainText('暂无 Markdown 文件');

    await page.keyboard.press('Control+S');
    await expect(page.locator('.dirty-dot')).toHaveCount(0);
    await expect(page.getByTestId('files-tree-view')).toContainText('virtual');
  });

  test('pin-outline 应联动顶栏与底栏位移', async ({ page }) => {
    await page.goto('/');

    const layout = page.getByTestId('app-layout');
    const sidebar = page.getByTestId('sidebar');
    const topTitlebar = page.locator('.top-titlebar');
    const statusBar = page.locator('.status-bar');
    const outlineToggle = page.locator('#outline-btn');

    await expect(layout).toHaveClass(/pin-outline/);
    await expect(topTitlebar).toBeVisible();
    await expect(statusBar).toBeVisible();

    const sidebarBox = await sidebar.boundingBox();
    expect(sidebarBox?.width ?? 0).toBeGreaterThan(220);

    const topBefore = await topTitlebar.boundingBox();
    const statusBefore = await statusBar.boundingBox();
    expect(topBefore?.x ?? 0).toBeGreaterThan(120);
    expect(statusBefore?.x ?? 0).toBeGreaterThan(120);

    await outlineToggle.click();
    await expect(layout).toHaveClass(/hide-sidebar/);

    await expect.poll(async () => {
      const box = await topTitlebar.boundingBox();
      return box?.x ?? 999;
    }).toBeLessThan(6);

    await expect.poll(async () => {
      const box = await statusBar.boundingBox();
      return box?.x ?? 999;
    }).toBeLessThan(6);

    await outlineToggle.click();
    await expect(layout).toHaveClass(/pin-outline/);

    await expect.poll(async () => {
      const box = await topTitlebar.boundingBox();
      return box?.x ?? 0;
    }).toBeGreaterThan(120);
  });

  test('侧栏宽度拖拽后应在刷新后保持', async ({ page }) => {
    await page.goto('/');

    const sidebar = page.getByTestId('sidebar');
    const resizer = page.getByTestId('sidebar-resizer');

    await expect(resizer).toBeVisible();

    const before = await sidebar.boundingBox();
    const resizeBox = await resizer.boundingBox();
    expect(before).not.toBeNull();
    expect(resizeBox).not.toBeNull();

    const startX = (resizeBox?.x ?? 0) + (resizeBox?.width ?? 0) / 2;
    const startY = (resizeBox?.y ?? 0) + (resizeBox?.height ?? 0) / 2;
    const targetX = startX + 120;

    await page.mouse.move(startX, startY);
    await page.mouse.down();
    await page.mouse.move(targetX, startY, { steps: 14 });
    await page.mouse.up();

    await expect.poll(async () => {
      const box = await sidebar.boundingBox();
      return box?.width ?? 0;
    }).toBeGreaterThan((before?.width ?? 0) + 70);

    const afterResize = await sidebar.boundingBox();
    const resizedWidth = afterResize?.width ?? 0;
    expect(resizedWidth).toBeGreaterThan((before?.width ?? 0) + 70);

    await page.reload();

    await expect.poll(async () => {
      const box = await sidebar.boundingBox();
      return box?.width ?? 0;
    }).toBeGreaterThan((before?.width ?? 0) + 70);

    const afterReload = await sidebar.boundingBox();
    const reloadWidth = afterReload?.width ?? 0;
    expect(Math.abs(reloadWidth - resizedWidth)).toBeLessThan(24);
  });

  test('Files 视图应默认展示树并支持按文档名搜索', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.type('# File A\n\nalpha');
    await page.getByTestId('tab-files').click();
    await page.keyboard.press('Control+S');

    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.type('# File B\n\nbeta');
    await page.keyboard.press('Control+Shift+S');

    const treeView = page.getByTestId('files-tree-view');
    await expect(treeView).toBeVisible();
    await expect(treeView.locator('.files-tree-row')).toHaveCount(6);
    await expect(treeView).toContainText('virtual');
    await expect(treeView).toContainText('最近打开');

    await expect(page.getByTestId('sidebar-search-input')).toHaveAttribute(
      'placeholder',
      '匹配文档名（当前目录 / 最近打开）',
    );
    const sidebarSearchInput = page.getByTestId('sidebar-search-input');
    await sidebarSearchInput.fill('document-2');
    await expect(treeView).toContainText('document-2.md');
    await expect(treeView).not.toContainText('document-1.md');

    await page.getByTestId('tab-outline').click();
    await expect(sidebarSearchInput).toHaveValue('');
    await sidebarSearchInput.fill('File B');

    await page.getByTestId('tab-files').click();
    await expect(sidebarSearchInput).toHaveValue('document-2');
  });

  test('主内容搜索替换面板应支持查找与 Replace All', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.type('foo bar\\nfoo baz\\nfoo');

    await page.keyboard.press('Control+F');
    await expect(page.getByTestId('doc-search-panel')).toBeVisible();

    await page.getByTestId('doc-search-query').fill('foo');
    await expect(page.getByTestId('doc-search-summary')).toContainText('1/3');
    await page.getByRole('button', { name: '↓' }).click();
    await expect(page.getByTestId('doc-search-summary')).toContainText('2/3');

    await page.getByTestId('doc-open-replace').click();
    await page.getByTestId('doc-replace-query').fill('bar');
    await page.getByRole('button', { name: 'Replace All' }).click();

    await expect(editor).toContainText('bar bar');
    await expect(editor).not.toContainText('foo');
  });

  test('Source 模式应支持 Ctrl+B/I/K 与缩进快捷键', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();

    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('Alpha');
    await page.keyboard.press('Control+A');
    await page.keyboard.press('Control+B');
    await expect(editor).toContainText('**Alpha**');

    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('Beta');
    await page.keyboard.press('Control+A');
    await page.keyboard.press('Control+I');
    await expect(editor).toContainText('*Beta*');

    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('OpenAI');
    await page.keyboard.press('Control+A');
    await page.keyboard.press('Control+K');
    await expect(editor).toContainText('[OpenAI](https://)');

    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('- item');
    await page.keyboard.press('Control+A');
    await page.keyboard.press('Tab');
    await expect(editor).toContainText('  - item');

    await page.keyboard.press('Shift+Tab');
    await expect(editor).toContainText('- item');
  });

  test('渲染态内联编辑应支持 Ctrl+B/I/K 与缩进快捷键', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('inline text');

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    await page.locator('.preview p').first().click();
    const inlineEditor = page.getByTestId('preview-inline-editor');
    const textarea = inlineEditor.locator('textarea');
    await expect(inlineEditor).toBeVisible();

    await textarea.press('Control+A');
    await textarea.press('Control+B');
    await expect(textarea).toHaveValue('**inline text**');

    await textarea.fill('inline text');
    await textarea.press('Control+A');
    await textarea.press('Control+I');
    await expect(textarea).toHaveValue('*inline text*');

    await textarea.fill('OpenAI');
    await textarea.press('Control+A');
    await textarea.press('Control+K');
    await expect(textarea).toHaveValue('[OpenAI](https://)');

    await textarea.fill('- item');
    await textarea.press('Control+A');
    await textarea.press('Tab');
    await expect(textarea).toHaveValue('  - item');

    await textarea.press('Shift+Tab');
    await expect(textarea).toHaveValue('- item');

    await textarea.press('Control+Enter');
    await expect(inlineEditor).toHaveCount(0);

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();
    await expect(editor).toContainText('- item');
  });

  test('渲染态应支持单击块编辑并回写源码', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText([
      '# Inline Edit',
      '',
      'old paragraph',
    ].join('\n'));

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const paragraph = page.locator('.preview p').first();
    await paragraph.click();

    const inlineEditor = page.getByTestId('preview-inline-editor');
    await expect(inlineEditor).toBeVisible();
    const textarea = inlineEditor.locator('textarea');
    await textarea.fill('new paragraph');
    await textarea.press('Control+Enter');

    await expect(inlineEditor).toHaveCount(0);
    await expect(page.locator('.preview p').first()).toContainText('new paragraph');

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();
    await expect(editor).toContainText('new paragraph');
  });

  test('渲染态内联编辑应实时生效并支持 Esc 回滚', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('draft paragraph');

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const paragraph = page.locator('.preview p').first();
    await paragraph.click();

    const inlineEditor = page.getByTestId('preview-inline-editor');
    await expect(inlineEditor).toBeVisible();
    const textarea = inlineEditor.locator('textarea');
    await textarea.fill('## live heading');

    await expect(page.locator('.preview h2', {
      hasText: 'live heading',
    })).toBeVisible();

    await textarea.press('Escape');
    await expect(inlineEditor).toHaveCount(0);
    await expect(page.locator('.preview p').first()).toContainText('draft paragraph');
    await expect(page.locator('.preview h2', {
      hasText: 'live heading',
    })).toHaveCount(0);

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();
    await expect(editor).toContainText('draft paragraph');
    await expect(editor).not.toContainText('## live heading');
  });

  test('渲染态内联编辑提交后应可通过撤销回到编辑前内容', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('undo baseline');

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    await page.locator('.preview p').first().click();
    const inlineEditor = page.getByTestId('preview-inline-editor');
    await expect(inlineEditor).toBeVisible();
    const textarea = inlineEditor.locator('textarea');
    await textarea.fill('');
    await textarea.type('undo target content with multiple updates');
    await textarea.press('Control+Enter');
    await expect(inlineEditor).toHaveCount(0);

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();
    await expect(editor).toContainText('undo target content with multiple updates');

    await editor.click();
    let restored = false;
    for (let index = 0; index < 12; index += 1) {
      await page.keyboard.press('Control+Z');
      const text = (await editor.textContent()) || '';
      if (text.includes('undo baseline')) {
        restored = true;
        break;
      }
    }

    expect(restored).toBe(true);
    await expect(editor).toContainText('undo baseline');
    await expect(editor).not.toContainText('undo target content with multiple updates');
  });

  test('渲染态表格应支持单元格直接编辑与 Tab/Enter 导航', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText(
      [
        '| A | B |',
        '| --- | --- |',
        '| 1 | 2 |',
        '| 3 | 4 |',
      ].join('\n'),
    );

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    await page.locator('.preview table tbody tr').first().locator('td').first().click();
    const cellEditor = page.getByTestId('preview-table-cell-editor');
    const cellInput = page.getByTestId('preview-table-cell-editor-input');
    await expect(cellEditor).toBeVisible();
    await expect(cellInput).toHaveValue('1');

    await cellInput.fill('11');
    await cellInput.press('Tab');
    await expect(cellInput).toHaveValue('2');
    await expect(page.locator('.preview table tbody tr').first().locator('td').first()).toContainText('11');

    await cellInput.fill('22');
    await cellInput.press('Enter');
    await expect(cellInput).toHaveValue('4');
    await expect(page.locator('.preview table tbody tr').first().locator('td').nth(1)).toContainText('22');

    await cellInput.press('Escape');
    await expect(cellEditor).toHaveCount(0);

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();
    await expect(editor).toContainText('| 11 | 22 |');
  });

  test('渲染态表格单元格编辑应支持结构快捷操作并同步 Markdown', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText(
      [
        '| A | B |',
        '| --- | --- |',
        '| 1 | 2 |',
        '| 3 | 4 |',
      ].join('\n'),
    );

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    await page.locator('.preview table tbody tr').first().locator('td').first().click();
    const cellInput = page.getByTestId('preview-table-cell-editor-input');
    await expect(cellInput).toBeVisible();

    await cellInput.press('Control+Alt+ArrowDown');
    await expect(page.locator('.preview table tbody tr').first().locator('td').first()).toContainText('3');

    await cellInput.press('Control+Alt+ArrowUp');
    await expect(page.locator('.preview table tbody tr').first().locator('td').first()).toContainText('1');

    await cellInput.press('Control+Alt+ArrowRight');
    await expect(page.locator('.preview table tr').first().locator('th').first()).toContainText('B');

    await cellInput.press('Control+Alt+ArrowLeft');
    await expect(page.locator('.preview table tr').first().locator('th').first()).toContainText('A');

    await cellInput.press('Control+Shift+Enter');
    await expect(page.locator('.preview table tbody tr')).toHaveCount(3);

    await cellInput.press('Control+Shift+Alt+Enter');
    await expect(page.locator('.preview table tr').first().locator('th')).toHaveCount(3);

    await cellInput.press('Control+Shift+Alt+Delete');
    await expect(page.locator('.preview table tr').first().locator('th')).toHaveCount(2);

    await cellInput.press('Control+Shift+Delete');
    await expect(page.locator('.preview table tbody tr')).toHaveCount(2);

    await cellInput.press('Escape');
    await expect(page.getByTestId('preview-table-cell-editor')).toHaveCount(0);

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();
    await expect(editor).toContainText('| A | B |');
    await expect(editor).toContainText('| 1 | 2 |');
  });

  test('渲染态表格右键操作应同步回写 Markdown', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText(
      [
        '| A | B |',
        '| --- | --- |',
        '| 1 | 2 |',
        '| 3 | 4 |',
      ].join('\n'),
    );

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const firstBodyCell = page.locator('.preview table tbody tr').first().locator('td').first();
    await firstBodyCell.click({
      button: 'right',
    });
    const tableMenu = page.getByTestId('preview-table-menu');
    await expect(tableMenu).toBeVisible();
    await tableMenu.getByRole('button', { name: '在下方插入行' }).click();
    await expect(page.locator('.preview table tbody tr')).toHaveCount(3);

    const secondColCell = page.locator('.preview table tbody tr').first().locator('td').nth(1);
    await secondColCell.click({
      button: 'right',
    });
    await expect(tableMenu).toBeVisible();
    await tableMenu.getByRole('button', { name: '删除当前列' }).click();
    await expect(page.locator('.preview table tr').first().locator('th')).toHaveCount(1);

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();
    await expect(editor).toContainText('| A |');
    await expect(editor).not.toContainText('| B |');
  });

  test('Source 右键菜单应支持查找与 ESC 关闭', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText('# Title\n\nsearch-target');

    const editorHost = page.getByTestId('editor-host');
    const sourceMenu = page.getByTestId('source-context-menu');

    await editorHost.click({
      button: 'right',
      position: {
        x: 56,
        y: 48,
      },
    });
    await expect(sourceMenu).toBeVisible();
    await sourceMenu.getByRole('button', { name: '查找' }).click();
    await expect(page.getByTestId('doc-search-panel')).toBeVisible();
    await expect(sourceMenu).toHaveCount(0);

    await editorHost.click({
      button: 'right',
      position: {
        x: 90,
        y: 76,
      },
    });
    await expect(sourceMenu).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(sourceMenu).toHaveCount(0);
  });

  test('扩展语法应渲染为 Typora 风格结构', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.type(
      [
        '[toc]',
        '',
        '- [x] done task',
        '',
        'inline math: $a+b$',
        'H~2~O and X^2^',
        '==highlight== :smile:',
        '',
        'ref[^1]',
        '',
        '[^1]: footnote text',
      ].join('\n'),
    );

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const preview = page.locator('.preview');
    await expect(preview.locator('.md-toc')).toBeVisible();
    await expect.poll(async () => {
      return await preview.locator('.task-list-item input[type="checkbox"]').count();
    }).toBeGreaterThan(0);
    await expect.poll(async () => {
      return await preview.locator('.task-list-item input[type="checkbox"]:checked').count();
    }).toBeGreaterThan(0);
    await expect.poll(async () => {
      const texts = await preview.locator('sub').allTextContents();
      return texts.some((item) => item.includes('2'));
    }).toBe(true);
    await expect.poll(async () => {
      const texts = await preview.locator('sup').allTextContents();
      return texts.some((item) => item.includes('2'));
    }).toBe(true);
    await expect(preview.locator('mark')).toContainText('highlight');
    await expect(preview.locator('.footnotes')).toContainText('footnote text');
  });

  test('复杂语料渲染态编辑后应保持 Markdown 结构有效', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText(
      [
        '# Complex Corpus',
        '',
        'Paragraph before render edit.',
        '',
        '| A | B |',
        '| --- | --- |',
        '| 1 | 2 |',
        '',
        '- [ ] task item',
        '',
        'inline math $a+b$',
        '',
        '$$',
        'c = \\sqrt{a^2 + b^2}',
        '$$',
        '',
        '```mermaid',
        'graph TD',
        '  A[Start] --> B[End]',
        '```',
        '',
        'ref[^1]',
        '',
        '[^1]: footnote text',
        '',
        '<details><summary>More</summary><p>html block</p></details>',
      ].join('\n'),
    );

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    await page.locator('.preview p', {
      hasText: 'Paragraph before render edit.',
    }).first().click();
    const inlineEditor = page.getByTestId('preview-inline-editor');
    const textarea = inlineEditor.locator('textarea');
    await expect(inlineEditor).toBeVisible();
    await textarea.fill('Paragraph after render edit.');
    await textarea.press('Control+Enter');
    await expect(inlineEditor).toHaveCount(0);

    const checkboxes = page.locator('.preview .task-list-item-checkbox');
    await expect(checkboxes).toHaveCount(1);
    await checkboxes.first().click();
    await expect(checkboxes.first()).toBeChecked();

    const firstCell = page.locator('.preview table tbody tr').first().locator('td').first();
    await firstCell.click();
    const cellInput = page.getByTestId('preview-table-cell-editor-input');
    await expect(cellInput).toBeVisible();
    await cellInput.fill('11');
    await cellInput.press('Tab');
    await page.keyboard.press('Escape');

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();
    const sourceText = await page.locator('.editor-host .cm-content').evaluate((node) => {
      const lines = Array.from(node.querySelectorAll('.cm-line'));
      return lines
        .map((line) => line.textContent || '')
        .join('\n');
    });
    const issues = validateMarkdownSyntax(sourceText);
    const blockingRules = new Set([
      'fence',
      'footnote-ref',
      'footnote-def',
      'block-math',
      'table-header',
      'diagram-mermaid',
      'diagram-flow',
      'diagram-sequence',
      'task-list',
      'link-target',
    ]);
    const blockingIssues = issues.filter((item) => blockingRules.has(item.rule));

    expect(blockingIssues).toEqual([]);
    expect(sourceText).toContain('```mermaid');
    expect(sourceText).toContain('[^1]: footnote text');
    expect(sourceText).toContain('$$');
    expect(sourceText).toContain('<details><summary>More</summary><p>html block</p></details>');
  });

  test('渲染态任务列表 checkbox 点击应回写 Markdown', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText(
      [
        '- [ ] todo item',
        '- [x] done item',
      ].join('\n'),
    );

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const checkboxes = page.locator('.preview .task-list-item-checkbox');
    await expect(checkboxes).toHaveCount(2);
    await expect(checkboxes.nth(0)).not.toBeChecked();
    await expect(checkboxes.nth(1)).toBeChecked();

    await checkboxes.nth(0).click();
    await expect(checkboxes.nth(0)).toBeChecked();

    await checkboxes.nth(1).click();
    await expect(checkboxes.nth(1)).not.toBeChecked();

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).toBeVisible();
    await expect(editor).toContainText('- [x] todo item');
    await expect(editor).toContainText('- [ ] done item');
  });

  test('图表代码块应支持 mermaid/sequence/flow', async ({ page }) => {
    await page.goto('/');
    await ensureSourceEditorVisible(page);

    const editor = page.locator('.cm-content').first();
    await editor.click();
    await page.keyboard.press('Control+A');
    await page.keyboard.insertText(
      [
        '```mermaid',
        'graph TD',
        '  A[Start] --> B[End]',
        '```',
        '',
        '```sequence',
        'Alice->>Bob: hello',
        '```',
        '',
        '```flow',
        'st=>start: Start',
        'op=>operation: Work',
        'e=>end: End',
        'st->op->e',
        '```',
      ].join('\n'),
    );

    await page.locator('#toggle-sourceview-btn').click();
    await expect(page.getByTestId('editor-host')).not.toBeVisible();

    const preview = page.locator('.preview');
    await expect.poll(async () => {
      return await preview.locator('.mermaid-block').count();
    }, {
      timeout: 20000,
    }).toBe(3);
    await expect.poll(async () => {
      return await preview.locator('.mermaid-block .mermaid-output svg').count();
    }, {
      timeout: 20000,
    }).toBeGreaterThan(2);
    await expect(preview.locator('.mermaid-block .mermaid-error:visible')).toHaveCount(0);
  });

  test('Source/Focus/Typewriter 模式应支持切换与会话恢复', async ({ page }) => {
    await page.goto('/');

    const layout = page.getByTestId('app-layout');
    const sourceBtn = page.locator('#toggle-sourceview-btn');
    const previewHost = page.getByTestId('preview-host');
    const editorHost = page.getByTestId('editor-host');
    const mainResizer = page.getByTestId('main-preview-resizer');
    const topTitlebar = page.locator('.top-titlebar');
    const statusBar = page.locator('.status-bar');
    const outlineBtn = page.locator('#outline-btn');
    const statusRight = page.locator('.status-right');

    await expect(layout).not.toHaveClass(/typora-sourceview-on/);
    await expect(layout).not.toHaveClass(/on-focus-mode/);
    await expect(layout).not.toHaveClass(/ty-on-typewriter-mode/);
    await expect(previewHost).toBeVisible();
    await expect(editorHost).not.toBeVisible();
    await expect(outlineBtn).toBeVisible();
    await expect(statusRight).toBeVisible();

    const topBefore = await topTitlebar.boundingBox();
    const statusBefore = await statusBar.boundingBox();
    expect(topBefore?.x ?? 0).toBeGreaterThan(120);
    expect(statusBefore?.x ?? 0).toBeGreaterThan(120);

    await sourceBtn.click();
    await expect(previewHost).not.toBeVisible();
    await expect(editorHost).toBeVisible();
    await expect(mainResizer).not.toBeVisible();
    await expect(outlineBtn).not.toBeVisible();
    await expect(statusRight).not.toBeVisible();

    await expect.poll(async () => {
      const box = await topTitlebar.boundingBox();
      return box?.x ?? 999;
    }).toBeLessThan(6);
    await expect.poll(async () => {
      const box = await statusBar.boundingBox();
      return box?.x ?? 999;
    }).toBeLessThan(6);

    await page.keyboard.press('Control+Alt+F');
    await page.keyboard.press('Control+Alt+T');

    await expect(layout).toHaveClass(/typora-sourceview-on/);
    await expect(layout).toHaveClass(/on-focus-mode/);
    await expect(layout).toHaveClass(/ty-on-typewriter-mode/);

    await page.reload();

    await expect(layout).not.toHaveClass(/typora-sourceview-on/);
    await expect(layout).toHaveClass(/on-focus-mode/);
    await expect(layout).toHaveClass(/ty-on-typewriter-mode/);
    await expect(editorHost).not.toBeVisible();
  });

  test('拼写检查状态应支持开关、语言选择与状态展示', async ({ page }) => {
    await page.goto('/');

    await page.keyboard.press('Control+Alt+L');
    await expect(page.getByTestId('spell-panel')).toBeVisible();
    await expect(page.getByTestId('spell-dictionary-status')).toContainText('Dictionary');

    const enableCheckbox = page.getByTestId('spell-panel').locator('input[type="checkbox"]');
    await enableCheckbox.uncheck();
    await expect(enableCheckbox).not.toBeChecked();
    await enableCheckbox.check();
    await expect(enableCheckbox).toBeChecked();

    const languageSelect = page.getByTestId('spell-panel').locator('select');
    await languageSelect.selectOption('zh-CN');
    await expect(page.getByTestId('spell-dictionary-status')).toContainText('Dictionary Ready');
  });
});
