import { expect, test, type Locator, type Page } from './fixtures';

const graph = 'flowchart TD\nsubgraph 系统架构\n'
  + 'A[开始 Start] -->|下一步 Next| B[处理 Process]\n'
  + 'B --> C{成功 Success?}\nC -->|是 Yes| D[结束 End]\nend';
const graphLabels = ['开始 Start', '处理 Process', '成功 Success?', '结束 End',
  '系统架构', '下一步 Next', '是 Yes'];

async function openGraphs(page: Page, sources: string[], theme = 'light') {
  await page.addInitScript(({ sources, theme }) => {
    localStorage.setItem('unireadmd.web-theme.v1', theme);
    localStorage.setItem('unireadmd.phase2.session.v1', JSON.stringify({
      content: sources.map((source) => '```mermaid\n' + source + '\n```').join('\n\n'),
      sourceMode: false, filePath: '', isDirty: true, cursorHint: 1,
    }));
  }, { sources, theme });
  await page.goto('/');
}

async function expectLabels(svg: Locator, labels: string[]) {
  await expect(svg).toBeVisible();
  for (const label of labels) {
    const text = svg.locator('text').filter({ hasText: label }).first();
    await expect(text).toBeVisible();
    const bounds = (await text.boundingBox())!;
    expect(bounds.width).toBeGreaterThan(0);
    expect(bounds.height).toBeGreaterThan(0);
    await expect(text).not.toHaveCSS('fill', 'rgba(0, 0, 0, 0)');
    await expect(text).not.toHaveCSS('fill', 'none');
  }
  await expect(svg.locator('foreignObject, script, [onload]')).toHaveCount(0);
}

for (const theme of ['light', 'dark']) {
  test(`${theme} 主题流程图节点、连线与分组均显示中英文文字`, async ({ page }) => {
    await openGraphs(page, [graph], theme);
    await expect(page.locator('html')).toHaveAttribute('data-theme', theme);
    await expectLabels(page.locator('.mermaid-output svg'), graphLabels);
    await page.screenshot({ path: test.info().outputPath(`mermaid-labels-${theme}.png`) });
  });
}

test('文档内 HTML 标签配置与新节点形状不会导致文字消失', async ({ page }) => {
  await openGraphs(page, [
    '%%{init: {"htmlLabels":true,"flowchart":{"htmlLabels":true}}}%%\n' + graph,
    '---\nconfig:\n  htmlLabels: true\n  flowchart:\n    htmlLabels: true\n---\n' + graph,
    'flowchart LR\nA@{shape: rect, label: "新矩形"} --> '
      + 'B@{shape: cyl, label: "数据库"}\nB --> C@{shape: diamond, label: "判断"}',
    'flowchart LR\nA["第一行<br/>第二行"] --> B["`**重要步骤**`"]',
  ]);
  const svgs = page.locator('.mermaid-output svg');
  await expect(svgs).toHaveCount(4);
  await expectLabels(svgs.nth(0), graphLabels);
  await expectLabels(svgs.nth(1), graphLabels);
  await expectLabels(svgs.nth(2), ['新矩形', '数据库', '判断']);
  await expectLabels(svgs.nth(3), ['第一行', '第二行', '重要步骤']);
});

test('状态、类、时序与思维导图继续显示节点文字', async ({ page }) => {
  await openGraphs(page, [
    'stateDiagram-v2\nstate "等待审核" as waiting\nstate "处理完成" as done\n'
      + '[*] --> waiting\nwaiting --> done: 审核通过',
    'classDiagram\nclass 用户\n用户 : +String 姓名\n用户 : +登录()',
    'sequenceDiagram\nparticipant A as 客户端\nparticipant B as 服务端\nA->>B: 请求数据',
    'mindmap\n  root((系统中心))\n    功能一\n    功能二',
  ]);
  const svgs = page.locator('.mermaid-output svg');
  await expect(svgs).toHaveCount(4);
  await expectLabels(svgs.nth(0), ['等待审核', '处理完成', '审核通过']);
  await expectLabels(svgs.nth(1), ['用户', '姓名', '登录']);
  await expectLabels(svgs.nth(2), ['客户端', '服务端', '请求数据']);
  await expectLabels(svgs.nth(3), ['系统中心', '功能一', '功能二']);
});

test('图表编辑提交、主题切换与 HTML 导出保留实际节点文字', async ({ page }) => {
  await openGraphs(page, [graph]);
  const svg = page.locator('.mermaid-output svg');
  await expectLabels(svg, graphLabels);
  await svg.locator('.node text').filter({ hasText: '开始 Start' }).click();
  const input = page.locator('.preview-inline-editor-input');
  await expect(input).toBeVisible();
  const edited = (await input.inputValue()).replace('开始 Start', '修改后的起点');
  await input.fill(edited);
  await input.press('Control+Enter');
  await expectLabels(svg, ['修改后的起点', '处理 Process']);
  await page.getByText('偏好设置', { exact: true }).click();
  await page.getByRole('combobox', { name: '主题', exact: true }).selectOption('dark');
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await expectLabels(svg, ['修改后的起点', '处理 Process']);
  await page.evaluate(() => {
    window.nativeBridge.exporter.exportDocument = async (payload) => {
      Reflect.set(window, '__mermaidExport', payload.html);
      return { filePath: '/virtual/mermaid.html', format: 'html', engine: 'builtin' };
    };
  });
  await page.getByRole('combobox', { name: '导出格式' }).selectOption('html');
  await page.getByRole('button', { name: '导出', exact: true }).click();
  await expect.poll(() => page.evaluate(() => {
    const document = new DOMParser().parseFromString(
      String(Reflect.get(window, '__mermaidExport') || ''), 'text/html',
    );
    return [...document.querySelectorAll('.mermaid-output svg text')]
      .map((node) => node.textContent).join(' ');
  })).toContain('修改后的起点');
});
