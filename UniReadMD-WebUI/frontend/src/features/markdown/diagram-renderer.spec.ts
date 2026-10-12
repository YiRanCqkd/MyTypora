import { beforeEach, describe, expect, it, vi } from 'vitest';
import { ref } from 'vue';
import MarkdownIt from 'markdown-it';
import { createDiagramRenderer } from './diagram-renderer';
import { createWorkbenchRuntime } from '../editor/workbench-runtime';

const mermaid = vi.hoisted(() => ({
  initialize: vi.fn(),
  render: vi.fn(),
  secure: ['secure', 'securityLevel', 'startOnLoad', 'maxTextSize',
    'suppressErrorRendering', 'maxEdges'],
}));
vi.mock('mermaid', () => ({
  default: { initialize: mermaid.initialize, render: mermaid.render,
    mermaidAPI: { defaultConfig: { secure: mermaid.secure } } },
}));

beforeEach(() => { vi.clearAllMocks(); });

function fixture() {
  const input = { ...createWorkbenchRuntime(), markdown: new MarkdownIt(),
    mathPluginRevision: ref(0) };
  return createDiagramRenderer(input as Parameters<typeof createDiagramRenderer>[0]);
}

describe('Mermaid 节点文字兼容性', () => {
  it('统一输出 SVG 文字，保留全部默认安全配置并禁止文档改回 HTML 标签', async () => {
    const renderer = fixture();
    const api = (await renderer.ensureMermaidLoaded())!;
    api.initialize({ startOnLoad: false, securityLevel: 'strict', theme: 'dark',
      flowchart: { htmlLabels: false }, htmlLabels: true, secure: [] });
    const config = mermaid.initialize.mock.calls[0][0];
    expect(config.htmlLabels).toBe(false);
    expect(config.theme).toBe('dark');
    expect(config.securityLevel).toBe('strict');
    expect(config.secure).toEqual([...mermaid.secure, 'htmlLabels']);
    expect(mermaid.secure).not.toContain('htmlLabels');
    expect(await renderer.ensureMermaidLoaded()).toBe(api);
  });

  it('节点 SVG 文字经过过滤保留，HTML 外部对象与事件仍被移除', async () => {
    const renderer = fixture();
    const api = (await renderer.ensureMermaidLoaded())!;
    const output = document.createElement('div');
    mermaid.render.mockResolvedValue({ svg: '<svg id="diagram"><g class="node">'
      + '<text onload="alert(1)">开始 Start</text></g>'
      + '<foreignObject><div>危险 HTML</div></foreignObject></svg>' });
    await renderer.renderMermaidSource(api, output, 'graph TD\nA[开始 Start]');
    expect(output.querySelector('.node text')?.textContent).toBe('开始 Start');
    expect(output.querySelector('foreignObject, [onload]')).toBeNull();
    expect(mermaid.render).toHaveBeenCalledWith('unireadmd-mermaid-1', 'graph TD\nA[开始 Start]');
  });
});
