import MarkdownIt from 'markdown-it';
import type { createMarkdownRenderer } from './renderer';
import type { usePreviewEditingState } from '../editor/preview-editing-state';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';
type CreateDiagramRendererDependencies =
  ReturnType<typeof createMarkdownRenderer> &
  ReturnType<typeof usePreviewEditingState> &
  ReturnType<typeof createWorkbenchRuntime>;

export function createDiagramRenderer(input: CreateDiagramRendererDependencies) {
  const { markdown,mathPluginRevision,runtime,
  } = input;

  async function ensureMathPluginLoaded() {
    if (mathPluginRevision.value > 0) {
      return;
    }

    if (runtime.mathPluginLoadPromise) {
      await runtime.mathPluginLoadPromise;
      return;
    }

    runtime.mathPluginLoadPromise = import('markdown-it-mathjax3')
      .then((module) => {
        const plugin = (module.default ?? module) as (instance: MarkdownIt) => void;
        markdown.use(plugin);
        mathPluginRevision.value += 1;
      })
      .finally(() => {
        runtime.mathPluginLoadPromise = null;
      });

    await runtime.mathPluginLoadPromise;
  }

  async function ensureMermaidLoaded() {
    if (runtime.mermaidApi) {
      return runtime.mermaidApi;
    }

    if (runtime.mermaidLoadPromise) {
      await runtime.mermaidLoadPromise;
      return runtime.mermaidApi;
    }

    runtime.mermaidLoadPromise = import('mermaid')
      .then((module) => {
        const api = module.default;
        runtime.mermaidApi = {
          initialize: api.initialize.bind(api),
          render: api.render.bind(api),
        };
      })
      .finally(() => {
        runtime.mermaidLoadPromise = null;
      });

    await runtime.mermaidLoadPromise;
    return runtime.mermaidApi;
  }

  async function ensureFlowchartLoaded() {
    if (runtime.flowchartApi) {
      return runtime.flowchartApi;
    }

    if (runtime.flowchartLoadPromise) {
      await runtime.flowchartLoadPromise;
      return runtime.flowchartApi;
    }

    runtime.flowchartLoadPromise = import('flowchart.js')
      .then((module) => {
        const api = (module.default ?? module) as {
          parse?: (code: string) => {
            drawSVG: (container: HTMLElement | string, options?: Record<string, unknown>) => void;
          };
        };

        if (typeof api.parse !== 'function') {
          throw new Error('flowchart.js 未提供 parse 方法');
        }

        runtime.flowchartApi = {
          parse: api.parse,
        };
      })
      .finally(() => {
        runtime.flowchartLoadPromise = null;
      });

    await runtime.flowchartLoadPromise;
    return runtime.flowchartApi;
  }

  function decodeDiagramSource(encoded: string) {
    try {
      return decodeURIComponent(encoded);
    } catch {
      return encoded;
    }
  }

  function toMermaidSource(kind: string, source: string) {
    const trimmed = source.trimStart();
    if (kind === 'sequence') {
      if (/^sequenceDiagram\b/.test(trimmed)) {
        return source;
      }

      return `sequenceDiagram\n${source}`;
    }

    if (kind === 'flow') {
      if (/^(flowchart|graph)\b/.test(trimmed)) {
        return source;
      }

      return `flowchart TD\n${source}`;
    }

    return source;
  }

  async function renderMermaidSource(
    mermaid: NonNullable<typeof runtime.mermaidApi>,
    outputNode: HTMLElement,
    source: string,
  ) {
    runtime.mermaidRenderSeed += 1;
    const renderId = `unireadmd-mermaid-${runtime.mermaidRenderSeed}`;
    const result = await mermaid.render(renderId, source);
    outputNode.innerHTML = result.svg;
  }

  function renderFlowSource(
    flowchart: NonNullable<typeof runtime.flowchartApi>,
    outputNode: HTMLElement,
    source: string,
  ) {
    outputNode.innerHTML = '';
    const chart = flowchart.parse(source);
    chart.drawSVG(outputNode, {
      'line-width': 2,
      'font-size': 13,
      'font-family': 'Consolas, Microsoft YaHei UI, sans-serif',
      'yes-text': 'yes',
      'no-text': 'no',
    });
  }

  return { ensureMathPluginLoaded,ensureMermaidLoaded,ensureFlowchartLoaded,
    decodeDiagramSource,toMermaidSource,renderMermaidSource,renderFlowSource,
  };
}
