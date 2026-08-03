import type { createDiagramRenderer } from './diagram-renderer';
import type { useEditorHosts } from '../editor/editor-hosts';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';
import type { usePreferences } from '../settings/preferences';
type UseDiagramBlocksDependencies =
  ReturnType<typeof createDiagramRenderer> &
  ReturnType<typeof useEditorHosts> &
  ReturnType<typeof createWorkbenchRuntime> &
  ReturnType<typeof usePreferences>;

export function useDiagramBlocks(input: UseDiagramBlocksDependencies) {
  const { decodeDiagramSource,ensureFlowchartLoaded,ensureMermaidLoaded,previewHost,
    renderFlowSource,renderMermaidSource,runtime,theme,toMermaidSource,
  } = input;

  async function renderMermaidBlocks() {
    if (!previewHost.value) {
      return;
    }

    const blocks = Array.from(previewHost.value.querySelectorAll<HTMLElement>('.mermaid-block'));
    if (blocks.length === 0) {
      return;
    }

    const mermaid = await ensureMermaidLoaded();
    let flowchart: NonNullable<typeof runtime.flowchartApi> | null = null;
    const hasFlowBlock = blocks.some((block) => {
      return (block.dataset.diagramKind || 'mermaid') === 'flow';
    });
    if (hasFlowBlock) {
      try {
        flowchart = await ensureFlowchartLoaded();
      } catch (error) {
        console.warn(
          '[diagram] flowchart.js 加载失败，将尝试走 Mermaid 兼容渲染。',
          error,
        );
        flowchart = null;
      }
    }

    if (mermaid) {
      const nextThemeKey = theme.value === 'dark' ? 'dark' : 'default';
      if (runtime.mermaidThemeKey !== nextThemeKey) {
        mermaid.initialize({
          startOnLoad: false,
          securityLevel: 'strict',
          theme: nextThemeKey,
        });
        runtime.mermaidThemeKey = nextThemeKey;
      }
    }

    for (const block of blocks) {
      const kind = block.dataset.diagramKind || 'mermaid';
      const encoded = block.dataset.diagramCode || '';
      const source = decodeDiagramSource(encoded);
      const outputNode = block.querySelector<HTMLElement>('.mermaid-output');
      const fallbackNode = block.querySelector<HTMLElement>('.mermaid-fallback');
      const errorNode = block.querySelector<HTMLElement>('.mermaid-error');

      if (!source || !outputNode || !fallbackNode || !errorNode) {
        continue;
      }

      try {
        let rendered = false;
        if (kind === 'flow' && flowchart) {
          renderFlowSource(flowchart, outputNode, source);
          rendered = true;
        }

        if (!rendered) {
          if (!mermaid) {
            throw new Error('Mermaid 未加载成功');
          }

          const mermaidSource = toMermaidSource(kind, source);
          await renderMermaidSource(mermaid, outputNode, mermaidSource);
        }

        outputNode.style.display = 'block';
        fallbackNode.style.display = 'none';
        errorNode.style.display = 'none';
        errorNode.textContent = '';
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);

        outputNode.innerHTML = '';
        outputNode.style.display = 'none';
        fallbackNode.style.display = 'block';
        errorNode.style.display = 'block';
        errorNode.textContent = `${kind} 渲染失败：${message}`;
      }
    }
  }

  return { renderMermaidBlocks,
  };
}
