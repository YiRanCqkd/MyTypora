import { onBeforeUnmount, onMounted } from 'vue';
import type { createDiagramRenderer } from './diagram-renderer';
import type { useEditorHosts } from '../editor/editor-hosts';
import type { createWorkbenchRuntime } from '../editor/workbench-runtime';
import type { usePreferences } from '../settings/preferences';

type Dependencies = ReturnType<typeof createDiagramRenderer> & ReturnType<typeof useEditorHosts>
  & ReturnType<typeof createWorkbenchRuntime> & ReturnType<typeof usePreferences>;

export function useDiagramBlocks(input: Dependencies) {
  const { decodeDiagramSource, ensureFlowchartLoaded, ensureMermaidLoaded, previewHost,
    renderFlowSource, renderMermaidSource, runtime, toMermaidSource } = input;
  let observer: IntersectionObserver | null = null;
  let queue = Promise.resolve();
  let disposed = false;
  const pending = new WeakSet<HTMLElement>();
  const rendered = new WeakMap<HTMLElement, string>();
  const cache = new Map<string, string>();
  let cacheBytes = 0;
  const themeKey = () => document.documentElement.dataset.theme === 'dark' ? 'dark' : 'default';
  const keyFor = (block: HTMLElement) => [themeKey(), block.dataset.sourceStart,
    block.dataset.diagramKind, block.dataset.diagramCode].join(':');

  async function renderBlock(block: HTMLElement) {
    if (disposed || !block.isConnected) return;
    const key = keyFor(block);
    if (rendered.get(block) === key) return;
    const kind = block.dataset.diagramKind || 'mermaid';
    const source = decodeDiagramSource(block.dataset.diagramCode || '');
    const output = block.querySelector<HTMLElement>('.mermaid-output');
    const fallback = block.querySelector<HTMLElement>('.mermaid-fallback');
    const errorNode = block.querySelector<HTMLElement>('.mermaid-error');
    if (!source || !output || !fallback || !errorNode) return;
    try {
      const cached = cache.get(key);
      if (cached !== undefined) output.innerHTML = cached;
      else {
        let flowchart: typeof runtime.flowchartApi = null;
        if (kind === 'flow') {
          try { flowchart = await ensureFlowchartLoaded(); }
          catch { /* 原流程图引擎不可用时继续使用 Mermaid 兼容路径。 */ }
        }
        if (flowchart) renderFlowSource(flowchart, output, source);
        else {
          const mermaid = await ensureMermaidLoaded();
          if (!mermaid) throw new Error('Mermaid 未加载成功');
          if (runtime.mermaidThemeKey !== themeKey()) {
            mermaid.initialize({ startOnLoad: false, securityLevel: 'strict',
              flowchart: { htmlLabels: false }, theme: themeKey() });
            runtime.mermaidThemeKey = themeKey();
          }
          await renderMermaidSource(mermaid, output, toMermaidSource(kind, source));
        }
        if (disposed || !block.isConnected) return;
        if (key !== keyFor(block)) return renderBlock(block);
        const html = output.innerHTML;
        cache.set(key, html);
        cacheBytes += key.length + html.length;
        while (cache.size > 64 || cacheBytes > 4 * 1024 * 1024) {
          const first = cache.keys().next().value!;
          cacheBytes -= first.length + cache.get(first)!.length;
          cache.delete(first);
        }
      }
      output.style.display = 'block';
      fallback.style.display = 'none';
      errorNode.style.display = 'none';
      errorNode.textContent = '';
      rendered.set(block, key);
    } catch (error) {
      output.innerHTML = '';
      output.style.display = 'none';
      fallback.style.display = 'block';
      errorNode.style.display = 'block';
      errorNode.textContent = `${kind} 渲染失败：${String(error)}`;
    }
  }
  function enqueue(block: HTMLElement) {
    if (pending.has(block) || rendered.get(block) === keyFor(block)) return;
    pending.add(block);
    // Mermaid 的全局配置和 SVG 生成串行执行，避免主题变更与异步渲染相互覆盖。
    queue = queue.then(() => renderBlock(block)).finally(() => { pending.delete(block); });
  }
  async function renderMermaidBlocks(all = false) {
    const host = previewHost.value;
    if (!host || disposed) return;
    observer?.disconnect();
    observer = null;
    const blocks = [...host.querySelectorAll<HTMLElement>('.mermaid-block')];
    const bounds = host.getBoundingClientRect();
    if (!all && typeof IntersectionObserver !== 'undefined') {
      observer = new IntersectionObserver((entries) => {
        for (const entry of entries) if (entry.isIntersecting) {
          enqueue(entry.target as HTMLElement);
          observer?.unobserve(entry.target);
        }
      }, { root: host, rootMargin: '600px' });
    }
    for (const block of blocks) {
      const rect = block.getBoundingClientRect();
      if (all || !observer || rect.top < bounds.bottom + 600 && rect.bottom > bounds.top - 600) {
        enqueue(block);
      } else observer.observe(block);
    }
    await queue;
  }
  const themeObserver = new MutationObserver(() => { void renderMermaidBlocks(); });
  onMounted(() => {
    themeObserver.observe(document.documentElement, {
      attributes: true, attributeFilter: ['data-theme'],
    });
  });
  onBeforeUnmount(() => {
    disposed = true; observer?.disconnect(); themeObserver.disconnect(); cache.clear();
  });
  return { renderMermaidBlocks };
}
