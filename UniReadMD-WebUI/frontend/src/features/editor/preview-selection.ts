import type { PreviewTaskItemContext } from './editor-contracts';
import type { useEditorHosts } from './editor-hosts';
type CreatePreviewSelectionDependencies =
  ReturnType<typeof useEditorHosts>;

export function createPreviewSelection(input: CreatePreviewSelectionDependencies) {
  const { previewHost,
  } = input;

  function findSourceMappedNode(target: EventTarget | null) {
    if (!(target instanceof Element)) {
      return null;
    }

    // SVG 节点和文字也可触发点击，但编辑映射必须落在外层 HTML 源码块。
    const mapped = target.closest('pre[data-source-start][data-source-end][data-source-token]')
      || target.closest('[data-source-start][data-source-end][data-source-token]');
    return mapped instanceof HTMLElement ? mapped : null;
  }

  function parseInlineEditSelection(node: HTMLElement) {
    const rawStart = Number(node.dataset.sourceStart || 0);
    const rawEnd = Number(node.dataset.sourceEnd || 0);
    const token = node.dataset.sourceToken || '';
    if (!Number.isFinite(rawStart) || rawStart <= 0) {
      return null;
    }

    if (!Number.isFinite(rawEnd) || rawEnd <= 0) {
      return null;
    }

    if (!token) {
      return null;
    }

    return {
      startLine: Math.round(rawStart),
      endLine: Math.round(rawEnd),
      token,
    };
  }

  function resolvePreviewTaskItemContext(target: HTMLElement | null) {
    if (!target) {
      return null;
    }

    let checkbox = target.closest<HTMLInputElement>('.task-list-item-checkbox');
    if (!checkbox) {
      const label = target.closest<HTMLLabelElement>('label.task-list-item-label[for]');
      if (label) {
        const labelFor = label.getAttribute('for');
        const candidate = labelFor ? document.getElementById(labelFor) : null;
        if (
          candidate instanceof HTMLInputElement
          && candidate.classList.contains('task-list-item-checkbox')
        ) {
          checkbox = candidate;
        }
      }
    }

    if (!checkbox) {
      return null;
    }

    let mappedNode = checkbox.closest<HTMLElement>(
      "[data-source-token='list_item_open'][data-source-start][data-source-end]",
    );
    if (!mappedNode) {
      mappedNode = checkbox.closest<HTMLElement>('[data-source-start][data-source-end]');
    }

    if (!mappedNode) {
      let parent = checkbox.parentElement;
      while (parent) {
        if (parent.dataset.sourceStart && parent.dataset.sourceEnd) {
          mappedNode = parent;
          break;
        }
        parent = parent.parentElement;
      }
    }

    if (!mappedNode) {
      return null;
    }

    const startLine = Number(mappedNode.dataset.sourceStart || 0);
    const endLine = Number(mappedNode.dataset.sourceEnd || 0);
    if (
      !Number.isFinite(startLine) ||
      !Number.isFinite(endLine) ||
      startLine <= 0 ||
      endLine <= 0
    ) {
      return null;
    }

    const taskCheckboxes = Array.from(
      mappedNode.querySelectorAll<HTMLInputElement>('.task-list-item-checkbox'),
    );
    const resolvedTaskIndex = Math.max(0, taskCheckboxes.indexOf(checkbox));
    const allTaskCheckboxes = previewHost.value
      ? Array.from(previewHost.value.querySelectorAll<HTMLInputElement>('.task-list-item-checkbox'))
      : [];
    const resolvedGlobalTaskIndex = Math.max(0, allTaskCheckboxes.indexOf(checkbox));

    return {
      startLine: Math.round(startLine),
      endLine: Math.round(endLine),
      taskIndex: resolvedTaskIndex,
      globalTaskIndex: resolvedGlobalTaskIndex,
    } as PreviewTaskItemContext;
  }

  return { findSourceMappedNode,parseInlineEditSelection,
    resolvePreviewTaskItemContext,
  };
}
