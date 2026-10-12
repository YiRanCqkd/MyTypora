const sizeProperties = ['--preview-edit-width', '--preview-edit-height'];

export function preservePreviewScroll(preview: HTMLElement) {
  const scrollTop = preview.scrollTop;
  if (scrollTop <= 0) return () => {};
  const padding = preview.style.getPropertyValue('padding-bottom');
  const priority = preview.style.getPropertyPriority('padding-bottom');
  preview.style.setProperty('padding-bottom', `${preview.scrollHeight}px`);
  let restored = false;
  return () => {
    if (restored) return;
    restored = true;
    if (padding) preview.style.setProperty('padding-bottom', padding, priority);
    else preview.style.removeProperty('padding-bottom');
    preview.scrollTop = scrollTop;
  };
}

export function getPreviewCaretOffset(textarea: Pick<HTMLTextAreaElement,
  'selectionStart' | 'selectionEnd' | 'selectionDirection'>) {
  return textarea.selectionDirection === 'backward'
    ? textarea.selectionStart : textarea.selectionEnd;
}

export function observePreviewLayout(container: HTMLElement, onResize: () => void) {
  if (typeof ResizeObserver === 'undefined') return () => {};
  let frame = 0;
  let stopped = false;
  const observer = new ResizeObserver(() => {
    if (frame || stopped) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      onResize();
    });
  });
  observer.observe(container);
  return () => {
    stopped = true;
    observer.disconnect();
    if (frame) cancelAnimationFrame(frame);
  };
}

export function clearPreviewEditingLayout(preview: HTMLElement) {
  preview.querySelectorAll<HTMLElement>('.preview-editing-block, .preview-editing-fragment')
    .forEach((node) => {
      node.classList.remove('preview-editing-block', 'preview-editing-fragment');
      sizeProperties.forEach((property) => node.style.removeProperty(property));
      if (!node.getAttribute('style')) node.removeAttribute('style');
    });
}

export function revealPreviewEditorCaret(
  preview: HTMLElement,
  textarea: HTMLTextAreaElement,
  anchor?: HTMLElement | null,
) {
  if (!textarea.isConnected) return;
  const viewport = preview.getBoundingClientRect();
  const rect = textarea.getBoundingClientRect();
  // 滚动后编辑层的 Vue 样式可能尚未提交，以正文锚点计算真实位置，避免使用上一帧坐标。
  const top = anchor?.getBoundingClientRect().top ?? rect.top;
  const padding = 16;
  if (top >= viewport.top + padding && top + rect.height <= viewport.bottom - padding) return;
  const computed = getComputedStyle(textarea);
  const mirror = document.createElement('div');
  Object.assign(mirror.style, {
    position: 'fixed', top: '0', left: '0', visibility: 'hidden', pointerEvents: 'none',
    width: `${textarea.clientWidth}px`, whiteSpace: 'pre-wrap', overflowWrap: 'break-word',
    fontFamily: computed.fontFamily, fontSize: computed.fontSize, fontWeight: computed.fontWeight,
    fontStyle: computed.fontStyle, lineHeight: computed.lineHeight,
    letterSpacing: computed.letterSpacing, tabSize: computed.tabSize,
  });
  // 纯文本镜像测量真实换行位置，Markdown、中文和空行均不作为 HTML 执行。
  mirror.textContent = textarea.value.slice(0, getPreviewCaretOffset(textarea));
  const caret = document.createElement('span');
  caret.textContent = '\u200b';
  mirror.append(caret);
  document.body.append(mirror);
  const caretTop = top + caret.getBoundingClientRect().top
    - mirror.getBoundingClientRect().top;
  const lineHeight = parseFloat(computed.lineHeight)
    || (parseFloat(computed.fontSize) || 16) * 1.75;
  mirror.remove();
  if (caretTop < viewport.top + padding) {
    preview.scrollTop = Math.max(0, preview.scrollTop + caretTop - viewport.top - padding);
  } else if (caretTop + lineHeight > viewport.bottom - padding) {
    preview.scrollTop += caretTop + lineHeight - viewport.bottom + padding;
  }
}

export function layoutPreviewEditor(
  container: HTMLElement,
  preview: HTMLElement,
  anchor: HTMLElement,
  textarea: HTMLTextAreaElement | null,
  fragment: boolean,
) {
  const textNode = anchor.matches('pre') ? anchor.querySelector('code') || anchor : anchor;
  const computed = getComputedStyle(textNode);
  const typography = {
    fontFamily: computed.fontFamily,
    fontSize: computed.fontSize,
    fontWeight: computed.fontWeight,
    fontStyle: computed.fontStyle,
    lineHeight: computed.lineHeight,
    letterSpacing: computed.letterSpacing,
    color: computed.color,
  };
  const initial = anchor.getBoundingClientRect();
  let width = initial.width;
  let height = Math.max(initial.height, parseFloat(computed.lineHeight) || 24);
  if (textarea) {
    Object.assign(textarea.style, typography);
    if (fragment) {
      const measure = document.createElement('span');
      Object.assign(measure.style, typography, { position: 'absolute', visibility: 'hidden',
        whiteSpace: 'pre', pointerEvents: 'none' });
      measure.textContent = textarea.value || ' ';
      container.append(measure);
      const mapped = anchor.parentElement?.closest<HTMLElement>('[data-source-start]');
      const available = mapped?.clientWidth || preview.clientWidth;
      width = Math.max(8, Math.min(available,
        Math.ceil(Math.max(measure.getBoundingClientRect().width, measure.scrollWidth)) + 2));
      measure.remove();
      anchor.style.setProperty('--preview-edit-width', `${width}px`);
    }
    // 用实际源码测量高度，包含换行和 Markdown 标记；不读取渲染后的字数猜测尺寸。
    textarea.style.width = `${width}px`;
    textarea.style.height = '0px';
    height = Math.max(Math.ceil(textarea.scrollHeight), parseFloat(computed.lineHeight) || 24);
    textarea.style.height = `${height}px`;
    if (fragment) {
      anchor.style.setProperty('--preview-edit-height', `${height}px`);
    } else {
      const blocks = preview.querySelectorAll<HTMLElement>('.preview-editing-block');
      const bottom = Math.max(initial.bottom,
        ...Array.from(blocks, (node) => node.getBoundingClientRect().bottom));
      const extra = Math.max(0, height - (bottom - initial.top));
      anchor.style.setProperty('--preview-edit-height', `${initial.height + extra}px`);
    }
  }
  const rect = anchor.getBoundingClientRect();
  const origin = container.getBoundingClientRect();
  return {
    ...typography,
    backgroundColor: getComputedStyle(preview).backgroundColor,
    left: `${rect.left - origin.left}px`,
    top: `${rect.top - origin.top}px`,
    width: `${width}px`,
    minHeight: `${height}px`,
  };
}
