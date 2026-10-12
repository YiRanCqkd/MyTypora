import { afterEach, describe, expect, it, vi } from 'vitest';
import { clearPreviewEditingLayout, layoutPreviewEditor, revealPreviewEditorCaret }
  from './preview-editor-layout';
import { getPreviewCaretOffset, observePreviewLayout, preservePreviewScroll }
  from './preview-editor-layout';

afterEach(() => {
  document.body.replaceChildren();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

function bounds(node: HTMLElement, top: number, width: number, height: number, left = 100) {
  vi.spyOn(node, 'getBoundingClientRect').mockReturnValue({
    top, left, width, height, right: left + width, bottom: top + height,
    x: left, y: top, toJSON: () => ({}),
  });
}

function fixture(tag = 'p', height = 80) {
  const container = document.createElement('section');
  const preview = document.createElement('article');
  const anchor = document.createElement(tag);
  const textarea = document.createElement('textarea');
  anchor.style.cssText = 'font-size: 20px; line-height: 28px; font-family: serif;';
  anchor.dataset.sourceStart = '1';
  preview.append(anchor);
  container.append(preview, textarea);
  document.body.append(container);
  bounds(container, 10, 700, 500, 40);
  bounds(anchor, 30, 600, height);
  Object.defineProperty(preview, 'clientWidth', { value: 600 });
  Object.defineProperty(anchor, 'clientWidth', { value: 600 });
  Object.defineProperty(textarea, 'scrollHeight', { value: 200, configurable: true });
  return { container, preview, anchor, textarea };
}

describe('预览编辑真实字体与占位布局', () => {
  it('DOM 更新前保留滚动空间，恢复已有内边距和优先级且重复清理不回跳', () => {
    const f = fixture();
    f.preview.scrollTop = 500;
    Object.defineProperty(f.preview, 'scrollHeight', { value: 2400 });
    f.preview.style.setProperty('padding-bottom', '80px', 'important');
    const restore = preservePreviewScroll(f.preview);
    expect(f.preview.style.paddingBottom).toBe('2400px');
    f.preview.scrollTop = 30;
    restore();
    expect(f.preview.scrollTop).toBe(500);
    expect(f.preview.style.paddingBottom).toBe('80px');
    expect(f.preview.style.getPropertyPriority('padding-bottom')).toBe('important');
    f.preview.scrollTop = 200;
    restore();
    expect(f.preview.scrollTop).toBe(200);
  });

  it('原来没有内边距时移除临时样式，顶部更新不干预后续用户滚动', () => {
    const f = fixture();
    const topCleanup = preservePreviewScroll(f.preview);
    f.preview.scrollTop = 50;
    topCleanup();
    expect(f.preview.scrollTop).toBe(50);
    const restore = preservePreviewScroll(f.preview);
    restore();
    expect(f.preview.style.getPropertyValue('padding-bottom')).toBe('');
  });

  it('选区光标按活动端定位，折叠选区保持原位置', () => {
    expect(getPreviewCaretOffset({ selectionStart: 2, selectionEnd: 8,
      selectionDirection: 'forward' })).toBe(8);
    expect(getPreviewCaretOffset({ selectionStart: 2, selectionEnd: 8,
      selectionDirection: 'backward' })).toBe(2);
    expect(getPreviewCaretOffset({ selectionStart: 4, selectionEnd: 4,
      selectionDirection: 'none' })).toBe(4);
  });

  it('容器变化合并到一帧，卸载后断开观察并忽略迟到通知', () => {
    let notify: () => void = () => {};
    let render: FrameRequestCallback = () => {};
    const observe = vi.fn();
    const disconnect = vi.fn();
    vi.stubGlobal('ResizeObserver', class {
      constructor(callback: () => void) { notify = callback; }
      observe = observe;
      disconnect = disconnect;
    });
    const request = vi.fn((callback: FrameRequestCallback) => { render = callback; return 7; });
    const cancel = vi.fn();
    vi.stubGlobal('requestAnimationFrame', request);
    vi.stubGlobal('cancelAnimationFrame', cancel);
    const resized = vi.fn();
    const f = fixture();
    const cleanup = observePreviewLayout(f.container, resized);
    expect(observe).toHaveBeenCalledWith(f.container);
    notify();
    notify();
    expect(request).toHaveBeenCalledTimes(1);
    render(0);
    expect(resized).toHaveBeenCalledTimes(1);
    notify();
    cleanup();
    notify();
    expect(disconnect).toHaveBeenCalledOnce();
    expect(cancel).toHaveBeenCalledWith(7);
    expect(request).toHaveBeenCalledTimes(2);
  });

  it('无容器观察能力时保留窗口事件回退，空闲卸载不取消无关动画', () => {
    vi.stubGlobal('ResizeObserver', undefined);
    expect(() => observePreviewLayout(fixture().container, vi.fn())()).not.toThrow();
    const disconnect = vi.fn();
    vi.stubGlobal('ResizeObserver', class { observe() {} disconnect = disconnect; });
    const cancel = vi.fn();
    vi.stubGlobal('cancelAnimationFrame', cancel);
    observePreviewLayout(fixture().container, vi.fn())();
    expect(disconnect).toHaveBeenCalledOnce();
    expect(cancel).not.toHaveBeenCalled();
  });

  it('编辑继承自定义正文颜色和写作主题背景', () => {
    const f = fixture();
    f.anchor.style.color = 'rgb(10, 20, 30)';
    f.preview.style.backgroundColor = 'rgb(240, 230, 210)';
    const style = layoutPreviewEditor(f.container, f.preview, f.anchor, f.textarea, false);
    expect(style.color).toBe('rgb(10, 20, 30)');
    expect(f.textarea.style.color).toBe(style.color);
    expect(style.backgroundColor).toBe('rgb(240, 230, 210)');
  });

  it('输入控件挂载前保持标题字号与原始坐标', () => {
    const f = fixture('h1', 48);
    const style = layoutPreviewEditor(f.container, f.preview, f.anchor, null, false);
    expect(style).toMatchObject({ fontSize: '20px', lineHeight: '28px',
      left: '60px', top: '20px', width: '600px', minHeight: '48px' });
    expect(f.anchor.style.getPropertyValue('--preview-edit-height')).toBe('');
  });

  it('长源码增加占位高度，后续块已有的高度不重复累加', () => {
    const f = fixture();
    const following = document.createElement('p');
    following.classList.add('preview-editing-block');
    f.preview.append(following);
    bounds(following, 120, 600, 60);
    const style = layoutPreviewEditor(f.container, f.preview, f.anchor, f.textarea, false);
    expect(style.minHeight).toBe('200px');
    expect(f.textarea.style.height).toBe('200px');
    expect(f.textarea.style.fontFamily).toBe('serif');
    expect(f.anchor.style.getPropertyValue('--preview-edit-height')).toBe('130px');
  });

  it('单块增长按实际输入高度扩展，不使用固定两行尺寸', () => {
    const f = fixture();
    layoutPreviewEditor(f.container, f.preview, f.anchor, f.textarea, false);
    expect(f.anchor.style.getPropertyValue('--preview-edit-height')).toBe('200px');
  });

  it('短片段按源码测量宽度，最长不超过所在正文且不插入 HTML', () => {
    const f = fixture('strong', 28);
    const paragraph = document.createElement('p');
    paragraph.dataset.sourceStart = '1';
    Object.defineProperty(paragraph, 'clientWidth', { value: 60 });
    f.preview.append(paragraph);
    paragraph.append(f.anchor);
    f.textarea.value = '<img src=x>**粗体**';
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockReturnValue({
      width: 90, height: 28, top: 0, bottom: 28, left: 0, right: 90,
      x: 0, y: 0, toJSON: () => ({}),
    });
    const style = layoutPreviewEditor(f.container, f.preview, f.anchor, f.textarea, true);
    expect(style.width).toBe('60px');
    expect(f.anchor.style.getPropertyValue('--preview-edit-width')).toBe('60px');
    expect(f.anchor.style.getPropertyValue('--preview-edit-height')).toBe('200px');
    expect(f.container.querySelector('img')).toBeNull();
    expect(f.container.querySelector('span')).toBeNull();
  });

  it('空片段至少保留光标宽度，缺少段落映射时按预览宽度限制', () => {
    const f = fixture('strong');
    Object.defineProperty(f.textarea, 'scrollHeight', { value: 0 });
    const style = layoutPreviewEditor(f.container, f.preview, f.anchor, f.textarea, true);
    expect(style.width).toBe('8px');
    expect(style.minHeight).toBe('28px');
  });

  it('代码块采用内部代码字体，未设置行高时提供安全高度', () => {
    const f = fixture('pre');
    const code = document.createElement('code');
    code.style.fontFamily = 'monospace';
    f.anchor.append(code);
    const style = layoutPreviewEditor(f.container, f.preview, f.anchor, null, false);
    expect(style.fontFamily).toBe('monospace');
    expect(style.minHeight).toBe('80px');
  });

  it('代码源码沿用原块内边距、圆角和背景，长行预留滚动条高度', () => {
    const f = fixture('pre');
    f.anchor.dataset.sourceToken = 'fence';
    f.anchor.style.cssText += 'padding: 8px 12px; border: 1px solid; border-radius: 3px;';
    Object.defineProperty(f.textarea, 'scrollWidth', { value: 1200 });
    Object.defineProperty(f.textarea, 'clientWidth', { value: 600 });
    Object.defineProperty(f.textarea, 'offsetHeight', { value: 200 });
    Object.defineProperty(f.textarea, 'clientHeight', { value: 183 });
    const style = layoutPreviewEditor(f.container, f.preview, f.anchor, f.textarea, false);
    expect(style.backgroundColor).toBe('var(--code-block-bg)');
    expect(style['--preview-code-padding']).toBe('9px 13px 9px 13px');
    expect(style.borderRadius).toBe('3px');
    expect(f.textarea.style.paddingLeft).toBe('13px');
    expect(f.textarea.style.height).toBe('217px');
  });

  it('图表源码也使用等宽字体，重复利用输入层不会把代码内边距带入正文', () => {
    const f = fixture('div');
    f.anchor.dataset.sourceToken = 'fence';
    const codeStyle = layoutPreviewEditor(f.container, f.preview, f.anchor, f.textarea, false);
    expect(codeStyle.fontFamily).toContain('Consolas');
    expect(f.textarea.style.paddingLeft).toBe('13px');
    f.anchor.dataset.sourceToken = 'paragraph_open';
    const textStyle = layoutPreviewEditor(f.container, f.preview, f.anchor, f.textarea, false);
    expect(textStyle.fontFamily).toBe('serif');
    expect(f.textarea.style.paddingLeft).toBe('');
  });

  it('退出编辑清理占位，保留正文已有样式', () => {
    const f = fixture();
    f.anchor.classList.add('preview-editing-block');
    f.anchor.style.setProperty('--preview-edit-height', '200px');
    const fragment = document.createElement('strong');
    fragment.classList.add('preview-editing-fragment');
    fragment.style.setProperty('--preview-edit-width', '64px');
    f.anchor.append(fragment);
    clearPreviewEditingLayout(f.preview);
    expect(f.anchor.classList.contains('preview-editing-block')).toBe(false);
    expect(f.anchor.style.getPropertyValue('--preview-edit-height')).toBe('');
    expect(f.anchor.style.fontSize).toBe('20px');
    expect(fragment.hasAttribute('style')).toBe(false);
  });

  it('可见编辑框和已卸载控件不改变用户滚动位置', () => {
    const f = fixture();
    bounds(f.preview, 0, 600, 500);
    bounds(f.textarea, 30, 600, 100);
    revealPreviewEditorCaret(f.preview, f.textarea);
    expect(f.preview.scrollTop).toBe(0);
    f.textarea.remove();
    revealPreviewEditorCaret(f.preview, f.textarea);
    expect(f.preview.scrollTop).toBe(0);
  });

  it('长编辑框按光标测量位置向下滚动，不执行用户输入的 HTML', () => {
    const f = fixture();
    f.textarea.style.lineHeight = '28px';
    f.textarea.value = '<img src=x>\n最后一行';
    f.textarea.selectionStart = f.textarea.value.length;
    bounds(f.preview, 0, 600, 300);
    bounds(f.textarea, 0, 600, 600);
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function(
      this: HTMLElement,
    ) {
      const top = this.tagName === 'SPAN' ? 560 : 0;
      return { top, bottom: top + 28, left: 0, right: 600, width: 600, height: 28,
        x: 0, y: top, toJSON: () => ({}) };
    });
    revealPreviewEditorCaret(f.preview, f.textarea);
    expect(f.preview.scrollTop).toBe(304);
    expect(document.body.querySelector('img, div')).toBeNull();
  });

  it('向上移动光标可恢复顶部，默认行高不会产生负数滚动', () => {
    const f = fixture();
    f.preview.scrollTop = 300;
    bounds(f.preview, 0, 600, 300);
    bounds(f.textarea, -300, 600, 600);
    revealPreviewEditorCaret(f.preview, f.textarea);
    expect(f.preview.scrollTop).toBe(0);
  });

  it('编辑层坐标尚未更新时按正文锚点定位，不被上一帧可见位置提前返回', () => {
    const f = fixture();
    f.preview.scrollTop = 500;
    bounds(f.preview, 0, 600, 300);
    bounds(f.textarea, 30, 600, 100);
    bounds(f.anchor, -500, 600, 100);
    revealPreviewEditorCaret(f.preview, f.textarea, f.anchor);
    expect(f.preview.scrollTop).toBe(0);
  });
});
