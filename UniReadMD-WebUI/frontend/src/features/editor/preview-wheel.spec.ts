import { afterEach, describe, expect, it, vi } from 'vitest';
import { ref } from 'vue';
import { usePreviewWheel } from './preview-wheel';

afterEach(() => { document.body.replaceChildren(); });

function fixture() {
  const host = document.createElement('section');
  const preview = document.createElement('article');
  const overlay = document.createElement('section');
  overlay.className = 'preview-inline-editor';
  const textarea = document.createElement('textarea');
  overlay.append(textarea);
  host.append(preview, overlay);
  document.body.append(host);
  const stepContentZoom = vi.fn();
  const input = { editorPreviewHost: ref(host), previewHost: ref(preview),
    previewInlineEditVisible: ref(true), previewTableCellEditVisible: ref(false),
    stepContentZoom } as unknown as Parameters<typeof usePreviewWheel>[0];
  const { handleGlobalWheel } = usePreviewWheel(input);
  host.addEventListener('wheel', handleGlobalWheel);
  const wheel = (target: HTMLElement, options: WheelEventInit = {}) => {
    const event = new WheelEvent('wheel', { bubbles: true, cancelable: true, deltaY: 60,
      ...options });
    target.dispatchEvent(event);
    return event;
  };
  return { host, preview, overlay, textarea, input, stepContentZoom, wheel };
}

describe('预览滚轮统一处理', () => {
  it('普通编辑框滚轮只转发一次到预览滚动区', () => {
    const f = fixture();
    f.preview.scrollTop = 100;
    expect(f.wheel(f.textarea).defaultPrevented).toBe(true);
    expect(f.preview.scrollTop).toBe(160);
    expect(f.stepContentZoom).not.toHaveBeenCalled();
  });

  it('Ctrl 和 Meta 缩放不混入普通滚动，正负方向一致', () => {
    const f = fixture();
    f.preview.scrollTop = 100;
    f.wheel(f.textarea, { ctrlKey: true });
    f.wheel(f.preview, { metaKey: true, deltaY: -60 });
    expect(f.stepContentZoom.mock.calls).toEqual([['out'], ['in']]);
    expect(f.preview.scrollTop).toBe(100);
  });

  it('零纵向滚动和已处理事件不会造成额外缩放', () => {
    const f = fixture();
    f.wheel(f.textarea, { deltaY: 0, ctrlKey: true });
    f.textarea.addEventListener('wheel', (event) => event.preventDefault());
    f.wheel(f.textarea, { ctrlKey: true });
    expect(f.stepContentZoom).not.toHaveBeenCalled();
    expect(f.preview.scrollTop).toBe(0);
  });

  it('非编辑内容保留浏览器滚动，无编辑会话时不拦截', () => {
    const f = fixture();
    expect(f.wheel(f.preview).defaultPrevented).toBe(false);
    f.input.previewInlineEditVisible.value = false;
    expect(f.wheel(f.textarea).defaultPrevented).toBe(false);
  });

  it('表格编辑框也统一转发，缺少预览时安全退出', () => {
    const f = fixture();
    f.input.previewInlineEditVisible.value = false;
    f.input.previewTableCellEditVisible.value = true;
    f.overlay.className = 'preview-table-cell-editor';
    expect(f.wheel(f.textarea).defaultPrevented).toBe(true);
    expect(f.preview.scrollTop).toBe(60);
    f.input.previewHost.value = null;
    expect(f.wheel(f.textarea).defaultPrevented).toBe(false);
  });
});
