import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import { Transaction } from '@codemirror/state';
import type { WorkbenchDomains } from '../editor/workbench-domains';
import { findPreviewImageSource, buildImageResizeSource } from '../editor/image-sizing';
import type { ImageSource } from '../editor/image-sizing';

export function usePreviewImageSizing(domains: WorkbenchDomains) {
  const { store, runtime, previewHost, errorMessage } = domains.editor.foundation;
  const sourceMode = domains.settings.foundation().sourceMode;
  const imageMenu = ref(false);
  const imageMenuPosition = ref({ left: '0px', top: '0px' });
  const imagePercent = ref<number | string>(100);
  const imageSizeError = ref('');
  const imageNaturalWidth = ref(0);
  const imageNaturalHeight = ref(0);
  const imageScaleAvailable = computed(() => imageNaturalWidth.value > 0);
  let selected: { range: ImageSource; source: string; path: string } | null = null;
  let scrollFrame = 0;
  let scrollReady = false;

  function closeImageMenu() {
    imageMenu.value = false;
    imageSizeError.value = '';
    selected = null;
    cancelAnimationFrame(scrollFrame);
    scrollReady = false;
  }
  function openImageMenu(event: MouseEvent) {
    const image = event.target instanceof Element ? event.target.closest('img') : null;
    const host = previewHost.value;
    if (!(image instanceof HTMLImageElement) || !host?.contains(image) || sourceMode.value) {
      closeImageMenu();
      return;
    }
    event.preventDefault();
    event.stopImmediatePropagation();
    runtime.flushVisualEdits?.();
    domains.editor.preview().closePreviewInlineEditor();
    domains.editor.table().closePreviewTableCellEditorAndCommit();
    domains.editor.preview().closePreviewTableMenu();
    domains.editor.preview().closeSourceContextMenu();
    const source = runtime.editorView?.state.doc.toString() || store.content;
    const range = findPreviewImageSource(source, image, host,
      domains.markdown.feature().markdown, store.resourceBaseUrl);
    if (!range) {
      closeImageMenu();
      errorMessage.value = '无法准确定位这张图片，请结束当前编辑后重试，或在源码中调整尺寸。';
      return;
    }
    selected = { range, source, path: store.filePath };
    imageNaturalWidth.value = image.naturalWidth;
    imageNaturalHeight.value = image.naturalHeight;
    imagePercent.value = image.naturalWidth
      ? Math.round((parseFloat(image.style.width) || image.width)
        / image.naturalWidth * 10000) / 100
      : 100;
    imageSizeError.value = '';
    imageMenuPosition.value = {
      left: `${Math.max(8, Math.min(event.clientX, window.innerWidth - 248))}px`,
      top: `${Math.max(8, Math.min(event.clientY, window.innerHeight - 362))}px`,
    };
    imageMenu.value = true;
    // 右键前浏览器自动滚入图片的 scroll 事件可能延后到菜单打开后才派发。
    scrollFrame = requestAnimationFrame(() => {
      scrollFrame = requestAnimationFrame(() => { scrollReady = true; });
    });
    void nextTick(() => document.querySelector<HTMLElement>('.image-size-menu button')
      ?.focus({ preventScroll: true }));
  }
  function applyImageScale(percent: number | null) {
    const view = runtime.editorView;
    if (!selected || !view || store.filePath !== selected.path
      || view.state.doc.toString() !== selected.source) {
      closeImageMenu();
      return;
    }
    try {
      const { from, to } = selected.range;
      const insert = buildImageResizeSource(selected.source, selected.range,
        imageNaturalWidth.value, percent);
      closeImageMenu();
      view.dispatch({ changes: { from, to, insert },
        annotations: Transaction.userEvent.of('input.image-size') });
      previewHost.value?.focus({ preventScroll: true });
    } catch (error) {
      imageSizeError.value = error instanceof Error ? error.message : String(error);
    }
  }
  function dismiss(event: Event) {
    if (event.type === 'scroll' && !scrollReady) return;
    const target = event.target;
    if (target instanceof Element && target.closest('.image-size-menu')) return;
    closeImageMenu();
  }
  function keydown(event: KeyboardEvent) {
    if (!imageMenu.value || event.key !== 'Escape') return;
    event.preventDefault();
    event.stopImmediatePropagation();
    closeImageMenu();
    previewHost.value?.focus({ preventScroll: true });
  }
  watch(() => [store.content, store.filePath, sourceMode.value], closeImageMenu, { flush: 'sync' });
  onMounted(() => {
    document.addEventListener('contextmenu', openImageMenu, true);
    document.addEventListener('pointerdown', dismiss, true);
    document.addEventListener('scroll', dismiss, true);
    document.addEventListener('wheel', dismiss, true);
    window.addEventListener('resize', closeImageMenu);
    window.addEventListener('blur', closeImageMenu);
    window.addEventListener('keydown', keydown, true);
  });
  onBeforeUnmount(() => {
    cancelAnimationFrame(scrollFrame);
    document.removeEventListener('contextmenu', openImageMenu, true);
    document.removeEventListener('pointerdown', dismiss, true);
    document.removeEventListener('scroll', dismiss, true);
    document.removeEventListener('wheel', dismiss, true);
    window.removeEventListener('resize', closeImageMenu);
    window.removeEventListener('blur', closeImageMenu);
    window.removeEventListener('keydown', keydown, true);
  });
  return { imageMenu, imageMenuPosition, imagePercent, imageSizeError, imageNaturalWidth,
    imageNaturalHeight, imageScaleAvailable, applyImageScale, closeImageMenu };
}
