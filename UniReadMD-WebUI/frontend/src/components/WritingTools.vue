<script setup lang='ts'>
import { nextTick, reactive, ref, watch } from 'vue';
import type { useWritingEnhancements } from '../features/settings/writing-enhancements';
const props = defineProps<{ model: ReturnType<typeof useWritingEnhancements> }>();
const model = reactive(props.model);
const panel = ref<HTMLElement | null>(null);
let returnFocus: HTMLElement | null = null;
watch(() => model.visible, async (visible) => {
  if (visible) returnFocus = document.activeElement as HTMLElement;
  await nextTick();
  if (visible) panel.value?.querySelector('select')?.focus();
  else returnFocus?.focus();
});
function keyboard(event: KeyboardEvent) {
  if (event.key === 'Escape') { event.stopPropagation(); model.visible = false; }
  if (event.key !== 'Tab') return;
  const controls = [...(panel.value?.querySelectorAll<HTMLElement>('button, input, select') || [])];
  if (event.shiftKey && document.activeElement === controls[0]) {
    event.preventDefault(); controls.at(-1)?.focus();
  } else if (!event.shiftKey && document.activeElement === controls.at(-1)) {
    event.preventDefault(); controls[0]?.focus();
  }
}
</script>

<template>
  <div v-if='model.visible' class='workbench-dialog-backdrop'>
    <section
      ref='panel' class='workbench-dialog writing-tools' role='dialog' aria-modal='true'
      aria-label='写作设置' @keydown='keyboard'
    >
      <header>
        <h2>写作设置</h2>
        <button
          class='dialog-close' type='button' aria-label='关闭写作设置'
          @click='model.visible = false'
        >
          ×
        </button>
      </header>
      <label>状态栏统计
        <select v-model='model.preferences.statUnit' autofocus>
          <option value='words'>字数</option><option value='characters'>字符数</option>
        </select>
      </label>
      <label>阅读速度（字／分钟）
        <input v-model.number='model.preferences.readingSpeed' type='number' min='50' max='2000' />
      </label>
      <label><input v-model='model.preferences.equationNumbering' type='checkbox' />公式自动编号</label>
      <label>浅色写作主题
        <select v-model='model.preferences.lightTheme'>
          <option value='paper'>白纸</option><option value='book'>书籍</option>
          <option value='sepia'>暖纸</option>
        </select>
      </label>
      <label>深色写作主题
        <select v-model='model.preferences.darkTheme'>
          <option value='graphite'>石墨</option><option value='midnight'>午夜</option>
        </select>
      </label>
      <p>设置自动保存。选中文字时，状态栏显示选区统计。</p>
      <footer><button type='button' @click='model.visible = false'>完成</button></footer>
    </section>
  </div>
</template>

<style>
.writing-tools button {
  font: inherit;
  color: var(--text);
  background: var(--sidebar);
  border: 1px solid var(--line);
  border-radius: 4px;
  padding: 6px 12px;
  cursor: pointer;
}
.writing-tools button:hover { background: var(--item-hover-bg); }
.writing-tools .dialog-close { padding: 4px 8px; font-size: 20px; line-height: 1; }
.writing-tools footer { display: flex; justify-content: flex-end; margin-top: 20px; }
html[data-writing-theme='book'] .preview {
  font-family: Georgia, 'Noto Serif SC', '宋体', serif;
  line-height: 1.9;
}
html[data-writing-theme='book'] .preview h1,
html[data-writing-theme='book'] .preview h2 { border-bottom: 1px solid var(--line); }
html[data-theme='light'][data-writing-theme='sepia'] {
  --panel: #faf4e8;
  --sidebar: #f0e8d8;
  --bg: #f0e8d8;
  --text: #4b4034;
  --line: #d7cbb8;
}
html[data-theme='dark'][data-writing-theme='midnight'] {
  --panel: #141e2e;
  --sidebar: #101827;
  --bg: #101827;
  --text: #dae5f4;
  --line: #334156;
}
</style>
