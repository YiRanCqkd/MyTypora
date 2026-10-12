<script setup lang='ts'>
import { nextTick, reactive } from 'vue';
import type { usePreviewInsertMenu } from '../features/workspace/preview-insert-menu';
import type { PreviewInsertKind } from '../features/editor/preview-insertion';
import { contextMenuKeys } from '../features/workspace/context-menu-position';

const props = defineProps<{ menu: ReturnType<typeof usePreviewInsertMenu> }>();
const model = reactive(props.menu);
const headingNames = ['一级标题', '二级标题', '三级标题', '四级标题', '五级标题', '六级标题'];
const blocks: Array<[PreviewInsertKind, string]> = [
  ['paragraph', '段落'], ['code', '代码块'], ['math', '公式块'], ['quote', '引用'],
  ['unordered', '无序列表'], ['ordered', '有序列表'], ['task', '任务列表'], ['rule', '水平分割线'],
];
async function show(section: 'main' | 'headings' | 'table') {
  model.section = section;
  await nextTick();
  await model.reposition();
  document.querySelector<HTMLElement>('.preview-insert-menu input, .preview-insert-menu button')
    ?.focus({ preventScroll: true });
}
function menuKeys(event: KeyboardEvent) {
  if (event.key === 'Tab') { model.close(); return; }
  if (event.key === 'ArrowLeft' && model.section !== 'main') {
    event.preventDefault();
    void show('main');
    return;
  }
  contextMenuKeys(event);
}
</script>

<template>
  <div
    v-if='model.visible' class='preview-insert-menu typora-context-menu' :style='model.position'
    role='menu' aria-label='工作区插入菜单' @keydown='menuKeys'
  >
    <template v-if='model.section === "main"'>
      <button type='button' role='menuitem' @click='show("headings")'>
        标题<span aria-hidden='true'>›</span>
      </button>
      <button type='button' role='menuitem' @click='show("table")'>表格…</button>
      <button type='button' role='menuitem' @click='model.image'>图像…</button>
      <hr role='separator' />
      <button
        v-for='[kind, name] in blocks' :key='kind' type='button' role='menuitem'
        @click='model.insert(kind)'
      >
        {{ name }}
      </button>
    </template>
    <template v-else-if='model.section === "headings"'>
      <button type='button' role='menuitem' @click='show("main")'>‹ 返回</button>
      <hr role='separator' />
      <button
        v-for='(name, index) in headingNames' :key='name' type='button' role='menuitem'
        @click='model.insert(("heading" + (index + 1)) as PreviewInsertKind)'
      >
        {{ name }}
      </button>
    </template>
    <form v-else @submit.prevent='model.insert("table")'>
      <button type='button' @click='show("main")'>‹ 返回</button>
      <label>数据行数
        <input v-model.number='model.rows' type='number' min='1' max='20' required />
      </label>
      <label>列数
        <input v-model.number='model.columns' type='number' min='1' max='12' required />
      </label>
      <button type='submit'>插入表格</button>
    </form>
  </div>
</template>

<style scoped>
.preview-insert-menu {
  z-index: 1100;
}
button span { float: right; }
label { display: flex; align-items: center; gap: 8px; padding: 8px 24px; }
input {
  width: 70px;
  color: inherit;
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 3px;
  padding: 3px 6px;
  font: inherit;
}
</style>
