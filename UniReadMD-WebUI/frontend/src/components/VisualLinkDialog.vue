<script setup lang='ts'>
import { reactive } from 'vue';
import type { useVisualPreviewEditing } from '../features/workspace/visual-preview-editing';
const props = defineProps<{ editor: ReturnType<typeof useVisualPreviewEditing> }>();
const model = reactive(props.editor);
</script>

<template>
  <div v-if='model.linkVisible' class='visual-link-backdrop' @mousedown.self='model.cancelLink'>
    <form
      class='visual-link-dialog' role='dialog' aria-label='插入链接'
      @submit.prevent='model.applyLink' @keydown.esc.prevent='model.cancelLink'
    >
      <label>链接地址
        <input v-model='model.linkUrl' placeholder='https:// 或相对路径' required />
      </label>
      <p v-if='model.linkError' role='alert'>{{ model.linkError }}</p>
      <div>
        <button type='button' @click='model.cancelLink'>取消</button>
        <button type='submit'>确定</button>
      </div>
    </form>
  </div>
</template>

<style scoped>
.visual-link-backdrop {
  position: fixed;
  inset: 0;
  z-index: 1200;
  display: grid;
  place-items: center;
  background: rgb(0 0 0 / 20%);
}
.visual-link-dialog {
  width: min(400px, calc(100vw - 40px));
  padding: 20px;
  border: 1px solid var(--line);
  border-radius: 8px;
  color: var(--text);
  background: var(--panel);
  box-shadow: 0 9px 24px rgb(0 0 0 / 15%);
}
label { display: grid; gap: 8px; }
input { padding: 8px; background: var(--panel); color: inherit; border: 1px solid var(--line); }
form > div { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
</style>
