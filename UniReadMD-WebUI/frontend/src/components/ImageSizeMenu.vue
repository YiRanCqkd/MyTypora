<script setup lang='ts'>
import { reactive } from 'vue';
import type { usePreviewImageSizing } from '../features/workspace/preview-image-sizing';

const props = defineProps<{ sizing: ReturnType<typeof usePreviewImageSizing> }>();
const model = reactive(props.sizing);
function menuKeys(event: KeyboardEvent) {
  if (event.key === 'Tab') { model.closeImageMenu(); return; }
  if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)
    || event.target instanceof HTMLInputElement) return;
  const host = event.currentTarget as HTMLElement;
  const controls = [...host.querySelectorAll<HTMLElement>('button:not(:disabled), input')];
  const index = controls.indexOf(document.activeElement as HTMLElement);
  const next = event.key === 'Home' ? 0 : event.key === 'End' ? controls.length - 1
    : (index + (event.key === 'ArrowDown' ? 1 : -1) + controls.length) % controls.length;
  event.preventDefault();
  controls[next]?.focus();
}
</script>

<template>
  <div
    v-if='model.imageMenu' class='image-size-menu' :style='model.imageMenuPosition'
    role='dialog' aria-label='图片显示比例' @keydown='menuKeys' @contextmenu.prevent
  >
    <strong>图片显示比例</strong>
    <p v-if='model.imageScaleAvailable'>
      原图 {{ model.imageNaturalWidth }} × {{ model.imageNaturalHeight }} 像素
    </p>
    <p v-else>图片尚未加载完成，暂不能计算比例。</p>
    <div class='image-scale-presets'>
      <button
        v-for='percent in [25, 50, 75, 100, 150, 200]' :key='percent'
        type='button' :disabled='!model.imageScaleAvailable'
        @click='model.applyImageScale(percent)'
      >
        {{ percent }}%
      </button>
    </div>
    <form @submit.prevent='model.applyImageScale(Number(model.imagePercent))'>
      <label>自定义比例（%）
        <input
          v-model='model.imagePercent' type='number' min='1' max='400' step='0.01' required
          :disabled='!model.imageScaleAvailable'
        />
      </label>
      <button type='submit' :disabled='!model.imageScaleAvailable'>应用比例</button>
    </form>
    <p class='image-scale-hint'>以原图大小为基准，保持纵横比，最大不超过内容区。</p>
    <p v-if='model.imageSizeError' role='alert'>{{ model.imageSizeError }}</p>
    <button type='button' @click='model.applyImageScale(null)'>恢复自动大小</button>
  </div>
</template>

<style scoped>
.image-size-menu {
  position: fixed;
  z-index: 1100;
  box-sizing: border-box;
  width: min(240px, calc(100vw - 16px));
  max-height: calc(100vh - 16px);
  overflow: auto;
  padding: 12px;
  color: var(--text);
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 6px;
  box-shadow: 0 4px 20px rgb(0 0 0 / 15%);
  font-size: 12px;
}
p { margin: 8px 0; }
.image-scale-hint { opacity: 0.75; line-height: 1.5; }
.image-scale-presets { display: grid; grid-template-columns: repeat(3, 1fr); gap: 4px; }
button, input {
  box-sizing: border-box;
  min-width: 0;
  padding: 6px;
  border: 1px solid var(--line);
  border-radius: 4px;
  color: inherit;
  background: var(--panel);
  font: inherit;
}
button { cursor: pointer; }
button:hover { background: var(--item-hover-bg); }
button:disabled { opacity: 0.45; cursor: default; }
button:focus-visible, input:focus-visible { outline: 2px solid var(--accent); }
label { display: grid; gap: 5px; margin: 10px 0 6px; }
form button, .image-size-menu > button { width: 100%; }
</style>
