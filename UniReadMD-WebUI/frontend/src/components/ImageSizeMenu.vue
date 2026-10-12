<script setup lang='ts'>
import { reactive } from 'vue';
import type { usePreviewImageSizing } from '../features/workspace/preview-image-sizing';
import { contextMenuKeys } from '../features/workspace/context-menu-position';

const props = defineProps<{ sizing: ReturnType<typeof usePreviewImageSizing> }>();
const model = reactive(props.sizing);
function menuKeys(event: KeyboardEvent) {
  if (event.key === 'Tab') { model.closeImageMenu(); return; }
  contextMenuKeys(event);
}
</script>

<template>
  <div
    v-if='model.imageMenu' class='image-size-menu typora-context-menu'
    :style='model.imageMenuPosition'
    role='dialog' aria-label='图片显示比例' @keydown='menuKeys' @contextmenu.prevent
  >
    <div class='image-menu-caption'>图片显示比例</div>
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
    <hr role='separator' />
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
    <hr role='separator' />
    <button type='button' @click='model.applyImageScale(null)'>恢复自动大小</button>
  </div>
</template>

<style scoped>
.image-size-menu {
  z-index: 1100;
}
p, .image-menu-caption { margin: 4px 24px; font-size: 12px; color: var(--muted); }
.image-menu-caption { padding: 3px 0; }
.image-scale-hint { opacity: 0.75; line-height: 1.5; }
input {
  box-sizing: border-box;
  min-width: 0;
  padding: 3px 6px;
  border: 1px solid var(--line);
  border-radius: 4px;
  color: inherit;
  background: var(--panel);
  font: inherit;
}
input:focus-visible { outline: 1px solid var(--accent); }
label { display: flex; align-items: center; gap: 8px; padding: 5px 24px; }
input { width: 65px; }
</style>
