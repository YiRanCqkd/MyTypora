<script setup lang='ts'>
import { nextTick, onBeforeUnmount, onMounted, reactive, ref, watch } from 'vue';
import ImageSizeMenu from './ImageSizeMenu.vue';
import PreviewInsertMenu from './PreviewInsertMenu.vue';
import VisualLinkDialog from './VisualLinkDialog.vue';
import type { ParityWorkflows } from '../features/workspace/parity-workflows';
import '../features/workspace/parity-workflows.css';

const props = defineProps<{ workflows: ParityWorkflows }>();
const model = reactive(props.workflows);
const firstInput = ref<HTMLInputElement | null>(null);
const pagePlaceholder = '{{page}}';
const pagesPlaceholder = '{{pages}}';
const toolbar = ref<HTMLElement | null>(null);
const moreButton = ref<HTMLButtonElement | null>(null);
const compact = ref(false);
const moreOpen = ref(false);
let toolbarObserver: ResizeObserver | undefined;
watch(() => model.documentToolsVisible, (visible) => {
  if (!visible) moreOpen.value = false;
});
function dismissTools(event: PointerEvent) {
  if (event.target instanceof Node && !toolbar.value?.contains(event.target)) {
    moreOpen.value = false;
  }
}
function toolbarKeydown(event: KeyboardEvent) {
  if (event.key !== 'Escape' || !moreOpen.value) return;
  event.preventDefault();
  event.stopPropagation();
  moreOpen.value = false;
  moreButton.value?.focus();
}
onMounted(() => {
  if (typeof ResizeObserver !== 'undefined' && toolbar.value) {
    toolbarObserver = new ResizeObserver(([entry]) => {
      compact.value = entry.contentRect.width > 0 && entry.contentRect.width < 840;
      if (!compact.value) moreOpen.value = false;
    });
    toolbarObserver.observe(toolbar.value);
  }
  document.addEventListener('pointerdown', dismissTools, true);
});
onBeforeUnmount(() => {
  toolbarObserver?.disconnect();
  document.removeEventListener('pointerdown', dismissTools, true);
});
let returnFocus: HTMLElement | null = null;
watch(() => model.dialog, async (value, previous) => {
  if (value && !previous) returnFocus = document.activeElement as HTMLElement;
  await nextTick();
  if (value) {
    const firstControl = firstInput.value
      || document.querySelector<HTMLElement>('.workbench-dialog button:not(:disabled)');
    firstControl?.focus();
  }
  else returnFocus?.focus();
});

function trapFocus(event: KeyboardEvent) {
  if (event.key !== 'Tab') return;
  const host = event.currentTarget as HTMLElement;
  const controls = [...host.querySelectorAll<HTMLElement>(
    'button:not(:disabled), input, select, textarea',
  )];
  const first = controls[0];
  const last = controls.at(-1);
  if (event.shiftKey && document.activeElement === first) {
    event.preventDefault(); last?.focus();
  } else if (!event.shiftKey && document.activeElement === last) {
    event.preventDefault(); first?.focus();
  }
}
</script>

<template>
  <ImageSizeMenu :sizing='props.workflows.imageSizing' />
  <PreviewInsertMenu :menu='props.workflows.previewInsertMenu' />
  <VisualLinkDialog :editor='props.workflows.visualEditing' />
  <div
    v-show='model.documentToolsVisible' ref='toolbar' class='document-tools' :class='{ compact }'
    role='toolbar' aria-label='文档工具' @keydown='toolbarKeydown'
  >
    <button type='button' @click='model.dialog = "quick"'>快速打开</button>
    <button type='button' @click='model.collectImages'>整理图片</button>
    <button type='button' @click='model.insertLocalImages()'>插入图片</button>
    <div
      v-if='!compact || moreOpen' id='document-secondary-tools' class='document-tools-secondary'
      role='group' aria-label='更多文档工具' @click='compact && (moreOpen = false)'
    >
      <button type='button' @click='model.insertLocalImages(true)'>粘贴截图</button>
      <button type='button' @click='model.setImageRoot'>图片根目录</button>
      <button type='button' @click='model.showDrafts'>历史草稿</button>
      <button type='button' :disabled='!model.pandocAvailable' @click='model.importDocument'>
        导入文档
      </button>
      <button type='button' @click='model.dialog = "pdf"'>PDF 设置</button>
      <button type='button' @click='model.dialog = "export"'>导出设置</button>
    </div>
    <div class='document-export-tools'>
      <select v-model='model.selectedExport' aria-label='导出格式'>
        <option value='html'>HTML</option><option value='pdf'>PDF</option>
        <option value='png'>PNG</option><option value='jpeg'>JPEG</option>
        <option value='docx'>Word</option><option value='odt'>ODT</option>
        <option value='rtf'>RTF</option><option value='epub'>EPUB</option>
        <option value='revealjs'>RevealJS 幻灯片</option><option value='mediawiki'>MediaWiki</option>
        <option value='rst'>reStructuredText</option><option value='textile'>Textile</option>
        <option value='opml'>OPML 大纲</option>
        <option value='latex'>LaTeX</option>
      </select>
      <button type='button' :disabled='!model.canExportSelected' @click='model.exportSelected'>
        导出
      </button>
    </div>
    <button
      v-if='compact' ref='moreButton' type='button' aria-label='更多文档工具'
      aria-controls='document-secondary-tools' :aria-expanded='moreOpen'
      @click='moreOpen = !moreOpen'
    >
      更多 ▾
    </button>
  </div>
  <div v-if='model.notice' class='document-notice' role='status'>
    {{ model.notice }}
    <button type='button' aria-label='关闭提示' @click='model.notice = ""'>×</button>
  </div>
  <div v-if='model.externalChanged' class='document-notice' role='status'>
    文件已在外部修改或移动。当前编辑内容已保留。
    <button type='button' @click='model.reloadCurrent'>重新载入</button>
    <button type='button' @click='model.externalChanged = false'>稍后处理</button>
  </div>
  <div
    v-if='model.fileMenu && model.selectedRow' class='workspace-context-menu'
    :style='model.fileMenuPosition' role='menu' aria-label='文件操作菜单'
  >
    <template v-if='model.selectedRow.type === "file"'>
      <button role='menuitem' @click='model.openFileAt(model.selectedRow.path)'>打开</button>
      <button role='menuitem' @click='model.requestEntry("duplicate")'>复制文件</button>
    </template>
    <template v-else>
      <button role='menuitem' @click='model.requestEntry("create")'>新建 Markdown 文件</button>
      <button role='menuitem' @click='model.requestEntry("mkdir")'>新建文件夹</button>
    </template>
    <template v-if='model.selectedRow.depth > 0'>
      <button role='menuitem' @click='model.requestEntry("rename")'>重命名</button>
      <button role='menuitem' @click='model.moveEntry'>移动到…</button>
      <button role='menuitem' @click='model.trashEntry'>移入回收站</button>
      <button role='menuitem' @click='model.showNativeFileMenu'>复制路径 / 资源管理器…</button>
    </template>
  </div>
  <div
    v-if='model.dialog' class='workbench-dialog-backdrop' @click.self='model.dialog = ""'
  >
    <section
      class='workbench-dialog' role='dialog' aria-modal='true'
      aria-labelledby='workbench-dialog-title'
      @keydown='trapFocus'
    >
      <header>
        <h2 id='workbench-dialog-title'>
          {{ { quick: '快速打开', drafts: '恢复历史草稿', name: '文件操作', pdf: 'PDF 导出设置',
               export: '导出设置' }
            [model.dialog] }}
        </h2>
        <button type='button' aria-label='关闭对话框' @click='model.dialog = ""'>×</button>
      </header>
      <p v-if='model.errorMessage' class='error-banner' role='alert'>{{ model.errorMessage }}</p>
      <template v-if='model.dialog === "quick"'>
        <input
          ref='firstInput' v-model='model.quickQuery' aria-label='查找文件名'
          placeholder='输入文件名，支持不连续字符匹配'
          @keydown.enter='model.quickFiles[0] && model.openFileAt(model.quickFiles[0])'
        />
        <div class='dialog-results'>
          <button v-for='file in model.quickFiles' :key='file' @click='model.openFileAt(file)'>
            {{ file }}
          </button>
          <p v-if='!model.quickFiles.length'>没有匹配的文件，请先打开文件夹或文档。</p>
        </div>
      </template>
      <template v-else-if='model.dialog === "drafts"'>
        <p>恢复为未命名副本，保存时可选择新的文件名。</p>
        <div class='dialog-results'>
          <button
            v-for='draft in model.drafts' :key='draft.id' @click='model.restoreDraft(draft.id)'
          >
            {{ new Date(draft.savedAt).toLocaleString('zh-CN') }}
            {{ draft.filePath || draft.title || '未命名文档' }}
          </button>
          <p v-if='!model.drafts.length'>暂无可恢复草稿。</p>
        </div>
      </template>
      <form
        v-else-if='model.dialog === "name"'
        @submit.prevent='model.applyFileOperation(model.entryAction)'
      >
        <label>名称
          <input ref='firstInput' v-model='model.entryName' required maxlength='240' />
        </label>
        <button type='submit' :disabled='model.busy'>确定</button>
      </form>
      <form v-else-if='model.dialog === "pdf"' @submit.prevent='model.savePdfOptions'>
        <label>纸张大小
          <select v-model='model.pdfOptions.pageSize'>
            <option>A4</option><option>A5</option><option>Letter</option><option>Legal</option>
            <option value='Custom'>自定义</option>
          </select>
        </label>
        <template v-if='model.pdfOptions.pageSize === "Custom"'>
          <label>纸张宽度（毫米）
            <input
              v-model.number='model.pdfOptions.widthMm' type='number' min='50' max='1000' required
            />
          </label>
          <label>纸张高度（毫米）
            <input
              v-model.number='model.pdfOptions.heightMm' type='number' min='50' max='1000' required
            />
          </label>
        </template>
        <label>页边距（毫米）
          <input
            ref='firstInput' v-model.number='model.pdfOptions.marginMm' type='number'
            required min='0' max='50' step='1'
          />
        </label>
        <label><input
          type='checkbox' :checked='Boolean(model.pdfOptions.marginsMm)'
          @change='model.setIndependentMargins'
        />分别设置四侧页边距</label>
        <template v-if='model.pdfOptions.marginsMm'>
          <label v-for='side in (["top", "bottom", "left", "right"] as const)' :key='side'>
            {{ { top: '上', bottom: '下', left: '左', right: '右' }[side] }}侧（毫米）
            <input
              v-model.number='model.pdfOptions.marginsMm[side]' type='number'
              min='0' max='50' required
            />
          </label>
        </template>
        <label><input v-model='model.pdfOptions.landscape' type='checkbox' />横向纸张</label>
        <label><input v-model='model.pdfOptions.bookmarks' type='checkbox' />生成 PDF 标题书签</label>
        <label>页眉<input v-model='model.pdfOptions.header' maxlength='200' /></label>
        <label>页脚<input v-model='model.pdfOptions.footer' maxlength='200' /></label>
        <p>
          页码使用 <code v-text='pagePlaceholder' />，总页数使用
          <code v-text='pagesPlaceholder' />。
        </p>
        <button type='submit'>保存设置</button>
      </form>
      <form v-else-if='model.dialog === "export"' @submit.prevent='model.saveExportOptions'>
        <label>标题
          <input ref='firstInput' v-model='model.exportOptions.title' maxlength='200' />
        </label>
        <label>作者<input v-model='model.exportOptions.author' maxlength='200' /></label>
        <label>配置预设
          <select @change='model.applyExportPreset($event)'>
            <option value='default'>默认</option><option value='article'>长文报告</option>
            <option value='ebook'>电子书</option>
          </select>
        </label>
        <label><input v-model='model.exportOptions.tableOfContents' type='checkbox' />生成目录</label>
        <label>目录标题层级
          <input v-model.number='model.exportOptions.tocDepth' type='number' min='1' max='6' />
        </label>
        <label>EPUB 章节拆分层级
          <input
            v-model.number='model.exportOptions.epubChapterLevel'
            type='number' min='1' max='6'
          />
        </label>
        <label>文档语言
          <input
            v-model='model.exportOptions.language'
            pattern='[a-z]{2,3}(-[a-zA-Z0-9]{2,8})*' required
          />
        </label>
        <label>导出配色
          <select v-model='model.exportOptions.theme'>
            <option value='current'>当前配色</option><option value='light'>浅色</option>
            <option value='dark'>深色</option>
          </select>
        </label>
        <label><input v-model='model.exportOptions.includeStyles' type='checkbox' />包含预览样式</label>
        <label>
          <input v-model='model.exportOptions.headingPageBreak' type='checkbox' />一级标题前分页
        </label>
        <label>图片宽度（像素）
          <input
            v-model.number='model.exportOptions.imageWidth' type='number'
            min='256' max='4096' required
          />
        </label>
        <label>JPEG 质量（1–100）
          <input
            v-model.number='model.exportOptions.jpegQuality' type='number'
            min='1' max='100' required
          />
        </label>
        <label>Word 参考文档
          <button type='button' @click='model.chooseExportResource("referenceDoc")'>
            选择 .docx
          </button>
          <button type='button' @click='model.exportOptions.referenceDoc = ""'>清除</button>
        </label>
        <p v-if='model.exportOptions.referenceDoc'>{{ model.exportOptions.referenceDoc }}</p>
        <label>EPUB 封面
          <button type='button' @click='model.chooseExportResource("epubCover")'>选择图片</button>
          <button type='button' @click='model.exportOptions.epubCover = ""'>清除</button>
        </label>
        <p v-if='model.exportOptions.epubCover'>{{ model.exportOptions.epubCover }}</p>
        <p>标题、作者和语言用于文档导出；配色与样式用于 HTML、PDF 和图片。</p>
        <button type='submit'>保存设置</button>
      </form>
    </section>
  </div>
</template>
