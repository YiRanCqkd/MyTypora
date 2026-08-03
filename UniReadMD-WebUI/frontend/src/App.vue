<script setup lang='ts'>
import { useWorkbench } from './features/editor/use-workbench';

const { activeHeadingLine,activeTab,applyPreviewTableAction,changeSpellCheckEnabled,
    changeSpellCheckLanguage,closeDocSearchPanel,closeWindow,contentCharCount,
    contentReadMinutes,contentZoomLabel,docReplaceQuery,docSearchCaseSensitive,docSearchInput,
    docSearchQuery,docSearchSummary,docSearchWholeWord,editorHost,editorPreviewHost,
    errorMessage,fileTreeRows,filesFilter,goToNextDocMatch,goToPrevDocMatch,
    handleDocSearchEnter,handleFileTreeContextMenu,handlePreviewClick,handlePreviewContextMenu,
    handlePreviewDoubleClick,handlePreviewInlineEditBlur,handlePreviewInlineEditCompositionEnd,
    handlePreviewInlineEditCompositionStart,handlePreviewInlineEditInput,
    handlePreviewInlineEditKeydown,handlePreviewInlineEditorWheel,handlePreviewScroll,
    handlePreviewTableCellEditBlur,handlePreviewTableCellEditKeydown,handleSourceContextAction,
    handleSourceContextMenu,handleWindowDragEnter,handleWindowDragLeave,handleWindowDragOver,
    handleWindowDrop,isDocumentLoading,isDragOverFile,jumpToHeading,layoutClassName,layoutStyle,
    minimizeWindow,openFileFromList,outlineDisplayItems,pinOutline,previewHost,
    previewInlineEditStyle,previewInlineEditTextarea,previewInlineEditTitle,
    previewInlineEditValue,previewInlineEditVisible,previewTableCellEditInput,
    previewTableCellEditStyle,previewTableCellEditValue,previewTableCellEditVisible,
    previewTableMenuStyle,previewTableMenuVisible,renderedHtml,replaceAllDocMatch,
    replaceCurrentDocMatch,resetContentZoom,searchText,setTab,showDocReplacePanel,
    showDocSearchPanel,showSpellPanel,sidebarSearchPlaceholder,sidebarSearchText,
    sourceContextMenuStyle,sourceContextMenuVisible,sourceMode,sourceToggleLabel,
    spellCheckEnabled,spellCheckLanguage,spellCheckLanguages,spellDictionaryLabel,
    startMainPreviewResize,startSidebarResize,statusLineLabel,statusWordLabel,store,
    toggleMaximizeWindow,togglePinOutline,toggleSourceMode,toggleTreeRow,windowTitle,
} = useWorkbench();
</script>

<template>
  <div
    class='layout'
    :class='layoutClassName'
    :style='layoutStyle'
    data-testid='app-layout'
    @dragenter='handleWindowDragEnter'
    @dragleave='handleWindowDragLeave'
    @dragover='handleWindowDragOver'
    @drop='handleWindowDrop'
  >
    <header class='top-titlebar'>
      <div class='title-left'>
        <button type='button' class='menu-icon' @click='setTab("files")'>
          <span></span>
          <span></span>
          <span></span>
        </button>
        <span class='title-name'>{{ windowTitle }}</span>
        <span v-if='store.isDirty' class='dirty-dot'>●</span>
      </div>
      <div class='title-actions'></div>
      <div class='window-actions'>
        <button type='button' aria-label='最小化' @click='minimizeWindow'>_</button>
        <button type='button' aria-label='最大化' @click='toggleMaximizeWindow'>□</button>
        <button type='button' aria-label='关闭' @click='closeWindow'>×</button>
      </div>
    </header>

    <div class='layout-body'>
      <aside class='sidebar' data-testid='sidebar'>
        <div class='sidebar-search-top'>
          <input
            v-model='sidebarSearchText'
            class='search-input'
            data-testid='sidebar-search-input'
            :placeholder='sidebarSearchPlaceholder'
          />
        </div>
        <header class='sidebar-header'>
          <button
            type='button'
            data-testid='tab-files'
            :class='{ active: activeTab === "files" }'
            @click='setTab("files")'
          >
            文件
          </button>
          <button
            type='button'
            data-testid='tab-outline'
            :class='{ active: activeTab === "outline" }'
            @click='setTab("outline")'
          >
            大纲
          </button>
        </header>

        <section v-if='activeTab === "files"' class='panel panel-files'>
          <div class='files-workbench'>
            <div class='files-tree' data-testid='files-tree-view'>
              <div
                v-for='row in fileTreeRows'
                :key='row.id'
                class='files-tree-row'
                :class='{ group: row.type === "dir" && row.depth === 0 }'
                :style='{ paddingLeft: `${row.depth * 16 + 8}px` }'
              >
                <button
                  type='button'
                  class='files-tree-toggle'
                  :class='{
                    group: row.type === "dir" && row.depth === 0,
                    hidden: !row.expandable,
                  }'
                  @click='toggleTreeRow(row)'
                >
                  {{ row.expanded ? '▾' : '▸' }}
                </button>
                <button
                  type='button'
                  class='files-tree-entry'
                  :title='row.label'
                  :class='{
                    dir: row.type === "dir",
                    file: row.type === "file",
                    group: row.type === "dir" && row.depth === 0,
                    active: row.filePath === store.filePath,
                  }'
                  @contextmenu.prevent='handleFileTreeContextMenu($event, row)'
                  @click='
                    row.type === "file" && row.filePath
                      ? openFileFromList(row.filePath)
                      : toggleTreeRow(row)
                  '
                >
                  {{ row.label }}
                </button>
              </div>
              <div v-if='fileTreeRows.length === 0' class='empty'>
                {{ filesFilter.trim().length > 0 ? '未找到匹配文档' : '暂无 Markdown 文件' }}
              </div>
            </div>
          </div>
        </section>

        <section v-else-if='activeTab === "outline"' class='panel'>
          <div v-if='outlineDisplayItems.length === 0' class='empty'>
            {{ searchText.trim().length > 0 ? '未找到匹配标题' : '当前文档无标题' }}
          </div>
          <button
            v-for='item in outlineDisplayItems'
            :key='`${item.line}-${item.text}`'
            type='button'
            data-testid='outline-item'
            class='outline-item'
            :class='[
              `outline-level-${item.level}`,
              { active: item.line === activeHeadingLine },
            ]'
            :style='{ paddingLeft: `${(item.level - 1) * 14 + 10}px` }'
            @click='jumpToHeading(item.line)'
          >
            <span class='outline-item-text'>{{ item.text }}</span>
          </button>
        </section>
        <footer class='sidebar-footer' aria-label='sidebar-footer'></footer>
      </aside>
      <div
        class='sidebar-resizer'
        data-testid='sidebar-resizer'
        role='separator'
        aria-orientation='vertical'
        @mousedown='startSidebarResize'
      ></div>

      <main class='main'>
        <section v-if='showDocSearchPanel' class='doc-search-panel' data-testid='doc-search-panel'>
          <div class='doc-search-row'>
            <input
              ref='docSearchInput'
              v-model='docSearchQuery'
              class='doc-search-input'
              data-testid='doc-search-query'
              placeholder='Search text'
              @keydown.enter.prevent='handleDocSearchEnter($event.shiftKey)'
            />
            <button
              type='button'
              class='doc-search-option'
              :class='{ active: docSearchCaseSensitive }'
              @click='docSearchCaseSensitive = !docSearchCaseSensitive'
            >
              Aa
            </button>
            <button
              type='button'
              class='doc-search-option'
              :class='{ active: docSearchWholeWord }'
              @click='docSearchWholeWord = !docSearchWholeWord'
            >
              W
            </button>
            <button type='button' class='doc-search-btn' @click='goToPrevDocMatch'>↑</button>
            <button type='button' class='doc-search-btn' @click='goToNextDocMatch'>↓</button>
            <button
              v-if='!showDocReplacePanel'
              type='button'
              class='doc-search-btn'
              data-testid='doc-open-replace'
              @click='showDocReplacePanel = true'
            >
              Replace
            </button>
            <button
              v-else
              type='button'
              class='doc-search-btn'
              data-testid='doc-close-replace'
              @click='showDocReplacePanel = false'
            >
              Search Only
            </button>
            <button
              type='button'
              class='doc-search-btn close'
              @click='closeDocSearchPanel'
            >
              ✕
            </button>
          </div>
          <div v-if='showDocReplacePanel' class='doc-search-row'>
            <input
              v-model='docReplaceQuery'
              class='doc-search-input'
              data-testid='doc-replace-query'
              placeholder='Replace with'
              @keydown.enter.prevent='replaceCurrentDocMatch'
            />
            <button type='button' class='doc-search-btn' @click='replaceCurrentDocMatch'>
              Replace
            </button>
            <button type='button' class='doc-search-btn' @click='replaceAllDocMatch'>
              Replace All
            </button>
          </div>
          <div class='doc-search-summary' data-testid='doc-search-summary'>
            {{ docSearchSummary }}
          </div>
        </section>
        <div v-if='errorMessage' class='error-banner'>{{ errorMessage }}</div>
        <section
          ref='editorPreviewHost'
          class='editor-preview'
          data-testid='editor-preview'
        >
          <section
            v-if='isDocumentLoading'
            class='document-loading-overlay'
            data-testid='document-loading-overlay'
          >
            <div class='document-loading-card'>加载中...</div>
          </section>
          <article
            id='write'
            ref='previewHost'
            class='preview'
            data-testid='preview-host'
            @click='handlePreviewClick'
            @dblclick='handlePreviewDoubleClick'
            @scroll='handlePreviewScroll'
            @contextmenu='handlePreviewContextMenu'
            v-html='renderedHtml'
          ></article>
          <section
            v-if='previewInlineEditVisible'
            class='preview-inline-editor'
            :style='previewInlineEditStyle'
            data-testid='preview-inline-editor'
          >
            <textarea
              ref='previewInlineEditTextarea'
              v-model='previewInlineEditValue'
              class='preview-inline-editor-input'
              :aria-label='previewInlineEditTitle'
              @input='handlePreviewInlineEditInput'
              @compositionstart='handlePreviewInlineEditCompositionStart'
              @compositionend='handlePreviewInlineEditCompositionEnd'
              @blur='handlePreviewInlineEditBlur'
              @keydown='handlePreviewInlineEditKeydown'
              @wheel='handlePreviewInlineEditorWheel'
            ></textarea>
          </section>
          <section
            v-if='previewTableCellEditVisible'
            class='preview-table-cell-editor'
            :style='previewTableCellEditStyle'
            data-testid='preview-table-cell-editor'
          >
            <input
              ref='previewTableCellEditInput'
              v-model='previewTableCellEditValue'
              class='preview-table-cell-editor-input'
              data-testid='preview-table-cell-editor-input'
              @blur='handlePreviewTableCellEditBlur'
              @keydown='handlePreviewTableCellEditKeydown'
            />
          </section>
          <section
            v-if='previewTableMenuVisible'
            class='preview-table-menu'
            :style='previewTableMenuStyle'
            data-testid='preview-table-menu'
          >
            <button type='button' @click='applyPreviewTableAction("insert-row-above")'>
              在上方插入行
            </button>
            <button type='button' @click='applyPreviewTableAction("insert-row-below")'>
              在下方插入行
            </button>
            <button type='button' @click='applyPreviewTableAction("delete-row")'>删除当前行</button>
            <button type='button' @click='applyPreviewTableAction("move-row-up")'>上移当前行</button>
            <button type='button' @click='applyPreviewTableAction("move-row-down")'>下移当前行</button>
            <div class='preview-table-menu-divider'></div>
            <button type='button' @click='applyPreviewTableAction("insert-col-left")'>
              在左侧插入列
            </button>
            <button type='button' @click='applyPreviewTableAction("insert-col-right")'>
              在右侧插入列
            </button>
            <button type='button' @click='applyPreviewTableAction("delete-col")'>删除当前列</button>
            <button type='button' @click='applyPreviewTableAction("move-col-left")'>左移当前列</button>
            <button type='button' @click='applyPreviewTableAction("move-col-right")'>右移当前列</button>
          </section>
          <section
            v-if='sourceContextMenuVisible'
            class='source-context-menu'
            :style='sourceContextMenuStyle'
            data-testid='source-context-menu'
          >
            <button type='button' @click='handleSourceContextAction("cut")'>剪切</button>
            <button type='button' @click='handleSourceContextAction("copy")'>复制</button>
            <button type='button' @click='handleSourceContextAction("paste")'>粘贴</button>
            <div class='source-context-menu-divider'></div>
            <button type='button' @click='handleSourceContextAction("select-all")'>全选</button>
            <button type='button' @click='handleSourceContextAction("search")'>查找</button>
          </section>
          <div
            class='main-preview-resizer'
            data-testid='main-preview-resizer'
            role='separator'
            aria-orientation='vertical'
            @mousedown='startMainPreviewResize'
          ></div>
          <div
            id='typora-source'
            ref='editorHost'
            class='editor-host'
            data-testid='editor-host'
            @contextmenu='handleSourceContextMenu'
          ></div>
        </section>
      </main>
    </div>

    <footer class='status-bar'>
      <div class='status-left'>
        <button
          id='outline-btn'
          type='button'
          class='status-action'
          :class='{ active: pinOutline }'
          @click='togglePinOutline'
        >
          {{ pinOutline ? '隐藏侧栏' : '显示侧栏' }}
        </button>
        <button
          id='toggle-sourceview-btn'
          type='button'
          class='status-action'
          :class='{ active: sourceMode }'
          @click='toggleSourceMode'
        >
          {{ sourceToggleLabel }}
        </button>
      </div>
      <div class='status-right'>
        <button
          id='footer-content-zoom'
          type='button'
          class='status-action status-zoom'
          title='按住 Ctrl + 鼠标滚轮缩放，单击重置为 100%'
          @click='resetContentZoom'
        >
          {{ contentZoomLabel }}
        </button>
        <button id='footer-word-count' type='button' class='status-action status-word'>
          {{ statusWordLabel }}
        </button>
        <span class='status-metric'>{{ contentReadMinutes }} min</span>
        <span class='status-metric'>{{ contentCharCount }} chars</span>
        <span class='status-metric'>{{ statusLineLabel }}</span>
      </div>
    </footer>
    <div
      v-if='showSpellPanel'
      class='spell-panel spell-panel-global'
      data-testid='spell-panel'
    >
      <label class='spell-option'>
        <input
          v-model='spellCheckEnabled'
          type='checkbox'
          @change='changeSpellCheckEnabled'
        />
        Enable
      </label>
      <label class='spell-option'>
        Language
        <select v-model='spellCheckLanguage' @change='changeSpellCheckLanguage'>
          <option
            v-for='item in spellCheckLanguages'
            :key='item'
            :value='item'
          >
            {{ item }}
          </option>
        </select>
      </label>
      <div class='spell-dictionary-status' data-testid='spell-dictionary-status'>
        {{ spellDictionaryLabel }}
      </div>
    </div>
    <div v-if='isDragOverFile' class='drag-drop-overlay' aria-hidden='true'>
      <div class='drag-drop-overlay-card'>
        <div class='drag-drop-overlay-title'>松开以打开 Markdown</div>
        <div class='drag-drop-overlay-subtitle'>支持 .md / .markdown 文件</div>
      </div>
    </div>
  </div>
</template>
