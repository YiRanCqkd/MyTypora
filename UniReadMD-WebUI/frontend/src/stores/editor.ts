import { defineStore } from 'pinia';

interface EditorState {
  filePath: string;
  content: string;
  isDirty: boolean;
  cursorHint: number;
  documentRevision: number;
  resourceBaseUrl: string;
  version: string;
  savedContent: string;
}

interface SessionPayload {
  filePath: string;
  content: string;
  isDirty: boolean;
  cursorHint: number;
}

function sanitizeLoadedContent(content: string): string {
  return content.replace(/^\uFEFF/, '');
}

export const useEditorStore = defineStore('editor', {
  state: (): EditorState => ({
    filePath: '',
    content: '# UniReadMD\n\n在这里输入 Markdown 内容。',
    isDirty: false,
    cursorHint: 0,
    documentRevision: 0,
    resourceBaseUrl: '',
    version: '',
    savedContent: '# UniReadMD\n\n在这里输入 Markdown 内容。',
  }),
  actions: {
    setContent(content: string) {
      this.content = content;
      this.isDirty = content !== this.savedContent;
    },
    setFile(filePath: string, content: string,
      metadata?: { resourceBaseUrl?: string; version?: string }) {
      this.filePath = filePath;
      this.content = sanitizeLoadedContent(content);
      this.isDirty = false;
      this.cursorHint = 0;
      this.documentRevision += 1;
      this.resourceBaseUrl = metadata?.resourceBaseUrl || '';
      this.version = metadata?.version || '';
      this.savedContent = this.content;
    },
    markSaved(filePath: string, savedContent?: string,
      metadata?: { resourceBaseUrl?: string; version?: string }) {
      this.filePath = filePath;
      this.savedContent = savedContent ?? this.content;
      this.isDirty = this.content !== this.savedContent;
      if (metadata?.resourceBaseUrl) this.resourceBaseUrl = metadata.resourceBaseUrl;
      if (metadata?.version) this.version = metadata.version;
    },
    setDirty(isDirty: boolean) {
      this.isDirty = isDirty;
    },
    setCursorHint(line: number) {
      this.cursorHint = line;
    },
    restoreSession(payload: SessionPayload) {
      this.filePath = payload.filePath;
      this.content = sanitizeLoadedContent(payload.content);
      this.isDirty = payload.isDirty;
      this.cursorHint = payload.cursorHint;
      this.documentRevision += 1;
      this.savedContent = payload.isDirty ? '' : this.content;
    },
  },
});
