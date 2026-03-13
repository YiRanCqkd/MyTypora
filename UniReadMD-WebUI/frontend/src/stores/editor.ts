import { defineStore } from 'pinia';

interface EditorState {
  filePath: string;
  content: string;
  isDirty: boolean;
  cursorHint: number;
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
  }),
  actions: {
    setContent(content: string) {
      this.content = content;
      this.isDirty = true;
    },
    setFile(filePath: string, content: string) {
      this.filePath = filePath;
      this.content = sanitizeLoadedContent(content);
      this.isDirty = false;
      this.cursorHint = 0;
    },
    markSaved(filePath: string) {
      this.filePath = filePath;
      this.isDirty = false;
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
    },
  },
});
