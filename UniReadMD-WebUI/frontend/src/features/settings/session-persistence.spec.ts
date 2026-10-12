import { afterEach, describe, expect, it, vi } from 'vitest';
import { reactive, ref } from 'vue';
import { useSessionPersistence } from './session-persistence';
import type { SessionSnapshot } from '../workspace/workspace-contracts';

const originalBridge = Object.getOwnPropertyDescriptor(window, 'nativeBridge');

afterEach(() => {
  if (originalBridge) Object.defineProperty(window, 'nativeBridge', originalBridge);
  else Reflect.deleteProperty(window, 'nativeBridge');
  localStorage.clear();
});

function createPersistence(save?: (snapshot: SessionSnapshot) => Promise<{ ok: true }>) {
  Object.defineProperty(window, 'nativeBridge', {
    configurable: true,
    value: save ? { session: { save } } : undefined,
  });
  const input = {
    SESSION_STORAGE_KEY: 'session-regression',
    store: reactive({
      filePath: 'I:/文档/示例.md',
      version: '版本一',
      content: '# 示例',
      isDirty: false,
      cursorHint: 1,
    }),
    workspaceDirectory: ref('I:/文档'),
    recentFiles: ref(['I:/文档/示例.md']),
    fileTreeExpanded: ref({ 'I:/文档': true }),
    contentZoom: ref(1),
    pinOutline: ref(true),
    sourceMode: ref(false),
    focusMode: ref(false),
    typewriterMode: ref(false),
    sidebarWidth: ref(270),
    mainPreviewWidth: ref(0),
    fileViewMode: ref('tree'),
    fileSortMode: ref('natural'),
    filesFilter: ref(''),
    showDocSearchPanel: ref(false),
    showDocReplacePanel: ref(false),
    docSearchQuery: ref(''),
    docReplaceQuery: ref(''),
    docSearchCaseSensitive: ref(false),
    docSearchWholeWord: ref(false),
    docSearchMatchIndex: ref(0),
    errorMessage: ref(''),
  };
  const persistence = useSessionPersistence(
    input as unknown as Parameters<typeof useSessionPersistence>[0],
  );
  return { input, ...persistence };
}

describe('恢复快照的 IPC 数据与写入队列', () => {
  it('打开文档的响应式状态能被结构化克隆，且快照不随后续状态变化', async () => {
    const saved: SessionSnapshot[] = [];
    const save = vi.fn(async (snapshot: SessionSnapshot) => {
      saved.push(structuredClone(snapshot));
      return { ok: true as const };
    });
    const persistence = createPersistence(save);
    const snapshot = persistence.buildSessionSnapshot();
    expect(structuredClone(snapshot)).toEqual(snapshot);
    await persistence.persistSessionSnapshot();
    expect(save).toHaveBeenCalledOnce();
    expect(persistence.input.errorMessage.value).toBe('');
    persistence.input.recentFiles.value.push('I:/文档/其他.md');
    persistence.input.fileTreeExpanded.value['I:/文档'] = false;
    expect(snapshot.recentFiles).toEqual(['I:/文档/示例.md']);
    expect(snapshot.fileTreeExpanded).toEqual({ 'I:/文档': true });
    expect(saved[0]).toEqual(snapshot);
  });

  it('写入期间连续编辑只追加最新快照，关闭等待涵盖最后一次修改', async () => {
    let finishFirst!: () => void;
    const firstWrite = new Promise<void>((resolve) => { finishFirst = resolve; });
    const saved: SessionSnapshot[] = [];
    const persistence = createPersistence(async (snapshot) => {
      saved.push(structuredClone(snapshot));
      if (saved.length === 1) await firstWrite;
      return { ok: true };
    });
    const first = persistence.persistSessionSnapshot();
    persistence.input.store.content = '中间内容';
    persistence.persistSessionSnapshot();
    persistence.input.store.content = '最终内容';
    const last = persistence.persistSessionSnapshot();
    expect(last).toBe(first);
    finishFirst();
    await last;
    expect(saved.map((snapshot) => snapshot.content)).toEqual(['# 示例', '最终内容']);
    expect(persistence.input.errorMessage.value).toBe('');
  });

  it('真实保存失败仍显示提示并保留文档，后续保存可以重试', async () => {
    const save = vi.fn()
      .mockRejectedValueOnce(new Error('模拟磁盘写入失败'))
      .mockResolvedValue({ ok: true });
    const persistence = createPersistence(save);
    await persistence.persistSessionSnapshot();
    expect(persistence.input.errorMessage.value).toContain('模拟磁盘写入失败');
    expect(persistence.input.store.content).toBe('# 示例');
    await persistence.persistSessionSnapshot();
    expect(save).toHaveBeenCalledTimes(2);
  });

  it('浏览器模式仍保存完整的最近文件和目录展开状态', () => {
    const persistence = createPersistence();
    persistence.persistSessionSnapshot();
    const snapshot = JSON.parse(localStorage.getItem('session-regression')!);
    expect(snapshot).toEqual(persistence.buildSessionSnapshot());
    expect(persistence.input.errorMessage.value).toBe('');
  });
});
