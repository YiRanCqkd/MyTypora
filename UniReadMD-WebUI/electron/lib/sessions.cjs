const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { atomicWrite } = require('./atomic-write.cjs');
const { createContractError } = require('./errors.cjs');

const MAX_SESSION_BYTES = 25 * 1024 * 1024;

function createSessionService({ fs, directory, authorization, dialog, getEventWindow }) {
  const owners = new Map();
  const recovery = [];
  const writes = new Map();
  const historyDirectory = path.join(directory, 'history');
  const archivedAt = new Map();
  async function archive(snapshot) {
    if (!snapshot?.isDirty || !snapshot.content) return;
    await fs.mkdir(historyDirectory, { recursive: true });
    const id = randomUUID();
    await atomicWrite(fs, path.join(historyDirectory, `${id}.json`),
      JSON.stringify({ ...snapshot, savedAt: Date.now() }));
    const entries = [];
    for (const name of await fs.readdir(historyDirectory)) {
      if (!/^[a-f0-9-]+\.json$/.test(name)) continue;
      const target = path.join(historyDirectory, name);
      const stat = await fs.stat(target);
      entries.push({ target, size: stat.size, time: stat.mtimeMs });
    }
    entries.sort((a, b) => b.time - a.time);
    let bytes = 0;
    for (let index = 0; index < entries.length; index += 1) {
      bytes += entries[index].size;
      if (index >= 30 || bytes > 100 * 1024 * 1024) await fs.unlink(entries[index].target);
    }
  }
  async function initialize() {
    await fs.mkdir(directory, { recursive: true });
    for (const name of await fs.readdir(directory)) {
      if (!/^[a-f0-9-]+\.json$/.test(name)) continue;
      try {
        const filePath = path.join(directory, name);
        const stat = await fs.stat(filePath);
        if (!stat.isFile() || stat.size > MAX_SESSION_BYTES) continue;
        const raw = await fs.readFile(filePath, 'utf8');
        if (Buffer.byteLength(raw) > MAX_SESSION_BYTES) continue;
        const snapshot = JSON.parse(raw);
        if (typeof snapshot.content === 'string' && typeof snapshot.filePath === 'string') {
          recovery.push({ id: name.slice(0, -5), snapshot });
        }
      } catch {
        // 损坏快照不会覆盖其他可恢复文档。
      }
    }
    recovery.sort((a, b) => Number(b.snapshot.savedAt) - Number(a.snapshot.savedAt));
    let keptClean = false;
    for (let index = 0; index < recovery.length;) {
      const entry = recovery[index];
      if (entry.snapshot.isDirty || !keptClean) {
        if (!entry.snapshot.isDirty) keptClean = true;
        index += 1;
      } else {
        // 清理失败不应阻止其他有效草稿恢复。
        await fs.unlink(path.join(directory, `${entry.id}.json`)).catch(() => {});
        recovery.splice(index, 1);
      }
    }
  }
  function register(win, hasLaunchFile) {
    const entry = hasLaunchFile ? null : recovery.shift();
    const owner = { id: entry?.id || randomUUID(), snapshot: entry?.snapshot || null };
    owners.set(win.webContents.id, owner);
    if (owner.snapshot?.filePath) authorization.grant(win, owner.snapshot.filePath);
    if (owner.snapshot?.workspaceDirectory) {
      authorization.grantDirectory(win, owner.snapshot.workspaceDirectory);
    }
  }
  async function load(event) {
    return owners.get(event.sender.id)?.snapshot || null;
  }
  async function save(event, payload) {
    const owner = owners.get(event.sender.id);
    const snapshot = payload?.snapshot;
    if (!owner || !snapshot || typeof snapshot.content !== 'string'
      || typeof snapshot.filePath !== 'string') {
      throw createContractError('ERR_INVALID_SESSION', '会话格式无效');
    }
    if (snapshot.filePath) authorization.assertAuthorized(event, snapshot.filePath);
    if (snapshot.workspaceDirectory) {
      authorization.assertAuthorized(event, snapshot.workspaceDirectory);
    }
    const raw = JSON.stringify({ ...owner.snapshot, ...snapshot, savedAt: Date.now() });
    if (Buffer.byteLength(raw) > MAX_SESSION_BYTES) {
      throw createContractError('ERR_SESSION_TOO_LARGE', '恢复快照超过 25 MiB，请保存文档');
    }
    const target = path.join(directory, `${owner.id}.json`);
    const pending = (writes.get(owner.id) || Promise.resolve()).catch(() => {}).then(async () => {
      const previous = owner.snapshot;
      const changedDocument = previous?.filePath !== snapshot.filePath;
      const discarded = previous?.isDirty && !snapshot.isDirty;
      if (changedDocument || discarded) await archive(previous);
      await atomicWrite(fs, target, raw);
      if (snapshot.isDirty && previous?.content !== snapshot.content
        && Date.now() - (archivedAt.get(owner.id) || 0) >= 30_000) {
        await archive(snapshot);
        archivedAt.set(owner.id, Date.now());
      }
      owner.snapshot = JSON.parse(raw);
    });
    writes.set(owner.id, pending);
    await pending;
    return { ok: true };
  }
  async function listDrafts() {
    await fs.mkdir(historyDirectory, { recursive: true });
    const drafts = [];
    for (const name of await fs.readdir(historyDirectory)) {
      if (!/^[a-f0-9-]+\.json$/.test(name)) continue;
      try {
        const target = path.join(historyDirectory, name);
        const stat = await fs.stat(target);
        if (!stat.isFile() || stat.size > MAX_SESSION_BYTES) continue;
        const raw = await fs.readFile(target, 'utf8');
        if (Buffer.byteLength(raw) > MAX_SESSION_BYTES) continue;
        const snapshot = JSON.parse(raw);
        if (typeof snapshot.content !== 'string') continue;
        drafts.push({ id: name.slice(0, -5), filePath: snapshot.filePath || '',
          savedAt: snapshot.savedAt, title: snapshot.content.split(/\r?\n/)[0].slice(0, 100) });
      } catch { /* 单个损坏草稿不会阻断恢复列表。 */ }
    }
    return { drafts: drafts.sort((a, b) => b.savedAt - a.savedAt) };
  }
  async function restoreDraft(event, payload) {
    const id = payload?.id;
    if (typeof id !== 'string' || !/^[a-f0-9-]+$/.test(id)) {
      throw createContractError('ERR_INVALID_DRAFT', '草稿标识无效');
    }
    const target = path.join(historyDirectory, `${id}.json`);
    if ((await fs.stat(target)).size > MAX_SESSION_BYTES) throw new Error('草稿超过大小限制');
    const snapshot = JSON.parse(await fs.readFile(target, 'utf8'));
    if (typeof snapshot.content !== 'string') throw new Error('草稿内容已损坏');
    // 恢复为未命名副本，防止历史草稿覆盖磁盘上的新版本或重新授权旧路径。
    return { content: snapshot.content, originalFilePath: snapshot.filePath || '',
      filePath: '', isDirty: true, savedAt: snapshot.savedAt };
  }
  async function confirmDiscard(event) {
    const result = await dialog.showMessageBox(getEventWindow(event), {
      type: 'question', title: '未保存的修改',
      message: '当前文档有未保存的修改。',
      buttons: ['保存', '放弃修改', '取消'], defaultId: 0, cancelId: 2, noLink: true,
    });
    return { action: ['save', 'discard', 'cancel'][result.response] || 'cancel' };
  }
  async function release(windowOrId) {
    const ownerId = typeof windowOrId === 'number' ? windowOrId : windowOrId.webContents.id;
    const owner = owners.get(ownerId);
    if (!owner) return;
    await writes.get(owner.id)?.catch(() => {});
    writes.delete(owner.id);
    owners.delete(ownerId);
  }
  return { initialize, register, load, save, confirmDiscard, release, listDrafts, restoreDraft,
    getRecoveryCount: () => recovery.length,
  };
}

module.exports = { createSessionService, MAX_SESSION_BYTES };
