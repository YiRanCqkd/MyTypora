const path = require('node:path');
const { constants } = require('node:fs');
const { createContractError, requireString } = require('./errors.cjs');
const { relocateDocumentImages } = require('./document-images.cjs');

function validateEntryName(value) {
  const name = requireString(value, 'name', 240);
  if (/[/\\<>:"|?*\x00-\x1f]/.test(name) || /[. ]$/.test(name)
    || /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name)) {
    throw createContractError('ERR_INVALID_FILE_NAME', '文件名包含非法字符或系统保留名称');
  }
  return name;
}

function isWithin(root, target) {
  const relative = path.relative(root, target);
  return relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

function createWorkspaceTools({ fs, files, assets, authorization, shell, dialog, getEventWindow }) {
  async function openDroppedFile(event, payload) {
    // 此入口仅由隔离 preload 的真实拖放 File 能力调用，不暴露字符串路径授权 API。
    const filePath = await fs.realpath(requireString(payload.filePath, 'filePath'));
    if (!/\.(md|markdown)$/i.test(filePath)) {
      throw createContractError('ERR_UNSUPPORTED_DOCUMENT', '拖放仅支持 Markdown 文档');
    }
    authorization.grant(event, filePath);
    return files.openFromPath(event, { filePath });
  }

  async function openLinkedFile(event, payload) {
    const source = requireString(payload.sourceFilePath, 'sourceFilePath');
    authorization.assertAuthorized(event, source);
    const href = requireString(payload.href, 'href');
    if (/^[a-z][a-z\d+.-]*:/i.test(href) && !/^[a-z]:[\\/]/i.test(href)) {
      throw createContractError('ERR_INVALID_LINK', '仅支持本地 Markdown 链接');
    }
    const root = await fs.realpath(files.getWorkspaceRoot(event) || path.dirname(source));
    const requested = path.resolve(path.dirname(source), decodeURIComponent(href.split('#')[0]));
    if (!/\.(md|markdown)$/i.test(requested)) {
      throw createContractError('ERR_UNSUPPORTED_DOCUMENT', '仅支持 Markdown 文档链接');
    }
    const target = await fs.realpath(requested);
    if (!isWithin(root, target) && !authorization.isAuthorized(event, target)) {
      throw createContractError('ERR_PATH_NOT_AUTHORIZED', '链接超出当前工作区，请通过打开菜单选择文件');
    }
    authorization.grant(event, target);
    return files.openFromPath(event, { filePath: target });
  }

  async function operate(event, payload) {
    const action = requireString(payload.action, 'action', 24);
    const source = await fs.realpath(requireString(payload.filePath, 'filePath'));
    authorization.assertAuthorized(event, source);
    const stat = await fs.stat(source);
    if (source === files.getWorkspaceRoot(event) && ['trash', 'move', 'rename'].includes(action)) {
      throw createContractError('ERR_WORKSPACE_ROOT_PROTECTED', '请在资源管理器中操作工作区根目录');
    }
    if (action === 'trash') {
      await shell.trashItem(source);
      return { filePath: source, action };
    }
    if (!['rename', 'duplicate', 'mkdir', 'move', 'create'].includes(action)) {
      throw createContractError('ERR_INVALID_FILE_OPERATION', '文件操作无效');
    }
    let parent = ['mkdir', 'create'].includes(action) ? source : path.dirname(source);
    if (['mkdir', 'create'].includes(action) && !stat.isDirectory()) {
      throw createContractError('ERR_NOT_DIRECTORY', '请选择目标文件夹');
    }
    if (action === 'move') {
      const directory = requireString(payload.directoryPath, 'directoryPath');
      authorization.assertAuthorized(event, directory);
      parent = await fs.realpath(directory);
      if (!(await fs.stat(parent)).isDirectory() || isWithin(source, parent)) {
        throw createContractError('ERR_INVALID_MOVE', '不能把目录移动到自身或子目录');
      }
    }
    const name = validateEntryName(payload.name || path.basename(source));
    const target = path.join(parent, name);
    if (!stat.isDirectory() || action === 'create') {
      if (!/\.(md|markdown|txt)$/i.test(target)) {
        throw createContractError('ERR_UNSUPPORTED_DOCUMENT', '文件扩展名须为 md、markdown 或 txt');
      }
    }
    try {
      await fs.lstat(target);
      throw createContractError('ERR_TARGET_EXISTS', '目标名称已存在，不会覆盖原文件');
    } catch (error) { if (error.code !== 'ENOENT') throw error; }
    if (action === 'mkdir') await fs.mkdir(target);
    else if (action === 'create') await fs.writeFile(target, '', { flag: 'wx', encoding: 'utf8' });
    else if (action === 'duplicate') {
      if (stat.isDirectory()) throw new Error('请选择 Markdown 文件进行复制');
      await fs.copyFile(source, target, constants.COPYFILE_EXCL);
    } else if (stat.isDirectory()) {
      // Windows 的目录重命名不会覆盖已有目标；跨卷移动保留源目录并报告错误。
      await fs.rename(source, target);
    } else {
      // 独占复制后移除源文件，同时支持跨卷文件移动，避免覆盖同名目标。
      const content = await fs.readFile(source, 'utf8');
      const relocated = await relocateDocumentImages({ content, source, target, event, assets, fs,
        rootDirectory: files.getWorkspaceRoot(event) });
      let created = false;
      try {
        await fs.writeFile(target, relocated.content, { flag: 'wx', encoding: 'utf8' });
        created = true;
        await fs.unlink(source);
      } catch (error) {
        if (created) await fs.unlink(target).catch(() => {});
        await relocated.cleanup();
        throw error;
      }
    }
    if (stat.isDirectory() && action !== 'create') authorization.grantDirectory(event, target);
    else authorization.grant(event, target);
    return { filePath: target, previousPath: source, action };
  }

  async function chooseMoveDirectory(event) {
    const result = await dialog.showOpenDialog(getEventWindow(event), {
      title: '选择移动目标文件夹', properties: ['openDirectory'],
    });
    if (result.canceled || !result.filePaths[0]) return null;
    const directoryPath = await fs.realpath(result.filePaths[0]);
    authorization.grantDirectory(event, directoryPath);
    return { directoryPath };
  }

  async function inspect(event, payload) {
    const filePath = requireString(payload.filePath, 'filePath');
    authorization.assertAuthorized(event, filePath);
    const { createHash } = require('node:crypto');
    const stat = await fs.stat(filePath);
    if (!stat.isFile() || stat.size > 20 * 1024 * 1024) throw new Error('文档超过大小限制');
    return { version: createHash('sha256').update(await fs.readFile(filePath)).digest('hex') };
  }
  return { openDroppedFile, openLinkedFile, operate, chooseMoveDirectory, inspect };
}

module.exports = { createWorkspaceTools, validateEntryName, isWithin };
