const { contextBridge, ipcRenderer, webUtils } = require('electron');

const trustedDrops = new Map();
document.addEventListener('drop', (event) => {
  if (!event.isTrusted) return;
  trustedDrops.clear();
  for (const file of [...(event.dataTransfer?.files || [])].slice(0, 16)) {
    const filePath = webUtils.getPathForFile(file);
    if (filePath) trustedDrops.set(filePath, Date.now());
  }
}, true);

function consumeDroppedFile(file) {
  const filePath = webUtils.getPathForFile(file);
  const receivedAt = trustedDrops.get(filePath);
  trustedDrops.delete(filePath);
  if (!filePath || !receivedAt || Date.now() - receivedAt > 5000) {
    throw new Error('拖放文件必须来自真实的系统拖放操作');
  }
  return filePath;
}

async function openDroppedFile(file) {
  return ipcRenderer.invoke('file:openDroppedFile', { filePath: consumeDroppedFile(file) });
}

function withTracePayload(payload, traceId) {
  return {
    ...(payload || {}),
    traceId,
  };
}

const api = {
  app: {
    getVersion: async (traceId) => ipcRenderer.invoke('app:getVersion', { traceId }),
    logEvent: async (payload) => ipcRenderer.invoke('app:logEvent', payload || {}),
    consumeLaunchFilePath: async (traceId) => {
      return ipcRenderer.invoke('app:consumeLaunchFilePath', { traceId });
    },
    onLaunchFile: (handler) => {
      if (typeof handler !== 'function') {
        return () => {};
      }

      const listener = (_event, payload) => {
        handler(payload || null);
      };

      ipcRenderer.on('app:launchFile', listener);

      return () => {
        ipcRenderer.removeListener('app:launchFile', listener);
      };
    },
  },
  file: {
    openDroppedFile,
    openLinkedFile: async (payload) => ipcRenderer.invoke('file:openLinkedFile', payload || {}),
    operate: async (payload) => ipcRenderer.invoke('file:operate', payload || {}),
    inspect: async (payload) => ipcRenderer.invoke('file:inspect', payload || {}),
    chooseMoveDirectory: async () => ipcRenderer.invoke('file:chooseMoveDirectory', {}),
    openDirectory: async (payload) => ipcRenderer.invoke('file:openDirectory', payload || {}),
    searchDirectory: async (payload) => ipcRenderer.invoke('file:searchDirectory', payload || {}),
    openMarkdown: async (payload) => ipcRenderer.invoke('file:openMarkdown', payload || {}),
    openFromPath: async (payload) => ipcRenderer.invoke('file:openFromPath', payload || {}),
    listMarkdownInDirectory: async (payload) => {
      return ipcRenderer.invoke('file:listMarkdownInDirectory', payload || {});
    },
    saveMarkdown: async (payload) => ipcRenderer.invoke('file:saveMarkdown', payload || {}),
    getPathForFile: (file) => {
      try {
        return webUtils.getPathForFile(file);
      } catch {
        return '';
      }
    },
  },
  exporter: {
    chooseResource: async (kind) => ipcRenderer.invoke('export:chooseResource', { kind }),
    importDocument: async () => ipcRenderer.invoke('export:importDocument', {}),
    exportHtml: async (payload) => ipcRenderer.invoke('export:html', payload || {}),
    exportDocument: async (payload) => ipcRenderer.invoke('export:document', payload || {}),
    getCapabilities: async (traceId) => {
      return ipcRenderer.invoke('export:getCapabilities', { traceId });
    },
  },
  session: {
    listDrafts: async () => ipcRenderer.invoke('session:listDrafts', {}),
    restoreDraft: async (id) => ipcRenderer.invoke('session:restoreDraft', { id }),
    load: async () => ipcRenderer.invoke('session:load', {}),
    save: async (snapshot) => ipcRenderer.invoke('session:save', { snapshot }),
    confirmDiscard: async () => ipcRenderer.invoke('document:confirmDiscard', {}),
  },
  images: {
    collect: async (payload) => ipcRenderer.invoke('image:collect', payload || {}),
    drop: async (file, payload) => ipcRenderer.invoke('image:drop', {
      ...(payload || {}), filePath: consumeDroppedFile(file),
    }),
    applyRoot: async (payload) => ipcRenderer.invoke('image:applyRoot', payload || {}),
    select: async (payload) => ipcRenderer.invoke('image:select', payload || {}),
    paste: async (payload) => ipcRenderer.invoke('image:paste', payload || {}),
    setRoot: async (payload) => ipcRenderer.invoke('image:setRoot', payload || {}),
  },
  window: {
    minimize: async (traceId) => ipcRenderer.invoke('window:minimize', { traceId }),
    toggleMaximize: async (traceId) => ipcRenderer.invoke('window:toggleMaximize', { traceId }),
    isMaximized: async (traceId) => ipcRenderer.invoke('window:isMaximized', { traceId }),
    close: async (traceId) => ipcRenderer.invoke('window:close', { traceId }),
    toggleDevTools: async (traceId) => ipcRenderer.invoke('window:toggleDevTools', { traceId }),
  },
  settings: {
    getUserCss: async () => ipcRenderer.invoke('settings:getUserCss', {}),
    getAll: async (traceId) => ipcRenderer.invoke('settings:getAll', { traceId }),
    getRecentFiles: async (traceId) => {
      return ipcRenderer.invoke('settings:getRecentFiles', { traceId });
    },
    get: async (key, traceId) => ipcRenderer.invoke('settings:get', { key, traceId }),
    set: async (key, value, traceId) => {
      return ipcRenderer.invoke('settings:set', { key, value, traceId });
    },
    setRecentFiles: async (filePaths, traceId) => {
      return ipcRenderer.invoke('settings:setRecentFiles', {
        filePaths,
        traceId,
      });
    },
    selectUserCssFile: async (traceId) => {
      return ipcRenderer.invoke('settings:selectUserCssFile', { traceId });
    },
  },
  context: {
    showFileTreeMenu: async (payload) => {
      return ipcRenderer.invoke('context:showFileTreeMenu', payload || {});
    },
    showSourceEditorMenu: async (payload) => {
      return ipcRenderer.invoke('context:showSourceEditorMenu', payload || {});
    },
  },
  menu: {
    onCommand: (handler) => {
      if (typeof handler !== 'function') {
        return () => {};
      }

      const listener = (_event, command) => {
        handler(command);
      };

      ipcRenderer.on('menu:command', listener);

      return () => {
        ipcRenderer.removeListener('menu:command', listener);
      };
    },
  },
  spell: {
    getState: async (traceId) => ipcRenderer.invoke('spell:getState', { traceId }),
    setEnabled: async (enabled, traceId) => {
      return ipcRenderer.invoke('spell:setEnabled', { enabled, traceId });
    },
    setLanguage: async (language, traceId) => {
      return ipcRenderer.invoke('spell:setLanguage', { language, traceId });
    },
  },

  // backward compatibility
  getAppVersion: async () => {
    const result = await ipcRenderer.invoke('app:getVersion', {});
    return result.version;
  },
  openMarkdownFile: async (traceId) => {
    return ipcRenderer.invoke('file:openMarkdown', withTracePayload({}, traceId));
  },
  saveMarkdownFile: async (payload) => {
    return ipcRenderer.invoke('file:saveMarkdown', payload || {});
  },
};

contextBridge.exposeInMainWorld('nativeBridge', api);
