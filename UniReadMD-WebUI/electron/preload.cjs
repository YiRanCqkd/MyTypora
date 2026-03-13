const { contextBridge, ipcRenderer, webUtils } = require('electron');

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
    exportHtml: async (payload) => ipcRenderer.invoke('export:html', payload || {}),
    exportDocument: async (payload) => ipcRenderer.invoke('export:document', payload || {}),
    getCapabilities: async (traceId) => {
      return ipcRenderer.invoke('export:getCapabilities', { traceId });
    },
  },
  window: {
    minimize: async (traceId) => ipcRenderer.invoke('window:minimize', { traceId }),
    toggleMaximize: async (traceId) => ipcRenderer.invoke('window:toggleMaximize', { traceId }),
    isMaximized: async (traceId) => ipcRenderer.invoke('window:isMaximized', { traceId }),
    close: async (traceId) => ipcRenderer.invoke('window:close', { traceId }),
    toggleDevTools: async (traceId) => ipcRenderer.invoke('window:toggleDevTools', { traceId }),
  },
  settings: {
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
