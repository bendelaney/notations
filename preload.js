const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('notations', {
  printDocument: (payload) => ipcRenderer.invoke('print-document', payload),
  exportPdf: (payload) => ipcRenderer.invoke('export-pdf', payload),
  exportText: (payload) => ipcRenderer.invoke('export-text', payload),
  loadState: () => ipcRenderer.invoke('load-state'),
  saveState: (payload) => ipcRenderer.invoke('save-state', payload),
  consumeInitialDeepLink: () => ipcRenderer.invoke('consume-initial-deep-link'),
  onDeepLink: (handler) => {
    if (typeof handler !== 'function') {
      return () => {};
    }
    const listener = (_event, routePath) => handler(routePath);
    ipcRenderer.on('open-deep-link', listener);
    return () => ipcRenderer.removeListener('open-deep-link', listener);
  },
  onMenuPrintRequest: (handler) => {
    if (typeof handler !== 'function') {
      return () => {};
    }
    const listener = () => handler();
    ipcRenderer.on('request-print-document', listener);
    return () => ipcRenderer.removeListener('request-print-document', listener);
  }
});
