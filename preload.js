const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  exportDocx: (markdown, options = {}) => ipcRenderer.invoke('export-docx', {
    markdown,
    ...options
  })
});
