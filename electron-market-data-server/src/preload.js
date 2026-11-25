const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('marketServer', {
  getStatus: () => ipcRenderer.invoke('get-server-status'),
  onStatus: (callback) => ipcRenderer.on('server-status', (_event, data) => callback(data))
});
