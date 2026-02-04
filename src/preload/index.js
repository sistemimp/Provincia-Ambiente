const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('electronApi', {
  convertWorkbook: (params) => ipcRenderer.invoke('excel:convert', params),
  revealFile: (params) => ipcRenderer.invoke('excel:reveal', params),
  importWorkbookToDb: (params) => ipcRenderer.invoke('excel:import-db', params),
  truncateTable: () => ipcRenderer.invoke('excel:truncate-db'),
  downloadJasperPdf: (params) => ipcRenderer.invoke('jasper:download-pdf', params),
})
