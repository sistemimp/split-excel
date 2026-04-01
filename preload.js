const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("excelSplitter", {
  pickInputFile: () => ipcRenderer.invoke("dialog:openXlsx"),
  pickOutputDir: () => ipcRenderer.invoke("dialog:openOutputDir"),
  analyzeFile: (inputPath) => ipcRenderer.invoke("xlsx:analyzeFile", inputPath),
  splitFile: (payload) => ipcRenderer.invoke("xlsx:splitFile", payload)
});
