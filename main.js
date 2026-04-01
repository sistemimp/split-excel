const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");

function createWindow() {
  const win = new BrowserWindow({
    width: 820,
    height: 640,
    minWidth: 760,
    minHeight: 560,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  win.loadFile("index.html");
}

function chunkArray(array, size) {
  const chunks = [];
  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }
  return chunks;
}

function sanitizeFilePart(value) {
  const sanitized = String(value || "")
    .replace(/[<>:"/\\|?*]+/g, "_")
    .replace(/[. ]+$/g, "")
    .trim();

  return sanitized || "sheet";
}

ipcMain.handle("dialog:openXlsx", async () => {
  const result = await dialog.showOpenDialog({
    title: "Seleziona file XLSX",
    properties: ["openFile"],
    filters: [{ name: "Excel", extensions: ["xlsx"] }]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return result.filePaths[0];
});

ipcMain.handle("dialog:openOutputDir", async () => {
  const result = await dialog.showOpenDialog({
    title: "Seleziona cartella di output",
    properties: ["openDirectory", "createDirectory"]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return result.filePaths[0];
});

ipcMain.handle("xlsx:analyzeFile", async (_event, inputPath) => {
  if (!inputPath || !fs.existsSync(inputPath)) {
    throw new Error("File XLSX non valido o non trovato.");
  }

  const workbook = XLSX.readFile(inputPath, {
    cellDates: true
  });

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error("Il file non contiene fogli validi.");
  }

  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = sheet
      ? XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: "",
          raw: false,
          blankrows: false
        })
      : [];
    const totalRows = rows.length;
    const dataRows = Math.max(totalRows - 1, 0);

    return {
      name: sheetName,
      totalRows,
      dataRows
    };
  });

  return {
    sheets
  };
});

ipcMain.handle("xlsx:splitFile", async (_event, payload) => {
  const { inputPath, outputDir, recordsPerFile, sheetName } = payload || {};
  const safeRecords = Number(recordsPerFile);

  if (!inputPath || !fs.existsSync(inputPath)) {
    throw new Error("File XLSX non valido o non trovato.");
  }

  if (!outputDir || !fs.existsSync(outputDir)) {
    throw new Error("Cartella di output non valida.");
  }

  if (!Number.isInteger(safeRecords) || safeRecords <= 0) {
    throw new Error("Il numero record deve essere un intero maggiore di zero.");
  }

  if (!sheetName || typeof sheetName !== "string") {
    throw new Error("Seleziona un foglio valido da splittare.");
  }

  const workbook = XLSX.readFile(inputPath, { cellDates: true });
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error("Il file non contiene fogli validi.");
  }

  if (!workbook.SheetNames.includes(sheetName)) {
    throw new Error("Il foglio selezionato non esiste nel file.");
  }

  const sourceSheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sourceSheet, {
    header: 1,
    defval: "",
    raw: false
  });

  if (rows.length === 0) {
    throw new Error("Il file è vuoto.");
  }

  const headerRow = rows[0];
  const dataRows = rows.slice(1);

  if (dataRows.length === 0) {
    throw new Error("Il file contiene solo intestazioni e nessun record.");
  }

  const chunks = chunkArray(dataRows, safeRecords);
  const inputBaseName = path.basename(inputPath, path.extname(inputPath));
  const safeSheetPart = sanitizeFilePart(sheetName);
  const outputFiles = [];

  chunks.forEach((chunk, index) => {
    const outWb = XLSX.utils.book_new();
    const outRows = [headerRow, ...chunk];
    const outSheet = XLSX.utils.aoa_to_sheet(outRows);
    XLSX.utils.book_append_sheet(outWb, outSheet, sheetName);

    const fileName = `${inputBaseName}_${safeSheetPart}_part_${String(index + 1).padStart(3, "0")}.xlsx`;
    const destination = path.join(outputDir, fileName);
    XLSX.writeFile(outWb, destination);
    outputFiles.push(destination);
  });

  return {
    createdFiles: outputFiles.length,
    outputFiles
  };
});

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
