const inputPathEl = document.getElementById("inputPath");
const outputDirEl = document.getElementById("outputDir");
const sheetSelectEl = document.getElementById("sheetSelect");
const sheetInfoEl = document.getElementById("sheetInfo");
const recordsPerFileEl = document.getElementById("recordsPerFile");
const resultBoxEl = document.getElementById("resultBox");

const pickInputBtn = document.getElementById("pickInputBtn");
const pickOutputBtn = document.getElementById("pickOutputBtn");
const splitBtn = document.getElementById("splitBtn");

function setResult(message) {
  resultBoxEl.textContent = message;
}

function resetSheetSelection(message) {
  sheetSelectEl.innerHTML = "";
  const option = document.createElement("option");
  option.value = "";
  option.textContent = message;
  sheetSelectEl.appendChild(option);
  sheetSelectEl.value = "";
  sheetSelectEl.disabled = true;
}

function populateSheetSelection(sheets) {
  sheetSelectEl.innerHTML = "";
  sheets.forEach((sheet, index) => {
    const option = document.createElement("option");
    option.value = sheet.name;
    option.textContent = `${sheet.name} (${sheet.dataRows} record)`;
    if (index === 0) {
      option.selected = true;
    }
    sheetSelectEl.appendChild(option);
  });
  sheetSelectEl.disabled = false;
}

pickInputBtn.addEventListener("click", async () => {
  const selected = await window.excelSplitter.pickInputFile();
  if (selected) {
    inputPathEl.value = selected;
    if (!outputDirEl.value) {
      const normalized = selected.replace(/\\/g, "/");
      outputDirEl.value = normalized.substring(0, normalized.lastIndexOf("/")).replace(/\//g, "\\");
    }

    sheetInfoEl.textContent = "Analisi file in corso...";
    resetSheetSelection("Caricamento fogli...");

    try {
      const analysis = await window.excelSplitter.analyzeFile(selected);
      const sheets = Array.isArray(analysis.sheets) ? analysis.sheets : [];

      if (sheets.length === 0) {
        resetSheetSelection("Nessun foglio disponibile");
        sheetInfoEl.textContent = "Il file non contiene fogli utilizzabili.";
        return;
      }

      populateSheetSelection(sheets);
      sheetInfoEl.textContent = `Trovati ${sheets.length} fogli. Seleziona quello da splittare.`;
      setResult(
        ["Analisi completata.", "", ...sheets.map((s) => `- ${s.name}: ${s.dataRows} record (esclusa intestazione)`)].join("\n")
      );
    } catch (error) {
      resetSheetSelection("Errore analisi file");
      sheetInfoEl.textContent = "Impossibile leggere i fogli dal file selezionato.";
      setResult(`Errore: ${error.message || "Analisi file non riuscita."}`);
    }
  }
});

pickOutputBtn.addEventListener("click", async () => {
  const selected = await window.excelSplitter.pickOutputDir();
  if (selected) {
    outputDirEl.value = selected;
  }
});

splitBtn.addEventListener("click", async () => {
  const inputPath = inputPathEl.value.trim();
  const outputDir = outputDirEl.value.trim();
  const sheetName = sheetSelectEl.value.trim();
  const recordsPerFile = Number(recordsPerFileEl.value);

  if (!inputPath) {
    setResult("Seleziona prima un file XLSX di input.");
    return;
  }

  if (!outputDir) {
    setResult("Seleziona una cartella di output.");
    return;
  }

  if (!sheetName) {
    setResult("Seleziona il foglio da splittare.");
    return;
  }

  if (!Number.isInteger(recordsPerFile) || recordsPerFile <= 0) {
    setResult("Il numero di record per file deve essere un intero maggiore di zero.");
    return;
  }

  splitBtn.disabled = true;
  setResult("Elaborazione in corso...");

  try {
    const result = await window.excelSplitter.splitFile({
      inputPath,
      outputDir,
      sheetName,
      recordsPerFile
    });

    const lines = [
      `Completato: creati ${result.createdFiles} file dal foglio "${sheetName}".`,
      "",
      ...result.outputFiles
    ];
    setResult(lines.join("\n"));
  } catch (error) {
    setResult(`Errore: ${error.message || "Operazione non riuscita."}`);
  } finally {
    splitBtn.disabled = false;
  }
});
