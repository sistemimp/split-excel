# Split Excel (Electron)

Applicazione desktop in Electron per dividere un file `.xlsx` in piu file,
impostando da interfaccia grafica il numero massimo di record per file.

## Comportamento

- Analizza il file Excel e mostra i fogli disponibili.
- Permette di scegliere da interfaccia il **foglio da splittare**.
- Considera la **prima riga come intestazione**.
- Divide le righe successive in blocchi da `N` record (valore impostato nella GUI).
- Ogni file generato contiene:
  - prima riga: intestazioni originali
  - righe dati del blocco corrente

## Avvio

```bash
npm install
npm start
```

## Uso

1. Seleziona il file `.xlsx` di input.
2. Seleziona il foglio da splittare (lista popolata dopo l'analisi).
3. Seleziona la cartella di output.
4. Inserisci il numero di record per file.
5. Clicca **Avvia split**.
