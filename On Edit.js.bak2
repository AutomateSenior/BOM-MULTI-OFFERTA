function Inserimento(e){
  Logger.log("Inserimento: Trigger OnEdit attivato");

  // Trigger automatico per cella L56 (PRIORITARIO - eseguito per primo)
  try {
    if (e && e.range) {
      var range = e.range;
      var sheet = range.getSheet();
      var sheetName = sheet.getName();
      var cellAddress = range.getA1Notation();

      Logger.log("Inserimento: Foglio=" + sheetName + ", Cella=" + cellAddress);

      if (sheetName === "Budget" && cellAddress === "L56") {
        Logger.log("Inserimento: Rilevata modifica L56, avvio allineamento BOM");
        CONFIG.LOG.info("Inserimento", "Modificata cella L56, avvio allineamento BOM");

        controlla(SpreadsheetApp.getActiveSpreadsheet(), false);  // Funzione locale in BOMCore.js

        Logger.log("Inserimento: Allineamento BOM completato");
        return; // Esci dopo aver gestito L56
      }
    }
  } catch (error) {
    Logger.log("Inserimento: ERRORE trigger L56 - " + error.toString());
    CONFIG.LOG.error("Inserimento", "Errore trigger L56", error);
  }

  // Gestione trigger standard BOM (solo se non è L56)
  try {
    InserimentoBOM(e);  // Funzione locale in BOMCore.js
  } catch (error) {
    CONFIG.LOG.error("Inserimento", "Errore InserimentoBOM", error);
  }
}





