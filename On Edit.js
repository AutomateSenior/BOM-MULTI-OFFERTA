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
        BOM8.CONFIG.LOG.info("Inserimento", "Modificata cella L56, avvio allineamento BOM");

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        BOM8.controlla(ss, false);  // Funzione locale in BOMCore.js

        try {
          var s56 = ss.getSheetByName("Budget").getRange("S56").getValue();
          if (s56 === "Commessa OK") {
            BOM8.gestisciDisallineamentoNomeFile(ss, e.oldValue || "");
          }
        } catch (errDisal) {
          Logger.log("Inserimento: Errore gestione disallineamento - " + errDisal.toString());
        }

        Logger.log("Inserimento: Allineamento BOM completato");
        return; // Esci dopo aver gestito L56
      }
    }
  } catch (error) {
    Logger.log("Inserimento: ERRORE trigger L56 - " + error.toString());
    BOM8.CONFIG.LOG.error("Inserimento", "Errore trigger L56", error);
  }

  // Gestione trigger standard BOM (solo se non è L56)
  try {
    BOM8.InserimentoBOM(e);  // Funzione locale in BOMCore.js
  } catch (error) {
    BOM8.CONFIG.LOG.error("Inserimento", "Errore InserimentoBOM", error);
  }
}





