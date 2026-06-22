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

      // L56 scritta in un foglio Off_XX: l'utente sta sbagliando foglio
      if (cellAddress === "L56" && /^Off_/.test(sheetName)) {
        Logger.log("Inserimento: L56 scritta in " + sheetName + " - reindirizzo a Budget");
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var nuovoValore = range.getValue();

        // UI disponibile solo se il trigger gira come l'utente che ha il foglio
        // aperto; in background (eseguito come Archive per modifiche altrui) no.
        var ui = null;
        try { ui = SpreadsheetApp.getUi(); } catch (eUi) { ui = null; }
        if (ui) {
          ui.alert(
            "Foglio errato",
            "Il codice commessa va inserito nel foglio Budget, non in " + sheetName + ".\n" +
            "Il valore verrà copiato automaticamente in Budget.",
            ui.ButtonSet.OK
          );
        }

        // Cancella il valore scritto nel posto sbagliato
        range.clearContent();

        // Copia in Budget e processa come se fosse stato scritto lì
        var budgetSheet = ss.getSheetByName("Budget");
        if (budgetSheet) {
          var oldBudgetL56 = String(budgetSheet.getRange("L56").getValue()).trim();
          budgetSheet.getRange("L56").setValue(nuovoValore);
          BOM8.controlla(ss, false);
          try {
            var s56off = budgetSheet.getRange("S56").getValue();
            if (s56off === "Commessa OK") {
              BOM8.propagaL56(ss, String(nuovoValore).trim());
              BOM8.gestisciDisallineamentoNomeFile(ss, oldBudgetL56);
            }
          } catch (errOff) {
            Logger.log("Inserimento: Errore gestione Off_XX L56 - " + errOff.toString());
          }
        }
        return;
      }

      if (sheetName === "Budget" && cellAddress === "L56") {
        Logger.log("Inserimento: Rilevata modifica L56, avvio allineamento BOM");
        BOM8.CONFIG.LOG.info("Inserimento", "Modificata cella L56, avvio allineamento BOM");

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        BOM8.controlla(ss, false);  // Funzione locale in BOMCore.js

        try {
          var s56 = ss.getSheetByName("Budget").getRange("S56").getValue();
          if (s56 === "Commessa OK") {
            BOM8.propagaL56(ss, String(ss.getSheetByName("Budget").getRange("L56").getValue()).trim());
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






