/**
 * Test.js - Funzioni di test per debug
 * v1.4.0: Aggiornato per Sviluppo Software
 */

/**
 * Test diretto della dialog
 */
function testDialogDiretto() {
  try {
    Logger.log("testDialogDiretto: Inizio");

    var html = HtmlService.createHtmlOutputFromFile('OffertaDialog')
      .setWidth(600)
      .setHeight(500);

    Logger.log("testDialogDiretto: HTML creato");

    SpreadsheetApp.getUi().showModalDialog(html, 'TEST - Gestione Rapida Offerte');

    Logger.log("testDialogDiretto: Dialog mostrata");

  } catch (error) {
    Logger.log("testDialogDiretto: ERRORE - " + error.toString());
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
  }
}

/**
 * Test getConfigurazioneOfferte diretta
 */
function testGetConfigurazioneOfferte() {
  try {
    Logger.log("testGetConfigurazioneOfferte: Inizio");

    var result = getConfigurazioneOfferte();

    Logger.log("testGetConfigurazioneOfferte: Risultato = " + JSON.stringify(result));

    SpreadsheetApp.getUi().alert("Offerte trovate: " + result.length + "\n\n" + JSON.stringify(result, null, 2));

  } catch (error) {
    Logger.log("testGetConfigurazioneOfferte: ERRORE - " + error.toString());
    Logger.log("Stack: " + error.stack);
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
  }
}

/**
 * Test completo: esegue getConfigurazioneOfferte e mostra risultato
 */
function testCompleto() {
  testGetConfigurazioneOfferte();
  Utilities.sleep(1000);
  testDialogDiretto();
}
