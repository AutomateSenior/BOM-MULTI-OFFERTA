/**
 * Funzione eseguita all'apertura del foglio - crea un menu personalizzato
 * Mostra il menu di gestione formule e label solo se il nome del file contiene "master"
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();

    // Se è versione BETA, aggiungi watermark (con controllo)
    try {
      if (BOMLib.CONFIG && BOMLib.CONFIG.VERSION && CONFIG.VERSION.IS_BETA) {
        aggiungiWatermarkBeta();
      }
    } catch (e) {
      // Ignora errori watermark
    }

    // Crea il menu principale (con indicatore BETA se necessario)
    var isBeta = false;
    try {
      isBeta = BOMLib.CONFIG && BOMLib.CONFIG.VERSION && CONFIG.VERSION.IS_BETA;
    } catch (e) {}

    var titoloMenu = isBeta ? '⚠️ Automate [BETA]' : 'Automate';
    var menuPrincipale = ui.createMenu(titoloMenu);

  // SEZIONE 1: Gestione Offerte (prima voce)
  var menuGestioneOfferte = ui.createMenu('Gestione Offerte')
    .addItem('⚡ Gestione rapida', 'mostraGestioneRapidaOfferte')
    .addItem('Configurazione avanzata...', 'mostraConfigurazioneOfferte')
    .addSeparator()
    .addItem('Rigenera Budget', 'rigeneraBudgetDaOfferte');

  // Aggiungi submenu azzeramento dinamico
  var menuAzzera = ui.createMenu('Azzera Offerte');
  menuAzzera.addItem('Azzera tutte le offerte', 'azzeraTutteLeOfferte');

  // Ottieni lista offerte per menu dinamico
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Configurazione");
    if (configSheet && configSheet.getLastRow() >= 2) {
      var lastRow = configSheet.getLastRow();
      var data = configSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      if (data.length > 0) {
        menuAzzera.addSeparator();
        for (var i = 0; i < data.length; i++) {
          if (data[i][0]) {
            var offertaId = data[i][0];
            var offertaNome = data[i][1] || offertaId;
            // Crea funzione dinamica per ogni offerta
            menuAzzera.addItem('Azzera ' + offertaNome, 'azzeraOfferta_' + offertaId);
          }
        }
      }
    }
  } catch (e) {
    // Sistema non inizializzato o libreria non ancora caricata, ignora
    try {
      if (BOMLib.CONFIG && BOMLib.CONFIG.LOG) {
        BOMLib.CONFIG.LOG.warn("onOpen", "Errore creazione menu dinamico: " + e.toString());
      }
    } catch (logError) {}
  }

    menuGestioneOfferte.addSubMenu(menuAzzera);

    menuPrincipale.addSubMenu(menuGestioneOfferte);

    // SEZIONE 2: Operazioni BOM
    menuPrincipale
      .addSeparator()
      .addItem('Allinea BOM alla commessa', 'controlla');

    // SEZIONE 3: Verifica e Ripristino (solo per file master)
    var nomeFile = SpreadsheetApp.getActiveSpreadsheet().getName().toLowerCase();

    if (nomeFile.indexOf("master") > -1) {
      menuPrincipale
        .addSeparator()
        .addSubMenu(ui.createMenu('Verifica Integrità')
          .addItem('Verifica formule', 'verificaFormuleTutteLeOfferte')
          .addItem('Verifica label', 'verificaLabelTutteLeOfferte')
          .addItem('Verifica formule e label', 'verificaFormuleELabelTutteLeOfferte'))
        .addSubMenu(ui.createMenu('Ripristina')
          .addItem('Ripristina formule', 'ripristinaFormuleTutteLeOfferte')
          .addItem('Ripristina label', 'ripristinaLabelTutteLeOfferte')
          .addItem('Ripristina formule e label', 'ripristinaFormuleELabelTutteLeOfferte'));
    }

    // Aggiungi il menu all'interfaccia utente
    menuPrincipale.addToUi();

  } catch (error) {
    // Se onOpen fallisce completamente, almeno mostra un menu minimo
    try {
      SpreadsheetApp.getUi()
        .createMenu('Automate')
        .addItem('Allinea BOM alla commessa', 'controlla')
        .addToUi();
    } catch (e) {}
  }
}

// Funzione che installa l'attivatore per l'inserimento
function installaAttivatore() {
  // Questa funzione richiederà automaticamente le autorizzazioni necessarie
  try {
    // Elimina eventuali attivatori esistenti per evitare duplicati
    var triggersPresenti = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggersPresenti.length; i++) {
      var trigger = triggersPresenti[i];
      if (trigger.getEventType() == ScriptApp.EventType.ON_EDIT && 
          trigger.getHandlerFunction() == "Inserimento") {
        ScriptApp.deleteTrigger(trigger);
      }
    }
    
    // Crea l'attivatore per l'inserimento/modifica
    ScriptApp.newTrigger('Inserimento')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
    
    // Mostra un messaggio di conferma
    SpreadsheetApp.getUi().alert('Attivatore installato con successo!');
  } catch (e) {
    // Se c'è un errore (probabilmente di autorizzazioni), mostra messaggio all'utente
    SpreadsheetApp.getUi().alert('Errore durante l\'installazione dell\'attivatore: ' + e.toString() +
                                '\nVerifica di aver autorizzato le autorizzazioni richieste.');
  }
}

/**
 * Aggiunge watermark BETA visibile nel foglio Budget
 */
function aggiungiWatermarkBeta() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var budget = ss.getSheetByName(BOMLib.CONFIG.SHEETS.BUDGET);

    if (!budget) {
      return;
    }

    // Aggiungi watermark in cella A1
    var cella = budget.getRange("A1");
    cella.setValue(BOMLib.CONFIG.VERSION.BETA_WARNING);
    cella.setBackground("#ff6b35");  // Arancione acceso
    cella.setFontColor("#ffffff");   // Bianco
    cella.setFontWeight("bold");
    cella.setFontSize(12);
    cella.setHorizontalAlignment("center");
    cella.setVerticalAlignment("middle");

    // Proteggi la cella per evitare modifiche accidentali
    var protection = cella.protect().setDescription("Watermark BETA - Non modificare");
    protection.setWarningOnly(true);

    BOMLib.CONFIG.LOG.info("aggiungiWatermarkBeta", "Watermark BETA aggiunto in Budget!A1");

  } catch (error) {
    BOMLib.CONFIG.LOG.error("aggiungiWatermarkBeta", "Errore", error);
  }
}

// ===== NOTA: Funzioni principali ora in BOMCore.js e OffertaManager.js =====
// controlla(), rigeneraBudgetDaOfferte(), azzeraTutteLeOfferte() sono definite nei rispettivi file

/**
 * Mostra dialog rapida per gestione offerte (usa file HTML locale)
 */
function mostraGestioneRapidaOfferte() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('OffertaDialog')
      .setWidth(600)
      .setHeight(500);

    SpreadsheetApp.getUi().showModalDialog(html, 'Gestione Rapida Offerte');

  } catch (error) {
    if (BOMLib.CONFIG && BOMLib.CONFIG.LOG) {
      BOMLib.CONFIG.LOG.error("mostraGestioneRapidaOfferte", "Errore", error);
    }
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
  }
}

/**
 * Mostra sidebar configurazione offerte (usa file HTML locale)
 */
function mostraConfigurazioneOfferte() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('OffertaConfigUI')
      .setTitle('Configurazione Offerte')
      .setWidth(350);

    SpreadsheetApp.getUi().showSidebar(html);

  } catch (error) {
    if (BOMLib.CONFIG && BOMLib.CONFIG.LOG) {
      BOMLib.CONFIG.LOG.error("mostraConfigurazioneOfferte", "Errore", error);
    }
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
  }
}

// ===== FUNZIONI UI PER HTML =====
// Funzioni chiamate da OffertaConfigUI.html e OffertaDialog.html

/**
 * Verifica se è versione BETA
 */
function isBetaVersion() {
  try {
    return CONFIG && CONFIG.VERSION && CONFIG.VERSION.IS_BETA;
  } catch (e) {
    return false;
  }
}

/**
 * Ottiene configurazione offerte
 */
function getConfigurazioneOfferte() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Configurazione");
    if (!configSheet) return [];

    var lastRow = configSheet.getLastRow();
    if (lastRow < 2) return [];

    var data = configSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var offerte = [];

    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) {
        var foglio = ss.getSheetByName(data[i][0]);
        if (foglio) {
          offerte.push({
            id: data[i][0],
            nome: data[i][1] || data[i][0],
            descrizione: data[i][2] || "",
            abilitata: data[i][3] === true || data[i][3] === "TRUE",
            ordine: data[i][4] || (i + 1)
          });
        }
      }
    }
    return offerte;
  } catch (error) {
    Logger.log("Errore getConfigurazioneOfferte: " + error.toString());
    throw error;
  }
}

/**
 * Abilita/disabilita offerta
 */
function toggleAbilitaOfferta(id, abilitata) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Configurazione");
    if (!configSheet) throw new Error("Foglio Configurazione non trovato");

    var lastRow = configSheet.getLastRow();
    var data = configSheet.getRange(2, 1, lastRow - 1, 4).getValues();

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        configSheet.getRange(i + 2, 4).setValue(abilitata);
        var foglio = ss.getSheetByName(id);
        if (foglio) foglio.setTabColor(abilitata ? "#00ff00" : "#cccccc");
        return true;
      }
    }
    throw new Error("Offerta " + id + " non trovata");
  } catch (error) {
    Logger.log("Errore toggleAbilitaOfferta: " + error.toString());
    throw error;
  }
}

/**
 * Aggiorna descrizione offerta
 */
function aggiornaDescrizioneOfferta(id, descrizione) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Configurazione");
    if (!configSheet) throw new Error("Foglio Configurazione non trovato");

    var lastRow = configSheet.getLastRow();
    var data = configSheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        configSheet.getRange(i + 2, 3).setValue(descrizione);
        return true;
      }
    }
    throw new Error("Offerta " + id + " non trovata");
  } catch (error) {
    Logger.log("Errore aggiornaDescrizioneOfferta: " + error.toString());
    throw error;
  }
}

/**
 * Rimuove offerta
 */
function rimuoviOfferta(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("Configurazione");
    if (!configSheet) throw new Error("Foglio Configurazione non trovato");

    var lastRow = configSheet.getLastRow();
    var data = configSheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        configSheet.deleteRow(i + 2);
        break;
      }
    }

    var foglio = ss.getSheetByName(id);
    if (foglio) ss.deleteSheet(foglio);

    return true;
  } catch (error) {
    Logger.log("Errore rimuoviOfferta: " + error.toString());
    throw error;
  }
}

/**
 * Azzera offerta specifica - chiama azzeraOffertaSingola da OffertaManager.js
 */
function azzeraOfferta(id) {
  try {
    return BOMLib.azzeraOffertaSingola(id);
  } catch (error) {
    Logger.log("Errore azzeraOfferta: " + error.toString());
    throw error;
  }
}

// ===== WRAPPER FUNZIONI VERIFICA/RIPRISTINO =====

/**
 * NOTA: Funzioni libreria temporaneamente disabilitate per debug
 */
// function verificaFormuleTutteLeOfferte() { return BOMLib.verificaFormuleTutteLeOfferte(); }
// function verificaLabelTutteLeOfferte() { return BOMLib.verificaLabelTutteLeOfferte(); }
// function verificaFormuleELabelTutteLeOfferte() { return BOMLib.verificaFormuleELabelTutteLeOfferte(); }
// function ripristinaFormuleTutteLeOfferte() { return BOMLib.ripristinaFormuleTutteLeOfferte(); }
// function ripristinaLabelTutteLeOfferte() { return BOMLib.ripristinaLabelTutteLeOfferte(); }
// function ripristinaFormuleELabelTutteLeOfferte() { return BOMLib.ripristinaFormuleELabelTutteLeOfferte(); }
// function azzeraOfferta_Off_01() { return BOMLib.azzeraOfferta("Off_01"); }
// function azzeraOfferta_Off_02() { return BOMLib.azzeraOfferta("Off_02"); }
// ... altre funzioni azzera commentate per debug


// ===== WRAPPER FUNZIONI DA LIBRERIA =====
// Queste funzioni sono chiamate dal menu e fanno da wrapper alle funzioni della libreria

/**
 * Funzioni chiamate dal menu - delegano alla libreria BOMLib
 */
function controlla() { return BOMLib.controlla(); }
function rigeneraBudgetDaOfferte() { return BOMLib.rigeneraBudgetDaOfferte(); }
function azzeraTutteLeOfferte() { return BOMLib.azzeraTutteLeOfferte(); }
function aggiungiNuovaOfferta(nome, descrizione) { return BOMLib.aggiungiNuovaOfferta(nome, descrizione); }
function rigeneraBudget() { return BOMLib.rigeneraBudget(); }

// Funzioni verifica e ripristino
function verificaFormuleTutteLeOfferte() { return BOMLib.verificaFormuleTutteLeOfferte(); }
function verificaLabelTutteLeOfferte() { return BOMLib.verificaLabelTutteLeOfferte(); }
function verificaFormuleELabelTutteLeOfferte() { return BOMLib.verificaFormuleELabelTutteLeOfferte(); }
function ripristinaFormuleTutteLeOfferte() { return BOMLib.ripristinaFormuleTutteLeOfferte(); }
function ripristinaLabelTutteLeOfferte() { return BOMLib.ripristinaLabelTutteLeOfferte(); }
function ripristinaFormuleELabelTutteLeOfferte() { return BOMLib.ripristinaFormuleELabelTutteLeOfferte(); }
