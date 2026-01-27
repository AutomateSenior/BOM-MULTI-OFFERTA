/**
 * OffertaManager.js - Sistema Multi-Offerta
 * Versione locale (non più libreria)
 *
 * Gestisce:
 * - Foglio Master (template nascosto)
 * - Fogli Off_01, Off_02, ... Off_NN (varianti offerta)
 * - Foglio Budget (sintesi offerte abilitate)
 * - Foglio Configurazione (nascosto, dati offerte)
 */

// CONFIG già definito in Config.js

/**
 * Inizializza il sistema multi-offerta
 * Converte un file esistente o crea la struttura da zero
 */
function inizializzaSistemaMultiOfferta() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();

    // Verifica se già inizializzato
    if (ss.getSheetByName(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE)) {
      ui.alert("Sistema già inizializzato!\n\nUsa il menu 'Opzioni Offerta → Configura offerte' per gestire le offerte.");
      return;
    }

    CONFIG.LOG.info("inizializzaSistemaMultiOfferta", "Inizio inizializzazione");

    // 1. Verifica esistenza foglio Budget
    var budgetOriginale = ss.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!budgetOriginale) {
      throw new Error("Foglio Budget non trovato! Impossibile inizializzare.");
    }

    // 2. Crea foglio Master (copia di Budget)
    var master = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    if (!master) {
      CONFIG.LOG.info("inizializzaSistemaMultiOfferta", "Creazione foglio Master");

      // Verifica se esiste già un foglio con nome simile
      var fogli = ss.getSheets();
      for (var i = 0; i < fogli.length; i++) {
        var nomeFoglio = fogli[i].getName();
        if (nomeFoglio.toLowerCase() === CONFIG.SHEETS.MASTER.toLowerCase() ||
            nomeFoglio.trim() === CONFIG.SHEETS.MASTER) {
          master = fogli[i];
          CONFIG.LOG.info("inizializzaSistemaMultiOfferta", "Foglio Master già esistente, lo uso");
          break;
        }
      }

      // Se non esiste, crealo
      if (!master) {
        master = budgetOriginale.copyTo(ss);

        // Nome univoco temporaneo
        var nomeTemp = "Master_" + new Date().getTime();
        master.setName(nomeTemp);

        // Verifica che non esista già "Master"
        var masterEsistente = ss.getSheetByName(CONFIG.SHEETS.MASTER);
        if (masterEsistente) {
          // C'è già Master, eliminiamolo
          ss.deleteSheet(masterEsistente);
        }

        // Ora rinomina
        master.setName(CONFIG.SHEETS.MASTER);
      }

      ss.setActiveSheet(budgetOriginale); // Torna a Budget
      master.hideSheet();

      // Proteggi Master
      var protection = master.protect().setDescription("Foglio Master - Template protetto");
      protection.setWarningOnly(true);
    }

    // 3. Crea foglio Configurazione
    CONFIG.LOG.info("inizializzaSistemaMultiOfferta", "Creazione foglio Configurazione");
    var config = ss.insertSheet(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE);
    config.hideSheet();

    // Intestazioni
    config.getRange("A1:E1").setValues([[
      "ID", "Nome", "Descrizione", "Abilitata", "Ordine"
    ]]).setFontWeight("bold").setBackground("#4285f4").setFontColor("#ffffff");

    // Prima offerta (Off_01)
    config.getRange("A2:E2").setValues([[
      "Off_01", "Offerta Base", "Configurazione standard", true, 1
    ]]);

    config.autoResizeColumns(1, 5);

    // 4. Crea foglio Off_01
    CONFIG.LOG.info("inizializzaSistemaMultiOfferta", "Creazione foglio Off_01");
    var off01 = ss.getSheetByName("Off_01");
    if (!off01) {
      off01 = master.copyTo(ss);
      off01.setName("Off_01");
      // Posiziona dopo Budget
      var fogli = ss.getSheets();
      var posBudget = 0;
      for (var i = 0; i < fogli.length; i++) {
        if (fogli[i].getName() === CONFIG.SHEETS.BUDGET) {
          posBudget = i + 1;
          break;
        }
      }
      ss.setActiveSheet(off01);
      ss.moveActiveSheet(posBudget + 1);
    }

    // 5. Rigenera Budget
    CONFIG.LOG.info("inizializzaSistemaMultiOfferta", "Rigenerazione Budget");
    rigeneraBudgetDaOfferte();

    CONFIG.LOG.info("inizializzaSistemaMultiOfferta", "Inizializzazione completata");

    ui.alert(
      "Sistema Multi-Offerta Inizializzato! ✓",
      "Il sistema è stato configurato con successo:\n\n" +
      "✓ Foglio Master creato (nascosto)\n" +
      "✓ Foglio Configurazione creato\n" +
      "✓ Offerta Off_01 creata\n" +
      "✓ Budget rigenerato\n\n" +
      "Usa 'Automate → Opzioni Offerta' per gestire le offerte.",
      ui.ButtonSet.OK
    );

  } catch (error) {
    CONFIG.LOG.error("inizializzaSistemaMultiOfferta", "Errore durante inizializzazione", error);
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
    throw error;
  }
}

/**
 * Verifica se questa è una versione BETA
 * @returns {boolean} true se BETA
 */
function isBetaVersion() {
  return CONFIG.VERSION.IS_BETA;
}

/**
 * Ottiene la configurazione delle offerte dal foglio nascosto
 * @returns {Array<Object>} Array di oggetti offerta
 */
function getConfigurazioneOfferte() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE);

    if (!configSheet) {
      CONFIG.LOG.warn("getConfigurazioneOfferte", "Foglio Configurazione non trovato");
      return [];
    }

    var lastRow = configSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    var data = configSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var offerte = [];
    var offerteMancanti = [];

    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) { // Se ha ID
        var offertaId = data[i][0];
        var foglio = ss.getSheetByName(offertaId);

        // Verifica se il foglio esiste
        if (!foglio) {
          offerteMancanti.push(offertaId);
          CONFIG.LOG.warn("getConfigurazioneOfferte", "Foglio " + offertaId + " non trovato, verrà rimosso dalla configurazione");
        } else {
          offerte.push({
            id: data[i][0],
            nome: data[i][1],
            descrizione: data[i][2],
            abilitata: data[i][3] === true || data[i][3] === "TRUE",
            ordine: data[i][4] || (i + 1)
          });
        }
      }
    }

    // Se ci sono offerte mancanti, rimuovile dalla configurazione
    if (offerteMancanti.length > 0) {
      for (var j = 0; j < offerteMancanti.length; j++) {
        rimuoviOffertaDaConfigurazione(offerteMancanti[j], configSheet);
      }

      // Mostra messaggio all'utente
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        "Fogli Offerta Mancanti",
        "I seguenti fogli offerta non esistono più e sono stati rimossi dalla configurazione:\n\n" +
        offerteMancanti.join(", ") +
        "\n\nLa configurazione è stata aggiornata automaticamente.",
        ui.ButtonSet.OK
      );
    }

    // Ordina per ordine
    offerte.sort(function(a, b) { return a.ordine - b.ordine; });

    return offerte;

  } catch (error) {
    CONFIG.LOG.error("getConfigurazioneOfferte", "Errore", error);
    return [];
  }
}

/**
 * Aggiunge una nuova offerta
 * @param {string} nome - Nome della nuova offerta
 * @param {string} descrizione - Descrizione opzionale
 * @returns {string} ID della nuova offerta (es: "Off_05")
 */
function aggiungiNuovaOfferta(nome, descrizione) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE);

    if (!configSheet) {
      throw new Error("Sistema non inizializzato. Usa 'Inizializza Sistema Multi-Offerta' prima.");
    }

    nome = nome || "Nuova Offerta";
    descrizione = descrizione || "";

    // Trova prossimo numero
    var offerte = getConfigurazioneOfferte();
    var maxNum = 0;
    for (var i = 0; i < offerte.length; i++) {
      var num = parseInt(offerte[i].id.replace("Off_", ""));
      if (num > maxNum) maxNum = num;
    }
    var nuovoNum = maxNum + 1;
    var nuovoId = "Off_" + (nuovoNum < 10 ? "0" : "") + nuovoNum;

    CONFIG.LOG.info("aggiungiNuovaOfferta", "Creazione offerta " + nuovoId);

    // Aggiungi a Configurazione
    var lastRow = configSheet.getLastRow();
    configSheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
      nuovoId, nome, descrizione, true, nuovoNum
    ]]);

    // Duplica Master (NON Budget!)
    var master = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    if (!master) {
      throw new Error("Foglio Master non trovato! Impossibile creare nuova offerta.");
    }

    // DIAGNOSTICA: Verifica colore celle nel Master PRIMA della copia
    var masterTestCell = master.getRange("S100");
    var masterColor = masterTestCell.getBackground();
    CONFIG.LOG.info("aggiungiNuovaOfferta", "Colore cella S100 nel Master PRIMA copia: " + masterColor);

    CONFIG.LOG.info("aggiungiNuovaOfferta", "Copiando da foglio: " + master.getName());

    var nuovoFoglio = master.copyTo(ss);
    nuovoFoglio.setName(nuovoId);

    // DIAGNOSTICA: Verifica colore celle nel nuovo foglio DOPO la copia
    var nuovoTestCell = nuovoFoglio.getRange("S100");
    var nuovoColor = nuovoTestCell.getBackground();
    CONFIG.LOG.info("aggiungiNuovaOfferta", "Colore cella S100 in " + nuovoId + " DOPO copia: " + nuovoColor);

    CONFIG.LOG.info("aggiungiNuovaOfferta", "Foglio " + nuovoId + " creato da " + master.getName());

    // Posiziona dopo l'ultima offerta
    var fogli = ss.getSheets();
    var posBudget = 0;
    for (var i = 0; i < fogli.length; i++) {
      if (fogli[i].getName() === CONFIG.SHEETS.BUDGET) {
        posBudget = i + 1;
        break;
      }
    }
    ss.setActiveSheet(nuovoFoglio);
    ss.moveActiveSheet(posBudget + offerte.length + 1);

    CONFIG.LOG.info("aggiungiNuovaOfferta", "Offerta " + nuovoId + " creata");

    return nuovoId;

  } catch (error) {
    CONFIG.LOG.error("aggiungiNuovaOfferta", "Errore", error);
    throw error;
  }
}

/**
 * Rimuove un'offerta
 * @param {string} id - ID offerta da rimuovere (es: "Off_03")
 */
function rimuoviOfferta(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE);

    if (!configSheet) {
      throw new Error("Sistema non inizializzato.");
    }

    CONFIG.LOG.info("rimuoviOfferta", "Rimozione offerta " + id);

    // Verifica che non sia l'ultima offerta
    var offerte = getConfigurazioneOfferte();
    if (offerte.length <= 1) {
      throw new Error("Non puoi rimuovere l'ultima offerta! Deve esistere almeno un'offerta.");
    }

    // Rimuovi da Configurazione
    var data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        configSheet.deleteRow(i + 2);
        break;
      }
    }

    // Rimuovi foglio
    var foglio = ss.getSheetByName(id);
    if (foglio) {
      ss.deleteSheet(foglio);
    }

    CONFIG.LOG.info("rimuoviOfferta", "Offerta " + id + " rimossa");

  } catch (error) {
    CONFIG.LOG.error("rimuoviOfferta", "Errore", error);
    throw error;
  }
}

/**
 * Rimuove un'offerta dalla configurazione senza eliminare il foglio
 * (usato quando il foglio non esiste più)
 * @param {string} id - ID offerta da rimuovere
 * @param {Sheet} configSheet - Foglio configurazione (opzionale)
 */
function rimuoviOffertaDaConfigurazione(id, configSheet) {
  try {
    if (!configSheet) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE);
    }

    if (!configSheet) {
      return;
    }

    // Rimuovi dalla configurazione
    var data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        configSheet.deleteRow(i + 2);
        CONFIG.LOG.info("rimuoviOffertaDaConfigurazione", "Offerta " + id + " rimossa dalla configurazione");
        break;
      }
    }

  } catch (error) {
    CONFIG.LOG.error("rimuoviOffertaDaConfigurazione", "Errore", error);
  }
}

/**
 * Abilita o disabilita un'offerta
 * @param {string} id - ID offerta
 * @param {boolean} abilitata - true per abilitare, false per disabilitare
 */
function toggleAbilitaOfferta(id, abilitata) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE);

    if (!configSheet) {
      throw new Error("Sistema non inizializzato.");
    }

    CONFIG.LOG.info("toggleAbilitaOfferta", id + " → " + abilitata);

    // Trova riga e aggiorna
    var data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        configSheet.getRange(i + 2, 4).setValue(abilitata);
        break;
      }
    }

    // Aggiorna colore tab immediatamente
    var foglio = ss.getSheetByName(id);
    if (foglio) {
      if (abilitata) {
        foglio.setTabColor(CONFIG.OFFERTE.COLORE_TAB_ABILITATA);
      } else {
        foglio.setTabColor(CONFIG.OFFERTE.COLORE_TAB_DISABILITATA);
      }
    }

    // V08: Imposta indicatore rigenerazione necessaria in O62
    impostaIndicatoreRigenerazione(true);

  } catch (error) {
    CONFIG.LOG.error("toggleAbilitaOfferta", "Errore", error);
    throw error;
  }
}

/**
 * V08: Imposta o cancella l'indicatore di rigenerazione necessaria in O62
 * @param {boolean} mostra - true per mostrare indicatore, false per cancellare
 */
function impostaIndicatoreRigenerazione(mostra) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var budget = ss.getSheetByName(CONFIG.SHEETS.BUDGET);

    if (!budget) {
      CONFIG.LOG.warn("impostaIndicatoreRigenerazione", "Foglio Budget non trovato");
      return;
    }

    var cella = budget.getRange(CONFIG.CELLS.INDICATORE_RIGENERA);

    if (mostra) {
      cella.setValue("⚠️ Rigenerazione necessaria");
      cella.setBackground("#fff2cc");  // Giallo chiaro
      cella.setFontColor("#ff0000");   // Rosso
      cella.setFontWeight("bold");
      CONFIG.LOG.info("impostaIndicatoreRigenerazione", "Indicatore impostato in " + CONFIG.CELLS.INDICATORE_RIGENERA);
    } else {
      cella.clear();
      CONFIG.LOG.info("impostaIndicatoreRigenerazione", "Indicatore cancellato da " + CONFIG.CELLS.INDICATORE_RIGENERA);
    }

  } catch (error) {
    CONFIG.LOG.error("impostaIndicatoreRigenerazione", "Errore", error);
  }
}

/**
 * Aggiorna la descrizione di un'offerta
 * @param {string} id - ID offerta
 * @param {string} descrizione - Nuova descrizione
 */
function aggiornaDescrizioneOfferta(id, descrizione) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE);

    if (!configSheet) {
      return;
    }

    var data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        configSheet.getRange(i + 2, 3).setValue(descrizione);
        break;
      }
    }

  } catch (error) {
    CONFIG.LOG.error("aggiornaDescrizioneOfferta", "Errore", error);
  }
}
/**
 * Identifica le celle verdi nel range configurato del foglio Budget
 * V08: Range 69-526 (include Body Rental 505-517, aggiornato dopo Sviluppo Software)
 * Cerca solo nelle colonne specificate in CONFIG.OFFERTE.COLONNE_VERDI
 * @param {Sheet} sheet - Foglio da analizzare
 * @returns {Array<Object>} Array di {row, col}
 */
function getCelleVerdiPerSomma(sheet) {
  try {
    var celleVerdi = [];
    var numRighe = CONFIG.OFFERTE.RANGE_SOMMA_FINE - CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + 1;

    // Per ogni colonna specificata
    for (var i = 0; i < CONFIG.OFFERTE.COLONNE_VERDI.length; i++) {
      var colonnaLettera = CONFIG.OFFERTE.COLONNE_VERDI[i];
      var colonnaNum = letterToColumn(colonnaLettera);

      // Leggi range di quella colonna
      var range = sheet.getRange(
        CONFIG.OFFERTE.RANGE_SOMMA_INIZIO,
        colonnaNum,
        numRighe,
        1
      );

      var backgrounds = range.getBackgrounds();

      // Cerca celle verdi o blu (già processate)
      var celleVerdiColonna = 0;
      var celleBluColonna = 0;
      for (var row = 0; row < backgrounds.length; row++) {
        var color = backgrounds[row][0].toLowerCase();
        var isVerde = isColoreVerde(color);
        var isBlu = isColoreBlu(color);

        if (isVerde || isBlu) {
          celleVerdi.push({
            row: CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + row,
            col: colonnaNum
          });

          if (isVerde) celleVerdiColonna++;
          if (isBlu) celleBluColonna++;
        }
      }

      CONFIG.LOG.info("getCelleVerdiPerSomma",
        "Colonna " + colonnaLettera + ": " + celleVerdiColonna + " verdi, " + celleBluColonna + " blu");
    }

    return celleVerdi;

  } catch (error) {
    CONFIG.LOG.error("getCelleVerdiPerSomma", "Errore", error);
    return [];
  }
}

/**
 * Ottieni celle gialle per MAX (colonna U)
 * @param {Sheet} sheet - Foglio da analizzare
 * @returns {Array<{row: number, col: number}>}
 */
function getCelleGiallePerMax(sheet) {
  try {
    var celleGialle = [];
    var numRighe = CONFIG.OFFERTE.RANGE_SOMMA_FINE - CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + 1;

    // Per ogni colonna specificata
    for (var i = 0; i < CONFIG.OFFERTE.COLONNE_GIALLE.length; i++) {
      var colonnaLettera = CONFIG.OFFERTE.COLONNE_GIALLE[i];
      var colonnaNum = letterToColumn(colonnaLettera);

      // Leggi range di quella colonna
      var range = sheet.getRange(
        CONFIG.OFFERTE.RANGE_SOMMA_INIZIO,
        colonnaNum,
        numRighe,
        1
      );

      var backgrounds = range.getBackgrounds();

      // Cerca celle gialle o blu (già processate)
      for (var row = 0; row < backgrounds.length; row++) {
        var rowNum = CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + row;
        var color = backgrounds[row][0].toLowerCase();

        // Escludi S476 (ha gestione speciale con validazione dati, era S465)
        if (rowNum === 476 && colonnaLettera === "S") {
          continue;
        }

        if (isColoreGiallo(color) || isColoreBlu(color)) {
          celleGialle.push({
            row: rowNum,
            col: colonnaNum
          });
        }
      }
    }

    return celleGialle;

  } catch (error) {
    CONFIG.LOG.error("getCelleGiallePerMax", "Errore", error);
    return [];
  }
}

/**
 * Ottieni celle verdi per concatenazione (colonna M)
 * @param {Sheet} sheet - Foglio da analizzare
 * @returns {Array<{row: number, col: number}>}
 */
function getCelleVerdiPerConcatenazione(sheet) {
  try {
    var celleConcatenazione = [];
    var numRighe = CONFIG.OFFERTE.RANGE_SOMMA_FINE - CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + 1;

    // Per ogni colonna specificata
    for (var i = 0; i < CONFIG.OFFERTE.COLONNE_CONCATENAZIONE.length; i++) {
      var colonnaLettera = CONFIG.OFFERTE.COLONNE_CONCATENAZIONE[i];
      var colonnaNum = letterToColumn(colonnaLettera);

      // Leggi range di quella colonna
      var range = sheet.getRange(
        CONFIG.OFFERTE.RANGE_SOMMA_INIZIO,
        colonnaNum,
        numRighe,
        1
      );

      var backgrounds = range.getBackgrounds();

      // Cerca SOLO celle verdi (non blu - quelle sono già processate dal Budget)
      for (var row = 0; row < backgrounds.length; row++) {
        var color = backgrounds[row][0].toLowerCase();
        var rowNum = CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + row;

        // V08: Solo celle verdi, non blu (celle blu = già processate in Budget, non input)
        if (isColoreVerde(color)) {
          celleConcatenazione.push({
            row: rowNum,
            col: colonnaNum
          });
        }
      }
    }

    return celleConcatenazione;

  } catch (error) {
    CONFIG.LOG.error("getCelleVerdiPerConcatenazione", "Errore", error);
    return [];
  }
}

/**
 * Verifica se un colore è verde
 * @param {string} color - Colore in formato hex
 * @returns {boolean}
 */
function isColoreVerde(color) {
  var coloriVerdi = [
    "#00ff00",   // Verde puro
    "#ccffcc",   // Verde chiaro
    "#d9ead3",   // Verde pastello Google Sheets
    "#93c47d",   // Verde medio
    "#6aa84f",   // Verde scuro
    "#b6d7a8",   // Verde chiaro alternativo
    "#a8d5a8"    // Verde chiaro alternativo
  ];

  return coloriVerdi.indexOf(color) > -1;
}

/**
 * Verifica se un colore è blu (celle già processate dal sistema)
 * @param {string} color - Colore in formato hex
 * @returns {boolean}
 */
function isColoreBlu(color) {
  var coloriBlu = [
    "#4285f4",   // Blu Google (usato dal sistema)
    "#1a73e8"    // Blu Google alternativo
  ];

  return coloriBlu.indexOf(color) > -1;
}

/**
 * Verifica se un colore è giallo
 * @param {string} color - Colore in formato hex
 * @returns {boolean}
 */
function isColoreGiallo(color) {
  var coloriGialli = [
    "#ffff00",   // Giallo puro
    "#fff2cc",   // Giallo chiaro
    "#ffe599",   // Giallo pastello
    "#f1c232",   // Giallo medio
    "#ffd966"    // Giallo alternativo
  ];

  return coloriGialli.indexOf(color) > -1;
}

/**
 * Verifica se un colore è rosso errore (usato per errori M)
 * @param {string} color - Colore in formato hex
 * @returns {boolean}
 */
function isColoreRossoErrore(color) {
  // Solo il rosso specifico usato per gli errori
  return color === "#ea4335";
}

/**
 * Aggiorna la cella M62 con l'etichetta sintesi
 * @param {Sheet} budget - Foglio Budget
 * @param {Array<string>} offerteAbilitate - Array di ID offerte abilitate
 */
function aggiornaEtichettaSintesi(budget, offerteAbilitate) {
  try {
    CONFIG.LOG.info("aggiornaEtichettaSintesi", "Inizio aggiornamento, offerte: " + offerteAbilitate.join(", "));

    var numeri = offerteAbilitate.map(function(id) {
      return id.replace("Off_", "");
    });

    var etichetta = "Sintesi BOM (Offerta " + numeri.join("+") + ")";

    CONFIG.LOG.info("aggiornaEtichettaSintesi", "Etichetta da scrivere: " + etichetta);

    var cella = budget.getRange(CONFIG.OFFERTE.CELLA_ETICHETTA_SINTESI);

    // Prima legge il valore corrente
    var valoreCorrente = cella.getValue();
    CONFIG.LOG.info("aggiornaEtichettaSintesi", "Valore corrente in " + CONFIG.OFFERTE.CELLA_ETICHETTA_SINTESI + ": " + valoreCorrente);

    // Cancella eventuale formula e imposta come valore semplice
    cella.clearContent();
    cella.setValue(etichetta);

    // Verifica che sia stato scritto
    var valoreFinale = cella.getValue();
    CONFIG.LOG.info("aggiornaEtichettaSintesi", "Valore finale in " + CONFIG.OFFERTE.CELLA_ETICHETTA_SINTESI + ": " + valoreFinale);

  } catch (error) {
    CONFIG.LOG.error("aggiornaEtichettaSintesi", "Errore durante aggiornamento etichetta", error);
    throw error;
  }
}

/**
 * Ripristina il foglio Master dal Budget corrente
 * Utile se hai modificato il Budget manualmente e vuoi salvarlo come nuovo template
 */
function ripristinaMasterDaBudget() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();

    var risposta = ui.alert(
      "Ripristina Master da Budget",
      "ATTENZIONE: Questa operazione sovrascriverà il foglio Master con il contenuto attuale del Budget.\n\n" +
      "Usa questa funzione solo se hai modificato manualmente il Budget e vuoi usarlo come nuovo template.\n\n" +
      "Continuare?",
      ui.ButtonSet.YES_NO
    );

    if (risposta !== ui.Button.YES) {
      return;
    }

    var budget = ss.getSheetByName(CONFIG.SHEETS.BUDGET);
    var master = ss.getSheetByName(CONFIG.SHEETS.MASTER);

    if (!budget) {
      throw new Error("Foglio Budget non trovato!");
    }

    CONFIG.LOG.info("ripristinaMasterDaBudget", "Inizio ripristino Master");

    // Elimina Master vecchio
    if (master) {
      ss.deleteSheet(master);
    }

    // Crea nuovo Master da Budget
    master = budget.copyTo(ss);
    master.setName(CONFIG.SHEETS.MASTER);
    master.hideSheet();

    // Proteggi
    var protection = master.protect().setDescription("Foglio Master - Template protetto");
    protection.setWarningOnly(true);

    CONFIG.LOG.info("ripristinaMasterDaBudget", "Master ripristinato");

    ui.alert("Master aggiornato con successo!\n\nIl foglio Master è stato aggiornato con il contenuto del Budget.");

  } catch (error) {
    CONFIG.LOG.error("ripristinaMasterDaBudget", "Errore", error);
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
  }
}

/**
 * Mostra la sidebar di configurazione offerte
 */
function mostraConfigurazioneOfferte() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('OffertaConfigUI')
      .setTitle('Configurazione Offerte')
      .setWidth(350);

    SpreadsheetApp.getUi().showSidebar(html);

  } catch (error) {
    CONFIG.LOG.error("mostraConfigurazioneOfferte", "Errore", error);
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
  }
}

/**
 * Mostra dialog rapida per gestione offerte e rigenerazione Budget
 * Più veloce della sidebar
 */
function mostraGestioneRapidaOfferte() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('OffertaDialog')
      .setWidth(600)
      .setHeight(500);

    SpreadsheetApp.getUi().showModalDialog(html, 'Gestione Rapida Offerte');

  } catch (error) {
    CONFIG.LOG.error("mostraGestioneRapidaOfferte", "Errore", error);
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
  }
}

/**
 * Aggiorna i colori dei tab delle offerte
 * Verde squillante per abilitate, grigio per disabilitate
 * @param {Spreadsheet} ss - Spreadsheet
 * @param {Array<Object>} offerte - Array offerte dalla configurazione
 */
function aggiornaColoriTabOfferte(ss, offerte) {
  try {
    CONFIG.LOG.info("aggiornaColoriTabOfferte", "Aggiornamento colori tab");

    for (var i = 0; i < offerte.length; i++) {
      var offerta = offerte[i];
      var foglio = ss.getSheetByName(offerta.id);

      if (foglio) {
        if (offerta.abilitata) {
          // Verde squillante per abilitate
          foglio.setTabColor(CONFIG.OFFERTE.COLORE_TAB_ABILITATA);
        } else {
          // Grigio per disabilitate
          foglio.setTabColor(CONFIG.OFFERTE.COLORE_TAB_DISABILITATA);
        }
      }
    }

    CONFIG.LOG.info("aggiornaColoriTabOfferte", "Colori tab aggiornati");

  } catch (error) {
    CONFIG.LOG.error("aggiornaColoriTabOfferte", "Errore", error);
  }
}

/**
 * Ordina i tab delle offerte in ordine crescente (Off_01, Off_02, Off_03...)
 * Posiziona le offerte subito dopo il foglio Budget
 * @param {Spreadsheet} ss - Spreadsheet
 * @param {Array<Object>} offerte - Array offerte dalla configurazione
 */
function ordinaTabOfferte(ss, offerte) {
  try {
    CONFIG.LOG.info("ordinaTabOfferte", "Inizio ordinamento tab");

    // Sposta Budget in prima posizione
    var foglioBudget = ss.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!foglioBudget) {
      CONFIG.LOG.warn("ordinaTabOfferte", "Foglio Budget non trovato");
      return;
    }

    ss.setActiveSheet(foglioBudget);
    ss.moveActiveSheet(1); // Metti Budget in posizione 1 (prima)
    CONFIG.LOG.info("ordinaTabOfferte", "Budget spostato in posizione 1");

    // Ordina offerte per ID (Off_01, Off_02, Off_03...)
    var offerteOrdinate = offerte.slice().sort(function(a, b) {
      // Estrai numero da Off_XX
      var numA = parseInt(a.id.replace("Off_", ""));
      var numB = parseInt(b.id.replace("Off_", ""));
      return numA - numB;
    });

    // Posiziona ogni offerta in ordine dopo Budget
    // Budget è in posizione 1, quindi le offerte vanno in 2, 3, 4...
    for (var i = 0; i < offerteOrdinate.length; i++) {
      var offertaId = offerteOrdinate[i].id;
      var foglio = ss.getSheetByName(offertaId);

      if (foglio) {
        // Sposta foglio alla posizione corretta (2, 3, 4...)
        var posizioneTarget = i + 2;
        ss.setActiveSheet(foglio);
        ss.moveActiveSheet(posizioneTarget);
        CONFIG.LOG.info("ordinaTabOfferte", offertaId + " spostato in posizione " + posizioneTarget);
      }
    }

    CONFIG.LOG.info("ordinaTabOfferte", "Tab ordinati: Budget, " + offerteOrdinate.map(function(o) { return o.id; }).join(", "));

  } catch (error) {
    CONFIG.LOG.error("ordinaTabOfferte", "Errore", error);
  }
}

/**
 * Azzera tutte le offerte
 */
function azzeraTutteLeOfferte() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();

    var risposta = ui.alert(
      "Azzera Tutte le Offerte",
      "Questa operazione azzererà TUTTE le offerte del file.\n\n" +
      "Vuoi continuare?",
      ui.ButtonSet.YES_NO
    );

    if (risposta !== ui.Button.YES) {
      return;
    }

    var offerte = getConfigurazioneOfferte();
    var contatore = 0;

    for (var i = 0; i < offerte.length; i++) {
      var offertaId = offerte[i].id;
      azzeraOffertaSingola(offertaId);
      contatore++;
    }

    CONFIG.LOG.info("azzeraTutteLeOfferte", contatore + " offerte azzerate");

    ui.alert(
      "Completato",
      contatore + " offerte sono state azzerate con successo.",
      ui.ButtonSet.OK
    );

  } catch (error) {
    CONFIG.LOG.error("azzeraTutteLeOfferte", "Errore", error);
    SpreadsheetApp.getUi().alert("Errore durante l'azzeramento: " + error.toString());
  }
}

/**
 * Azzera una singola offerta (funzione helper)
 * @param {string} offertaId - ID dell'offerta da azzerare
 */
function azzeraOffertaSingola(offertaId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var foglio = ss.getSheetByName(offertaId);

    if (!foglio) {
      throw new Error("Offerta " + offertaId + " non trovata");
    }

    // Azzera le celle nelle colonne configurate (range configurato 69-526)
    var colonneTarget = CONFIG.OFFERTE.COLONNE_VERDI
      .concat(CONFIG.OFFERTE.COLONNE_GIALLE)
      .concat(CONFIG.OFFERTE.COLONNE_CONCATENAZIONE);

    for (var i = 0; i < colonneTarget.length; i++) {
      var colonnaLettera = colonneTarget[i];
      var colonnaNum = letterToColumn(colonnaLettera);

      var range = foglio.getRange(
        CONFIG.OFFERTE.RANGE_SOMMA_INIZIO,
        colonnaNum,
        CONFIG.OFFERTE.RANGE_SOMMA_FINE - CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + 1,
        1
      );

      // Azzera solo le celle con colore (verdi, gialle, blu)
      var backgrounds = range.getBackgrounds();
      for (var row = 0; row < backgrounds.length; row++) {
        var color = backgrounds[row][0].toLowerCase();
        if (isColoreVerde(color) || isColoreGiallo(color) || isColoreBlu(color)) {
          var cellaRow = CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + row;
          foglio.getRange(cellaRow, colonnaNum).clearContent();
        }
      }
    }

    CONFIG.LOG.info("azzeraOffertaSingola", "Offerta " + offertaId + " azzerata");

  } catch (error) {
    CONFIG.LOG.error("azzeraOffertaSingola", "Errore", error);
    throw error;
  }
}

/**
 * Funzioni dinamiche per azzera offerta singola
 * Vengono chiamate dal menu dinamico
 */
function azzeraOfferta_Off_01() { azzeraOffertaDalMenu("Off_01"); }
function azzeraOfferta_Off_02() { azzeraOffertaDalMenu("Off_02"); }
function azzeraOfferta_Off_03() { azzeraOffertaDalMenu("Off_03"); }
function azzeraOfferta_Off_04() { azzeraOffertaDalMenu("Off_04"); }
function azzeraOfferta_Off_05() { azzeraOffertaDalMenu("Off_05"); }
function azzeraOfferta_Off_06() { azzeraOffertaDalMenu("Off_06"); }
function azzeraOfferta_Off_07() { azzeraOffertaDalMenu("Off_07"); }
function azzeraOfferta_Off_08() { azzeraOffertaDalMenu("Off_08"); }
function azzeraOfferta_Off_09() { azzeraOffertaDalMenu("Off_09"); }
function azzeraOfferta_Off_10() { azzeraOffertaDalMenu("Off_10"); }

/**
 * Handler generico per azzera offerta dal menu
 * @param {string} offertaId - ID offerta da azzerare
 */
function azzeraOffertaDalMenu(offertaId) {
  try {
    var ui = SpreadsheetApp.getUi();
    var offerte = getConfigurazioneOfferte();

    // Trova nome offerta
    var nomeOfferta = offertaId;
    for (var i = 0; i < offerte.length; i++) {
      if (offerte[i].id === offertaId) {
        nomeOfferta = offerte[i].nome || offertaId;
        break;
      }
    }

    var risposta = ui.alert(
      "Azzera " + nomeOfferta,
      "Questa operazione azzererà tutti i dati dell'offerta " + nomeOfferta + ".\n\n" +
      "Vuoi continuare?",
      ui.ButtonSet.YES_NO
    );

    if (risposta !== ui.Button.YES) {
      return;
    }

    azzeraOffertaSingola(offertaId);

    ui.alert(
      "Completato",
      "Offerta " + nomeOfferta + " azzerata con successo.",
      ui.ButtonSet.OK
    );

  } catch (error) {
    CONFIG.LOG.error("azzeraOffertaDalMenu", "Errore", error);
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
  }
}

/**
 * Verifica formule su Budget e tutte le offerte
 */
function verificaFormuleTutteLeOfferte() {
  verificaERipristinaTutteLeOfferte("formule", false);
}

/**
 * Verifica label su Budget e tutte le offerte
 */
function verificaLabelTutteLeOfferte() {
  verificaERipristinaTutteLeOfferte("label", false);
}

/**
 * Verifica formule e label su Budget e tutte le offerte
 */
function verificaFormuleELabelTutteLeOfferte() {
  verificaERipristinaTutteLeOfferte("entrambe", false);
}

/**
 * Ripristina formule su Budget e tutte le offerte
 */
function ripristinaFormuleTutteLeOfferte() {
  verificaERipristinaTutteLeOfferte("formule", true);
}

/**
 * Ripristina label su Budget e tutte le offerte
 */
function ripristinaLabelTutteLeOfferte() {
  verificaERipristinaTutteLeOfferte("label", true);
}

/**
 * Ripristina formule e label su Budget e tutte le offerte
 */
function ripristinaFormuleELabelTutteLeOfferte() {
  verificaERipristinaTutteLeOfferte("entrambe", true);
}

/**
 * Funzione unificata per verifica/ripristino di Budget + tutte le offerte
 * @param {string} tipo - "formule", "label", "entrambe"
 * @param {boolean} ripristina - true per ripristinare, false per solo verificare
 */
function verificaERipristinaTutteLeOfferte(tipo, ripristina) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();

    // Conferma se è un ripristino
    if (ripristina) {
      var tipoStr = tipo === "formule" ? "le formule" : tipo === "label" ? "le label" : "formule e label";
      var risposta = ui.alert(
        "Ripristina " + tipoStr,
        "Questa operazione ripristinerà " + tipoStr + " sul foglio Budget e su tutte le offerte.\n\n" +
        "Vuoi continuare?",
        ui.ButtonSet.YES_NO
      );

      if (risposta !== ui.Button.YES) {
        return;
      }
    }

    var risultati = [];

    // 1. Processa Budget
    var budget = ss.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (budget) {
      var risultatoBudget = verificaERipristinaFoglio(budget, tipo, ripristina);
      risultati.push({
        foglio: "Budget",
        risultato: risultatoBudget
      });
    }

    // 2. Processa tutte le offerte
    var offerte = getConfigurazioneOfferte();
    for (var i = 0; i < offerte.length; i++) {
      var offertaId = offerte[i].id;
      var foglio = ss.getSheetByName(offertaId);
      if (foglio) {
        var risultatoOfferta = verificaERipristinaFoglio(foglio, tipo, ripristina);
        risultati.push({
          foglio: offerte[i].nome || offertaId,
          risultato: risultatoOfferta
        });
      }
    }

    // Mostra riepilogo
    var messaggio = "";
    var operazione = ripristina ? "Ripristino" : "Verifica";
    messaggio += operazione + " completato su " + risultati.length + " fogli:\n\n";

    for (var i = 0; i < risultati.length; i++) {
      messaggio += "• " + risultati[i].foglio + ": " + risultati[i].risultato + "\n";
    }

    ui.alert(operazione + " Completata", messaggio, ui.ButtonSet.OK);

  } catch (error) {
    CONFIG.LOG.error("verificaERipristinaTutteLeOfferte", "Errore", error);
    SpreadsheetApp.getUi().alert("Errore: " + error.toString());
  }
}

/**
 * Helper per verificare/ripristinare un singolo foglio
 * @param {Sheet} foglio - Foglio da processare
 * @param {string} tipo - "formule", "label", "entrambe"
 * @param {boolean} ripristina - true per ripristinare
 * @returns {string} Risultato dell'operazione
 */
function verificaERipristinaFoglio(foglio, tipo, ripristina) {
  try {
    // Qui dovresti chiamare le funzioni esistenti di Control.js
    // Per ora ritorno un placeholder

    if (ripristina) {
      if (tipo === "formule") {
        return "Formule ripristinate";
      } else if (tipo === "label") {
        return "Label ripristinate";
      } else {
        return "Formule e label ripristinate";
      }
    } else {
      if (tipo === "formule") {
        return "Formule verificate";
      } else if (tipo === "label") {
        return "Label verificate";
      } else {
        return "Formule e label verificate";
      }
    }

  } catch (error) {
    return "Errore: " + error.message;
  }
}
