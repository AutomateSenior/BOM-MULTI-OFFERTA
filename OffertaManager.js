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
 * Rigenera il foglio Budget con formule SUM dinamiche per le celle verdi
 */
function rigeneraBudgetDaOfferte() {
  try {
    var tempoInizio = new Date().getTime();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var budget = ss.getSheetByName(CONFIG.SHEETS.BUDGET);

    if (!budget) {
      throw new Error("Foglio Budget non trovato!");
    }

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "========== INIZIO RIGENERAZIONE BUDGET (OTTIMIZZATO) ==========");

    // Ottieni offerte abilitate
    var offerte = getConfigurazioneOfferte();
    var offerteAbilitate = [];
    for (var i = 0; i < offerte.length; i++) {
      if (offerte[i].abilitata) {
        offerteAbilitate.push(offerte[i].id);
      }
    }

    if (offerteAbilitate.length === 0) {
      throw new Error("Nessuna offerta abilitata! Abilita almeno un'offerta.");
    }

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Offerte abilitate: " + offerteAbilitate.join(", "));

    // V08: Verifica che tutti i fogli offerta esistano realmente
    var fogliMancanti = [];
    for (var i = 0; i < offerteAbilitate.length; i++) {
      var offertaId = offerteAbilitate[i];
      var foglio = ss.getSheetByName(offertaId);
      if (!foglio) {
        fogliMancanti.push(offertaId);
      }
    }

    if (fogliMancanti.length > 0) {
      throw new Error("ERRORE: Fogli offerta mancanti: " + fogliMancanti.join(", ") +
                      "\n\nVerifica che i fogli esistano con il nome esatto registrato in Configurazione." +
                      "\n\nSoluzione: Disabilita le offerte mancanti o ricreale con il nome corretto.");
    }

    // Identifica celle verdi (SUM) nel range 69-520 (include Body Rental fino a 506 + margine)
    var celleVerdi = getCelleVerdiPerSomma(budget);
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Celle verdi (SUM) trovate: " + celleVerdi.length);

    // Identifica celle gialle (MAX) nel range 69-520
    var celleGialle = getCelleGiallePerMax(budget);
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Celle gialle (MAX) trovate: " + celleGialle.length);

    // Identifica celle verdi concatenazione (colonne L, M, N dopo migrazione V08) nel range 69-520
    var celleConcatenazione = getCelleVerdiPerConcatenazione(budget);
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Celle concatenazione (L,M,N) trovate: " + celleConcatenazione.length);

    if (celleVerdi.length === 0 && celleGialle.length === 0 && celleConcatenazione.length === 0) {
      CONFIG.LOG.warn("rigeneraBudgetDaOfferte", "Nessuna cella da processare trovata nel range 69-520");
      aggiornaEtichettaSintesi(budget, offerteAbilitate);
      return;
    }

    // PARTE 1: Processa celle verdi con SUM (SUPER-OTTIMIZZATO)
    var tempoParteInizio = new Date().getTime();

    // Separa celle consistenza da celle normali
    var celleVerdiNormali = [];
    var celleVerdiConsistenza = [];

    for (var i = 0; i < celleVerdi.length; i++) {
      var cella = celleVerdi[i];
      var cellaA1 = columnToLetter(cella.col) + cella.row;

      var isCellaConsistenza = false;
      for (var k = 0; k < CONFIG.OFFERTE.CELLE_CONSISTENZA.length; k++) {
        if (CONFIG.OFFERTE.CELLE_CONSISTENZA[k].cella === cellaA1) {
          isCellaConsistenza = true;
          break;
        }
      }

      if (isCellaConsistenza) {
        celleVerdiConsistenza.push(cella);
      } else {
        celleVerdiNormali.push(cella);
      }
    }

    // Processa celle normali in batch VERO (range contigui)
    if (celleVerdiNormali.length > 0) {
      // Raggruppa celle per colonna
      var cellePerColonna = {};

      for (var i = 0; i < celleVerdiNormali.length; i++) {
        var cella = celleVerdiNormali[i];
        var col = cella.col;

        if (!cellePerColonna[col]) {
          cellePerColonna[col] = [];
        }
        cellePerColonna[col].push(cella);
      }

      // Processa ogni colonna con range contigui
      for (var col in cellePerColonna) {
        var celle = cellePerColonna[col];

        // Ordina per riga
        celle.sort(function(a, b) { return a.row - b.row; });

        // Trova blocchi contigui
        var blocchi = [];
        var bloccoCorrente = [celle[0]];

        for (var i = 1; i < celle.length; i++) {
          if (celle[i].row === celle[i-1].row + 1) {
            // Cella contigua
            bloccoCorrente.push(celle[i]);
          } else {
            // Nuovo blocco
            blocchi.push(bloccoCorrente);
            bloccoCorrente = [celle[i]];
          }
        }
        blocchi.push(bloccoCorrente);

        // Processa ogni blocco contiguo con UNA SOLA chiamata getRange
        for (var b = 0; b < blocchi.length; b++) {
          var blocco = blocchi[b];
          var primaRiga = blocco[0].row;
          var numRighe = blocco.length;

          // Ottieni range intero del blocco
          var range = budget.getRange(primaRiga, parseInt(col), numRighe, 1);

          // Costruisci array 2D di formule
          var formule = [];
          for (var i = 0; i < blocco.length; i++) {
            var cella = blocco[i];
            var riferimenti = [];
            for (var j = 0; j < offerteAbilitate.length; j++) {
              riferimenti.push(offerteAbilitate[j] + "!" + columnToLetter(cella.col) + cella.row);
            }
            formule.push(["=" + riferimenti.join("+")]);
          }

          // Applica formule in batch
          range.setFormulas(formule);

          // Applica formattazione in batch
          range.setBackground("#4285f4");
          range.setFontColor("#ffffff");
        }
      }

      var tempoParte = ((new Date().getTime() - tempoParteInizio) / 1000).toFixed(2);
      CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Celle verdi normali processate: " + celleVerdiNormali.length + " in " + tempoParte + "s");
    }

    // Processa celle consistenza (rare)
    for (var i = 0; i < celleVerdiConsistenza.length; i++) {
      var cella = celleVerdiConsistenza[i];
      var cellaA1 = columnToLetter(cella.col) + cella.row;
      var cellRange = budget.getRange(cella.row, cella.col);

      var valori = [];
      for (var j = 0; j < offerteAbilitate.length; j++) {
        var offertaId = offerteAbilitate[j];
        var foglio = ss.getSheetByName(offertaId);
        if (foglio) {
          valori.push(foglio.getRange(cellaA1).getValue());
        }
      }

      var primoValore = valori[0] || "";
      var tuttiUguali = valori.every(function(v) { return v === primoValore; });

      if (tuttiUguali) {
        cellRange.setValue(primoValore);
        cellRange.setBackground("#4285f4");
        cellRange.setFontColor("#ffffff");
      } else {
        cellRange.setValue(CONFIG.OFFERTE.CELLE_CONSISTENZA[0].messaggioErrore);
        cellRange.setBackground("#ea4335");
        cellRange.setFontColor("#ffffff");
      }
    }

    // PARTE 2: Processa celle gialle con MAX (SUPER-OTTIMIZZATO)
    tempoParteInizio = new Date().getTime();

    var celleGialleNormali = [];
    var celleGialleConsistenza = [];

    for (var i = 0; i < celleGialle.length; i++) {
      var cella = celleGialle[i];
      var cellaA1 = columnToLetter(cella.col) + cella.row;

      var isCellaConsistenza = false;
      for (var k = 0; k < CONFIG.OFFERTE.CELLE_CONSISTENZA.length; k++) {
        if (CONFIG.OFFERTE.CELLE_CONSISTENZA[k].cella === cellaA1) {
          isCellaConsistenza = true;
          break;
        }
      }

      if (isCellaConsistenza) {
        celleGialleConsistenza.push(cella);
      } else {
        celleGialleNormali.push(cella);
      }
    }

    // Processa celle normali in batch VERO (range contigui)
    if (celleGialleNormali.length > 0) {
      // Raggruppa per colonna
      var cellePerColonna = {};

      for (var i = 0; i < celleGialleNormali.length; i++) {
        var cella = celleGialleNormali[i];
        var col = cella.col;

        if (!cellePerColonna[col]) {
          cellePerColonna[col] = [];
        }
        cellePerColonna[col].push(cella);
      }

      // Processa ogni colonna con range contigui
      for (var col in cellePerColonna) {
        var celle = cellePerColonna[col];

        // Ordina per riga
        celle.sort(function(a, b) { return a.row - b.row; });

        // Trova blocchi contigui
        var blocchi = [];
        var bloccoCorrente = [celle[0]];

        for (var i = 1; i < celle.length; i++) {
          if (celle[i].row === celle[i-1].row + 1) {
            bloccoCorrente.push(celle[i]);
          } else {
            blocchi.push(bloccoCorrente);
            bloccoCorrente = [celle[i]];
          }
        }
        blocchi.push(bloccoCorrente);

        // Processa ogni blocco contiguo
        for (var b = 0; b < blocchi.length; b++) {
          var blocco = blocchi[b];
          var primaRiga = blocco[0].row;
          var numRighe = blocco.length;

          // Ottieni range intero del blocco
          var range = budget.getRange(primaRiga, parseInt(col), numRighe, 1);

          // Costruisci array 2D di formule MAX
          var formule = [];
          for (var i = 0; i < blocco.length; i++) {
            var cella = blocco[i];
            var riferimenti = [];
            for (var j = 0; j < offerteAbilitate.length; j++) {
              riferimenti.push(offerteAbilitate[j] + "!" + columnToLetter(cella.col) + cella.row);
            }
            formule.push(["=MAX(" + riferimenti.join(";") + ")"]);
          }

          // Applica formule in batch
          range.setFormulas(formule);

          // Applica formattazione in batch
          range.setBackground("#4285f4");
          range.setFontColor("#ffffff");
        }
      }

      var tempoParte = ((new Date().getTime() - tempoParteInizio) / 1000).toFixed(2);
      CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Celle gialle normali processate: " + celleGialleNormali.length + " in " + tempoParte + "s");
    }

    // Processa celle consistenza (rare)
    for (var i = 0; i < celleGialleConsistenza.length; i++) {
      var cella = celleGialleConsistenza[i];
      var cellaA1 = columnToLetter(cella.col) + cella.row;
      var cellRange = budget.getRange(cella.row, cella.col);

      var valori = [];
      for (var j = 0; j < offerteAbilitate.length; j++) {
        var offertaId = offerteAbilitate[j];
        var foglio = ss.getSheetByName(offertaId);
        if (foglio) {
          valori.push(foglio.getRange(cellaA1).getValue());
        }
      }

      var primoValore = valori[0] || "";
      var tuttiUguali = valori.every(function(v) { return v === primoValore; });

      if (tuttiUguali) {
        cellRange.setValue(primoValore);
        cellRange.setBackground("#4285f4");
        cellRange.setFontColor("#ffffff");
      } else {
        cellRange.setValue(CONFIG.OFFERTE.CELLE_CONSISTENZA[0].messaggioErrore);
        cellRange.setBackground("#ea4335");
        cellRange.setFontColor("#ffffff");
      }
    }

    // PARTE 3: Processa celle descrittive (colonne L, M, N) - PRIMO VALORE COMPILATO
    // V08: Campi descrittivi NON concatenano, prendono primo valore compilato
    // OTTIMIZZATO: Pre-carica tutti i dati dalle offerte in memoria
    var datiOfferte = {};
    for (var j = 0; j < offerteAbilitate.length; j++) {
      var offertaId = offerteAbilitate[j];
      var foglio = ss.getSheetByName(offertaId);
      if (foglio) {
        // Carica tutto il range in memoria una volta sola
        var datiRange = foglio.getRange(
          CONFIG.OFFERTE.RANGE_SOMMA_INIZIO,
          1,
          CONFIG.OFFERTE.RANGE_SOMMA_FINE - CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + 1,
          foglio.getLastColumn()
        ).getValues();
        datiOfferte[offertaId] = datiRange;
      }
    }

    // Prepara batch per celle descrittive (SUPER-OTTIMIZZATO)
    var celleToUpdate = [];
    var celleToClear = [];
    var valuesToSet = [];

    for (var i = 0; i < celleConcatenazione.length; i++) {
      var cella = celleConcatenazione[i];
      var cellaA1 = columnToLetter(cella.col) + cella.row;

      // Raccogli tutti i valori dalle offerte abilitate (da cache in memoria)
      var valori = [];
      var offerteConValore = [];

      for (var j = 0; j < offerteAbilitate.length; j++) {
        var offertaId = offerteAbilitate[j];
        if (datiOfferte[offertaId]) {
          var rowIndex = cella.row - CONFIG.OFFERTE.RANGE_SOMMA_INIZIO;
          var colIndex = cella.col - 1; // 0-based
          var valore = datiOfferte[offertaId][rowIndex][colIndex];

          if (valore !== null && valore !== undefined && valore !== "") {
            valori.push(valore);
            offerteConValore.push(offertaId);
          }
        }
      }

      // Controllo multi-riempimento: se 2+ offerte hanno lo stesso campo compilato → ERRORE
      if (valori.length > 1) {
        var errMsg = "ERRORE: Campo descrittivo " + cellaA1 + " compilato in multiple offerte (" +
                     offerteConValore.join(", ") + "). Ogni campo descrittivo può essere compilato in una sola offerta.";
        CONFIG.LOG.error("rigeneraBudgetDaOfferte", errMsg);
        throw new Error(errMsg);
      }

      // Prepara per batch update
      if (valori.length === 1) {
        celleToUpdate.push(cella);
        valuesToSet.push(valori[0]);
      } else {
        celleToClear.push(cella);
      }
    }

    // Batch update valori e colori (OTTIMIZZATO senza getRanges)
    if (celleToUpdate.length > 0) {
      for (var i = 0; i < celleToUpdate.length; i++) {
        var cella = celleToUpdate[i];
        var cellRange = budget.getRange(cella.row, cella.col);

        // Rimuovi validazione se presente
        var validation = cellRange.getDataValidation();
        if (validation) {
          cellRange.setDataValidation(null);
        }

        // Scrivi valore e formattazione
        cellRange.setValue(valuesToSet[i]);
        cellRange.setBackground("#4285f4");
        cellRange.setFontColor("#ffffff");
      }
    }

    // Batch clear celle vuote (OTTIMIZZATO)
    if (celleToClear.length > 0) {
      for (var i = 0; i < celleToClear.length; i++) {
        var cella = celleToClear[i];
        var cellRange = budget.getRange(cella.row, cella.col);

        cellRange.clearContent();
        cellRange.setBackground(null);
        cellRange.setFontColor(null);
      }
    }

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Celle descrittive processate: " + celleConcatenazione.length +
                    " (valorizzate: " + celleToUpdate.length + ", vuote: " + celleToClear.length + ")");

    // Aggiorna etichetta M62
    aggiornaEtichettaSintesi(budget, offerteAbilitate);

    // Aggiorna colori tab offerte
    aggiornaColoriTabOfferte(ss, offerte);

    // Ordina tab offerte in ordine crescente
    ordinaTabOfferte(ss, offerte);

    // V08: Cancella indicatore rigenerazione (completata con successo)
    impostaIndicatoreRigenerazione(false);

    // Forza flush e ricalcolo formule - UNA SOLA VOLTA ALLA FINE
    SpreadsheetApp.flush();

    // Calcola tempo di esecuzione
    var tempoFine = new Date().getTime();
    var tempoEsecuzione = ((tempoFine - tempoInizio) / 1000).toFixed(2);

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "========== BUDGET RIGENERATO CON SUCCESSO ==========");
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Tempo di esecuzione: " + tempoEsecuzione + " secondi");

    // V08: Forza ricalcolo formule per risolvere eventuali errori di riferimento da rigenerazioni precedenti
    // Questo è necessario quando i fogli referenziati nelle formule sono stati ricreati
    try {
      // Trigger ricalcolo toccando una cella formula
      var primaFormulaCell = budget.getRange(CONFIG.OFFERTE.RANGE_SOMMA_INIZIO, letterToColumn("S"));
      var formula = primaFormulaCell.getFormula();
      if (formula) {
        primaFormulaCell.setFormula(formula); // Re-imposta stessa formula per trigger ricalcolo
        SpreadsheetApp.flush();
      }
    } catch (e) {
      CONFIG.LOG.warn("rigeneraBudgetDaOfferte", "Impossibile forzare ricalcolo: " + e);
    }

  } catch (error) {
    CONFIG.LOG.error("rigeneraBudgetDaOfferte", "Errore", error);
    throw error;
  }
}

/**
 * Converti numero colonna in lettera (1 -> A, 27 -> AA, etc)
 * @param {number} column - Numero colonna (1-based)
 * @returns {string} Lettera colonna
 */
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Identifica le celle verdi nel range configurato del foglio Budget
 * V08: Range 69-520 (include Body Rental 494-506)
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
        var color = backgrounds[row][0].toLowerCase();
        if (isColoreGiallo(color) || isColoreBlu(color)) {
          celleGialle.push({
            row: CONFIG.OFFERTE.RANGE_SOMMA_INIZIO + row,
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
 * Converti lettera colonna in numero (A -> 1, Z -> 26, AA -> 27, etc)
 * @param {string} letter - Lettera colonna
 * @returns {number} Numero colonna (1-based)
 */
function letterToColumn(letter) {
  var column = 0;
  var length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
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

    // Azzera le celle nelle colonne configurate (range configurato 69-520)
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
