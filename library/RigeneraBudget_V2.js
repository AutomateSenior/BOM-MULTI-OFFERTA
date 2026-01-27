/**
 * NUOVA VERSIONE rigeneraBudgetDaOfferte() - V2 con logica complessa
 *
 * Logica:
 * - Per ogni riga con celle verdi riempite
 * - Controlla colonna M in tutte le offerte
 * - Se M uguale → processa L (concat), N (sum se >519, concat se <=519), S/T (sum)
 * - Se M diverso → ERRORE in M, salta altre colonne
 * - Celle gialle U → MAX (invariato)
 * - Controllo consistenza S476
 */


/**
 * Rigenera Budget con nuova logica complessa - VERSIONE OTTIMIZZATA
 */
function rigeneraBudgetDaOfferte() {
  try {
    var tempoInizio = new Date().getTime();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var budget = ss.getSheetByName(CONFIG.SHEETS.BUDGET);

    if (!budget) {
      throw new Error("Foglio Budget non trovato!");
    }

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "========== INIZIO RIGENERAZIONE BUDGET V2 (LOGICA COMPLESSA) ==========");
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Versione Libreria BOM8: " + CONFIG.LIB_VERSION);

    // 1. Ottieni offerte abilitate
    var offerte = getConfigurazioneOfferte();
    var offerteAbilitate = [];
    for (var i = 0; i < offerte.length; i++) {
      if (offerte[i].abilitata) {
        offerteAbilitate.push(offerte[i].id);
      }
    }

    if (offerteAbilitate.length === 0) {
      throw new Error("Nessuna offerta abilitata!");
    }

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Offerte abilitate: " + offerteAbilitate.join(", "));

    // 2. Verifica esistenza fogli
    for (var i = 0; i < offerteAbilitate.length; i++) {
      if (!ss.getSheetByName(offerteAbilitate[i])) {
        throw new Error("Foglio mancante: " + offerteAbilitate[i]);
      }
    }

    // 3. PRE-CARICA TUTTI I DATI DALLE OFFERTE (OTTIMIZZAZIONE)
    var rangeInizio = CONFIG.OFFERTE.RANGE_SOMMA_INIZIO;
    var rangeFine = CONFIG.OFFERTE.RANGE_SOMMA_FINE;
    var numRighe = rangeFine - rangeInizio + 1;

    var offerteData = {};
    var colL = letterToColumn("L");
    var colM = letterToColumn("M");
    var colN = letterToColumn("N");
    var colS = letterToColumn("S");
    var colT = letterToColumn("T");
    var colU = letterToColumn("U");

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Pre-caricamento dati da " + offerteAbilitate.length + " offerte...");

    for (var i = 0; i < offerteAbilitate.length; i++) {
      var offertaId = offerteAbilitate[i];
      var foglio = ss.getSheetByName(offertaId);
      var dati = foglio.getRange(rangeInizio, 1, numRighe, foglio.getLastColumn()).getValues();
      offerteData[offertaId] = dati;
    }

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Dati pre-caricati in memoria");

    // 4. PROCESSA CELLE VERDI (L, M, N, S, T) - LOGICA COMPLESSA
    var tempoVerdiInizio = new Date().getTime();

    // Prepara array per batch update
    var aggiornamenti = {
      L: [], M: [], N: [], S: [], T: []
    };
    var formattazione = {
      L: [], M: [], N: [], S: [], T: []
    };

    // PRE-CARICA COLORI DELLE CELLE NEL BUDGET (per controllo verde/blu)
    var coloriBudget = {};
    var colonne = [colL, colM, colN, colS, colT];
    for (var c = 0; c < colonne.length; c++) {
      var col = colonne[c];
      var backgrounds = budget.getRange(rangeInizio, col, numRighe, 1).getBackgrounds();
      for (var i = 0; i < backgrounds.length; i++) {
        var riga = rangeInizio + i;
        var chiave = riga + "_" + col;
        coloriBudget[chiave] = backgrounds[i][0].toLowerCase();
      }
    }

    // Identifica righe con celle verdi/blu/rosse in Budget
    var righeConDati = {};
    for (var i = 0; i < numRighe; i++) {
      var riga = rangeInizio + i;

      // Controlla se almeno una cella è verde, blu o rossa-errore (errori da correggere) nel Budget
      var haVerdeBluRosso = false;
      for (var c = 0; c < colonne.length; c++) {
        var col = colonne[c];
        var chiave = riga + "_" + col;
        var colore = coloriBudget[chiave];
        // Include anche celle rosse-errore (errori precedenti da correggere)
        if (isColoreVerde(colore) || isColoreBlu(colore) || isColoreRossoErrore(colore)) {
          haVerdeBluRosso = true;
          break;
        }
      }

      if (haVerdeBluRosso) {
        righeConDati[riga] = true;
      }
    }

    var righe = Object.keys(righeConDati).map(function(r) { return parseInt(r); }).sort(function(a,b) { return a-b; });
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Righe con celle verdi/blu/rosse: " + righe.length);

    // Processa ogni riga
    var righeErrore = 0;
    var righeOk = 0;
    var erroriM = [];  // Lista errori M per dialog finale

    for (var r = 0; r < righe.length; r++) {
      var riga = righe[r];
      var indiceDati = riga - rangeInizio;

      // STEP 1: Controlla consistenza colonna M
      var valoriM = [];
      for (var i = 0; i < offerteAbilitate.length; i++) {
        var offertaId = offerteAbilitate[i];
        valoriM.push(offerteData[offertaId][indiceDati][colM-1]);
      }

      // Filtra valori non vuoti per il controllo di consistenza
      var valoriMNonVuoti = valoriM.filter(function(v) { return v !== "" && v !== null && v !== undefined; });

      var mConsistente = true;
      var primoM = "";

      if (valoriMNonVuoti.length > 0) {
        // Ci sono valori non vuoti, controlla che siano tutti uguali
        primoM = valoriMNonVuoti[0];
        mConsistente = valoriMNonVuoti.every(function(v) { return v === primoM; });
      } else {
        // Tutti vuoti, consideralo consistente
        primoM = "";
        mConsistente = true;
      }

      if (!mConsistente) {
        // M DIVERSO → ERRORE
        // Crea messaggio dettagliato con valori diversi
        var dettagliValori = [];
        for (var i = 0; i < offerteAbilitate.length; i++) {
          dettagliValori.push(offerteAbilitate[i] + ": \"" + valoriM[i] + "\"");
        }
        var messaggioErrore = "Tipo risorse diverse nelle diverse offerte nella riga " + riga;

        aggiornamenti.M.push({riga: riga, valore: messaggioErrore, isErrore: true});
        erroriM.push({
          riga: riga,
          valori: dettagliValori.join(", ")
        });
        righeErrore++;
        continue;  // Salta altre colonne per questa riga
      }

      // M UGUALE → Processa le colonne che sono verdi/blu
      righeOk++;

      // Controllo colore per ogni colonna PRIMA di aggiungere aggiornamento

      // Colonna M - Valore unico (SEMPRE processata quando valori sono uguali)
      var chiaveM = riga + "_" + colM;
      var coloreM = coloriBudget[chiaveM];

      // Log dettagliato per debug riga 119
      if (riga === 119) {
        CONFIG.LOG.info("rigeneraBudgetDaOfferte", "DEBUG M119 - coloreM: " + coloreM + ", primoM: '" + primoM + "', verde: " + isColoreVerde(coloreM) + ", blu: " + isColoreBlu(coloreM) + ", rosso: " + isColoreRossoErrore(coloreM));
      }

      // Processa sempre quando i valori M sono tutti uguali (non c'è errore)
      aggiornamenti.M.push({riga: riga, valore: primoM, isErrore: false, colonnaM: true});
      if (riga === 119) {
        CONFIG.LOG.info("rigeneraBudgetDaOfferte", "DEBUG M119 - Aggiunto aggiornamento M con valore: '" + primoM + "'");
      }

      // Colonna L - CONCATENAZIONE (solo se verde/blu)
      var chiaveL = riga + "_" + colL;
      var coloreL = coloriBudget[chiaveL];
      if (isColoreVerde(coloreL) || isColoreBlu(coloreL)) {
        var valoriL = [];
        for (var i = 0; i < offerteAbilitate.length; i++) {
          var val = offerteData[offerteAbilitate[i]][indiceDati][colL-1];
          if (val && String(val).trim() !== "") {
            valoriL.push(String(val));
          }
        }
        aggiornamenti.L.push({riga: riga, valore: valoriL.join(", ")});
      }

      // Colonna N - SUM se >519, CONCATENAZIONE altrimenti (solo se verde/blu)
      var chiaveN = riga + "_" + colN;
      var coloreN = coloriBudget[chiaveN];
      if (isColoreVerde(coloreN) || isColoreBlu(coloreN)) {
        if (riga > 519) {
          // SUM - Formula
          var riferimenti = [];
          for (var i = 0; i < offerteAbilitate.length; i++) {
            riferimenti.push(offerteAbilitate[i] + "!N" + riga);
          }
          aggiornamenti.N.push({riga: riga, formula: "=" + riferimenti.join("+")});
        } else {
          // CONCATENAZIONE - Valore
          var valoriN = [];
          for (var i = 0; i < offerteAbilitate.length; i++) {
            var val = offerteData[offerteAbilitate[i]][indiceDati][colN-1];
            if (val && String(val).trim() !== "") {
              valoriN.push(String(val));
            }
          }
          aggiornamenti.N.push({riga: riga, valore: valoriN.join(", ")});
        }
      }

      // Colonne S e T - SUM (Formula) - solo se verde/blu ed escludi S476
      ["S", "T"].forEach(function(colLetter) {
        var colNum = letterToColumn(colLetter);
        var chiave = riga + "_" + colNum;
        var colore = coloriBudget[chiave];

        // Escludi S476 (ha gestione speciale, era S476 prima di Sviluppo Software)
        if (riga === 476 && colLetter === "S") {
          return;
        }

        if (isColoreVerde(colore) || isColoreBlu(colore)) {
          var riferimenti = [];
          for (var i = 0; i < offerteAbilitate.length; i++) {
            riferimenti.push(offerteAbilitate[i] + "!" + colLetter + riga);
          }
          aggiornamenti[colLetter].push({riga: riga, formula: "=" + riferimenti.join("+")});
        }
      });
    }

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Righe OK: " + righeOk + ", Righe ERRORE: " + righeErrore);

    // APPLICA AGGIORNAMENTI IN BATCH (SUPER-OTTIMIZZATO)
    ["L", "M", "N", "S", "T"].forEach(function(colLetter) {
      var col = letterToColumn(colLetter);
      var updates = aggiornamenti[colLetter];

      if (updates.length === 0) return;

      for (var i = 0; i < updates.length; i++) {
        var upd = updates[i];
        var cell = budget.getRange(upd.riga, col);

        // Usa il colore pre-caricato
        var chiave = upd.riga + "_" + col;
        var coloreAttuale = coloriBudget[chiave];
        var eraVerde = isColoreVerde(coloreAttuale);

        // Log debug per riga 119
        if (upd.riga === 119 && upd.colonnaM) {
          CONFIG.LOG.info("rigeneraBudgetDaOfferte", "DEBUG M119 SCRITTURA - isErrore: " + upd.isErrore + ", valore: '" + upd.valore + "', coloreAttuale: " + coloreAttuale);
        }

        // Gestione errori: usa NOTE invece di setValue per evitare conflitti con validazione dati
        if (upd.isErrore) {
          // Cancella contenuto e metti messaggio come NOTA
          cell.clearContent();
          cell.setNote(upd.valore);
          cell.setBackground("#ea4335");  // Rosso per errori
          cell.setFontColor("#ffffff");
        } else {
          // Rimuovi eventuali note precedenti (da errori corretti)
          cell.clearNote();

          // Scrivi valore o formula normalmente
          if (upd.formula) {
            cell.setFormula(upd.formula);
          } else {
            cell.setValue(upd.valore || "");
            if (upd.riga === 119 && upd.colonnaM) {
              CONFIG.LOG.info("rigeneraBudgetDaOfferte", "DEBUG M119 SCRITTURA - Valore scritto: '" + upd.valore + "'");
            }
          }

          // Formattazione speciale per colonna M
          if (upd.colonnaM) {
            // Se era verde o blu, mantieni colore originale
            if (isColoreVerde(coloreAttuale) || isColoreBlu(coloreAttuale)) {
              // Non cambiare colore
              if (upd.riga === 119) {
                CONFIG.LOG.info("rigeneraBudgetDaOfferte", "DEBUG M119 SCRITTURA - Mantenuto colore verde/blu");
              }
            } else if (isColoreRossoErrore(coloreAttuale)) {
              // Era rosso-errore, ora corretto → grigio chiaro
              cell.setBackground("#d9d9d9");  // Grigio chiaro
              cell.setFontColor("#000000");   // Testo nero
              if (upd.riga === 119) {
                CONFIG.LOG.info("rigeneraBudgetDaOfferte", "DEBUG M119 SCRITTURA - Cambiato da rosso a grigio");
              }
            } else {
              // Altri colori → grigio chiaro
              cell.setBackground("#d9d9d9");  // Grigio chiaro
              cell.setFontColor("#000000");   // Testo nero
              if (upd.riga === 119) {
                CONFIG.LOG.info("rigeneraBudgetDaOfferte", "DEBUG M119 SCRITTURA - Impostato grigio (altro colore)");
              }
            }
          } else {
            // Altre colonne: colora in BLU solo se era VERDE
            if (eraVerde) {
              cell.setBackground("#4285f4");  // Blu solo se era verde
              cell.setFontColor("#ffffff");
            }
            // Se era già blu, mantieni il colore originale
          }
        }
      }
    });

    var tempoVerdi = ((new Date().getTime() - tempoVerdiInizio) / 1000).toFixed(2);
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Celle verdi processate in " + tempoVerdi + "s");

    // 5. PROCESSA CELLE GIALLE (U) - MAX (INVARIATO)
    var tempoGialleInizio = new Date().getTime();
    var celleGialle = getCelleGiallePerMax(budget);
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Celle gialle (MAX): " + celleGialle.length);

    for (var i = 0; i < celleGialle.length; i++) {
      var cella = celleGialle[i];
      var cellAddress = columnToLetter(cella.col) + cella.row;

      var riferimenti = [];
      for (var j = 0; j < offerteAbilitate.length; j++) {
        riferimenti.push(offerteAbilitate[j] + "!" + cellAddress);
      }

      var cellRange = budget.getRange(cella.row, cella.col);
      cellRange.setFormula("=MAX(" + riferimenti.join(";") + ")");
      cellRange.setBackground("#4285f4");
      cellRange.setFontColor("#ffffff");
    }

    var tempoGialle = ((new Date().getTime() - tempoGialleInizio) / 1000).toFixed(2);
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Celle gialle processate in " + tempoGialle + "s");

    // 5b. GESTIONE SPECIALE S64 e S65 (fuori dal range principale 69-526)
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "=== INIZIO GESTIONE S64 e S65 ===");
    [64, 65].forEach(function(riga) {
      var cellS = budget.getRange("S" + riga);
      var coloreS = cellS.getBackground().toLowerCase();
      var eraVerdeBluS = isColoreVerde(coloreS) || isColoreBlu(coloreS);
      var eraVerdeS = isColoreVerde(coloreS);

      CONFIG.LOG.info("rigeneraBudgetDaOfferte", "S" + riga + " - Colore: " + coloreS + ", isVerde: " + eraVerdeS + ", isVerdeBlu: " + eraVerdeBluS);

      if (eraVerdeBluS) {
        // Calcola la somma dai dati delle offerte
        var somma = 0;
        var dettagliValori = [];

        for (var i = 0; i < offerteAbilitate.length; i++) {
          var offertaId = offerteAbilitate[i];
          var foglio = ss.getSheetByName(offertaId);
          var valore = foglio.getRange("S" + riga).getValue();

          CONFIG.LOG.info("rigeneraBudgetDaOfferte", "  " + offertaId + "!S" + riga + " = '" + valore + "' (tipo: " + typeof valore + ", isNaN: " + isNaN(valore) + ")");

          if (valore && !isNaN(valore)) {
            somma += Number(valore);
            dettagliValori.push(offertaId + ":" + valore);
          } else {
            dettagliValori.push(offertaId + ":VUOTO_O_NON_NUMERO");
          }
        }

        CONFIG.LOG.info("rigeneraBudgetDaOfferte", "  Somma calcolata: " + somma + " [" + dettagliValori.join(", ") + "]");

        // Scrivi il VALORE (non la formula)
        cellS.setValue(somma);
        CONFIG.LOG.info("rigeneraBudgetDaOfferte", "  Valore scritto in Budget!S" + riga + ": " + somma);

        // Mantieni il colore VERDE
        if (isColoreVerde(coloreS)) {
          cellS.setBackground("#00ff00");
          cellS.setFontColor("#000000");
          CONFIG.LOG.info("rigeneraBudgetDaOfferte", "  Colore mantenuto VERDE");
        } else {
          CONFIG.LOG.info("rigeneraBudgetDaOfferte", "  Colore NON cambiato (era BLU)");
        }
      } else {
        CONFIG.LOG.warn("rigeneraBudgetDaOfferte", "S" + riga + " SALTATA - colore non verde/blu: " + coloreS);
      }
    });
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "=== FINE GESTIONE S64 e S65 ===");

    // 6. CONTROLLO CONSISTENZA S476
    var cellConsistenza = budget.getRange("S476");
    var valoriS476 = [];
    for (var i = 0; i < offerteAbilitate.length; i++) {
      var foglio = ss.getSheetByName(offerteAbilitate[i]);
      valoriS476.push(foglio.getRange("S476").getValue());
    }

    // Filtra valori non vuoti
    var valoriS476NonVuoti = [];
    for (var i = 0; i < valoriS476.length; i++) {
      if (valoriS476[i] && String(valoriS476[i]).trim() !== "") {
        valoriS476NonVuoti.push(String(valoriS476[i]).trim());
      }
    }

    if (valoriS476NonVuoti.length === 0) {
      // Nessuna offerta ha un valore → lascia vuoto in Budget SENZA formattazione
      cellConsistenza.clearContent();
      // Non applicare formattazione per evitare conflitti con validazione dati
      CONFIG.LOG.info("rigeneraBudgetDaOfferte", "S476: nessun valore nelle offerte, lasciato vuoto");
    } else {
      // Controlla se tutti i valori non vuoti sono uguali
      var primoValore = valoriS476NonVuoti[0];
      var tuttiUguali = valoriS476NonVuoti.every(function(v) { return v === primoValore; });

      if (tuttiUguali) {
        // Tutti uguali → scrivi il valore
        cellConsistenza.setValue(primoValore);
        cellConsistenza.setBackground("#4285f4");
        cellConsistenza.setFontColor("#ffffff");
        CONFIG.LOG.info("rigeneraBudgetDaOfferte", "S476: " + primoValore);
      } else {
        // Valori diversi → ERRORE
        cellConsistenza.setValue("Tipo di assistenza incoerente tra le diverse offerte");
        cellConsistenza.setBackground("#ea4335");
        cellConsistenza.setFontColor("#ffffff");
        CONFIG.LOG.warn("rigeneraBudgetDaOfferte", "S476: ERRORE - valori diversi: " + valoriS476NonVuoti.join(" ≠ "));
      }
    }

    // 7. Aggiorna etichetta sintesi
    aggiornaEtichettaSintesi(budget, offerteAbilitate);

    // 8. Rilancia controlla() per allineamento BOM (come quando si edita L56)
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Avvio allineamento BOM (controlla)...");
    try {
      controlla(ss, false);
      CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Allineamento BOM completato");
    } catch (errorControlla) {
      CONFIG.LOG.error("rigeneraBudgetDaOfferte", "Errore in controlla()", errorControlla);
      // Non bloccare l'esecuzione se controlla() fallisce
    }

    // 9. Mostra dialog con errori M se presenti
    if (erroriM.length > 0) {
      var messaggioDialog = "ATTENZIONE: Trovati " + erroriM.length + " errori di tipo risorse diverse:\n\n";
      for (var i = 0; i < erroriM.length; i++) {
        messaggioDialog += "• Riga " + erroriM[i].riga + ": " + erroriM[i].valori + "\n";
      }
      messaggioDialog += "\nLe celle con errori sono evidenziate in ROSSO nella colonna M.";

      var ui = SpreadsheetApp.getUi();
      ui.alert(
        "Errori Tipo Risorse",
        messaggioDialog,
        ui.ButtonSet.OK
      );
      CONFIG.LOG.warn("rigeneraBudgetDaOfferte", "Dialog errori M mostrata: " + erroriM.length + " errori");
    }

    // 10. Posiziona sul foglio Budget
    ss.setActiveSheet(budget);

    var tempoTotale = ((new Date().getTime() - tempoInizio) / 1000).toFixed(2);
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "========== RIGENERAZIONE COMPLETATA IN " + tempoTotale + "s ==========");

  } catch (error) {
    CONFIG.LOG.error("rigeneraBudgetDaOfferte", "ERRORE CRITICO", error);
    throw error;
  }
}
