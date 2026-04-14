/**
 * RigeneraBudget_V2.js
 *
 * Logica rigenerazione Budget dalle offerte Off_XXX:
 * - O, P, Q (righe 78-527): formula SUM di tutte le Off_XXX
 * - L (celle blu): concatenazione deduplicata delle Off_XXX
 * - M (se Q ≠ 0): tutte le Off_XXX con Q ≠ 0 devono avere stesso M → scritto in Budget; altrimenti errore bloccante
 * - N fino a riga 519: come M (stesso controllo, errore bloccante)
 * - N da riga 521: se Q ≠ 0 → somma N di tutte le Off_XXX
 */

function rigeneraBudgetDaOfferte() {
  var RANGE_INIZIO = 78;
  var RANGE_FINE   = 527;
  var numRighe     = RANGE_FINE - RANGE_INIZIO + 1;

  var tempoInizio = new Date().getTime();

  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var budget = ss.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!budget) throw new Error("Foglio Budget non trovato!");

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "========== INIZIO RIGENERAZIONE BUDGET ==========");

    // ── 1. Offerte abilitate ─────────────────────────────────────────────────
    var offerte = getConfigurazioneOfferte();
    var offerteAbilitate = [];
    for (var i = 0; i < offerte.length; i++) {
      if (offerte[i].abilitata) offerteAbilitate.push(offerte[i].id);
    }
    if (offerteAbilitate.length === 0) throw new Error("Nessuna offerta abilitata!");

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Offerte abilitate: " + offerteAbilitate.join(", "));

    // Verifica esistenza fogli
    for (var i = 0; i < offerteAbilitate.length; i++) {
      if (!ss.getSheetByName(offerteAbilitate[i])) {
        throw new Error("Foglio mancante: " + offerteAbilitate[i]);
      }
    }

    // ── 2. Pre-carica dati offerte (colonne L…Q) ─────────────────────────────
    // L=12, M=13, N=14, O=15, P=16, Q=17
    var colL = letterToColumn("L");
    var colM = letterToColumn("M");
    var colN = letterToColumn("N");
    var colO = letterToColumn("O");
    var colP = letterToColumn("P");
    var colQ = letterToColumn("Q");
    var numCols = colQ - colL + 1; // 6

    // Indici relativi all'array (0-based, partendo da colL)
    var iL = 0, iM = 1, iN = 2, iO = 3, iP = 4, iQ = 5;

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Pre-caricamento dati da " + offerteAbilitate.length + " offerte...");
    var offerteData = {};
    for (var i = 0; i < offerteAbilitate.length; i++) {
      var foglio = ss.getSheetByName(offerteAbilitate[i]);
      offerteData[offerteAbilitate[i]] = foglio
        .getRange(RANGE_INIZIO, colL, numRighe, numCols)
        .getValues();
    }

    // ── 3. Pre-carica colori colonna L nel Budget (celle blu) ────────────────
    var coloriL = budget.getRange(RANGE_INIZIO, colL, numRighe, 1).getBackgrounds();

    // ── 4. Ciclo su tutte le righe: calcola valori e individua errori ─────────
    var erroriBloccanti = [];
    var aggiornL = [];          // {riga, valore}
    var aggiornM = [];          // {riga, valore}
    var aggiornN = [];          // {riga, valore | formula}
    var formulaO = [];          // array di array per setFormulas
    var formulaP = [];
    var formulaQ_arr = [];

    for (var r = 0; r < numRighe; r++) {
      var riga = RANGE_INIZIO + r;

      // ── O, P, Q: formula SUM ──────────────────────────────────────────────
      var rifO = [], rifP = [], rifQ = [];
      for (var i = 0; i < offerteAbilitate.length; i++) {
        var id = offerteAbilitate[i];
        rifO.push(id + "!O" + riga);
        rifP.push(id + "!P" + riga);
        rifQ.push(id + "!Q" + riga);
      }
      formulaO.push(["=" + rifO.join("+")]);
      formulaP.push(["=" + rifP.join("+")]);
      formulaQ_arr.push(["=" + rifQ.join("+")]);

      // ── Identifica offerte con Q ≠ 0 per questa riga ──────────────────────
      var indiciConQNonZero = [];
      for (var i = 0; i < offerteAbilitate.length; i++) {
        var q = offerteData[offerteAbilitate[i]][r][iQ];
        if (q !== null && q !== "" && q !== 0) {
          indiciConQNonZero.push(i);
        }
      }
      var rigaAttiva = indiciConQNonZero.length > 0;

      // ── L: concatenazione deduplicata (solo celle blu nel Budget) ──────────
      var coloreL_cella = coloriL[r][0].toLowerCase();
      if (isColoreBlu(coloreL_cella)) {
        var visti = {};
        var parti = [];
        for (var i = 0; i < offerteAbilitate.length; i++) {
          var val = String(offerteData[offerteAbilitate[i]][r][iL] || "").trim();
          if (val && !visti[val]) {
            visti[val] = true;
            parti.push(val);
          }
        }
        aggiornL.push({ riga: riga, valore: parti.join(", ") });
      }

      if (!rigaAttiva) continue; // Nessuna offerta ha Q ≠ 0: salta M e N

      // ── M: controllo consistenza tra offerte con Q ≠ 0 ────────────────────
      var mValori = [];
      for (var i = 0; i < indiciConQNonZero.length; i++) {
        mValori.push(offerteData[offerteAbilitate[indiciConQNonZero[i]]][r][iM]);
      }
      var primoM = mValori[0];
      var mConsistente = mValori.every(function(v) { return v === primoM; });

      if (!mConsistente) {
        var dettagliM = indiciConQNonZero.map(function(i) {
          return offerteAbilitate[i] + ':"' + offerteData[offerteAbilitate[i]][r][iM] + '"';
        });
        erroriBloccanti.push("Riga " + riga + " col M — valori diversi: " + dettagliM.join(", "));
      } else {
        aggiornM.push({ riga: riga, valore: primoM });
      }

      // ── N: fino a riga 519 come M; da riga 521 somma ─────────────────────
      if (riga <= 519) {
        var nValori = [];
        for (var i = 0; i < indiciConQNonZero.length; i++) {
          nValori.push(offerteData[offerteAbilitate[indiciConQNonZero[i]]][r][iN]);
        }
        var primoN = nValori[0];
        var nConsistente = nValori.every(function(v) { return v === primoN; });

        if (!nConsistente) {
          var dettagliN = indiciConQNonZero.map(function(i) {
            return offerteAbilitate[i] + ':"' + offerteData[offerteAbilitate[i]][r][iN] + '"';
          });
          erroriBloccanti.push("Riga " + riga + " col N — valori diversi: " + dettagliN.join(", "));
        } else {
          aggiornN.push({ riga: riga, valore: primoN });
        }

      } else { // riga >= 521 (row 520 gestita come ≤519)
        var rifN = offerteAbilitate.map(function(id) { return id + "!N" + riga; });
        aggiornN.push({ riga: riga, formula: "=" + rifN.join("+") });
      }
    }

    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Analisi completata — errori bloccanti: " + erroriBloccanti.length);

    // ── 5. Errori bloccanti → interrompi ─────────────────────────────────────
    if (erroriBloccanti.length > 0) {
      var msg = "La rigenerazione è stata interrotta: " + erroriBloccanti.length + " errore/i bloccante/i.\n\n" +
                erroriBloccanti.join("\n");
      CONFIG.LOG.warn("rigeneraBudgetDaOfferte", "INTERRUZIONE per errori bloccanti:\n" + erroriBloccanti.join("\n"));
      SpreadsheetApp.getUi().alert("Errori Bloccanti — Rigenerazione Interrotta", msg, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // ── 6. Scrivi O, P, Q in batch ───────────────────────────────────────────
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Scrittura formule O, P, Q...");
    budget.getRange(RANGE_INIZIO, colO, numRighe, 1).setFormulas(formulaO);
    budget.getRange(RANGE_INIZIO, colP, numRighe, 1).setFormulas(formulaP);
    budget.getRange(RANGE_INIZIO, colQ, numRighe, 1).setFormulas(formulaQ_arr);

    // ── 7. Scrivi L ──────────────────────────────────────────────────────────
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Scrittura colonna L (" + aggiornL.length + " celle)...");
    for (var i = 0; i < aggiornL.length; i++) {
      budget.getRange(aggiornL[i].riga, colL).setValue(aggiornL[i].valore);
    }

    // ── 8. Scrivi M ──────────────────────────────────────────────────────────
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Scrittura colonna M (" + aggiornM.length + " celle)...");
    for (var i = 0; i < aggiornM.length; i++) {
      budget.getRange(aggiornM[i].riga, colM).setValue(aggiornM[i].valore);
    }

    // ── 9. Scrivi N ──────────────────────────────────────────────────────────
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Scrittura colonna N (" + aggiornN.length + " celle)...");
    for (var i = 0; i < aggiornN.length; i++) {
      var upd = aggiornN[i];
      var cell = budget.getRange(upd.riga, colN);
      if (upd.formula) {
        cell.setFormula(upd.formula);
      } else {
        cell.setValue(upd.valore);
      }
    }

    // ── 10. Allineamento BOM (controlla commessa) ─────────────────────────────
    try {
      controlla(ss, false);
      CONFIG.LOG.info("rigeneraBudgetDaOfferte", "Allineamento BOM completato");
    } catch (e) {
      CONFIG.LOG.warn("rigeneraBudgetDaOfferte", "Allineamento BOM fallito (non bloccante): " + e.message);
    }

    // ── 11. Cancella indicatore "rigenerazione necessaria" ────────────────────
    impostaIndicatoreRigenerazione(false);

    // ── 12. Torna al foglio Budget ────────────────────────────────────────────
    ss.setActiveSheet(budget);

    var tempoTotale = ((new Date().getTime() - tempoInizio) / 1000).toFixed(2);
    CONFIG.LOG.info("rigeneraBudgetDaOfferte", "========== RIGENERAZIONE COMPLETATA IN " + tempoTotale + "s ==========");

  } catch (error) {
    CONFIG.LOG.error("rigeneraBudgetDaOfferte", "ERRORE CRITICO", error);
    throw error;
  }
}
