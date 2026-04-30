/**
 * BOMCore.js - Funzioni Core
 * Versione locale (non più libreria)
 */

// CONFIG già definito in Config.js

// Cache dati ultima commessa caricata (valida nell'esecuzione corrente)
var _cacheCommessa = null;

/**
 * InserimentoBOM - Controlla inserimenti celle colorate
 * Chiamata dal trigger OnEdit
 */
function InserimentoBOM(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var cellColor = range.getBackground();
    var value = e.value;

    // Mappa dei colori
    var yellowColor = "#ffff00"; // Giallo
    var greenColor = "#00ff00";  // Verde

    // Mappa soglie per celle gialle (U432 a U463, dopo inserimento riga 77)
    var thresholds = {
      "U432": 0.10, "U433": 0.10, "U434": 0.10, "U435": 0.10, "U436": 0.10,
      "U437": 0.10, "U438": 0.10, "U439": 0.10, "U440": 0.10, "U441": 0.10,
      "U442": 0.10, "U443": 0.10, "U444": 0.10, "U445": 0.10, "U446": 0.10,
      "U447": 0.10, "U448": 0.10, "U449": 0.10, "U450": 0.10, "U451": 0.10,
      "U452": 0.10, "U453": 0.10, "U454": 0.10, "U455": 0.10, "U456": 0.10,
      "U457": 0.10, "U458": 0.10, "U459": 0.10, "U460": 0.10, "U461": 0.10,
      "U462": 0.10, "U463": 0.10
    };

    var cellAddress = range.getA1Notation();

    // Blocca modifica se non è cella gialla
    if (cellColor != yellowColor) {
      return; // Permetti solo celle gialle
    }

    // Controlla soglia per celle gialle
    if (cellColor === yellowColor) {
      var threshold = thresholds[cellAddress];

      if (threshold !== undefined) {
        var numericValue = parseFloat(String(value).replace(",", "."));

        if (isNaN(numericValue) || numericValue < threshold) {
          sheet.getRange(cellAddress).setValue(threshold);
        }
      }
      return;
    }

  } catch (error) {
    CONFIG.LOG.error("InserimentoBOM", "Errore", error);
  }
}

/**
 * Utility: Converti numero colonna in lettera
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
 * Utility: Converti lettera colonna in numero (A=1, Z=26, AA=27, etc)
 */
function letterToColumn(letter) {
  var column = 0;
  for (var i = 0; i < letter.length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, letter.length - i - 1);
  }
  return column;
}

/**
 * Utility: Converti notazione A1 in coordinate
 */
function convertiA1InCoordinate(riferimentoA1) {
  var match = riferimentoA1.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;

  var colLetters = match[1];
  var riga = parseInt(match[2], 10);
  var colonna = 0;

  for (var i = 0; i < colLetters.length; i++) {
    colonna = colonna * 26 + (colLetters.charCodeAt(i) - 64);
  }

  return { riga: riga, colonna: colonna };
}

/**
 * Utility: Converti coordinate in notazione A1
 */
function convertiCoordinateInA1(colonna, riga) {
  return columnToLetter(colonna) + riga;
}

function controlla(spreadsheet, mostraAlert) {
  mostraAlert = (mostraAlert !== false); // Default true
  try {
    // Se non viene fornito lo spreadsheet, usa quello attivo
    spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();

    CONFIG.LOG.info("controlla", "Inizio allineamento BOM alla commessa");

    // Verifica che il foglio Budget esista
    var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!foglioBudget) {
      throw new Error(CONFIG.ERRORS.format("FOGLIO_NON_TROVATO", {foglio: CONFIG.SHEETS.BUDGET}));
    }

    // Legge codice commessa da L56 e cerca nel file commesse
    cercaNelFileCommessa(spreadsheet);

    CONFIG.LOG.info("controlla", "Allineamento BOM completato con successo");

    // Mostra conferma all'utente solo se richiesto (non per trigger automatico)
    if (mostraAlert) {
      var codiceLetto = foglioBudget.getRange("L56").getValue();
      var risultato = foglioBudget.getRange("S56").getValue();

      SpreadsheetApp.getUi().alert(
        "Allineamento BOM Completato",
        "Codice commessa: " + codiceLetto + "\n" +
        "Risultato: " + risultato,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

  } catch (error) {
    CONFIG.LOG.error("controlla", "Errore durante allineamento BOM", error);
    if (mostraAlert) {
      SpreadsheetApp.getUi().alert("Errore: " + error.toString());
    }
    throw error;
  }
}

/*
function scriviNomeFile(spreadsheet) {
  try {
    spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var nomeFile = spreadsheet.getName();
    var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!foglioBudget) {
      throw new Error(CONFIG.ERRORS.format("FOGLIO_NON_TROVATO", {foglio: CONFIG.SHEETS.BUDGET}));
    }
    foglioBudget.getRange(CONFIG.CELLS.NOME_FILE).setValue(nomeFile);
    CONFIG.LOG.info("scriviNomeFile", "Nome file scritto in " + CONFIG.CELLS.NOME_FILE + ": " + nomeFile);
  } catch (error) {
    CONFIG.LOG.error("scriviNomeFile", "Errore nella scrittura del nome file", error);
    throw error;
  }
}

function analizzaEScriviCommessa(spreadsheet) {
  try {
    spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var nomeFile = spreadsheet.getName();
    var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!foglioBudget) {
      throw new Error(CONFIG.ERRORS.format("FOGLIO_NON_TROVATO", {foglio: CONFIG.SHEETS.BUDGET}));
    }
    var risultatoAnalisi = analizzaNomeFile(nomeFile);
    foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).setValue(risultatoAnalisi);
    var valoreVerifica = foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).getValue();
    if (valoreVerifica != risultatoAnalisi) {
      throw new Error("Errore nella scrittura: atteso '" + risultatoAnalisi + "', trovato '" + valoreVerifica + "'");
    }
  } catch (error) {
    CONFIG.LOG.error("analizzaEScriviCommessa", "Errore nell'analisi commessa", error);
    throw error;
  }
}

function analizzaNomeFile(nomeFile) {
  try {
    if (!nomeFile || nomeFile.trim() === "") return CONFIG.ERRORS.FILE_VUOTO;
    const revisioneMatch = nomeFile.match(CONFIG.NAMING.REVISIONE_PATTERN);
    if (!revisioneMatch) return CONFIG.ERRORS.REVISIONE_NON_VALIDA;
    const nomeFilePulito = nomeFile.trim().replace(/\s+/g, '_').replace(/[^\w\d\._\-#]/g, '');
    if (!nomeFilePulito.match(/\w+_BOM_/)) return CONFIG.ERRORS.FORMATO_NON_VALIDO;
    const commessaMatch = nomeFilePulito.match(CONFIG.NAMING.COMMESSA_PATTERN);
    if (!commessaMatch) return CONFIG.ERRORS.COMMESSA_NON_TROVATA;
    const codiceCommessaCompleto = commessaMatch[0];
    if (!codiceCommessaCompleto.startsWith('2')) return CONFIG.ERRORS.COMMESSA_NON_TROVATA;
    return codiceCommessaCompleto;
  } catch (error) {
    return "Errore durante l'analisi del nome file: " + error.message;
  }
}
*/

/**
 * Controlla all'apertura se il nome del file contiene il codice commessa di L56.
 * Non esegue nulla se il nome file contiene "MASTER".
 * In caso di disallineamento mostra solo un avviso, non blocca.
 */
function controllaNomeFileVsCommessa(spreadsheet) {
  try {
    spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    var nomeFile = spreadsheet.getName();

    if (nomeFile.toUpperCase().indexOf("MASTER") > -1) return;

    var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!foglioBudget) return;

    var codiceL56 = String(foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).getValue()).trim();
    if (!codiceL56) {
      SpreadsheetApp.getUi().alert(
        "Codice commessa mancante",
        "La cella L56 è vuota: inserire il codice commessa.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    if (nomeFile.indexOf(codiceL56) === -1) {
      SpreadsheetApp.getUi().alert(
        "Nome file non allineato",
        "Il nome del file non corrisponde al codice commessa inserito in L56.\n\n" +
        "Nome file:  " + nomeFile + "\n" +
        "Codice L56: " + codiceL56,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (e) {
    CONFIG.LOG.warn("controllaNomeFileVsCommessa", "Errore (non bloccante): " + e.toString());
  }
}

/**
 * Cerca il codice commessa nel file commesse e aggiorna i valori correlati
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - Foglio di calcolo opzionale
 */
function cercaNelFileCommessa(spreadsheet) {
  var valoreDaCercare = null;

  try {
    spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();

    var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!foglioBudget) {
      throw new Error(CONFIG.ERRORS.format("FOGLIO_NON_TROVATO", {foglio: CONFIG.SHEETS.BUDGET}));
    }

    // Ottiene il valore da cercare dalla cella M56
    valoreDaCercare = foglioBudget.getRange("L56").getValue();
    CONFIG.LOG.info("cercaNelFileCommessa", "Valore da cercare in " + CONFIG.CELLS.CODICE_COMMESSA + ": " + valoreDaCercare);

    // Verifica che il valore non sia vuoto
    if (!valoreDaCercare || valoreDaCercare.toString().trim() === "") {
      throw new Error(CONFIG.ERRORS.format("CELLA_VUOTA", {cella: CONFIG.CELLS.CODICE_COMMESSA}));
    }

    // Converte a stringa e rimuove spazi
    valoreDaCercare = valoreDaCercare.toString().trim();

    // Apre il file delle commesse usando l'ID da CONFIG
    CONFIG.LOG.info("cercaNelFileCommessa", "Apertura file commesse...");
    var fileCommesse;
    try {
      fileCommesse = SpreadsheetApp.openById(CONFIG.FILES.COMMESSE_ID);
    } catch (error) {
      throw new Error(CONFIG.ERRORS.FILE_COMMESSE_NON_ACCESSIBILE + ": " + error.toString());
    }

    var foglioCommesse = fileCommesse.getSheetByName(CONFIG.FILES.COMMESSE_SHEET_NAME);
    if (!foglioCommesse) {
      throw new Error(CONFIG.ERRORS.format("FOGLIO_NON_TROVATO", {foglio: CONFIG.FILES.COMMESSE_SHEET_NAME}));
    }

    CONFIG.LOG.info("cercaNelFileCommessa", "File commesse aperto, lettura dati...");

    // Leggi tutte le colonne necessarie
    var ultimaRiga = foglioCommesse.getLastRow();
    var colonnaA = foglioCommesse.getRange("A1:A" + ultimaRiga).getValues();
    var colonnaB = foglioCommesse.getRange("B1:B" + ultimaRiga).getValues();  // PM
    var colonnaC = foglioCommesse.getRange("C1:C" + ultimaRiga).getValues();  // Descrizione
    var colonnaD = foglioCommesse.getRange("D1:D" + ultimaRiga).getValues();  // Cliente
    var colonnaH = foglioCommesse.getRange("H1:H" + ultimaRiga).getValues();  // Soluzione
    var colonnaAF = foglioCommesse.getRange("AF1:AF" + ultimaRiga).getValues(); // Società
    var colonnaAH = foglioCommesse.getRange("AH1:AH" + ultimaRiga).getValues(); // Linea

    CONFIG.LOG.info("cercaNelFileCommessa", "Lette " + ultimaRiga + " righe, inizio ricerca...");

    // Cerca il valore
    var rigaTrovata = -1;
    var datiCommessa = null;

    for (var i = 0; i < colonnaA.length; i++) {
      var valoreColonnaA = colonnaA[i][0];

      if (valoreColonnaA && valoreColonnaA.toString().trim() === valoreDaCercare) {
        rigaTrovata = i + 1;
        datiCommessa = {
          codice: valoreDaCercare,
          pm: colonnaB[i][0] || "",
          descrizione: colonnaC[i][0] || "",
          cliente: colonnaD[i][0] || "",
          soluzione: colonnaH[i][0] || "",
          societa: colonnaAF[i][0] || "",
          linea: colonnaAH[i][0] || "",
          riga: rigaTrovata
        };
        _cacheCommessa = datiCommessa;
        CONFIG.LOG.info("cercaNelFileCommessa", "Commessa trovata alla riga " + rigaTrovata);
        break;
      }
    }

    // Scrivi risultato ricerca in S56
    if (datiCommessa) {
      foglioBudget.getRange("S56").setValue("Commessa OK");
      CONFIG.LOG.info("cercaNelFileCommessa", "Commessa OK: " + JSON.stringify(datiCommessa));

      // Calcola valori celle in base alla linea
      var params = CONFIG.COMMESSA_TYPES.getParams(datiCommessa.linea);
      var valori = {
        perc: params.perc,
        cost1: params.cost[0],
        cost2: params.cost[1],
        cost3: params.cost[2],
        vend1: params.vend[0],
        vend2: params.vend[1],
        vend3: params.vend[2],
        // Costi di struttura Body Rental
        costBR1: params.costBodyRental[0],
        costBR2: params.costBodyRental[1],
        costBR3: params.costBodyRental[2]
      };

      // Aggiorna foglio Budget con i dati della commessa
      aggiornaFoglioBudget(foglioBudget, datiCommessa.linea, valori);
    } else {
      foglioBudget.getRange("S56").setValue("Commessa inesistente");
      CONFIG.LOG.warn("cercaNelFileCommessa", "Commessa '" + valoreDaCercare + "' non trovata");

      // Usa valori di default
      var params = CONFIG.COMMESSA_TYPES.DEFAULT;
      var valori = {
        perc: params.perc,
        cost1: params.cost[0],
        cost2: params.cost[1],
        cost3: params.cost[2],
        vend1: params.vend[0],
        vend2: params.vend[1],
        vend3: params.vend[2],
        // Costi di struttura Body Rental
        costBR1: params.costBodyRental[0],
        costBR2: params.costBodyRental[1],
        costBR3: params.costBodyRental[2]
      };

      aggiornaFoglioBudget(foglioBudget, params.nome, valori);
    }

    CONFIG.LOG.info("cercaNelFileCommessa", "Aggiornamento foglio Budget completato");

  } catch (error) {
    CONFIG.LOG.error("cercaNelFileCommessa", "Errore (valore cercato: '" + valoreDaCercare + "')", error);
    throw error;
  }
}

/**
 * Aggiorna il foglio Budget con i valori calcolati
 * @param {SpreadsheetApp.Sheet} foglioBudget - Foglio Budget
 * @param {string} tipo - Tipo di commessa
 * @param {Object} valori - Valori da scrivere
 */
function aggiornaFoglioBudget(foglioBudget, tipo, valori) {
  try {
    CONFIG.LOG.info("aggiornaFoglioBudget", "Inizio aggiornamento Budget e Offerte con tipo: " + tipo);

    // Leggi il codice commessa da Budget L56
    var codiceCommessa = foglioBudget.getRange("L56").getValue();

    // Lista celle da aggiornare
    var rangeList = [
      {range: CONFIG.CELLS.TIPO_COMMESSA, value: tipo},
      {range: CONFIG.CELLS.PERC_GENERALE, value: valori.perc},
      {range: CONFIG.CELLS.VEND_1, value: valori.vend1},
      {range: CONFIG.CELLS.VEND_2, value: valori.vend2},
      {range: CONFIG.CELLS.VEND_3, value: valori.vend3},
      {range: CONFIG.CELLS.COST_1, value: valori.cost1},
      {range: CONFIG.CELLS.COST_2, value: valori.cost2},
      {range: CONFIG.CELLS.COST_3, value: valori.cost3},
      // Costi di struttura Body Rental
      {range: CONFIG.CELLS.COST_BR_1, value: valori.costBR1},
      {range: CONFIG.CELLS.COST_BR_2, value: valori.costBR2},
      {range: CONFIG.CELLS.COST_BR_3, value: valori.costBR3},
      // Percentuali per durata assistenza (AB3:AB6)
      {range: CONFIG.CELLS.PERC_1MESE,  value: CONFIG.MESI_PERC[1]},
      {range: CONFIG.CELLS.PERC_3MESI,  value: CONFIG.MESI_PERC[3]},
      {range: CONFIG.CELLS.PERC_6MESI,  value: CONFIG.MESI_PERC[6]},
      {range: CONFIG.CELLS.PERC_12MESI, value: CONFIG.MESI_PERC[12]}
    ];

    // Aggiorna Budget
    for (var i = 0; i < rangeList.length; i++) {
      var item = rangeList[i];
      foglioBudget.getRange(item.range).setValue(item.value);
      CONFIG.LOG.info("aggiornaFoglioBudget", "Budget - Cella " + item.range + " = " + item.value);
    }

    // Aggiorna TUTTE le offerte esistenti
    var ss = foglioBudget.getParent();
    var configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE);

    if (configSheet) {
      var lastRow = configSheet.getLastRow();
      if (lastRow >= 2) {
        var offerte = configSheet.getRange(2, 1, lastRow - 1, 1).getValues();

        for (var j = 0; j < offerte.length; j++) {
          var offertaId = offerte[j][0];
          if (offertaId) {
            var foglioOfferta = ss.getSheetByName(offertaId);
            if (foglioOfferta) {
              // Aggiorna L56 con il codice commessa
              foglioOfferta.getRange("L56").setValue(codiceCommessa);

              // Aggiorna le altre celle nell'offerta
              for (var k = 0; k < rangeList.length; k++) {
                var item = rangeList[k];
                foglioOfferta.getRange(item.range).setValue(item.value);
              }
              CONFIG.LOG.info("aggiornaFoglioBudget", offertaId + " - L56 e celle parametri aggiornate");
            }
          }
        }
      }
    }
    CONFIG.LOG.info("aggiornaFoglioBudget", "Aggiornamento completato per Budget e tutte le offerte");
  } catch (error) {
    CONFIG.LOG.error("aggiornaFoglioBudget", "Errore nell'aggiornamento", error);
    throw error;
  }
}

/**
 * Gestisce il disallineamento tra nome file e codice commessa in L56.
 * Chiamata dopo controlla() quando S56 = "Commessa OK".
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} oldValue - Vecchio valore di L56 (da e.oldValue)
 */
function gestisciDisallineamentoNomeFile(spreadsheet, oldValue) {
  try {
    var ui = SpreadsheetApp.getUi();
    var nomeFile = spreadsheet.getName();

    // File Master: crea nuovo file BOM e aprilo
    if (nomeFile.toUpperCase().indexOf("MASTER") > -1) {
      var foglioBudgetM = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
      if (!foglioBudgetM) return;
      var nuovaCommessaM = String(foglioBudgetM.getRange(CONFIG.CELLS.CODICE_COMMESSA).getValue()).trim();
      if (!nuovaCommessaM) return;

      var datiM = _cacheCommessa;
      if (!datiM || datiM.codice !== nuovaCommessaM) {
        ui.alert("Errore", "Dati commessa non disponibili. Riprovare.", ui.ButtonSet.OK);
        return;
      }

      var rispostaMaster = ui.alert(
        "Crea nuovo file BOM",
        "Vuoi creare un nuovo file BOM per la commessa " + nuovaCommessaM + "?",
        ui.ButtonSet.YES_NO
      );

      if (rispostaMaster !== ui.Button.YES) {
        foglioBudgetM.getRange(CONFIG.CELLS.CODICE_COMMESSA).setValue("");
        return;
      }

      var societaCodiceM = getSocietaCodice(datiM.societa);
      if (!societaCodiceM) {
        ui.alert("Errore", "Società non riconosciuta: \"" + datiM.societa + '"', ui.ButtonSet.OK);
        return;
      }

      var maxRevM = cercaMassimaRevisioneBOM(nuovaCommessaM);
      var nuovaRevM = ('00' + (maxRevM + 1)).slice(-2) + '.08';
      var nuovoNomeM = CONFIG.NAMING.generaNomeFile(
        societaCodiceM, nuovaCommessaM, datiM.pm, datiM.cliente, datiM.descrizione, nuovaRevM
      );

      var nuovoFileM = creaEConfiguraNuovoFileBOM(spreadsheet, nuovoNomeM, nuovaCommessaM, datiM.riga, "");

      try {
        var nuovoSsM = SpreadsheetApp.openById(nuovoFileM.getId());
        var nuovoBudgetM = nuovoSsM.getSheetByName(CONFIG.SHEETS.BUDGET);
        if (nuovoBudgetM) {
          nuovoBudgetM.getRange(CONFIG.CELLS.CODICE_COMMESSA).setValue(nuovaCommessaM);
          _cacheCommessa = null;
          controlla(nuovoSsM, false);
          propagaL56(nuovoSsM, nuovaCommessaM);
        }
      } catch (eCfg) {
        CONFIG.LOG.warn("gestisciDisallineamentoNomeFile", "Configurazione nuovo file fallita: " + eCfg.toString());
      }

      var urlM = "https://docs.google.com/spreadsheets/d/" + nuovoFileM.getId();
      apriUrlInNuovaScheda(urlM, ui);
      ui.alert("Completato", "Nuovo file creato e aperto:\n" + nuovoNomeM, ui.ButtonSet.OK);
      return;
    }

    var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!foglioBudget) return;

    var nuovaCommessa = String(foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).getValue()).trim();
    if (!nuovaCommessa) return;

    // Nome file già allineato: niente da fare
    if (nomeFile.indexOf(nuovaCommessa) > -1) return;

    // Mostra disallineamento e chiedi azione
    var risposta = ui.alert(
      "Nome file non allineato",
      "Il nome del file non corrisponde al codice commessa.\n\n" +
      "Nome file:       " + nomeFile + "\n" +
      "Codice commessa: " + nuovaCommessa + "\n\n" +
      "Sì      → Crea nuovo file\n" +
      "No      → Rinomina il file attuale\n" +
      "Annulla → Non fare nulla",
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (risposta === ui.Button.CANCEL || risposta === ui.Button.CLOSE) return;

    // Usa dati dalla cache (caricata da cercaNelFileCommessa nella stessa esecuzione)
    var datiCommessa = _cacheCommessa;
    if (!datiCommessa || datiCommessa.codice !== nuovaCommessa) {
      ui.alert("Errore", "Dati commessa non disponibili. Riprovare.", ui.ButtonSet.OK);
      return;
    }

    // Reverse-map società: nome esteso → codice A-E
    var societaCodice = getSocietaCodice(datiCommessa.societa);
    if (!societaCodice) {
      ui.alert("Errore", "Società non riconosciuta: \"" + datiCommessa.societa + '"', ui.ButtonSet.OK);
      return;
    }

    // Cerca revisione massima su Drive e calcola la nuova
    var maxRev = cercaMassimaRevisioneBOM(nuovaCommessa);
    var nuovaRev = ('00' + (maxRev + 1)).slice(-2) + '.08';

    // Genera nuovo nome file con la funzione esistente
    var nuovoNome = CONFIG.NAMING.generaNomeFile(
      societaCodice,
      nuovaCommessa,
      datiCommessa.pm,
      datiCommessa.cliente,
      datiCommessa.descrizione,
      nuovaRev
    );

    if (risposta === ui.Button.NO) {
      // RINOMINA
      rinominaFileBOM(spreadsheet, nuovoNome);
      ui.alert("Completato", "File rinominato:\n" + nuovoNome, ui.ButtonSet.OK);

    } else if (risposta === ui.Button.YES) {
      // NUOVO FILE
      var nuovoFile = creaEConfiguraNuovoFileBOM(spreadsheet, nuovoNome, nuovaCommessa, datiCommessa.riga, oldValue);
      var urlNuovo = "https://docs.google.com/spreadsheets/d/" + nuovoFile.getId();
      apriUrlInNuovaScheda(urlNuovo, ui);
      ui.alert("Completato", "Nuovo file creato e aperto:\n" + nuovoNome, ui.ButtonSet.OK);
    }

  } catch (e) {
    CONFIG.LOG.error("gestisciDisallineamentoNomeFile", "Errore", e);
    SpreadsheetApp.getUi().alert("Errore: " + e.toString());
  }
}

/**
 * Cerca su Drive la revisione principale massima (#Rev XX.YY → XX) per una commessa BOM.
 * @param {string} commessa
 * @returns {number} Massimo trovato, 0 se nessun file esistente
 */
function cercaMassimaRevisioneBOM(commessa) {
  try {
    var queryParts = ["A_", "B_", "C_", "D_", "E_"].map(function(p) {
      return 'title contains "' + p + 'BOM_' + commessa + '"';
    });
    var query = "trashed = false and (" + queryParts.join(" or ") + ")";
    var maxRev = 0;
    var pageToken;

    do {
      var files = Drive.Files.list({
        q: query,
        maxResults: 100,
        corpora: 'allDrives',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        pageToken: pageToken
      });

      if (files.items && files.items.length > 0) {
        for (var i = 0; i < files.items.length; i++) {
          var title = files.items[i].title;
          if (title.indexOf(commessa) > -1) {
            var revMatch = title.match(/#Rev\s+(\d+)\.\d+/i);
            if (revMatch) {
              var rev = parseInt(revMatch[1], 10);
              if (rev > maxRev) maxRev = rev;
            }
          }
        }
      }
      pageToken = files.nextPageToken;
    } while (pageToken);

    CONFIG.LOG.info("cercaMassimaRevisioneBOM", commessa + " → maxRev=" + maxRev);
    return maxRev;

  } catch (e) {
    CONFIG.LOG.error("cercaMassimaRevisioneBOM", "Errore ricerca Drive", e);
    return 0;
  }
}

/**
 * Restituisce il codice società (A-E) dal nome esteso.
 * @param {string} nomeEsteso - Es. "Automate", "Bridger"
 * @returns {string|null}
 */
function getSocietaCodice(nomeEsteso) {
  var societa = CONFIG.NAMING.SOCIETA;
  for (var codice in societa) {
    if (societa[codice] === nomeEsteso) return codice;
  }
  return null;
}

/**
 * Rinomina il file corrente su Drive.
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} nuovoNome
 */
function rinominaFileBOM(spreadsheet, nuovoNome) {
  var file = DriveApp.getFileById(spreadsheet.getId());
  file.setName(nuovoNome);
  CONFIG.LOG.info("rinominaFileBOM", "Rinominato: " + nuovoNome);
}

/**
 * Crea un nuovo file BOM (copia del corrente), ripristina il vecchio codice
 * nel file originale e aggiorna l'hyperlink in col Q del file commesse.
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - File originale
 * @param {string} nuovoNome
 * @param {string} nuovaCommessa
 * @param {number} rigaCommessa - Riga della commessa nel foglio commesse
 * @param {string} oldValue - Vecchio codice da ripristinare in L56 (vuoto = cancella)
 */
function creaEConfiguraNuovoFileBOM(spreadsheet, nuovoNome, nuovaCommessa, rigaCommessa, oldValue) {
  // 1. Copia nella cartella destinazione
  var cartella = DriveApp.getFolderById(CONFIG.FOLDERS.NUOVI_BOM);
  var nuovoFile = DriveApp.getFileById(spreadsheet.getId()).makeCopy(nuovoNome, cartella);
  CONFIG.LOG.info("creaEConfiguraNuovoFileBOM", "Creato: " + nuovoNome + " ID=" + nuovoFile.getId());

  // 2. Ripristina/cancella L56 nel file originale
  var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
  foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).setValue(oldValue || "");
  if (oldValue) {
    try {
      controlla(spreadsheet, false);
    } catch (e) {
      CONFIG.LOG.warn("creaEConfiguraNuovoFileBOM", "Riallineamento fallito: " + e.toString());
    }
  }

  // 3. Aggiorna hyperlink col Q nel file commesse per la nuova commessa
  try {
    var url = "https://docs.google.com/spreadsheets/d/" + nuovoFile.getId();
    var foglioCommesse = SpreadsheetApp.openById(CONFIG.FILES.COMMESSE_ID)
                          .getSheetByName(CONFIG.FILES.COMMESSE_SHEET_NAME);
    if (foglioCommesse && rigaCommessa > 0) {
      foglioCommesse.getRange(rigaCommessa, 17).setFormula('=HYPERLINK("' + url + '","BOM")');
      CONFIG.LOG.info("creaEConfiguraNuovoFileBOM", "Hyperlink aggiornato riga " + rigaCommessa);
    }
  } catch (e) {
    CONFIG.LOG.warn("creaEConfiguraNuovoFileBOM", "Aggiornamento hyperlink fallito: " + e.toString());
  }

  return nuovoFile;
}

/**
 * Apre un URL in una nuova scheda del browser tramite un dialog HTML invisibile.
 * @param {string} url
 * @param {GoogleAppsScript.Base.Ui} ui
 */
function apriUrlInNuovaScheda(url, ui) {
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '", "_blank"); google.script.host.close();</script>'
  ).setWidth(1).setHeight(1);
  ui.showModelessDialog(html, 'Apertura...');
}

/**
 * Propaga il valore di L56 (codice commessa) a tutti i fogli Off_XX e al foglio Master.
 * Va chiamata dopo controlla() quando S56 = "Commessa OK".
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} valore - Codice commessa da propagare
 */
function propagaL56(spreadsheet, valore) {
  var sheets = spreadsheet.getSheets();
  var count = 0;
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (/^Off_/.test(name) || name === CONFIG.SHEETS.MASTER) {
      sheets[i].getRange("L56").setValue(valore);
      count++;
    }
  }
  CONFIG.LOG.info("propagaL56", "Propagato '" + valore + "' a " + count + " fogli");
}

/**
 * Punto di ingresso headless per script schedulati/esterni.
 * Apre una nuova commessa a partire dal Master BOM:
 *   1. Scrive codiceCommessa in L56 del Master
 *   2. Valida la commessa (controlla)
 *   3. Crea una copia del Master nella cartella NUOVI_BOM con nome corretto
 *   4. Configura L56 nel nuovo file e lo allinea
 *   5. Svuota L56 nel Master
 *   6. Aggiorna hyperlink col Q nel file commesse
 *
 * Nessuna chiamata UI — compatibile con trigger temporali e script esterni.
 *
 * @param {string} masterSpreadsheetId - ID del file Master BOM
 * @param {string} codiceCommessa - Codice commessa da aprire (es. "12345.1")
 * @returns {{ nome: string, id: string, url: string }} Dati del nuovo file creato
 * @throws {Error} Se la commessa non esiste o il file non è un Master
 */
function apriNuovaCommessa(masterSpreadsheetId, codiceCommessa) {
  CONFIG.LOG.info("apriNuovaCommessa", "Inizio: commessa=" + codiceCommessa + " master=" + masterSpreadsheetId);

  // 1. Apri il Master
  var master = SpreadsheetApp.openById(masterSpreadsheetId);
  var nomeFileMaster = master.getName();

  if (nomeFileMaster.toUpperCase().indexOf("MASTER") === -1) {
    throw new Error("Il file '" + nomeFileMaster + "' non è un Master BOM (manca 'MASTER' nel nome)");
  }

  // 2. Scrivi il codice commessa in L56 e valida
  var foglioBudget = master.getSheetByName(CONFIG.SHEETS.BUDGET);
  if (!foglioBudget) {
    throw new Error(CONFIG.ERRORS.format("FOGLIO_NON_TROVATO", { foglio: CONFIG.SHEETS.BUDGET }));
  }

  foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).setValue(codiceCommessa);
  _cacheCommessa = null; // Azzera cache prima di controlla
  controlla(master, false);

  // 3. Verifica che la commessa sia stata riconosciuta
  var s56 = foglioBudget.getRange("S56").getValue();
  if (s56 !== "Commessa OK") {
    foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).setValue("");
    throw new Error("Commessa non valida (S56='" + s56 + "'). L56 ripristinata vuota.");
  }

  // 4. Leggi dati dalla cache popolata da controlla()
  var datiCommessa = _cacheCommessa;
  if (!datiCommessa || datiCommessa.codice !== String(codiceCommessa).trim()) {
    foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).setValue("");
    throw new Error("Cache commessa non disponibile dopo controlla()");
  }

  // 5. Genera nome file
  var societaCodice = getSocietaCodice(datiCommessa.societa);
  if (!societaCodice) {
    throw new Error("Società non riconosciuta: \"" + datiCommessa.societa + '"');
  }

  var maxRev = cercaMassimaRevisioneBOM(codiceCommessa);
  var nuovaRev = ('00' + (maxRev + 1)).slice(-2) + '.08';

  var nuovoNome = CONFIG.NAMING.generaNomeFile(
    societaCodice,
    codiceCommessa,
    datiCommessa.pm,
    datiCommessa.cliente,
    datiCommessa.descrizione,
    nuovaRev
  );

  CONFIG.LOG.info("apriNuovaCommessa", "Nome generato: " + nuovoNome + " (rev " + nuovaRev + ")");

  // 6. Crea copia, svuota L56 nel Master e aggiorna hyperlink col Q nel file commesse
  var nuovoFile = creaEConfiguraNuovoFileBOM(master, nuovoNome, codiceCommessa, datiCommessa.riga, "");

  // 7. Configura il nuovo file: scrivi L56 e allinea
  try {
    var nuovoSs = SpreadsheetApp.openById(nuovoFile.getId());
    var nuovoBudget = nuovoSs.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (nuovoBudget) {
      nuovoBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).setValue(codiceCommessa);
      _cacheCommessa = null;
      controlla(nuovoSs, false);
      propagaL56(nuovoSs, codiceCommessa);
      CONFIG.LOG.info("apriNuovaCommessa", "Nuovo file allineato con commessa " + codiceCommessa);
    }
  } catch (e) {
    CONFIG.LOG.warn("apriNuovaCommessa", "Allineamento nuovo file fallito: " + e.toString());
  }

  var url = "https://docs.google.com/spreadsheets/d/" + nuovoFile.getId();
  CONFIG.LOG.info("apriNuovaCommessa", "Completato: " + nuovoNome);

  return { nome: nuovoNome, id: nuovoFile.getId(), url: url };
}

