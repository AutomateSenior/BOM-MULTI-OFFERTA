/**
 * BOMCore.js - Funzioni Core
 * Versione locale (non più libreria)
 */

// CONFIG già definito in Config.js

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

    // Mappa soglie per celle gialle (U430 a U461)
    var thresholds = {
      "U430": 0.10, "U431": 0.10, "U432": 0.10, "U433": 0.10, "U434": 0.10,
      "U435": 0.10, "U436": 0.10, "U437": 0.10, "U438": 0.10, "U439": 0.10,
      "U440": 0.10, "U441": 0.10, "U442": 0.10, "U443": 0.10, "U444": 0.10,
      "U445": 0.10, "U446": 0.10, "U447": 0.10, "U448": 0.10, "U449": 0.10,
      "U450": 0.10, "U451": 0.10, "U452": 0.10, "U453": 0.10, "U454": 0.10,
      "U455": 0.10, "U456": 0.10, "U457": 0.10, "U458": 0.10, "U459": 0.10,
      "U460": 0.10, "U461": 0.10
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

/**
 * Scrive il nome del file nella cella T56 del foglio Budget
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - Foglio di calcolo opzionale
 */
function scriviNomeFile(spreadsheet) {
  try {
    spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();

    var nomeFile = spreadsheet.getName();
    var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);

    if (!foglioBudget) {
      throw new Error(CONFIG.ERRORS.format("FOGLIO_NON_TROVATO", {foglio: CONFIG.SHEETS.BUDGET}));
    }

    // Scrive il nome del file nella cella configurata
    foglioBudget.getRange(CONFIG.CELLS.NOME_FILE).setValue(nomeFile);
    CONFIG.LOG.info("scriviNomeFile", "Nome file scritto in " + CONFIG.CELLS.NOME_FILE + ": " + nomeFile);

  } catch (error) {
    CONFIG.LOG.error("scriviNomeFile", "Errore nella scrittura del nome file", error);
    throw error;
  }
}

/**
 * Analizza il nome del file ed estrae il codice commessa, scrivendolo in S56
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - Foglio di calcolo
 */
function analizzaEScriviCommessa(spreadsheet) {
  try {
    spreadsheet = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();

    var nomeFile = spreadsheet.getName();
    CONFIG.LOG.info("analizzaEScriviCommessa", "Analisi del nome file: " + nomeFile);

    var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (!foglioBudget) {
      throw new Error(CONFIG.ERRORS.format("FOGLIO_NON_TROVATO", {foglio: CONFIG.SHEETS.BUDGET}));
    }

    // Esegue l'analisi del nome file
    var risultatoAnalisi = analizzaNomeFile(nomeFile);
    CONFIG.LOG.info("analizzaEScriviCommessa", "Risultato analisi: " + risultatoAnalisi);

    // Scrive il risultato nella cella configurata
    foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).setValue(risultatoAnalisi);

    // Verifica che il valore sia stato scritto correttamente
    var valoreVerifica = foglioBudget.getRange(CONFIG.CELLS.CODICE_COMMESSA).getValue();
    if (valoreVerifica != risultatoAnalisi) {
      throw new Error("Errore nella scrittura: atteso '" + risultatoAnalisi + "', trovato '" + valoreVerifica + "'");
    }

    CONFIG.LOG.info("analizzaEScriviCommessa", "Valore scritto in " + CONFIG.CELLS.CODICE_COMMESSA + " con successo");

  } catch (error) {
    CONFIG.LOG.error("analizzaEScriviCommessa", "Errore nell'analisi commessa", error);
    throw error;
  }
}

/**
 * Analizza il nome del file per estrarre il codice commessa
 * Funzione consolidata da Utility.js e Esecuzione.js
 * @param {string} nomeFile - Nome del file da analizzare
 * @returns {string} Codice commessa o messaggio di errore
 */
function analizzaNomeFile(nomeFile) {
  try {
    CONFIG.LOG.info("analizzaNomeFile", "Analisi nome file: " + nomeFile);

    // Controllo se il nome file è vuoto o non definito
    if (!nomeFile || nomeFile.trim() === "") {
      CONFIG.LOG.warn("analizzaNomeFile", CONFIG.ERRORS.FILE_VUOTO);
      return CONFIG.ERRORS.FILE_VUOTO;
    }

    // Verifica della presenza e del formato della revisione
    const revisioneMatch = nomeFile.match(CONFIG.NAMING.REVISIONE_PATTERN);
    if (!revisioneMatch) {
      CONFIG.LOG.warn("analizzaNomeFile", CONFIG.ERRORS.REVISIONE_NON_VALIDA);
      return CONFIG.ERRORS.REVISIONE_NON_VALIDA;
    }
    CONFIG.LOG.info("analizzaNomeFile", "Revisione trovata: #Rev " + revisioneMatch[1] + "." + revisioneMatch[2]);

    // Pulizia del nome file
    const nomeFilePulito = nomeFile
      .trim()
      .replace(/\s+/g, '_')
      .replace(/[^\w\d\._\-#]/g, '');

    CONFIG.LOG.info("analizzaNomeFile", "Nome file pulito: " + nomeFilePulito);

    // Verifica la presenza del pattern *_BOM_
    if (!nomeFilePulito.match(/\w+_BOM_/)) {
      CONFIG.LOG.warn("analizzaNomeFile", CONFIG.ERRORS.FORMATO_NON_VALIDO + " - Pattern _BOM_ non trovato");
      return CONFIG.ERRORS.FORMATO_NON_VALIDO;
    }
    CONFIG.LOG.info("analizzaNomeFile", "Pattern _BOM_ trovato");

    // Cerca il codice commessa usando il pattern da CONFIG
    const commessaMatch = nomeFilePulito.match(CONFIG.NAMING.COMMESSA_PATTERN);

    if (!commessaMatch) {
      CONFIG.LOG.warn("analizzaNomeFile", CONFIG.ERRORS.COMMESSA_NON_TROVATA);
      return CONFIG.ERRORS.COMMESSA_NON_TROVATA;
    }

    // Controlla se il codice commessa inizia con 2
    const codiceCommessaCompleto = commessaMatch[0];
    if (!codiceCommessaCompleto.startsWith('2')) {
      CONFIG.LOG.warn("analizzaNomeFile", CONFIG.ERRORS.COMMESSA_NON_TROVATA + " - Non inizia con 2");
      return CONFIG.ERRORS.COMMESSA_NON_TROVATA;
    }

    CONFIG.LOG.info("analizzaNomeFile", "Codice commessa trovato: " + codiceCommessaCompleto);
    return codiceCommessaCompleto;

  } catch (error) {
    var messaggioErrore = "Errore durante l'analisi del nome file: " + error.message;
    CONFIG.LOG.error("analizzaNomeFile", messaggioErrore, error);
    return messaggioErrore;
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
          linea: colonnaAH[i][0] || ""
        };
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
        vend3: params.vend[2]
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
        vend3: params.vend[2]
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
      {range: CONFIG.CELLS.COST_3, value: valori.cost3}
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

