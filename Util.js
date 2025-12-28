/*
function ripristinaFormuleInTuttiIFiles () {
  verificaFormuleInTuttiIFiles(true) 
}

function verificaFormuleInTuttiIFiles(ripristina) {
  salvaFormule()
  // Se il parametro non è specificato, imposta il default a false
  ripristina = (ripristina === true);
  
  // Ottieni le formule di riferimento dal foglio attivo (da cui si lancia la funzione)
  var formuleDiRiferimento = caricaFormuleDiRiferimento();
  
  if (Object.keys(formuleDiRiferimento.formuleMappa).length === 0) {
    Logger.log("Nessuna formula di riferimento trovata nel foglio attivo.");
    return "Nessuna formula di riferimento trovata.";
  }
  
  Logger.log("Caricate " + Object.keys(formuleDiRiferimento.formuleMappa).length + " formule di riferimento.");
  
  // Cerca i file su Drive
  var files = cercaFilesByPattern();
  
  if (files.length === 0) {
    Logger.log("Nessun file trovato che corrisponde al pattern richiesto.");
    return "Nessun file trovato.";
  }
  
  Logger.log("Trovati " + files.length + " files da verificare.");
  
  // Array per raccogliere i risultati
  var risultatiComplessivi = [];
  var contenutoFileDifferenze = "REPORT DIFFERENZE FORMULE\n";
  contenutoFileDifferenze += "Data: " + new Date().toLocaleString() + "\n";
  contenutoFileDifferenze += "Modalità: " + (ripristina ? "VERIFICA E RIPRISTINO" : "SOLO VERIFICA") + "\n";
  contenutoFileDifferenze += "==================================================\n\n";
  
  // Verifica ogni file trovato
  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    var nomeFile = file.getName();
    //Logger.log("Analizzo il file: " + nomeFile);
    contenutoFileDifferenze += "FILE: " + nomeFile + "\n";
    
    try {
      var spreadsheet = SpreadsheetApp.open(file);
      var foglioBudget = spreadsheet.getSheetByName("Budget");
      
      if (!foglioBudget) {
        var messaggio = "File '" + nomeFile + "': Foglio 'Budget' non trovato.";
        risultatiComplessivi.push(messaggio);
        contenutoFileDifferenze += messaggio + "\n\n";
        continue;
      }
      
      // Verifica ultra-veloce con le formule di riferimento già caricate
      var risultato = verificaFormuleUltraVeloce(spreadsheet, nomeFile, formuleDiRiferimento, ripristina);
      risultatiComplessivi.push(risultato.messaggio);
      contenutoFileDifferenze += risultato.reportFormattato + "\n\n";
    } catch (e) {
      var errore = "Errore nell'elaborazione del file '" + nomeFile + "': " + e.message;
      risultatiComplessivi.push(errore);
      contenutoFileDifferenze += errore + "\n\n";
    }
  }
  
  // Scrive i risultati nel log
  for (var j = 0; j < risultatiComplessivi.length; j++) {
    Logger.log(risultatiComplessivi[j]);
  }
  
  // Salva il file delle differenze
  salvaSuDrive("differenze.txt", contenutoFileDifferenze);
  
  var messaggioFinale = "Analisi completata. " + files.length + " file verificati. ";
  if (ripristina) {
    messaggioFinale += "Le formule errate sono state corrette. ";
  }
  messaggioFinale += "Report salvato come 'differenze.txt' nella cartella principale di Drive.";
  
  return messaggioFinale;
}

function caricaFormuleDiRiferimento() {
  var spreadsheetAttivo = SpreadsheetApp.getActiveSpreadsheet();
  var foglioFormule = spreadsheetAttivo.getSheetByName("Formule");
  
  if (!foglioFormule) {
    return { formuleMappa: {}, minRiga: 0, maxRiga: 0, minColonna: 0, maxColonna: 0 };
  }
  
  // Carica i dati dal foglio Formule
  var ultimaRiga = foglioFormule.getLastRow();
  if (ultimaRiga <= 1) {
    return { formuleMappa: {}, minRiga: 0, maxRiga: 0, minColonna: 0, maxColonna: 0 };
  }
  
  var datiFormule = foglioFormule.getRange(2, 1, ultimaRiga - 1, 2).getValues();
  var formuleDaVerificare = foglioFormule.getRange(2, 2, ultimaRiga - 1, 1).getFormulas();
  
  // Costruisci una mappa per le formule attese e trova i range estremi
  var formuleMappa = {};
  var minRiga = 9999, maxRiga = 0, minColonna = 9999, maxColonna = 0;
  
  for (var i = 0; i < datiFormule.length; i++) {
    if (!datiFormule[i][0]) continue;
    
    var riferimentoCella = datiFormule[i][0].toString().trim();
    var formulaAttesa = formuleDaVerificare[i][0].toString().trim();
    
    formuleMappa[riferimentoCella] = formulaAttesa;
    
    // Calcola le coordinate della cella per determinare il range complessivo
    var coordinate = convertiA1InCoordinate(riferimentoCella);
    if (coordinate) {
      minRiga = Math.min(minRiga, coordinate.riga);
      maxRiga = Math.max(maxRiga, coordinate.riga);
      minColonna = Math.min(minColonna, coordinate.colonna);
      maxColonna = Math.max(maxColonna, coordinate.colonna);
    }
  }
  
  return {
    formuleMappa: formuleMappa,
    minRiga: minRiga,
    maxRiga: maxRiga,
    minColonna: minColonna,
    maxColonna: maxColonna
  };
}

function cercaFilesByPattern() {
  var query = "title contains '_BOM_' and title contains '#Rev 0' and title contains '.07'";
  var files = DriveApp.searchFiles(query);
  
  var risultati = [];
  while (files.hasNext()) {
    var file = files.next();
    var nomeFile = file.getName();
    
    // Verifica pattern specifico
    if (nomeFile.indexOf("#Rev 0") > -1 && nomeFile.lastIndexOf(".07") > nomeFile.indexOf("#Rev 0") + 6) {
      risultati.push(file);
    }
  }
  
  return risultati;
}

function verificaFormuleUltraVeloce(spreadsheet, nomeFile, formuleDiRiferimento, ripristina) {
  var foglioBudget = spreadsheet.getSheetByName("Budget");
  var formuleMappa = formuleDiRiferimento.formuleMappa;
  
  if (Object.keys(formuleMappa).length === 0) {
    return {
      messaggio: "File '" + nomeFile + "': Nessuna formula da verificare.",
      reportFormattato: "Nessuna formula da verificare."
    };
  }
  
  // 1. Carica TUTTE le formule del range rilevante nel foglio Budget in una volta sola
  var rangeFormule = foglioBudget.getRange(
    formuleDiRiferimento.minRiga, 
    formuleDiRiferimento.minColonna, 
    formuleDiRiferimento.maxRiga - formuleDiRiferimento.minRiga + 1, 
    formuleDiRiferimento.maxColonna - formuleDiRiferimento.minColonna + 1
  ).getFormulas();
  
  // 2. Costruisci una mappa delle formule effettive nel foglio Budget
  var formuleEffettive = {};
  
  for (var i = 0; i < rangeFormule.length; i++) {
    for (var j = 0; j < rangeFormule[i].length; j++) {
      var riga = formuleDiRiferimento.minRiga + i;
      var colonna = formuleDiRiferimento.minColonna + j;
      var cellaA1 = convertiCoordinateInA1(colonna, riga);
      
      if (rangeFormule[i][j]) {
        formuleEffettive[cellaA1] = rangeFormule[i][j];
      }
    }
  }
  
  // 3. Confronta le formule in memoria (senza più chiamate API)
  var discrepanze = [];
  
  for (var cella in formuleMappa) {
    var formulaAttesa = formuleMappa[cella];
    var formulaEffettiva = formuleEffettive[cella] || "";
    
    if (normalizzaFormula(formulaAttesa) !== normalizzaFormula(formulaEffettiva)) {
      discrepanze.push({
        cella: cella,
        attesa: formulaAttesa,
        trovata: formulaEffettiva,
        coordinate: convertiA1InCoordinate(cella)
      });
    }
  }
  
  // 4. Se richiesto e ci sono discrepanze, ripristina le formule corrette
  var correzioniEffettuate = 0;
  
  if (ripristina && discrepanze.length > 0) {
    // Per ogni discrepanza, imposta la formula corretta
    for (var i = 0; i < discrepanze.length; i++) {
      var d = discrepanze[i];
      try {
        var cella = foglioBudget.getRange(d.coordinate.riga, d.coordinate.colonna);
        cella.setFormula(d.attesa);
        correzioniEffettuate++;
        
        // Aggiorna le informazioni sulla discrepanza per il report
        discrepanze[i].corretto = true;
      } catch (e) {
        // Se c'è un errore nella correzione, lo segna nel report
        discrepanze[i].corretto = false;
        discrepanze[i].erroreCorrezione = e.message;
      }
    }
  }
  
  // 5. Genera il report
  var messaggio = "";
  var reportFormattato = "";
  
  if (discrepanze.length > 0) {
    var totalFormule = Object.keys(formuleMappa).length;
    messaggio = "File '" + nomeFile + "': Trovate " + discrepanze.length + 
                " discrepanze su " + totalFormule + " formule verificate.";
    
    if (ripristina) {
      messaggio += " Corrette: " + correzioniEffettuate + "/" + discrepanze.length + ".";
    }
    
    reportFormattato = "RISULTATO: " + discrepanze.length + " discrepanze su " + totalFormule + " formule verificate.\n";
    if (ripristina) {
      reportFormattato += "CORREZIONI EFFETTUATE: " + correzioniEffettuate + "/" + discrepanze.length + ".\n";
    }
    reportFormattato += "--------------------------------------------------\n";
    
    for (var i = 0; i < discrepanze.length; i++) {
      var d = discrepanze[i];
      reportFormattato += "CELLA: " + d.cella + "\n";
      reportFormattato += "ATTESA: " + d.attesa + "\n";
      reportFormattato += "TROVATA: " + d.trovata + "\n";
      
      if (ripristina) {
        if (d.corretto) {
          reportFormattato += "AZIONE: Formula corretta con successo\n";
        } else {
          reportFormattato += "AZIONE: Errore nella correzione - " + (d.erroreCorrezione || "Motivo sconosciuto") + "\n";
        }
      }
      
      reportFormattato += "--------------------------------------------------\n";
    }
  } else {
    var totalFormule = Object.keys(formuleMappa).length;
    messaggio = "File '" + nomeFile + "': Tutte le " + totalFormule + " formule corrispondono!";
    reportFormattato = "RISULTATO: Tutte le " + totalFormule + " formule corrispondono!";
  }
  
  return {
    messaggio: messaggio,
    reportFormattato: reportFormattato
  };
}
*/

function convertiA1InCoordinate(riferimentoA1) {
  // Regex per separare lettere da numeri
  var match = riferimentoA1.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
  if (!match) return null;
  
  var colStr = match[2];
  var rigaStr = match[4];
  
  // Calcola l'indice di colonna (A=1, B=2, ...)
  var colonna = 0;
  for (var i = 0; i < colStr.length; i++) {
    colonna = colonna * 26 + (colStr.charCodeAt(i) - 64);
  }
  
  return {
    colonna: colonna,
    riga: parseInt(rigaStr)
  };
}

function convertiCoordinateInA1(colonna, riga) {
  var colStr = "";
  
  while (colonna > 0) {
    var resto = (colonna - 1) % 26;
    colStr = String.fromCharCode(65 + resto) + colStr;
    colonna = Math.floor((colonna - resto) / 26);
  }
  
  return colStr + riga;
}

function normalizzaFormula(formula) {
  if (!formula) return "";
  
  // Rimuove tutti gli spazi per il confronto
  var normalizzata = formula.replace(/\s+/g, "");
  
  // Converte virgolette singole in virgolette doppie per uniformità
  normalizzata = normalizzata.replace(/'/g, "\"");
  
  return normalizzata;
}

function salvaSuDrive(nomeFile, contenuto) {
  try {
    // Cerca se esiste già un file con lo stesso nome
    var files = DriveApp.getFilesByName(nomeFile);
    var file;
    
    if (files.hasNext()) {
      // Se esiste, aggiorna il contenuto
      file = files.next();
      file.setContent(contenuto);
    } else {
      // Altrimenti crea un nuovo file
      file = DriveApp.createFile(nomeFile, contenuto, MimeType.PLAIN_TEXT);
    }
    
    Logger.log("File '" + nomeFile + "' salvato con successo. URL: " + file.getUrl());
    return file;
  } catch (e) {
    Logger.log("Errore nel salvare il file '" + nomeFile + "': " + e.message);
    return null;
  }
}
