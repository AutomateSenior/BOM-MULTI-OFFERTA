/**
 * Funzione per ripristinare le formule in tutti i file
 */
function ripristinaFormuleInTuttiIFiles() {
  verificaERipristinaInTuttiIFiles("formule", true);
}

/**
 * Funzione per ripristinare le label in tutti i file
 */
function ripristinaLabelInTuttiIFiles() {
  verificaERipristinaInTuttiIFiles("label", true);
}

/**
 * Funzione per ripristinare sia formule che label in tutti i file
 */
function ripristinaFormuleELabelInTuttiIFiles() {
  verificaERipristinaInTuttiIFiles("entrambi", true);
}

/**
 * Funzione per verificare le formule in tutti i file senza ripristinarle
 */
function verificaFormuleInTuttiIFiles() {
  verificaERipristinaInTuttiIFiles("formule", false);
}

/**
 * Funzione per verificare le label in tutti i file senza ripristinarle
 */
function verificaLabelInTuttiIFiles() {
  verificaERipristinaInTuttiIFiles("label", false);
}

/**
 * Funzione per verificare sia formule che label in tutti i file senza ripristinarle
 */
function verificaFormuleELabelInTuttiIFiles() {
  verificaERipristinaInTuttiIFiles("entrambi", false);
}

/**
 * Carica le formule di riferimento dal foglio attivo.
 */
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

/**
 * Cerca file in Google Drive che corrispondono al pattern.
 */
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

/**
 * Funzione per ripristinare il valore originale dalle label,
 * mantenendo il formato originale (inclusi gli zeri iniziali)
 * @param {string} cellValue - Il valore salvato nel foglio Label
 * @return {string|number} - Il valore originale ripristinato
 */
function ripristinaFormatoOriginale(cellValue) {
  // Controlla se il valore inizia con il marcatore "TEXT:"
  if (typeof cellValue === 'string' && cellValue.startsWith("TEXT:")) {
    // Restituisci il valore originale senza il marcatore
    return cellValue.substring(5);
  }
  
  // Se non è un valore marcato, restituiscilo così com'è
  return cellValue;
}

/**
 * Funzione aggiornata per caricare le label di riferimento.
 * Include anche il formato numerico per ogni cella.
 */
function caricaLabelDiRiferimento() {
  var spreadsheetAttivo = SpreadsheetApp.getActiveSpreadsheet();
  var foglioLabel = spreadsheetAttivo.getSheetByName("Label");
  
  if (!foglioLabel) {
    return { labelMappa: {}, formattiMappa: {}, minRiga: 0, maxRiga: 0, minColonna: 0, maxColonna: 0 };
  }
  
  // Carica i dati dal foglio Label
  var ultimaRiga = foglioLabel.getLastRow();
  if (ultimaRiga <= 1) {
    return { labelMappa: {}, formattiMappa: {}, minRiga: 0, maxRiga: 0, minColonna: 0, maxColonna: 0 };
  }
  
  // Ottieni l'ultima colonna basandosi sulle intestazioni
  var ultimaColonna = 2; // Default: posizione e contenuto
  var intestazioni = foglioLabel.getRange(1, 1, 1, 10).getValues()[0]; // Controlla fino a 10 colonne
  for (var i = 0; i < intestazioni.length; i++) {
    if (intestazioni[i] === "Formato") {
      ultimaColonna = i + 1;
      break;
    }
  }
  
  // Carica i dati includendo il formato se disponibile
  var datiLabel = foglioLabel.getRange(2, 1, ultimaRiga - 1, ultimaColonna).getValues();
  
  // Costruisci una mappa per le label attese e i formati e trova i range estremi
  var labelMappa = {};
  var formattiMappa = {};
  var minRiga = 9999, maxRiga = 0, minColonna = 9999, maxColonna = 0;
  
  for (var i = 0; i < datiLabel.length; i++) {
    if (!datiLabel[i][0]) continue;
    
    var riferimentoCella = datiLabel[i][0].toString().trim();
    var labelAttesa = datiLabel[i][1] !== null ? datiLabel[i][1].toString() : "";
    
    // Salva il formato se disponibile (colonna 3)
    if (ultimaColonna >= 3 && datiLabel[i][2]) {
      formattiMappa[riferimentoCella] = datiLabel[i][2].toString();
    }
    
    labelMappa[riferimentoCella] = labelAttesa;
    
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
    labelMappa: labelMappa,
    formattiMappa: formattiMappa,
    minRiga: minRiga,
    maxRiga: maxRiga,
    minColonna: minColonna,
    maxColonna: maxColonna
  };
}

/**
 * Funzione principale che cerca i file e verifica/ripristina formule e/o label.
 * @param {string} tipoOperazione - Tipo di operazione: "formule", "label" o "entrambi"
 * @param {boolean} ripristina - Se true, ripristina le formule/label errate con quelle di riferimento
 * @return {string} Messaggio con il risultato dell'operazione
 */
function verificaERipristinaInTuttiIFiles(tipoOperazione, ripristina) {
  // Se i parametri non sono specificati, imposta default
  tipoOperazione = tipoOperazione || "formule";
  ripristina = (ripristina === true);
  
  // Salva formule e/o label nel foglio attivo
  if (tipoOperazione === "formule" || tipoOperazione === "entrambi") {
    salvaFormule();
  }
  if (tipoOperazione === "label" || tipoOperazione === "entrambi") {
    salvaLabel();
  }
  
  // Ottieni le formule e/o label di riferimento dal foglio attivo
  var formuleDiRiferimento = (tipoOperazione === "formule" || tipoOperazione === "entrambi") ? 
                             caricaFormuleDiRiferimento() : 
                             { formuleMappa: {}, minRiga: 0, maxRiga: 0, minColonna: 0, maxColonna: 0 };
  
  var labelDiRiferimento = (tipoOperazione === "label" || tipoOperazione === "entrambi") ? 
                          caricaLabelDiRiferimento() : 
                          { labelMappa: {}, minRiga: 0, maxRiga: 0, minColonna: 0, maxColonna: 0 };
  
  // Verifica se ci sono formule o label da controllare
  var nessunRiferimento = true;
  
  if (tipoOperazione === "formule" || tipoOperazione === "entrambi") {
    if (Object.keys(formuleDiRiferimento.formuleMappa).length === 0) {
      Logger.log("Nessuna formula di riferimento trovata nel foglio attivo.");
      if (tipoOperazione === "formule") {
        return "Nessuna formula di riferimento trovata.";
      }
    } else {
      nessunRiferimento = false;
      Logger.log("Caricate " + Object.keys(formuleDiRiferimento.formuleMappa).length + " formule di riferimento.");
    }
  }
  
  if (tipoOperazione === "label" || tipoOperazione === "entrambi") {
    if (Object.keys(labelDiRiferimento.labelMappa).length === 0) {
      Logger.log("Nessuna label di riferimento trovata nel foglio attivo.");
      if (tipoOperazione === "label") {
        return "Nessuna label di riferimento trovata.";
      }
    } else {
      nessunRiferimento = false;
      Logger.log("Caricate " + Object.keys(labelDiRiferimento.labelMappa).length + " label di riferimento.");
    }
  }
  
  if (nessunRiferimento) {
    return "Nessuna formula o label di riferimento trovata.";
  }
  
  // Cerca i file su Drive
  var files = cercaFilesByPattern();
  
  if (files.length === 0) {
    Logger.log("Nessun file trovato che corrisponde al pattern richiesto.");
    return "Nessun file trovato.";
  }
  
  Logger.log("Trovati " + files.length + " files da verificare.");
  
  // Array per raccogliere i risultati
  var risultatiComplessivi = [];
  var contenutoFileDifferenze = "REPORT DIFFERENZE\n";
  contenutoFileDifferenze += "Data: " + new Date().toLocaleString() + "\n";
  contenutoFileDifferenze += "Tipo: " + tipoOperazione.toUpperCase() + "\n";
  contenutoFileDifferenze += "Modalità: " + (ripristina ? "VERIFICA E RIPRISTINO" : "SOLO VERIFICA") + "\n";
  contenutoFileDifferenze += "==================================================\n\n";
  
  // Verifica ogni file trovato
  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    var nomeFile = file.getName();
    contenutoFileDifferenze += "FILE: " + nomeFile + "\n";
    
    try {
      var spreadsheet = SpreadsheetApp.open(file);
      var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
      
      if (!foglioBudget) {
        var messaggio = "File '" + nomeFile + "': Foglio 'Budget' non trovato.";
        risultatiComplessivi.push(messaggio);
        contenutoFileDifferenze += messaggio + "\n\n";
        continue;
      }
      
      // Verifica ultra-veloce con le formule e/o label di riferimento già caricate
      var risultatoFormule, risultatoLabel;
      
      if (tipoOperazione === "formule" || tipoOperazione === "entrambi") {
        risultatoFormule = verificaFormuleUltraVeloce(spreadsheet, nomeFile, formuleDiRiferimento, ripristina);
        risultatiComplessivi.push(risultatoFormule.messaggio);
        contenutoFileDifferenze += "VERIFICA FORMULE:\n" + risultatoFormule.reportFormattato + "\n\n";
      }
      
      if (tipoOperazione === "label" || tipoOperazione === "entrambi") {
        // Utilizziamo una versione non modificata della funzione e aggiungiamo il logging separatamente
        risultatoLabel = verificaLabelUltraVeloce(spreadsheet, nomeFile, labelDiRiferimento, ripristina);
        risultatiComplessivi.push(risultatoLabel.messaggio);
        contenutoFileDifferenze += "VERIFICA LABEL:\n" + risultatoLabel.reportFormattato + "\n\n";
      }
      
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
  var nomeFileReport = "differenze_" + tipoOperazione + ".txt";
  salvaSuDrive(nomeFileReport, contenutoFileDifferenze);
  
  var messaggioFinale = "Analisi completata. " + files.length + " file verificati. ";
  if (ripristina) {
    messaggioFinale += "Gli elementi errati sono stati corretti. ";
  }
  messaggioFinale += "Report salvato come '" + nomeFileReport + "' nella cartella principale di Drive.";
  
  return messaggioFinale;
}

/**
 * Funzione che verifica le label tenendo conto del formato
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - Il foglio di calcolo da verificare
 * @param {string} nomeFile - Il nome del file
 * @param {Object} labelDiRiferimento - Oggetto con le label di riferimento
 * @param {boolean} ripristina - Se true, ripristina le label errate
 * @return {Object} Oggetto con messaggio e reportFormattato
 */
function verificaLabelUltraVeloce(spreadsheet, nomeFile, labelDiRiferimento, ripristina) {
  var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
  var labelMappa = labelDiRiferimento.labelMappa;
  var minRiga = labelDiRiferimento.minRiga;
  var maxRiga = labelDiRiferimento.maxRiga;
  var minColonna = labelDiRiferimento.minColonna;
  var maxColonna = labelDiRiferimento.maxColonna;
  
  // Ottieni tutte le celle nell'intervallo in una sola operazione
  var intervallo = foglioBudget.getRange(minRiga, minColonna, maxRiga - minRiga + 1, maxColonna - minColonna + 1);
  var valori = intervallo.getValues();
  
  var labelErrate = 0;
  var labelRipristinate = 0;
  var reportFormattato = "";
  
  // Verifica ogni cella nell'intervallo
  for (var riga = 0; riga < valori.length; riga++) {
    for (var colonna = 0; colonna < valori[0].length; colonna++) {
      var rigaEffettiva = minRiga + riga;
      var colonnaEffettiva = minColonna + colonna;
      
      // Ottieni la notazione A1 della cella
      var posizioneCella = foglioBudget.getRange(rigaEffettiva, colonnaEffettiva).getA1Notation();
      
      // Se esiste una label di riferimento per questa posizione
      if (labelMappa[posizioneCella]) {
        var valoreAttuale = valori[riga][colonna];
        var valoreRiferimento = labelMappa[posizioneCella];
        
        // Verifica se i valori sono semanticamente equivalenti
        var risultatoConfronto = confrontaValoriConFormato(valoreAttuale, valoreRiferimento);
        
        if (!risultatoConfronto.sonoEquivalenti) {
          labelErrate++;
          
          // Log dettagliato delle differenze
          Logger.log("Label diversa in " + nomeFile + " - Cella " + posizioneCella + 
                     "- Attuale='" + valoreAttuale + "' (" + typeof valoreAttuale + ")" + 
                     "- Master='" + valoreRiferimento + "' (" + typeof valoreRiferimento + ")" +
                     "- Motivo: " + risultatoConfronto.motivo);
          
          reportFormattato += "Cella " + posizioneCella + 
                            "- Label attuale '" + valoreAttuale + 
                            "' - differisce dalla label di riferimento '" + valoreRiferimento + "'\n";
          
          // Ripristina il valore se richiesto
          if (ripristina) {
            foglioBudget.getRange(rigaEffettiva, colonnaEffettiva).setValue(valoreRiferimento);
            labelRipristinate++;
          }
        }
      }
    }
  }
  
  // Prepara il messaggio di riepilogo
  var messaggio = "File '" + nomeFile + "': ";
  if (labelErrate === 0) {
    messaggio += "Tutte le label sono corrette.";
    reportFormattato = "Nessuna label errata trovata.\n";
  } else {
    messaggio += "Trovate " + labelErrate + " label errate";
    if (ripristina) {
      messaggio += ", " + labelRipristinate + " ripristinate.";
    } else {
      messaggio += ".";
    }
  }
  
  return {
    messaggio: messaggio,
    reportFormattato: reportFormattato
  };
}

/**
 * Confronta i valori tenendo conto dei formati e prefissi
 * @param {*} valoreAttuale - Il valore attuale della cella
 * @param {*} valoreRiferimento - Il valore di riferimento con possibile prefisso TEXT:
 * @return {Object} Oggetto con proprietà sonoEquivalenti e motivo della differenza
 */
function confrontaValoriConFormato(valoreAttuale, valoreRiferimento) {
  // Se i valori sono identici (incluso il tipo), sono equivalenti
  if (valoreAttuale === valoreRiferimento) {
    return { sonoEquivalenti: true, motivo: "Valori identici" };
  }

  // Converti entrambi i valori in stringhe
  var strAttuale = String(valoreAttuale);
  var strRiferimento = String(valoreRiferimento);
  
  // Se le stringhe sono identiche dopo la conversione, sono equivalenti
  if (strAttuale === strRiferimento) {
    return { sonoEquivalenti: true, motivo: "Valori identici come stringhe" };
  }

  // Gestisci il caso in cui il riferimento è nel formato TEXT:
  if (typeof valoreRiferimento === 'string' && valoreRiferimento.startsWith('TEXT:')) {
    // Estrai il valore dopo TEXT:
    var valoreFormattato = valoreRiferimento.substring(5);
    
    // Se il valore formattato è vuoto, confronta con stringa vuota o 0
    if (valoreFormattato === '') {
      return { 
        sonoEquivalenti: valoreAttuale === '' || valoreAttuale === 0, 
        motivo: "Confronto con valore vuoto" 
      };
    }
    
    // Gestisci formati monetari (con €)
    if (valoreFormattato.includes('€')) {
      // Rimuovi il simbolo dell'euro e spazi
      var valoreNumerico = valoreFormattato.replace('€', '').trim();
      
      // Gestione speciale per formato €0
      if (valoreNumerico === '0' || valoreFormattato === '€0') {
        return { 
          sonoEquivalenti: valoreAttuale === 0, 
          motivo: "Confronto con zero monetario" 
        };
      }
      
      // Gestisci formati con punto come separatore delle migliaia (1.600€)
      if (valoreNumerico.includes('.') && !valoreNumerico.includes(',')) {
        valoreNumerico = valoreNumerico.replace(/\./g, '');
      } 
      // Gestisci formati con sia punto che virgola (es. 1.234,56€)
      else if (valoreNumerico.includes('.') && valoreNumerico.includes(',')) {
        valoreNumerico = valoreNumerico.replace(/\./g, '').replace(',', '.');
      }
      // Altrimenti, converti virgola in punto per numeri decimali (se ci sono)
      else {
        valoreNumerico = valoreNumerico.replace(',', '.');
      }
      
      // Converti in numero per confronto
      var numeroRiferimento = parseFloat(valoreNumerico);
      
      // Gestisci il caso in cui il valore attuale è già un numero
      if (typeof valoreAttuale === 'number') {
        // Confronta con una piccola tolleranza per errori di arrotondamento
        var sonoSimili = Math.abs(valoreAttuale - numeroRiferimento) < 0.01;
        return { 
          sonoEquivalenti: sonoSimili, 
          motivo: sonoSimili ? "Equivalenza numerica" : "Valori numerici differenti: " + valoreAttuale + " vs " + numeroRiferimento
        };
      }
    }
    
    // Gestisci formati percentuali (con %)
    if (valoreFormattato.includes('%')) {
      // Rimuovi il simbolo di percentuale e spazi
      var valorePercentuale = valoreFormattato.replace('%', '').trim();
      
      // Converti virgola in punto per numeri decimali (se ci sono)
      valorePercentuale = valorePercentuale.replace(',', '.');
      
      // Gestisci zero percentuale
      if (valorePercentuale === '0' || valorePercentuale === '0.00' || valorePercentuale === '0.0') {
        return { 
          sonoEquivalenti: valoreAttuale === 0, 
          motivo: "Zero percentuale" 
        };
      }
      
      // Converti in numero per confronto (dividi per 100 se necessario)
      var numeroPercentuale = parseFloat(valorePercentuale) / 100;
      
      // Gestisci il caso in cui il valore attuale è già un numero
      if (typeof valoreAttuale === 'number') {
        // Il valore potrebbe essere già stato convertito da percentuale a decimale
        if (valoreAttuale <= 1) {
          var sonoSimiliDecimale = Math.abs(valoreAttuale - numeroPercentuale) < 0.0001;
          return { 
            sonoEquivalenti: sonoSimiliDecimale, 
            motivo: sonoSimiliDecimale ? "Equivalenza percentuale (decimale)" : "Valori percentuali differenti (decimale)" 
          };
        } else {
          // O potrebbe essere ancora in formato percentuale (es. 50 invece di 0.5)
          var sonoSimiliPerc = Math.abs(valoreAttuale - parseFloat(valorePercentuale)) < 0.01;
          return { 
            sonoEquivalenti: sonoSimiliPerc, 
            motivo: sonoSimiliPerc ? "Equivalenza percentuale" : "Valori percentuali differenti" 
          };
        }
      }
    }
    
    // Gestisci formati numerici generici (es. "2,0")
    if (/^-?\d+[,\.]\d+$/.test(valoreFormattato)) {
      // Converti virgola in punto per numeri decimali
      var valoreParsato = valoreFormattato.replace(',', '.');
      var numeroFormattato = parseFloat(valoreParsato);
      
      if (typeof valoreAttuale === 'number') {
        var sonoSimiliNum = Math.abs(valoreAttuale - numeroFormattato) < 0.01;
        return { 
          sonoEquivalenti: sonoSimiliNum, 
          motivo: sonoSimiliNum ? "Equivalenza numerica decimale" : "Valori decimali differenti" 
        };
      }
    }
    
    // Gestisci interi semplici
    if (/^\d+$/.test(valoreFormattato)) {
      var numeroIntero = parseInt(valoreFormattato, 10);
      
      if (typeof valoreAttuale === 'number') {
        var sonoUgualiInt = valoreAttuale === numeroIntero;
        return { 
          sonoEquivalenti: sonoUgualiInt, 
          motivo: sonoUgualiInt ? "Equivalenza numerica intera" : "Valori interi differenti" 
        };
      }
    }
    
    // Se arriva qui, non abbiamo trovato un'equivalenza specifica, ma proviamo un confronto di stringhe
    return { 
      sonoEquivalenti: strAttuale.trim() === valoreFormattato.trim(), 
      motivo: "Confronto stringa semplice dopo TEXT:" 
    };
  }
  
  // Confronto di stringhe come fallback
  var stringheEquivalenti = strAttuale.trim() === strRiferimento.trim();
  return { 
    sonoEquivalenti: stringheEquivalenti, 
    motivo: stringheEquivalenti ? "Equivalenza stringhe" : "Stringhe differenti" 
  };
}

/**
 * Verifica le formule con un approccio ultra-veloce:
 * 1. Usa le formule di riferimento già caricate
 * 2. Carica TUTTE le formule del foglio Budget in una volta sola
 * 3. Fa tutti i confronti in memoria
 * 4. Se ripristina=true, corregge le formule errate
 * 
 * @param {Spreadsheet} spreadsheet - Il foglio di calcolo da verificare
 * @param {string} nomeFile - Nome del file per il report
 * @param {Object} formuleDiRiferimento - Oggetto con le formule di riferimento
 * @param {boolean} ripristina - Se true, ripristina le formule errate
 * @return {Object} Oggetto con messaggio e report formattato
 */
function verificaFormuleUltraVeloce(spreadsheet, nomeFile, formuleDiRiferimento, ripristina) {
  var foglioBudget = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
  var formuleMappa = formuleDiRiferimento.formuleMappa;
  
  if (Object.keys(formuleMappa).length === 0) {
    return {
      messaggio: "File '" + nomeFile + "': Nessuna formula da verificare.",
      reportFormattato: "Nessuna formula da verificare."
    };
  }
  
  // Calcola il range che contiene tutte le formule da verificare
  var minRiga = formuleDiRiferimento.minRiga;
  var maxRiga = formuleDiRiferimento.maxRiga;
  var minColonna = formuleDiRiferimento.minColonna;
  var maxColonna = formuleDiRiferimento.maxColonna;
  
  // Carica in memoria tutte le formule nel range interessato
  var rangeFormule = foglioBudget.getRange(
    minRiga, 
    minColonna, 
    maxRiga - minRiga + 1, 
    maxColonna - minColonna + 1
  ).getFormulas();
  
  // Costruisci una mappa delle formule effettivamente presenti nel foglio Budget
  var formuleEffettive = {};
  
  for (var i = 0; i < rangeFormule.length; i++) {
    for (var j = 0; j < rangeFormule[i].length; j++) {
      var riga = minRiga + i;
      var colonna = minColonna + j;
      var cellaA1 = convertiCoordinateInA1(colonna, riga);
      
      // Aggiungi solo le celle che contengono formule
      if (rangeFormule[i][j] !== "") {
        formuleEffettive[cellaA1] = rangeFormule[i][j];
      }
    }
  }
  
  // Confronta le formule in memoria
  var discrepanze = [];
  
  for (var cella in formuleMappa) {
    var formulaAttesa = formuleMappa[cella];
    var formulaEffettiva = formuleEffettive[cella] || "";
    
    // Confronto esatto delle formule
    if (formulaAttesa !== formulaEffettiva) {
      discrepanze.push({
        cella: cella,
        attesa: formulaAttesa,
        trovata: formulaEffettiva,
        coordinate: convertiA1InCoordinate(cella)
      });
    }
  }
  
  // Se richiesto e ci sono discrepanze, ripristina le formule corrette
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
  
  // Genera il report
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
      reportFormattato += "TROVATA: " + (d.trovata !== undefined ? d.trovata : "(formula mancante)") + "\n";
      
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