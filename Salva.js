/**
 * Funzione che legge tutte le formule dal foglio Budget e le salva nel foglio Formule.
 * Se il foglio Formule non esiste, lo crea. Se esiste, lo ripulisce prima.
 * Il foglio Formule viene impostato come nascosto.
 */
function salvaFormule() {
  // Ottieni il foglio di calcolo attivo
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ottieni il foglio Budget
  var foglioSource = ss.getSheetByName("Budget");
  if (!foglioSource) {
    throw new Error("Il foglio Budget non esiste!");
  }
  
  // Controlla se il foglio Formule esiste
  var foglioTarget = ss.getSheetByName("Formule");
  if (!foglioTarget) {
    // Crea il foglio Formule se non esiste
    foglioTarget = ss.insertSheet("Formule");
  } else {
    // Pulisci il foglio se esiste già
    foglioTarget.clear();
  }
  
  // Nascondi il foglio Formule
  foglioTarget.hideSheet();
  
  // Ottieni tutte le formule dal foglio Budget
  var range = foglioSource.getDataRange();
  var values = range.getValues();
  var formulas = range.getFormulas();
  
  // Prepara un array per le formule trovate
  var formuleTrovate = [];
  
  // Scorri tutte le celle per trovare le formule
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j] !== "") {
        // Se la cella contiene una formula, aggiungi le informazioni all'array
        formuleTrovate.push([
          foglioSource.getRange(i+1, j+1).getA1Notation(),  // Posizione della cella in notazione A1
          formulas[i][j],                                   // La formula stessa
          values[i][j]                                      // Il valore calcolato
        ]);
      }
    }
  }
  
  // Scrivi le formule trovate nel foglio Formule
  if (formuleTrovate.length > 0) {
    // Aggiungi intestazioni
    foglioTarget.getRange(1, 1).setValue("Posizione");
    foglioTarget.getRange(1, 2).setValue("Formula");
    foglioTarget.getRange(1, 3).setValue("Valore");
    
    // Scrivi i dati delle formule
    foglioTarget.getRange(2, 1, formuleTrovate.length, 3).setValues(formuleTrovate);
    
    // Formatta la colonna delle formule per adattarsi al contenuto
    foglioTarget.autoResizeColumn(2);
  }
}

/**
 * Funzione che legge tutti i contenuti delle celle dal foglio Budget e li salva nel foglio Label.
 * Preserva il formato originale anche per valori che sembrano numerici (es. "001" o "001.001.001").
 * Se il foglio Label non esiste, lo crea. Se esiste, lo ripulisce prima.
 * Il foglio Label viene impostato come nascosto.
 * La funzione ignora celle con formule, celle evidenziate in verde o giallo, e tutte le celle della colonna M.
 */
function salvaLabel() {
  // Ottieni il foglio di calcolo attivo
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ottieni il foglio Budget
  var foglioSource = ss.getSheetByName("Budget");
  if (!foglioSource) {
    throw new Error("Il foglio Budget non esiste!");
  }
  
  // Controlla se il foglio Label esiste
  var foglioTarget = ss.getSheetByName("Label");
  if (!foglioTarget) {
    // Crea il foglio Label se non esiste
    foglioTarget = ss.insertSheet("Label");
  } else {
    // Pulisci il foglio se esiste già
    foglioTarget.clear();
  }
  
  // Nascondi il foglio Label
  foglioTarget.hideSheet();
  
  // Ottieni i dati dal foglio Budget (limitato dalla colonna A alla colonna AB)
  var range = foglioSource.getRange(1, 1, foglioSource.getLastRow(), 28); // 28 è l'indice della colonna AB
  var values = range.getValues();
  var formulas = range.getFormulas();
  var backgrounds = range.getBackgrounds();
  var numberFormats = range.getNumberFormats(); // Ottieni i formati numerici per preservarli
  
  // Prepara un array per i contenuti delle celle trovati
  var labelTrovate = [];
  
  // Scorri tutte le celle per trovare i contenuti
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      // Converti indici a notazione A1 per controlli specifici
      var cellA1 = foglioSource.getRange(i+1, j+1).getA1Notation();
      var colLetter = cellA1.match(/[A-Z]+/)[0]; // Estrae la lettera della colonna
      
      // Salta le celle S2, T2, T56, S4, S5, S6, T4, T5 e T6 specificamente
      if (cellA1 === "T2" || cellA1 === "T56" || cellA1 === "S2" || 
          cellA1 === "S4" || cellA1 === "S5" || cellA1 === "S6" || 
          cellA1 === "T4" || cellA1 === "T5" || cellA1 === "T6") {
        continue;
      }
      
      // Salta tutte le celle della colonna M
      if (colLetter === "M") {
        continue;
      }
      
      // Controlla se la cella ha un valore non vuoto E non è una formula E non è verde o gialla
      if (values[i][j] !== "" && 
          formulas[i][j] === "" && 
          !isGreenOrYellow(backgrounds[i][j])) {
        
        // Ottieni il valore come appare nella cella
        var cellValue = values[i][j];
        var numberFormat = numberFormats[i][j];
        
        // Conserva il formato originale del valore
        var preservedValue;
        
        // Ottieni SEMPRE il valore visualizzato, indipendentemente dal tipo
        var displayedValue = foglioSource.getRange(i+1, j+1).getDisplayValue();
        
        // Controlla se il valore ha un formato da preservare (numeri con zeri iniziali, formati speciali, ecc.)
        if (typeof cellValue === 'number' || 
            (typeof cellValue === 'string' && !isNaN(parseFloat(cellValue))) ||
            displayedValue.indexOf('.') > -1 || 
            (displayedValue.length > 0 && displayedValue.charAt(0) === '0')) {
            
          // Preserva sempre il valore visualizzato per qualsiasi dato numerico o con formato speciale
          preservedValue = "TEXT:" + displayedValue;
        } else {
          // Valore non numerico, conservalo così com'è
          preservedValue = cellValue;
        }
        
        // Se la cella contiene un valore e soddisfa le condizioni, aggiungi le informazioni all'array
        labelTrovate.push([
          cellA1,                // Posizione della cella in notazione A1
          preservedValue,        // Il contenuto della cella preservato
          numberFormat           // Il formato numerico della cella per il ripristino
        ]);
      }
    }
  }
  
  // Scrivi i contenuti trovati nel foglio Label
  if (labelTrovate.length > 0) {
    // Aggiungi intestazioni
    foglioTarget.getRange(1, 1).setValue("Posizione");
    foglioTarget.getRange(1, 2).setValue("Contenuto");
    foglioTarget.getRange(1, 3).setValue("Formato");
    
    // Scrivi i dati dei contenuti
    foglioTarget.getRange(2, 1, labelTrovate.length, 3).setValues(labelTrovate);
    
    // Formatta le colonne per adattarsi al contenuto
    foglioTarget.autoResizeColumn(2);
    foglioTarget.autoResizeColumn(3);
  }
}

/**
 * Controlla se un colore di sfondo è verde o giallo
 * @param {string} backgroundColor - Il colore di sfondo in formato RGB (es. "#ffff00")
 * @return {boolean} True se il colore è verde o giallo, altrimenti false
 */
function isGreenOrYellow(backgroundColor) {
  // Colori verdi e gialli in varie tonalità
  var greenYellowColors = [
    "#00ff00", // Verde
    "#ccffcc", // Verde chiaro
    "#d9ead3", // Verde chiaro (usato spesso in Google Sheets)
    "#93c47d", // Verde medio
    "#6aa84f", // Verde scuro
    "#ffff00", // Giallo
    "#fff2cc", // Giallo chiaro
    "#ffe599", // Giallo medio
    "#f1c232"  // Giallo scuro/oro
  ];
  
  return greenYellowColors.indexOf(backgroundColor.toLowerCase()) > -1;
}

/**
 * Funzione di supporto per determinare se un colore è verde o giallo
 * @param {string} colorHex - Il colore in formato HEX (#RRGGBB)
 * @return {boolean} - True se il colore è verde o giallo, false altrimenti
 */
function isGreenOrYellow(colorHex) {
  // Verifica esatta dei codici colore forniti
  return colorHex === "#00ff00" || colorHex === "#ffff00";
}


