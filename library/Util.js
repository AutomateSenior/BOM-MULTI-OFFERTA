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
