/**
 * Config.js - Configurazione Estesa BOM Master
 * Estende CONFIG con configurazioni specifiche
 * Caricato dopo 0_ConfigBase.js
 */

/**
 * Versione e modalità BETA
 * ⚠️ IMPORTANTE: Impostare IS_BETA a false per versione produzione
 */
CONFIG.VERSION = {
  NUMBER: "0.08",
  IS_BETA: false,  // Produzione
  BETA_WARNING: "⚠️ VERSIONE BETA - Solo per test",
  RELEASE_DATE: "2025-11-26"
};

/**
 * Elaborazione batch (specifico file)
 */
CONFIG.BATCH = {
  SIZE: 5,
  DELAY_MS: 60000,
  PROPERTIES_PREFIX: "FILE_PROCESSOR_"
};

/**
 * Notifiche email (specifico file)
 */
CONFIG.NOTIFICATIONS = {
  EMAIL: function() {
    return PropertiesService.getScriptProperties().getProperty("ADMIN_EMAIL") || null;
  },
  SUBJECT_PREFIX: "[CONTROLLO BOM]",
  FROM_NAME: "Sistema Automatico di Controllo"
};

/**
 * Pattern ricerca file su Drive (specifico file)
 */
CONFIG.SEARCH = {
  BOM_PATTERN: "_BOM_",
  REV_PATTERN: "#Rev 0",
  VERSION: ".08",
  buildQuery: function() {
    return "title contains '" + this.BOM_PATTERN + "' and " +
           "title contains '" + this.REV_PATTERN + "' and " +
           "title contains '" + this.VERSION + "'";
  },
  validatePattern: function(nomeFile) {
    return nomeFile.indexOf(this.REV_PATTERN) > -1 &&
           nomeFile.lastIndexOf(this.VERSION) > nomeFile.indexOf(this.REV_PATTERN) + 6;
  }
};

/**
 * Pattern per naming file (specifico file)
 */
CONFIG.NAMING = {
  SOCIETA: {
    A: "Automate",
    B: "Bridger",
    C: "Automate Consulting",
    D: "Latitudo",
    E: "NewBridger"
  },
  generaNomeFile: function(societaCodice, codiceCommessa, pmName, cliente, descrizione, revisione) {
    if (!this.SOCIETA[societaCodice]) {
      throw new Error("Codice società non valido: " + societaCodice);
    }
    return societaCodice + "_BOM_" + codiceCommessa + "_" + pmName + "_" +
           cliente + "_" + descrizione + " #Rev " + revisione;
  },
  PATTERN: /^[A-E]_BOM_\d{5}\.\d{1,2}\.[A-Z]{3}_[^_]+_[^_]+_[^#]+#Rev\s+\d{2}\.\d{2}$/,
  COMMESSA_PATTERN: /(\d{5})\.(\d{1,2})\.[A-Z]{3}/,
  REVISIONE_PATTERN: /#Rev\s+(\d{2})\.(\d{2})/
};

/**
 * Test file ID (specifico file)
 */
CONFIG.TEST = {
  FILE_ID: "18sXHc0oiIcpumQrgACtTSHPdO17t3JrJLQXLqsrmg3I"
};

/**
 * Funzione helper per impostare l'email admin
 */
function setAdminEmail(email) {
  PropertiesService.getScriptProperties().setProperty("ADMIN_EMAIL", email);
  Logger.log("Email admin impostata: " + email);
}

/**
 * Funzione helper per ottenere l'email admin corrente
 */
function getAdminEmail() {
  var email = CONFIG.NOTIFICATIONS.EMAIL();
  if (!email) {
    Logger.log("ATTENZIONE: Email admin non configurata. Usa setAdminEmail('your@email.com') per configurarla.");
  }
  return email;
}
