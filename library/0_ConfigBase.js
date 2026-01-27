/**
 * 0_ConfigBase.js - Configurazione Base BOM
 * Caricato per primo (ordine alfabetico 0_)
 */

var CONFIG = {

  /**
   * Versione libreria
   */
  LIB_VERSION: "1.5.0",

  /**
   * File esterni
   */
  FILES: {
    COMMESSE_ID: "1SkWjnjyS2hoxq9Wf7t0BlPeWoQEAJ2R-_a2D9nvwFQ4",
    COMMESSE_SHEET_NAME: "commesse",
    COMMESSE_COLUMN_CODE: "A",
    COMMESSE_COLUMN_TYPE: "AH"
  },

  /**
   * Nomi fogli
   */
  SHEETS: {
    BUDGET: "Budget",
    MASTER: "Master",
    CONFIGURAZIONE_OFFERTE: "Configurazione",
    FORMULE: "Formule",
    LABEL: "Label"
  },

  /**
   * Configurazione sistema multi-offerta
   */
  OFFERTE: {
    RANGE_SOMMA_INIZIO: 69,
    RANGE_SOMMA_FINE: 526,  // V08: Ultima riga dei fogli (aggiornato dopo inserimento Sviluppo Software)
    CELLA_ETICHETTA_SINTESI: "L62",
    PREFISSO_ID: "Off_",
    COLONNE_VERDI: ["L", "N", "S", "T"],
    COLONNE_GIALLE: ["U"],
    COLONNE_CONCATENAZIONE: ["L", "M", "N"],  // V08: Campi descrittivi verdi in L, M e N
    COLORE_TAB_ABILITATA: "#00ff00",
    COLORE_TAB_DISABILITATA: "#cccccc",
    CELLE_CONSISTENZA: [
      {
        cella: "S476",
        nome: "Tipo di assistenza",
        messaggioErrore: "Tipo di assistenza incoerente tra le diverse offerte"
      }
    ]
  },

  /**
   * Tipi di commessa e relativi parametri
   */
  COMMESSA_TYPES: {
    AUT: {
      nome: "AUT",
      perc: 0.2,
      percLinea: 0.25,
      cost: [89, 59, 42],
      vend: [90, 70, 56],
      vendBodyRental: [4000, 3000, 2500],  // Ricavi Body Rental: PM, Senior, Expert
      costBodyRental: [3500, 2500, 2000]   // Costi Body Rental: PM, Senior, Expert
    },
    MES: {
      nome: "MES",
      perc: 0.25,
      percLinea: 0.35,
      cost: [84, 68, 45],
      vend: [90, 75, 65],
      vendBodyRental: [4000, 3000, 2500],
      costBodyRental: [3500, 2500, 2000]
    },
    BI: {
      nome: "BI",
      perc: 0.25,
      percLinea: 0.35,
      cost: [84, 68, 45],
      vend: [90, 75, 65],
      vendBodyRental: [4000, 3000, 2500],
      costBodyRental: [3500, 2500, 2000]
    },
    "SIM/AI": {
      nome: "SIM/AI",
      perc: 0.3,
      percLinea: 0.40,
      cost: [65, 61, 45],
      vend: [86, 80, 65],
      vendBodyRental: [4000, 3000, 2500],
      costBodyRental: [3500, 2500, 2000]
    },
    DEFAULT: {
      nome: "Commessa non esistente",
      perc: -1,
      percLinea: -1,
      cost: [100, 200, 300],
      vend: [4, 5, 6],
      vendBodyRental: [4000, 3000, 2500],
      costBodyRental: [3500, 2500, 2000]
    },
    getParams: function(tipo) {
      return this[tipo] || this.DEFAULT;
    }
  },

  /**
   * Celle da gestire nel foglio Budget
   */
  CELLS: {
    NOME_FILE: "T56",
    CODICE_COMMESSA: "L56",
    INDICATORE_RIGENERA: "O62",  // V08: Indicatore "rigenerazione necessaria"
    TIPO_COMMESSA: "T2",
    PERC_GENERALE: "S2",
    VEND_1: "S4",
    VEND_2: "S5",
    VEND_3: "S6",
    COST_1: "T4",
    COST_2: "T5",
    COST_3: "T6",
    // Body Rental: Ricavi in Y, Costi in Z
    VEND_BR_1: "Y4",  // Ricavo Body Rental PM
    VEND_BR_2: "Y5",  // Ricavo Body Rental Senior
    VEND_BR_3: "Y6",  // Ricavo Body Rental Expert
    COST_BR_1: "Z4",  // Costo Body Rental PM
    COST_BR_2: "Z5",  // Costo Body Rental Senior
    COST_BR_3: "Z6",  // Costo Body Rental Expert
    EXCLUDED_FROM_LABEL: ["T2", "T56", "S56", "S2", "S4", "S5", "S6", "T4", "T5", "T6", "Y4", "Y5", "Y6", "Z4", "Z5", "Z6", "O62"]
  },

  /**
   * Righe specifiche V08 (dopo migrazione colonne e Body Rental)
   */
  RIGHE_V08: {
    MATERIALI: {
      HEADER: 440,         // Riga header L3 "Acquisti" (+11 dopo Sviluppo Software)
      INIZIO: 441,         // Prima riga materiali L4
      FINE: 459,           // Ultima riga materiali L4
      NUM_RIGHE: 19        // Totale righe materiali
    },
    BODY_RENTAL: {
      HEADER: 505,         // Riga header L2 "Body Rental" (+11 dopo Sviluppo Software)
      PM: {
        HEADER: 506,       // Header L3 "Project Manager"
        DETTAGLIO_INIZIO: 507,
        DETTAGLIO_FINE: 509,
        CELLE_COSTI: { RICAVO: "Y4", COSTO: "Z4" }
      },
      SENIOR: {
        HEADER: 510,       // Header L3 "Senior Consultant"
        DETTAGLIO_INIZIO: 511,
        DETTAGLIO_FINE: 513,
        CELLE_COSTI: { RICAVO: "Y5", COSTO: "Z5" }
      },
      EXPERT: {
        HEADER: 514,       // Header L3 "Expert Consultant"
        DETTAGLIO_INIZIO: 515,
        DETTAGLIO_FINE: 517,
        CELLE_COSTI: { RICAVO: "Y6", COSTO: "Z6" }
      },
      COLONNA_MESI: "S"    // Colonna per input mesi (gialla)
    },
    ONERI_SICUREZZA: 518   // V08: Era 507 prima di Sviluppo Software
  },

  /**
   * Colonne da escludere (campi descrittivi verdi)
   */
  COLUMNS: {
    EXCLUDED_FROM_LABEL: ["L", "M", "N"]  // V08: Campi descrittivi in L, M e N
  },

  /**
   * Colori per identificare celle
   */
  COLORS: {
    GREEN: ["#00ff00", "#ccffcc", "#d9ead3", "#93c47d", "#6aa84f"],
    YELLOW: ["#ffff00", "#fff2cc", "#ffe599", "#f1c232"],
    isGreenOrYellow: function(colorHex) {
      if (!colorHex) return false;
      var colorLower = colorHex.toLowerCase();
      return this.GREEN.concat(this.YELLOW).indexOf(colorLower) > -1;
    }
  },

  /**
   * Messaggi di errore standard
   */
  ERRORS: {
    FILE_VUOTO: "Nome file vuoto o non definito",
    FORMATO_NON_VALIDO: "Formato del nome file non valido",
    COMMESSA_NON_TROVATA: "Codice commessa non trovato",
    FOGLIO_NON_TROVATO: "Foglio '{{foglio}}' non trovato",
    FILE_COMMESSE_NON_ACCESSIBILE: "Impossibile accedere al file delle commesse",
    CELLA_VUOTA: "Il valore nella cella {{cella}} è vuoto o non valido",
    format: function(errorType, params) {
      var message = this[errorType] || errorType;
      if (params) {
        for (var key in params) {
          message = message.replace("{{" + key + "}}", params[key]);
        }
      }
      return message;
    }
  },

  /**
   * Utilità per logging
   */
  LOG: {
    ENABLED: true,
    info: function(tag, message) {
      if (this.ENABLED) {
        Logger.log("[" + tag + "] " + message);
      }
    },
    error: function(tag, message, error) {
      Logger.log("[ERROR:" + tag + "] " + message);
      if (error) {
        Logger.log("Stack trace: " + error.stack);
      }
    },
    warn: function(tag, message) {
      if (this.ENABLED) {
        Logger.log("[WARNING:" + tag + "] " + message);
      }
    }
  }
};
