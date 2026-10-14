/**
 * ImportComputo.js - Import del computo metrico (foglio Export_BOM) nel BOM
 *
 * Requisiti: Requisiti_Import_Computo_BOM.docx
 * - ogni lotto del computo diventa un'offerta Off_<lotto> (prezzo del lotto in S64)
 * - ore nette in PLC / SCADA (righe libere), FAT e Commissioning (righe fisse)
 * - righe fisse con più figure: quantità sommate, figura di livello più alto
 * - acquisti aggregati per tipologia, una riga fissa per tipologia
 * - righe insufficienti in una sezione: errore, nessuna riga aggiunta
 * - prima dell'import il BOM viene azzerato (offerte eliminate e ricreate dal Master)
 *
 * Flusso: analizzaComputoDaFile / analizzaComputoDaDrive (nessuna modifica al BOM)
 *         → il dialog mostra il riepilogo → eseguiImportComputo(analisi)
 */

var IMPORT_COMPUTO = {
  FOGLIO_EXPORT: "Export_BOM",

  // Figure interne: figura del computo → figura del BOM (colonna M)
  FIGURE_INTERNE: {
    "Programmatore PLC/SCADA": "3) Expert",
    "Tech Leader software": "2) Senior",
    "Project Manager": "1) PM"
  },
  // Livello per la regola "figura di livello più alto"
  LIVELLO_FIGURA: { "1) PM": 3, "2) Senior": 2, "3) Expert": 1 },

  // Figure esterne: vanno negli acquisti
  FIGURE_ESTERNE: ["Progettista elettrico", "Cablatore", "Tecnico esterno in sito"],

  // Tariffe del computo da confrontare con il Budget (riga S/T 4-6)
  CELLE_TARIFFE: {
    "1) PM": { ricavo: "S4", costo: "T4" },
    "2) Senior": { ricavo: "S5", costo: "T5" },
    "3) Expert": { ricavo: "S6", costo: "T6" }
  },

  // Ordine fisso delle tipologie nelle righe Materiali (stessa riga in tutte le offerte)
  TIPOLOGIE: ["Hardware", "Licenze", "Elettrici", "Servizi", "Consulenze"],

  ZONE: ["Italia", "Europa", "Mondo"],

  // Sezione dell'export → sottosezione di Sviluppo con righe libere
  SEZIONI_LIBERE: { "PLC": "PLC", "HMI_SCADA": "SCADA" },

  // Sezione dell'export → riga fissa di Test Interni
  SEZIONI_FAT: {
    "FAT_PLC": "PLC FAT",
    "FAT_HMI_SCADA": "SCADA FAT",
    "FAT_ELETTRICO": "ELETTRICO FAT"
  },

  COLONNE: ["Record", "Chiave", "Valore", "Lotto", "Nome_lotto", "Prezzo_lotto", "Zona",
            "Rientri", "Km_AR", "Sezione", "Attivita", "Descrizione", "Figura", "Quantita",
            "Unita", "Tipologia", "Costo_totale", "Ore", "Costo_orario", "Ricavo_orario"]
};

// ============================================================================
// Entry point chiamati dal dialog
// ============================================================================

/**
 * Analizza un file caricato dal computer (xlsx in base64). Non modifica il BOM.
 * @param {string} base64 - Contenuto del file
 * @param {string} nomeFile - Nome del file
 * @returns {Object} analisi (vedi analizzaExportBOM)
 */
function analizzaComputoDaFile(base64, nomeFile) {
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), MimeType.MICROSOFT_EXCEL, nomeFile);
  var valori = leggiExportBOMDaBlob_(blob);
  return analizzaComputo_(valori, nomeFile);
}

/**
 * Analizza un file su Drive (xlsx o Google Sheets), dato link o ID. Non modifica il BOM.
 * @param {string} linkOId - URL Drive o ID del file
 * @returns {Object} analisi
 */
function analizzaComputoDaDrive(linkOId) {
  var id = estraiIdDrive_(linkOId);
  if (!id) throw new Error("Link o ID Drive non valido: " + linkOId);

  var file = DriveApp.getFileById(id);
  var valori;
  if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
    valori = leggiExportBOMDaSpreadsheet_(SpreadsheetApp.openById(id));
  } else {
    valori = leggiExportBOMDaBlob_(file.getBlob());
  }
  return analizzaComputo_(valori, file.getName());
}

/**
 * Esegue l'import: azzera il BOM, crea un'offerta per lotto, scrive i valori, rigenera il Budget.
 * La conferma dell'azzeramento è data dall'utente nel dialog prima di questa chiamata.
 * @param {Object} analisi - Risultato di analizzaComputoDaFile / analizzaComputoDaDrive
 * @returns {Object} {offerte: [id], avvisi: [string]}
 */
function eseguiImportComputo(analisi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(CONFIG.SHEETS.MASTER);
  var configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIGURAZIONE_OFFERTE);
  if (!master || !configSheet) {
    throw new Error("Sistema multi-offerta non inizializzato (mancano Master o Configurazione).");
  }

  // Ricontrollo lato server: il dialog potrebbe essere rimasto aperto su un BOM cambiato
  var ricontrollo = analizzaComputo_(analisi.righe, analisi.nomeFile);
  if (ricontrollo.errori.length > 0) {
    throw new Error("Import annullato:\n" + ricontrollo.errori.join("\n"));
  }
  var piano = ricontrollo.piano;

  CONFIG.LOG.info("eseguiImportComputo", "Import " + analisi.nomeFile + " — lotti: " +
    piano.offerte.map(function(o) { return o.id; }).join(", "));

  azzeraBOMPerImport_(ss, configSheet);

  var ordine = 1;
  piano.offerte.forEach(function(off) {
    var foglio = master.copyTo(ss);
    foglio.setName(off.id);
    foglio.showSheet();
    configSheet.appendRow([off.id, off.nome, "Import " + analisi.nomeFile, true, ordine++]);
    scriviOfferta_(foglio, off);
  });

  var offerte = getConfigurazioneOfferte();
  ordinaTabOfferte(ss, offerte);
  SpreadsheetApp.flush();

  rigeneraBudgetDaOfferte();

  var avvisi = ricontrollo.avvisi.slice();
  var budget = ss.getSheetByName(CONFIG.SHEETS.BUDGET);
  if (budget.getRange(CONFIG.CELLS.INDICATORE_RIGENERA).getValue() !== "Budget allineato alle offerte") {
    avvisi.push("Offerte importate, ma la rigenerazione del Budget non è stata completata: rilanciarla dal menu.");
  }

  return {
    offerte: piano.offerte.map(function(o) { return o.id; }),
    avvisi: avvisi
  };
}

// ============================================================================
// Lettura del file
// ============================================================================

function estraiIdDrive_(linkOId) {
  var s = String(linkOId || "").trim();
  var m = s.match(/\/d\/([A-Za-z0-9_-]{20,})/) || s.match(/[?&]id=([A-Za-z0-9_-]{20,})/);
  if (m) return m[1];
  return /^[A-Za-z0-9_-]{20,}$/.test(s) ? s : null;
}

/**
 * Converte un xlsx in Google Sheets temporaneo (Drive v2), legge Export_BOM e cestina il temporaneo.
 */
function leggiExportBOMDaBlob_(blob) {
  var temp = Drive.Files.insert(
    { title: "TEMP_import_computo_" + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS },
    blob,
    { convert: true }
  );
  try {
    return leggiExportBOMDaSpreadsheet_(SpreadsheetApp.openById(temp.id));
  } finally {
    try { DriveApp.getFileById(temp.id).setTrashed(true); } catch (e) {
      CONFIG.LOG.warn("leggiExportBOMDaBlob_", "File temporaneo non cestinato: " + temp.id);
    }
  }
}

function leggiExportBOMDaSpreadsheet_(spreadsheet) {
  var foglio = spreadsheet.getSheetByName(IMPORT_COMPUTO.FOGLIO_EXPORT);
  if (!foglio) throw new Error("Il file non contiene il foglio '" + IMPORT_COMPUTO.FOGLIO_EXPORT + "'.");
  // Le date non passano da google.script.run: il dialog riceve e rimanda solo testo e numeri
  return foglio.getDataRange().getValues().map(function(r) {
    return r.map(function(v) { return (v instanceof Date) ? v.toISOString() : v; });
  });
}

// ============================================================================
// Analisi (senza modifiche al BOM)
// ============================================================================

function analizzaComputo_(righe, nomeFile) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(CONFIG.SHEETS.MASTER);
  var budget = ss.getSheetByName(CONFIG.SHEETS.BUDGET);
  if (!master) throw new Error("Foglio Master non trovato: inizializza prima il sistema multi-offerta.");

  var struttura = mappaStrutturaOfferta_(master.getRange(1, 1, master.getLastRow(), 13).getValues());

  var tariffeBudget = {};
  Object.keys(IMPORT_COMPUTO.CELLE_TARIFFE).forEach(function(fig) {
    var c = IMPORT_COMPUTO.CELLE_TARIFFE[fig];
    tariffeBudget[fig] = { ricavo: budget.getRange(c.ricavo).getValue(), costo: budget.getRange(c.costo).getValue() };
  });

  var contesto = {
    commessaBOM: String(budget.getRange(CONFIG.CELLS.CODICE_COMMESSA).getValue() || "").trim(),
    tariffeBudget: tariffeBudget,
    offerteEsistenti: getConfigurazioneOfferte().map(function(o) { return o.id; })
  };

  var analisi = analizzaExportBOM(righe, struttura, contesto);
  analisi.nomeFile = nomeFile;
  analisi.righe = righe;
  return analisi;
}

/**
 * Logica pura (nessuna chiamata Apps Script): valida l'export e costruisce il piano di scrittura.
 * @param {Array<Array>} righe - Valori del foglio Export_BOM (riga 1 = intestazioni)
 * @param {Object} struttura - Da mappaStrutturaOfferta_
 * @param {Object} contesto - {commessaBOM, tariffeBudget, offerteEsistenti}
 * @returns {{errori: string[], avvisi: string[], riepilogo: Object[], piano: Object}}
 */
function analizzaExportBOM(righe, struttura, contesto) {
  var errori = [];
  var avvisi = [];
  var C = IMPORT_COMPUTO;

  // ── Intestazioni ──────────────────────────────────────────────────────────
  var col = {};
  (righe[0] || []).forEach(function(h, i) { col[String(h).trim()] = i; });
  var mancanti = C.COLONNE.filter(function(n) { return col[n] === undefined; });
  if (mancanti.length > 0) {
    return { errori: ["Colonne mancanti in Export_BOM: " + mancanti.join(", ")], avvisi: [], riepilogo: [], piano: null };
  }
  var recs = righe.slice(1).map(function(r, i) {
    var o = { riga: i + 2 };
    C.COLONNE.forEach(function(n) {
      var v = r[col[n]];
      o[n] = (typeof v === "string") ? v.trim() : v;
    });
    return o;
  }).filter(function(o) { return o.Record; });

  var num = function(v) { return (v === "" || v === null || v === undefined) ? 0 : Number(v); };

  // ── GENERALE ──────────────────────────────────────────────────────────────
  var gen = {};
  recs.filter(function(r) { return r.Record === "GENERALE"; }).forEach(function(r) { gen[r.Chiave] = r.Valore; });
  ["Commessa", "PM_percentuale", "Rischio_software", "Imprevisti_acquisti", "Margine_acquisti"].forEach(function(k) {
    if (gen[k] === undefined || gen[k] === "") errori.push("GENERALE: manca la chiave '" + k + "'.");
  });

  var commessa = String(gen.Commessa || "").trim();
  if (commessa && contesto.commessaBOM !== commessa) {
    errori.push("Il codice commessa dell'export (" + commessa + ") non coincide con quello del BOM in " +
                CONFIG.CELLS.CODICE_COMMESSA + " (" + (contesto.commessaBOM || "vuoto") + ").");
  }

  // ── TARIFFA (solo avviso) ─────────────────────────────────────────────────
  recs.filter(function(r) { return r.Record === "TARIFFA"; }).forEach(function(r) {
    var fig = C.FIGURE_INTERNE[r.Figura];
    if (!fig) return;
    var b = contesto.tariffeBudget[fig];
    if (b && (Math.abs(num(r.Costo_orario) - num(b.costo)) > 0.005 || Math.abs(num(r.Ricavo_orario) - num(b.ricavo)) > 0.005)) {
      avvisi.push("Tariffa " + r.Figura + ": computo " + r.Costo_orario + "/" + r.Ricavo_orario +
                  " €/h, BOM " + b.costo + "/" + b.ricavo + " €/h. Vengono usate quelle del BOM.");
    }
  });

  // ── LOTTO ─────────────────────────────────────────────────────────────────
  var lotti = [];
  var perLotto = {};
  recs.filter(function(r) { return r.Record === "LOTTO"; }).forEach(function(r) {
    var codice = String(r.Lotto);
    if (!/^[A-Za-z0-9_]+$/.test(codice)) {
      errori.push("Riga " + r.riga + ": codice lotto '" + codice + "' non valido (solo lettere, numeri, underscore).");
      return;
    }
    if (perLotto[codice]) { errori.push("Riga " + r.riga + ": lotto " + codice + " ripetuto."); return; }
    if (C.ZONE.indexOf(r.Zona) < 0) errori.push("Lotto " + codice + ": zona '" + r.Zona + "' non valida (Italia, Europa, Mondo).");
    if (!(num(r.Prezzo_lotto) > 0)) avvisi.push("Lotto " + codice + ": prezzo del lotto mancante o zero.");
    var lotto = {
      codice: codice, id: CONFIG.OFFERTE.PREFISSO_ID + codice, nome: r.Nome_lotto || codice,
      prezzo: num(r.Prezzo_lotto), zona: r.Zona, rientri: num(r.Rientri), km: num(r.Km_AR),
      libere: { PLC: [], SCADA: [] },    // {descrizione, figura, ore}
      fisse: {},                         // chiave → {riga, ore|giorni, figure: {}}
      acquisti: {},                      // tipologia → {costo, voci}
      zeroSaltate: 0
    };
    perLotto[codice] = lotto;
    lotti.push(lotto);
  });
  if (lotti.length === 0) errori.push("Export_BOM non contiene righe LOTTO.");

  var lottoDi = function(r) {
    var l = perLotto[String(r.Lotto)];
    if (!l) errori.push("Riga " + r.riga + ": lotto '" + r.Lotto + "' non dichiarato nelle righe LOTTO.");
    return l;
  };

  // ── ATTIVITA ──────────────────────────────────────────────────────────────
  recs.filter(function(r) { return r.Record === "ATTIVITA"; }).forEach(function(r) {
    var lotto = lottoDi(r);
    if (!lotto) return;
    var q = num(r.Quantita);
    if (q === 0) { lotto.zeroSaltate++; return; }

    var fig = C.FIGURE_INTERNE[r.Figura];
    if (!fig) {
      errori.push("Riga " + r.riga + ": figura '" + r.Figura + "' non prevista per le attività. " +
                  (C.FIGURE_ESTERNE.indexOf(r.Figura) >= 0 ? "Le figure esterne vanno nelle righe ACQUISTO." : "Indicare come trattarla."));
      return;
    }

    var sezLibera = C.SEZIONI_LIBERE[r.Sezione];
    if (sezLibera) {
      if (r.Unita !== "h") { errori.push("Riga " + r.riga + ": nella sezione " + r.Sezione + " l'unità deve essere 'h'."); return; }
      lotto.libere[sezLibera].push({ descrizione: r.Descrizione, figura: fig, ore: q });
      return;
    }

    var chiave, rigaBOM;
    if (C.SEZIONI_FAT[r.Sezione]) {
      if (r.Unita !== "h") { errori.push("Riga " + r.riga + ": nella sezione " + r.Sezione + " l'unità deve essere 'h'."); return; }
      chiave = r.Sezione;
      rigaBOM = struttura.fat[C.SEZIONI_FAT[r.Sezione]];
    } else if (r.Sezione === "COMMISSIONING") {
      if (r.Unita !== "giorni") { errori.push("Riga " + r.riga + ": nel commissioning l'unità deve essere 'giorni'."); return; }
      var blocco = struttura.commissioning[lotto.zona];
      if (!blocco) return; // zona non valida: errore già segnalato sul lotto
      chiave = "COMMISSIONING:" + r.Attivita;
      rigaBOM = blocco[r.Attivita];
      if (!rigaBOM) {
        errori.push("Riga " + r.riga + ": attività di commissioning '" + r.Attivita + "' non presente nel BOM.");
        return;
      }
    } else {
      errori.push("Riga " + r.riga + ": sezione '" + r.Sezione + "' non riconosciuta.");
      return;
    }
    if (!rigaBOM) { errori.push("Riga " + r.riga + ": riga '" + chiave + "' non trovata nel Master."); return; }

    var f = lotto.fisse[chiave] || (lotto.fisse[chiave] = { riga: rigaBOM, quantita: 0, figure: {} });
    f.quantita += q;
    f.figure[fig] = true;
  });

  // ── ACQUISTO ──────────────────────────────────────────────────────────────
  recs.filter(function(r) { return r.Record === "ACQUISTO"; }).forEach(function(r) {
    var lotto = lottoDi(r);
    if (!lotto) return;
    if (C.TIPOLOGIE.indexOf(r.Tipologia) < 0) {
      errori.push("Riga " + r.riga + ": tipologia d'acquisto '" + r.Tipologia + "' non valida (" + C.TIPOLOGIE.join(", ") + ").");
      return;
    }
    if (r.Figura && C.FIGURE_ESTERNE.indexOf(r.Figura) < 0) {
      errori.push("Riga " + r.riga + ": figura esterna '" + r.Figura + "' non prevista. Indicare come trattarla.");
      return;
    }
    var costo = num(r.Costo_totale);
    if (costo === 0 && r.Figura) costo = num(r.Ore) * num(r.Costo_orario);
    if (r.Quantita !== "" && r.Quantita !== null && r.Quantita !== undefined && num(r.Quantita) === 0 && !r.Figura) {
      lotto.zeroSaltate++;
      return;
    }
    if (costo === 0) avvisi.push("Lotto " + lotto.codice + ": acquisto a costo 0 – " + r.Descrizione);
    var a = lotto.acquisti[r.Tipologia] || (lotto.acquisti[r.Tipologia] = { costo: 0, voci: 0 });
    a.costo += costo;
    a.voci++;
  });

  // ── Disposizione comune a tutte le offerte ────────────────────────────────
  // Il Budget richiede la stessa figura (colonna M) nella stessa riga di tutte le offerte.
  var layout = { PLC: [], SCADA: [] };
  ["PLC", "SCADA"].forEach(function(sez) {
    var maxPerFigura = {};
    lotti.forEach(function(l) {
      var conta = {};
      l.libere[sez].forEach(function(x) { conta[x.figura] = (conta[x.figura] || 0) + 1; });
      Object.keys(conta).forEach(function(fig) { maxPerFigura[fig] = Math.max(maxPerFigura[fig] || 0, conta[fig]); });
    });
    var figureOrdinate = Object.keys(maxPerFigura).sort(function(a, b) { return C.LIVELLO_FIGURA[b] - C.LIVELLO_FIGURA[a]; });
    figureOrdinate.forEach(function(fig) {
      for (var i = 0; i < maxPerFigura[fig]; i++) layout[sez].push(fig);
    });
    var disponibili = struttura.libere[sez].length;
    if (layout[sez].length > disponibili) {
      var dettaglio = figureOrdinate.map(function(fig) {
        return fig.replace(/^\d\)\s*/, "") + " " + maxPerFigura[fig];
      }).join(", ");
      errori.push("Sezione " + (sez === "SCADA" ? "HMI_SCADA" : sez) + ": troppe righe. Servono " + layout[sez].length +
                  " righe (" + dettaglio + ", massimo tra i lotti per figura), il BOM ne ha " + disponibili +
                  ". Raggruppare le voci nell'export.");
    }
  });

  // Righe fisse: figura di livello più alto tra tutte le offerte che usano la riga
  var figuraRigaFissa = {};
  lotti.forEach(function(l) {
    Object.keys(l.fisse).forEach(function(k) {
      Object.keys(l.fisse[k].figure).forEach(function(fig) {
        var att = figuraRigaFissa[k];
        if (!att || C.LIVELLO_FIGURA[fig] > C.LIVELLO_FIGURA[att]) figuraRigaFissa[k] = fig;
      });
    });
  });

  // ── Piano di scrittura per offerta ────────────────────────────────────────
  var parametri = {
    pm: num(gen.PM_percentuale), rischio: num(gen.Rischio_software),
    imprevisti: num(gen.Imprevisti_acquisti), margine: num(gen.Margine_acquisti)
  };
  var riepilogo = [];
  var offertePiano = lotti.map(function(l) {
    var celle = [];   // {a1, valore}
    var put = function(colonna, riga, valore) { celle.push({ a1: colonna + riga, valore: valore }); };

    put("S", struttura.prezzoVendita, l.prezzo);
    put("S", struttura.pmRisorseInterne, parametri.pm);
    put("S", struttura.rischio.sviluppo, parametri.rischio);
    put("S", struttura.rischio.test, parametri.rischio);
    put("S", struttura.rischio.commissioning, parametri.rischio);

    ["PLC", "SCADA"].forEach(function(sez) {
      var slot = struttura.libere[sez];
      var usati = {};
      layout[sez].forEach(function(fig, i) { if (slot[i]) put("M", slot[i], fig); });
      l.libere[sez].forEach(function(x) {
        // primo slot libero con la figura giusta
        for (var i = 0; i < layout[sez].length && i < slot.length; i++) {
          if (layout[sez][i] === x.figura && !usati[i]) {
            usati[i] = true;
            put("L", slot[i], x.descrizione);
            put("S", slot[i], x.ore);
            return;
          }
        }
      });
    });

    var oreFisse = 0, giorniComm = 0;
    Object.keys(l.fisse).forEach(function(k) {
      var f = l.fisse[k];
      put("M", f.riga, figuraRigaFissa[k]);
      put("S", f.riga, arrotonda_(f.quantita));
      if (k.indexOf("COMMISSIONING:") === 0) giorniComm += f.quantita; else oreFisse += f.quantita;
    });

    var tr = struttura.trasferte[l.zona];
    if (tr && giorniComm > 0) {
      put("S", tr.km, l.rientri);
      put("T", tr.km, l.km);
      put("S", tr.alloggio, l.rientri);
      put("S", tr.vitto, l.rientri);
    }

    var costoAcquisti = 0;
    C.TIPOLOGIE.forEach(function(tip, i) {
      var a = l.acquisti[tip];
      if (!a) return;
      var riga = struttura.materiali[i];
      put("L", riga, tip + " – " + a.voci + (a.voci === 1 ? " voce" : " voci") + " da computo");
      put("M", riga, tip);
      put("S", riga, arrotonda_(a.costo));
      put("T", riga, parametri.imprevisti);
      put("U", riga, parametri.margine);
      costoAcquisti += a.costo;
    });

    var oreLibere = 0;
    ["PLC", "SCADA"].forEach(function(s) { l.libere[s].forEach(function(x) { oreLibere += x.ore; }); });
    riepilogo.push({
      lotto: l.codice, offerta: l.id, nome: l.nome, prezzo: l.prezzo, zona: l.zona,
      righePLC: l.libere.PLC.length, righeSCADA: l.libere.SCADA.length,
      oreUfficio: arrotonda_(oreLibere + oreFisse), giorniCommissioning: arrotonda_(giorniComm),
      acquisti: arrotonda_(costoAcquisti), zeroSaltate: l.zeroSaltate
    });

    return { id: l.id, nome: l.nome, celle: celle };
  });

  if (struttura.materiali.length < C.TIPOLOGIE.length) {
    errori.push("Il Master ha " + struttura.materiali.length + " righe Materiali, ne servono " + C.TIPOLOGIE.length + ".");
  }

  return {
    errori: errori, avvisi: avvisi, riepilogo: riepilogo,
    commessa: commessa, revisione: gen.Revisione || "",
    offerteDaEliminare: contesto.offerteEsistenti || [],
    piano: errori.length > 0 ? null : { offerte: offertePiano }
  };
}

function arrotonda_(x) { return Math.round(x * 10000) / 10000; }

// ============================================================================
// Struttura del foglio offerta (righe trovate per livello BBS ed etichetta)
// ============================================================================

/**
 * @param {Array<Array>} v - Valori del Master, colonne A..M (H = livello, L = etichetta, M = risorsa)
 * @returns {Object} numeri di riga (1-based) delle celle usate dall'import
 */
function mappaStrutturaOfferta_(v) {
  var H = 7, L = 11, M = 12;
  var livello = function(r) { return Number(v[r - 1][H]) || 0; };
  var etichetta = function(r) { return String(v[r - 1][L] || "").trim(); };
  var n = v.length;

  var trova = function(label, liv, da, a) {
    for (var r = da || 1; r <= (a || n); r++) {
      if (etichetta(r) === label && livello(r) === liv) return r;
    }
    throw new Error("Struttura del Master: riga '" + label + "' (livello " + liv + ") non trovata.");
  };
  // Ultima riga del blocco che inizia in r (righe successive con livello maggiore)
  var fine = function(r) {
    var liv = livello(r);
    for (var i = r + 1; i <= n; i++) {
      if (livello(i) > 0 && livello(i) <= liv) return i - 1;
    }
    return n;
  };
  var figli = function(r) {
    var out = [];
    for (var i = r + 1; i <= fine(r); i++) if (livello(i) === livello(r) + 1) out.push(i);
    return out;
  };
  var figliPerEtichetta = function(r) {
    var m = {};
    figli(r).forEach(function(i) { if (etichetta(i)) m[etichetta(i)] = i; });
    return m;
  };

  var sviluppo = trova("Sviluppo", 2);
  var test = trova("Test Interni", 2);
  var comm = trova("Commissioning", 2);

  var fat = {};
  ["PLC FAT", "HMI FAT", "SCADA FAT", "ELETTRICO FAT"].forEach(function(lbl) {
    fat[lbl] = trova(lbl, 3, test, fine(test));
  });

  var commissioning = {};
  var trasferte = {};
  var trasf = trova("Trasferte", 3, comm, fine(comm));
  IMPORT_COMPUTO.ZONE.forEach(function(z) {
    commissioning[z] = figliPerEtichetta(trova(z, 3, comm, fine(comm)));
    var blocco = trova(z, 4, trasf, fine(trasf));
    var t = figliPerEtichetta(blocco);
    trasferte[z] = { km: t["KM spostamento"], alloggio: t["Alloggio"], vitto: t["Vitto"] };
    if (!t["KM spostamento"] || !t["Alloggio"] || !t["Vitto"]) {
      throw new Error("Struttura del Master: righe trasferte " + z + " incomplete.");
    }
  });

  // Righe Materiali: figli della riga "Acquisti" con M = "Tipo"
  var materiali = null;
  for (var r = 1; r <= n; r++) {
    if (etichetta(r) === "Acquisti" && String(v[r - 1][M]).trim() === "Tipo") { materiali = figli(r); break; }
  }
  if (!materiali) throw new Error("Struttura del Master: righe Materiali non trovate.");

  return {
    prezzoVendita: trova("Prezzo di vendita", 2),
    pmRisorseInterne: trova("PM risorse interne", 3),
    rischio: {
      sviluppo: trova("Rischio", 3, sviluppo, fine(sviluppo)),
      test: trova("Rischio", 3, test, fine(test)),
      commissioning: trova("Rischio", 3, comm, fine(comm))
    },
    libere: {
      PLC: figli(trova("PLC", 3, sviluppo, fine(sviluppo))),
      SCADA: figli(trova("SCADA", 3, sviluppo, fine(sviluppo)))
    },
    fat: fat,
    commissioning: commissioning,
    trasferte: trasferte,
    materiali: materiali
  };
}

// ============================================================================
// Scrittura
// ============================================================================

/**
 * Elimina tutte le offerte Off_XX e le loro righe in Configurazione.
 */
function azzeraBOMPerImport_(ss, configSheet) {
  getConfigurazioneOfferte().forEach(function(o) {
    var foglio = ss.getSheetByName(o.id);
    if (foglio) ss.deleteSheet(foglio);
  });
  if (configSheet.getLastRow() > 1) {
    configSheet.deleteRows(2, configSheet.getLastRow() - 1);
  }
  CONFIG.LOG.info("azzeraBOMPerImport_", "Offerte eliminate");
}

function scriviOfferta_(foglio, offerta) {
  offerta.celle.forEach(function(c) {
    foglio.getRange(c.a1).setValue(c.valore);
  });
  CONFIG.LOG.info("scriviOfferta_", offerta.id + ": " + offerta.celle.length + " celle scritte");
}
