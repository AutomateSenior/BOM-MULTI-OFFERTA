# BOM Multi Offerta

Sistema completo di gestione Bill of Materials (BOM) per Google Sheets con supporto multi-offerta, costruito con Google Apps Script.

## Caratteristiche Principali

### Sistema Multi-Offerta
- Creazione e gestione di multiple varianti di offerta (Off_01, Off_02, ...)
- Foglio Master nascosto come riferimento principale
- Budget rigenerato dinamicamente dalle offerte selezionate
- Configurazione intuitiva tramite sidebar UI

### Gestione Completa
- **Materiali/Acquisti** con categorizzazione automatica
- **Body Rental** (PM, Senior, Expert)
- **Configurazione Costi Globali** centralizzata
- **WBS/BBS/HIE** (Work/Budget/Hierarchical Breakdown Structure)
- **Generazione automatica nome file**
- **Salvataggio e ripristino formule ed etichette**

## Architettura

### File JavaScript Core
- `0_ConfigBase.js` - Configurazioni base del sistema
- `BOMCore.js` - Funzioni core del BOM
- `Config.js` - Configurazioni specifiche per fogli e offerte
- `Control.js` - Controlli e validazioni
- `OffertaManager.js` - Gestione completa sistema multi-offerta
- `On Edit.js` - Gestione eventi onEdit
- `On Open.js` - Menu e inizializzazione
- `Salva.js` - Funzioni di salvataggio formule/etichette
- `Test.js` - Funzioni di test
- `Util.js` - Utility generali

### Interfaccia Utente
- `OffertaConfigUI.html` - Sidebar configurazione offerte
- `OffertaDialog.html` - Dialog gestione offerte

### Configurazione
- `appsscript.json` - Manifest Google Apps Script
- `.clasp.json` - Configurazione clasp per deploy
- `.claspignore` - File esclusi dal deploy

## Installazione e Deploy

### Prerequisiti
- Account Google
- [clasp](https://github.com/google/clasp) installato e configurato
- Accesso al foglio Google Sheets di destinazione

### Deploy con clasp

```bash
# Clone del repository
git clone https://github.com/[your-username]/BOM-MULTI-OFFERTA.git
cd BOM-MULTI-OFFERTA

# Login a clasp (se non già fatto)
clasp login

# Crea nuovo progetto Apps Script o collega esistente
clasp create --type sheets --title "BOM Multi Offerta"
# oppure
clasp clone [SCRIPT_ID]

# Deploy del codice
clasp push

# Apri editor Apps Script
clasp open
```

### Configurazione Manuale

1. Apri il foglio Google Sheets di destinazione
2. Vai su **Estensioni → Apps Script**
3. Crea i file seguendo la struttura del repository
4. Copia il contenuto di ogni file
5. Salva e ricarica il foglio

## Utilizzo

### Primo Avvio

1. **Apri il foglio Google Sheets**
2. **Ricarica la pagina** (F5) per caricare il menu
3. Menu **Automate → Opzioni Offerta → Inizializza sistema multi-offerta**
4. Il sistema creerà automaticamente:
   - Foglio **Master** (nascosto, copia di Budget)
   - Foglio **Configurazione** (nascosto, settings)
   - Foglio **Off_01** (prima offerta)
   - Rigenera **Budget** con i dati della prima offerta

### Gestione Offerte

**Aprire la configurazione:**
- Menu: **Automate → Opzioni Offerta → Configura offerte...**
- Si aprirà una sidebar a destra

**Aggiungere nuova offerta:**
1. Click su **+ Aggiungi Nuova Offerta**
2. Inserisci nome e descrizione
3. Il sistema crea automaticamente un nuovo foglio Off_XX

**Modificare valori:**
1. Apri il foglio dell'offerta (es. Off_01, Off_02)
2. Modifica le celle VERDI (celle editabili)
3. Le altre celle contengono formule e sono protette

**Rigenerare Budget:**
1. Seleziona/deseleziona le offerte attive dalla sidebar
2. Click su **Rigenera Budget**
3. Il Budget si aggiorna con la somma delle offerte selezionate

### Menu Disponibili

#### Menu Automate → Opzioni Offerta
- **Configura offerte...** - Apre sidebar configurazione
- **Inizializza sistema multi-offerta** - Setup iniziale (una volta sola)
- **Rigenera Budget** - Aggiorna Budget dalle offerte attive
- **Ripristina Master da Budget** - Ripristina Master dal Budget corrente

#### Macro Disponibili
- **BBS_All**, **BBS_Check**, **BBS_Raggruppamento** - Gestione BBS
- **BOM_All**, **BOM_Check**, **BOM_Group** - Gestione BOM
- Vedi `appsscript.json` per lista completa macro

## Struttura Fogli Google Sheets

```
Spreadsheet
├── Budget              (visibile) - Foglio principale, somma offerte attive
├── Off_01              (visibile) - Prima offerta
├── Off_02              (visibile) - Seconda offerta (se creata)
├── Off_XX              (visibile) - Altre offerte
├── Formule             (visibile) - Salvataggio formule
├── Label               (visibile) - Salvataggio etichette
├── Master              (nascosto) - Riferimento principale
└── Configurazione      (nascosto) - Settings sistema multi-offerta
```

## Workflow Tipico

1. **Inizializzazione** (una volta)
   - Inizializza sistema multi-offerta
   - Verifica creazione fogli

2. **Creazione varianti**
   - Aggiungi nuove offerte dalla sidebar
   - Popola valori nelle celle verdi

3. **Gestione combinazioni**
   - Seleziona/deseleziona offerte attive
   - Rigenera Budget per ogni combinazione
   - Budget mostra la somma delle offerte selezionate

4. **Export e condivisione**
   - Esporta Budget come PDF o Excel
   - Condividi foglio Google Sheets

## Documentazione Aggiuntiva

- **[MANUALE_UTENTE.md](./MANUALE_UTENTE.md)** - Manuale completo utente
- **[GUIDA_SISTEMA_MULTI_OFFERTA.md](./GUIDA_SISTEMA_MULTI_OFFERTA.md)** - Guida dettagliata sistema multi-offerta
- **[QUICK_START.md](./QUICK_START.md)** - Quick start con troubleshooting
- **[Idee per il BOM.docx](./Idee%20per%20il%20BOM.docx)** - Documento requisiti originali

## Foglio di Riferimento

**BOM V08 Produzione:**
https://docs.google.com/spreadsheets/d/1puDsHyN-bQCBQVwf2dhkK3HgNGfshExoThiXw2r5ROY

Questo foglio contiene la versione V08 completa e funzionante del sistema.

## API Google utilizzate

- **Google Sheets API v4** - Manipolazione fogli
- **Google Drive API v3** - Gestione file
- **Google Apps Script** - Runtime V8

## Sviluppo

### Struttura Codice

```javascript
// Esempio di utilizzo API
const config = Config.OFFERTE;
const offerteAttive = OffertaManager.getOfferteAttive();
OffertaManager.rigeneraBudget();
```

### Testing

Il file `Test.js` contiene funzioni di test per:
- Verifica configurazione
- Test creazione offerte
- Validazione formule
- Check integrità dati

### Comandi clasp Utili

```bash
# Pull dal Google Apps Script
clasp pull

# Push modifiche locali
clasp push

# Watch per auto-push
clasp push --watch

# Visualizza versioni
clasp versions

# Apri nel browser
clasp open
```

## Troubleshooting

### Menu non appare
1. Vai su **Estensioni → Apps Script**
2. Seleziona funzione `onOpen`
3. Click **Esegui** (▶)
4. Autorizza permessi
5. Ricarica foglio (F5)

### Sistema già inizializzato
- Verifica esistenza fogli nascosti: **Visualizza → Fogli nascosti**
- Se esiste "Configurazione", sistema è già inizializzato

### Errori durante operazioni
1. Apri **Estensioni → Apps Script**
2. Vai su **Esecuzioni** (icona orologio)
3. Cerca errori in rosso
4. Verifica log per dettagli

## Versione

**Versione attuale:** V08
**Data ultima modifica:** Dicembre 2025
**Stato:** Production Ready

## Licenza

[Specificare licenza]

## Autori

[Specificare autori/contributori]

## Supporto

Per problemi o domande:
1. Controlla la documentazione in `/docs`
2. Verifica esempi nel foglio di riferimento
3. Apri issue su GitHub
