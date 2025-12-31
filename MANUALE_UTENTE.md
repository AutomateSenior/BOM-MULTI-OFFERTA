# BOM Multi-Offerta - Manuale Utente

**Versione:** 1.3.0 (Libreria BOM8 v38)
**Data:** Dicembre 2025  
**Stato:** Production Ready

## Indice

1. [Introduzione](#introduzione)
2. [Primo Avvio](#primo-avvio)
3. [Struttura del Sistema](#struttura-del-sistema)
4. [Gestione Offerte](#gestione-offerte)
5. [Rigenerazione Budget](#rigenerazione-budget)
6. [Celle e Colori](#celle-e-colori)
7. [Funzionalità Avanzate](#funzionalità-avanzate)
8. [FAQ e Risoluzione Problemi](#faq-e-risoluzione-problemi)

---

## Introduzione

Il sistema **BOM Multi-Offerta** permette di gestire multiple varianti di un preventivo (Bill of Materials) all'interno di un unico foglio Google Sheets. Ogni variante (offerta) può avere valori diversi per materiali, risorse, costi, e il sistema può generare automaticamente un Budget che rappresenta la somma di una o più offerte selezionate.

### Caratteristiche Principali

- **Multi-Offerta**: Crea e gestisci N varianti di preventivo (Off_01, Off_02, ...)
- **Rigenerazione Automatica**: Il Budget si rigenera automaticamente dalle offerte attive
- **Controllo Consistenza**: Verifica automatica della coerenza tra offerte (tipo risorse, valori)
- **Gestione Intelligente Colori**:
  - Verde = celle da compilare
  - Blu = celle processate dal sistema
  - Giallo = celle per valori MAX
  - Rosso = errori di consistenza
- **Integrazione Commesse**: Sincronizzazione automatica con file commesse esterno

---

## Primo Avvio

### Passo 1: Verifica Menu

1. Apri il foglio Google Sheets
2. Verifica presenza menu **Automate** nella barra superiore
3. Se il menu non appare:
   - Ricarica il foglio (F5 o Cmd+R)
   - Vai su **Estensioni → Apps Script**
   - Seleziona funzione `onOpen` e clicca **Esegui** (▶)
   - Autorizza i permessi richiesti
   - Ricarica il foglio

### Passo 2: Inizializzazione Sistema

**IMPORTANTE**: Questa operazione va eseguita **UNA SOLA VOLTA** su un foglio nuovo.

1. Menu: **Automate → Opzioni Offerta → Inizializza sistema multi-offerta**
2. Il sistema crea automaticamente:
   - **Master** (nascosto): copia del Budget originale, riferimento principale
   - **Configurazione** (nascosto): configurazione sistema multi-offerta
   - **Off_01**: prima offerta, copia del Master
3. Il **Budget** viene rigenerato con i dati di Off_01

### Passo 3: Verifica Creazione Fogli

1. Guarda i tab in basso: dovresti vedere **Budget**, **Off_01**
2. Per vedere i fogli nascosti: **Visualizza → Fogli nascosti**
3. Dovresti vedere **Master** e **Configurazione**

---

## Struttura del Sistema

### Fogli Visibili

| Foglio | Descrizione | Editabile |
|--------|-------------|-----------|
| **Budget** | Foglio principale, somma delle offerte attive | ⚠️ Solo celle verdi prima di rigenerare |
| **Off_01** | Prima offerta | ✅ Celle verdi |
| **Off_02** | Seconda offerta (se creata) | ✅ Celle verdi |
| **Off_XX** | Altre offerte | ✅ Celle verdi |
| **Formule** | Backup formule Budget | ❌ Gestito automaticamente |
| **Label** | Backup etichette Budget | ❌ Gestito automaticamente |

### Fogli Nascosti

| Foglio | Descrizione | Scopo |
|--------|-------------|-------|
| **Master** | Template di riferimento | Base per nuove offerte |
| **Configurazione** | Settings sistema | Elenco offerte e stati |

### Libreria BOM8

Il codice del sistema è organizzato in una libreria Google Apps Script condivisa:

**ID Libreria**: `1Q4giGpH67WkMfrWabm8QfTvvg70ueC81044yDDFe1OyHpHFA7d3_gvb1`  
**Versione Attuale**: 38 (1.3.0)  
**Percorso**: `/library/`

#### File della Libreria

| File | Descrizione |
|------|-------------|
| `0_ConfigBase.js` | Configurazioni globali del sistema |
| `BOMCore.js` | Funzioni core del BOM |
| `Config.js` | Configurazioni specifiche fogli |
| `Control.js` | Controlli e validazioni |
| `OffertaManager.js` | Gestione completa sistema multi-offerta |
| `RigeneraBudget_V2.js` | Logica rigenerazione Budget con controlli avanzati |
| `Salva.js` | Salvataggio/ripristino formule ed etichette |
| `Util.js` | Funzioni di utilità |

---

## Gestione Offerte

### Aprire la Configurazione

**Menu**: **Automate → Opzioni Offerta → Configura offerte...**

Si apre una **sidebar a destra** con:
- Elenco offerte esistenti
- Stato (abilitata/disabilitata)
- Pulsanti di gestione

### Creare Nuova Offerta

1. Nella sidebar, clicca **+ Aggiungi Nuova Offerta**
2. Inserisci:
   - **Nome**: es. "Offerta Cliente XYZ"
   - **Descrizione**: es. "Variante con materiali premium"
3. Clicca **Aggiungi**
4. Il sistema crea automaticamente:
   - Nuovo foglio **Off_XX** (dove XX è il numero progressivo)
   - Foglio popolato con struttura del Master
   - Offerta automaticamente **abilitata**

### Modificare Valori in un'Offerta

1. Apri il foglio dell'offerta (es. **Off_01**, **Off_02**)
2. Modifica SOLO le **celle VERDI** (celle editabili):
   - Colonna **L**: Descrizione materiale/attività
   - Colonna **M**: Tipo risorsa (PM, Senior, Expert, ...)
   - Colonna **N**: Descrizione dettagliata
   - Colonna **S**: Quantità, ore, mesi
   - Colonna **T**: Costi unitari
   - Colonna **U**: Valori MAX (gialle)
3. Le **celle BLU** contengono formule e sono protette

### Abilitare/Disabilitare Offerta

1. Apri sidebar configurazione
2. Trova l'offerta nell'elenco
3. Usa il toggle per abilitare/disabilitare
4. **Offerta abilitata** = contribuisce al Budget
5. **Offerta disabilitata** = esclusa dal Budget

### Eliminare Offerta

1. Apri sidebar configurazione
2. Clicca **Elimina** sull'offerta da rimuovere
3. Conferma l'operazione
4. Il foglio viene nascosto (non cancellato definitivamente)

---

## Rigenerazione Budget

### Quando Rigenerare

Il Budget deve essere rigenerato ogni volta che:
- Modifichi valori in un'offerta
- Abiliti/disabiliti un'offerta
- Crei una nuova offerta
- Modifichi il codice commessa in **Budget\!L56**

### Come Rigenerare

**Metodo 1 - Da Sidebar:**
1. Apri sidebar configurazione
2. Seleziona le offerte da includere (abilita/disabilita)
3. Clicca **Rigenera Budget**

**Metodo 2 - Da Menu:**
1. Menu: **Automate → Opzioni Offerta → Rigenera Budget**

**Metodo 3 - Da Codice Commessa:**
1. Modifica cella **Budget\!L56** (codice commessa)
2. Il sistema rigenera automaticamente

### Cosa Succede Durante la Rigenerazione

Il sistema esegue le seguenti operazioni in sequenza:

#### 1. Pre-caricamento Dati
- Carica tutti i dati dalle offerte abilitate in memoria
- Carica i colori delle celle del Budget
- Identifica le righe con celle verdi/blu/rosse

#### 2. Controllo Consistenza Colonna M (Tipo Risorsa)
Per ogni riga con dati:
- Raccoglie i valori della colonna M da tutte le offerte abilitate
- **Ignora i valori vuoti** ("", null, undefined)
- Controlla che i valori **non vuoti** siano tutti uguali
- Se **diversi** → ERRORE:
  - Cancella contenuto cella M
  - Imposta nota con dettaglio errore
  - Sfondo ROSSO (#ea4335)
- Se **uguali** → OK:
  - Scrive il valore comune
  - Se cella verde/blu → mantiene colore
  - Se cella grigia/rossa → imposta grigio chiaro (#d9d9d9)

**Esempi Controllo M:**
- Off_01: "PM", Off_02: "PM" → ✅ Scrive "PM"
- Off_01: "PM", Off_02: "" → ✅ Scrive "PM" (ignora vuoto)
- Off_01: "", Off_02: "" → ✅ Scrive "" (tutti vuoti)
- Off_01: "PM", Off_02: "Senior" → ❌ ERRORE con nota

#### 3. Processamento Celle Verdi/Blu

Per ogni colonna (L, N, S, T):
- **Solo celle verdi o blu** vengono processate
- Colonna **L** e **N**: CONCATENAZIONE
  - Unisce i valori dalle offerte con " + "
  - Esempio: "Mat A" + "Mat B" = "Mat A + Mat B"
- Colonna **S** e **T**: SOMMA
  - Somma i valori numerici
  - Esempio: 100 + 200 = 300
- **Celle VERDI** → diventano **BLU** (#4285f4)
- **Celle BLU** → rimangono **BLU**

#### 4. Celle Gialle (Colonna U) - MAX
- Seleziona il valore **MASSIMO** tra le offerte
- Esclude cella **S465** (ha validazione dati)
- Rimangono **GIALLE**

#### 5. Celle Speciali S64 e S65
- **S64**: Valore anticipo
- **S65**: Percentuale anticipo
- Scrive il **VALORE** della somma (non la formula)
- Mantiene colore **VERDE**

#### 6. Cella S465 (Tipo Assistenza)
- Se **tutte le offerte** hanno valore → scrive il MAX
- Se **nessuna offerta** ha valore → lascia vuoto
- **NON** genera errori di validazione

#### 7. Aggiornamento Etichetta Sintesi
- Cella **Budget\!L62**
- Scrive: "Sintesi BOM (Offerta 01+02+...)"
- Indica quali offerte sono incluse nel Budget

#### 8. Allineamento BOM (controlla)
- Legge codice commessa da **Budget\!L56**
- Cerca nel file commesse esterno
- Aggiorna parametri in Budget e tutte le offerte:
  - **T2**: Tipo commessa (AUT, MES, SIM/AI)
  - **S2**: Percentuale generale
  - **S4, S5, S6**: Tariffe vendita (PM, Senior, Expert)
  - **T4, T5, T6**: Costi (PM, Senior, Expert)

#### 9. Dialog Finale
- Se ci sono errori nella colonna M:
  - Mostra dialog con elenco dettagliato errori
  - Indica riga e valori diversi
- Si posiziona sul foglio **Budget**

### Tempi di Esecuzione

Per un Budget con ~450 righe e 2 offerte:
- Pre-caricamento: ~2s
- Celle verdi: ~7s
- Celle gialle: ~1s
- Totale: **~10-15 secondi**

---

## Celle e Colori

### Sistema Colori

Il sistema usa i colori delle celle per identificare il tipo di cella:

| Colore | Hex | Significato | Comportamento |
|--------|-----|-------------|---------------|
| **Verde** | #00ff00 | Cella da compilare | Diventa blu dopo rigenerazione |
| **Blu** | #4285f4 | Cella processata | Rimane blu |
| **Giallo** | #ffff00 | Cella MAX | Rimane giallo |
| **Rosso** | #ea4335 | Errore consistenza | Deve essere corretto |
| **Grigio chiaro** | #d9d9d9 | Valore scritto (colonna M) | Cella con tipo risorsa comune |
| **Grigio** | #efefef | Cella neutra | Processata se M è consistente |

### Colonne Importanti

| Colonna | Nome | Tipo | Colore Iniziale | Comportamento |
|---------|------|------|-----------------|---------------|
| **L** | Descrizione | Concatenazione | Verde | Verde → Blu |
| **M** | Tipo Risorsa | Valore unico | Vari | Verde/Blu mantiene, altri → Grigio |
| **N** | Descrizione Dettagliata | Concatenazione | Verde | Verde → Blu |
| **S** | Quantità/Ore/Mesi | Somma | Verde | Verde → Blu |
| **T** | Costo Unitario | Somma | Verde | Verde → Blu |
| **U** | Valore MAX | Massimo | Giallo | Rimane Giallo |

---

## FAQ e Risoluzione Problemi

### Il menu Automate non appare

**Soluzione**:
1. Ricarica foglio (F5)
2. **Estensioni → Apps Script** → Seleziona `onOpen` → **Esegui**
3. Autorizza permessi
4. Ricarica foglio

### Celle Budget diventano tutte blu anche se non erano verdi

**Soluzione**: Verifica versione >= 1.2.9 nel log (Estensioni → Apps Script → Esecuzioni)

### "Tipo risorse diverse" ma sono uguali

**Soluzione**: Dalla v38, i valori vuoti vengono ignorati. "PM" + "" = OK

### Budget non si rigenera automaticamente

**Soluzione**: Modifica cella Budget\!L56 o usa Rigenera Budget manuale

### Come vedere i log dettagliati

1. **Estensioni → Apps Script**
2. **Esecuzioni** (icona orologio)
3. Clicca sull'esecuzione recente
4. Cerca `[rigeneraBudgetDaOfferte]`

---

## Versioning

### Storia Versioni Recenti

| Versione Libreria | Deploy | Descrizione |
|-------------------|--------|-------------|
| **1.3.0** | v38 | Ignora valori vuoti in controllo M |
| **1.2.9** | v37 | Colonna M sempre processata |
| **1.2.7** | v35 | Fix formattazione colonna M |
| **1.2.3** | v31 | Note invece di setValue per errori M |

---

**Fine Manuale Utente**
