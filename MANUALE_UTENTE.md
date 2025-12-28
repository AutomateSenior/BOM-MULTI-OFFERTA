# Manuale Utente - Sistema Multi-Offerta BOM v08

**Versione**: 08
**Data**: 26 Novembre 2025
**Autore**: Sistema BOM Master

---

## 📋 Indice

1. [Introduzione](#introduzione)
2. [Concetti Base](#concetti-base)
3. [Primo Avvio](#primo-avvio)
4. [Uso Quotidiano](#uso-quotidiano)
5. [Funzioni Avanzate](#funzioni-avanzate)
6. [Risoluzione Problemi](#risoluzione-problemi)
7. [Domande Frequenti](#domande-frequenti)

---

## 🎯 Introduzione

### Cos'è il Sistema Multi-Offerta?

Il sistema Multi-Offerta ti permette di **gestire più varianti** di un preventivo BOM nello stesso file Google Sheets, senza dover creare copie separate.

### Perché usarlo?

**Scenario tradizionale (PRIMA):**
- Cliente vuole vedere 3 configurazioni diverse
- Devi creare 3 file separati
- Difficile mantenere aggiornamenti sincronizzati
- Rischio di errori tra versioni

**Con il Sistema Multi-Offerta (ADESSO):**
- **Un solo file** con 3 varianti (Off_01, Off_02, Off_03)
- **Un clic** per generare il Budget con qualsiasi combinazione
- Modifiche al template applicate a tutte le varianti
- Massima flessibilità e zero errori

---

## 📚 Concetti Base

### Struttura del File

Dopo l'inizializzazione, il tuo file avrà questa struttura:

```
📊 Il Tuo File BOM
│
├── 📄 Budget           ← Sintesi automatica (quello che mostri al cliente)
├── 📄 Off_01          ← Prima variante (es: "Configurazione Base")
├── 📄 Off_02          ← Seconda variante (es: "Configurazione Premium")
├── 📄 Off_03          ← Terza variante (es: "Configurazione Custom")
│
├── 🔒 Master          ← Template nascosto (NON modificare)
├── 🔒 Configurazione  ← Dati sistema (nascosto)
├── 🔒 Formule         ← Formule di riferimento (nascosto)
└── 🔒 Label           ← Etichette di riferimento (nascosto)
```

### Cosa Fa Ogni Foglio?

#### 1. **Budget** (Sintesi)
- Contiene la **somma** o il **massimo** delle offerte abilitate
- Si aggiorna **manualmente** quando clicchi "Rigenera Budget"
- È quello che **esporti e mostri al cliente**
- **Cella M62** mostra quali offerte sono incluse

#### 2. **Off_01, Off_02, Off_03...** (Varianti)
- Fogli dove **lavori** e inserisci dati
- Ogni foglio è una **variante indipendente**
- Puoi averne **quante ne vuoi** (Off_01...Off_99)
- Nome personalizzabile (es: "Base", "Premium", "Enterprise")

#### 3. **Master** (Template)
- Foglio **nascosto e protetto**
- È il modello da cui partono tutte le nuove offerte
- **Non modificarlo** direttamente

#### 4. **Configurazione** (Sistema)
- Foglio **nascosto**
- Contiene i metadati delle offerte (nome, descrizione, abilitata/disabilitata)
- Gestito automaticamente dal sistema

---

## 🚀 Primo Avvio

### Prerequisiti

- File Google Sheets con foglio **"Budget"** già compilato
- Accesso al menu "Automate" (se non appare, ricarica la pagina con F5)

### Passo 1: Inizializzazione (Solo la Prima Volta)

1. **Fai un BACKUP del file** (File → Crea una copia)
2. Apri il file originale
3. Menu: **Automate → Opzioni Offerta → Inizializza sistema multi-offerta**
4. Attendi 10-15 secondi
5. Vedrai una conferma: "Sistema Multi-Offerta Inizializzato! ✓"

**Cosa è successo?**
- ✅ Creato foglio Master (nascosto)
- ✅ Creato foglio Configurazione (nascosto)
- ✅ Creata prima offerta "Off_01"
- ✅ Budget aggiornato con i dati di Off_01

### Passo 2: Verifica

Controlla che in basso, nei tab dei fogli, ci sia:
- **Budget** (originale)
- **Off_01** (nuovo!)

Nella cella **M62** del Budget dovresti vedere:
```
Sintesi BOM (Offerta 01)
```

✅ **Sistema pronto!** Ora puoi iniziare a usarlo.

---

## 💼 Uso Quotidiano

### Gestione Rapida Offerte (Metodo Consigliato)

Usa la **dialog rapida** per gestire le offerte in modo veloce:

1. Menu: **Automate → Opzioni Offerta → ⚡ Gestione rapida offerte**
2. Si apre una finestra elegante con:
   - Lista di tutte le offerte
   - Checkbox per abilitare/disabilitare
   - Preview della sintesi Budget
   - Pulsante "Rigenera Budget"

**Vantaggi:**
- ⚡ Velocissima (più della sidebar)
- 👁️ Preview in tempo reale
- 🎯 Un solo clic per rigenerare
- ✨ Chiusura automatica al termine

### Come Lavorare con le Offerte

#### 1. **Creare una Nuova Offerta**

**Metodo A - Dialog Rapida:**
1. ⚡ Gestione rapida offerte
2. (In arrivo: pulsante "Aggiungi Offerta" nella dialog)

**Metodo B - Sidebar:**
1. Menu: **Automate → Opzioni Offerta → Configura offerte...**
2. Sidebar a destra si apre
3. Click: **+ Aggiungi Nuova Offerta**
4. Inserisci nome (es: "Premium")
5. Inserisci descrizione (es: "Configurazione con moduli extra")
6. Nuovo foglio **Off_02** creato automaticamente

**Tab colorati:**
- 🟢 Verde squillante = Offerta **abilitata**
- ⚫ Grigio = Offerta **disabilitata**

#### 2. **Modificare i Dati di un'Offerta**

1. Apri il foglio **Off_01** (o Off_02, Off_03...)
2. Modifica i valori come faresti normalmente
3. **IMPORTANTE:** Modifica solo le celle nelle colonne **S, T, U** (righe 69-495)
4. Quando hai finito: Menu → **⚡ Gestione rapida offerte** → Rigenera Budget

#### 3. **Abilitare/Disabilitare Offerte**

Puoi scegliere **quali offerte** includere nel Budget:

**Esempio:**
- Off_01 = Configurazione Base ✅ (abilitata)
- Off_02 = Configurazione Premium ❌ (disabilitata)
- Off_03 = Modulo Opzionale ✅ (abilitata)

Budget mostrerà: `Sintesi BOM (Offerta 01+03)`

**Come fare:**
1. ⚡ Gestione rapida offerte
2. Spunta/deseleziona le checkbox
3. Click "Rigenera Budget"
4. Fatto! Budget aggiornato in 5-10 secondi

#### 4. **Rigenerare il Budget**

Il Budget **non si aggiorna automaticamente**. Devi rigenerarlo manualmente quando vuoi.

**Quando rigenerare:**
- ✅ Dopo aver modificato dati in Off_XX
- ✅ Dopo aver abilitato/disabilitato offerte
- ✅ Dopo aver aggiunto/rimosso offerte
- ✅ Prima di esportare/condividere il file

**Come rigenerare:**

**Metodo 1 - Dialog Rapida (⚡ più veloce):**
1. Menu: **Automate → Opzioni Offerta → ⚡ Gestione rapida offerte**
2. (Eventualmente cambia selezione offerte)
3. Click: **Rigenera Budget**
4. Attendi 5-15 secondi
5. Dialog si chiude automaticamente
6. Budget aggiornato! ✅

**Metodo 2 - Menu Diretto:**
1. Menu: **Automate → Opzioni Offerta → Rigenera Budget**
2. Attendi completamento

### Come Funziona la Rigenerazione?

Il sistema aggiorna il Budget in base alle offerte abilitate:

#### Celle Verdi (Colonne S, T) → **SOMMA**
Se hai 3 offerte abilitate con questi valori in S100:
- Off_01: 1000€
- Off_02: 1500€
- Off_03: 500€

Budget in S100 mostrerà:
```
Formula: =Off_01!S100+Off_02!S100+Off_03!S100
Risultato: 3000€
```

#### Celle Gialle (Colonna U) → **MASSIMO**
Se hai 3 offerte abilitate con questi valori in U100:
- Off_01: 30 giorni
- Off_02: 45 giorni
- Off_03: 20 giorni

Budget in U100 mostrerà:
```
Formula: =MAX(Off_01!U100,Off_02!U100,Off_03!U100)
Risultato: 45 giorni
```

**Cosa cambia visivamente:**
- Celle verdi/gialle → **BLU** con testo **BIANCO**
- Cella M62 → Aggiornata con lista offerte

---

## 🎨 Funzioni Avanzate

### Configurare Offerte Dettagliata (Sidebar)

Per modifiche avanzate usa la sidebar completa:

1. Menu: **Automate → Opzioni Offerta → Configura offerte...**
2. Sidebar si apre sulla destra

**Funzioni disponibili:**
- ✏️ Modificare descrizione offerta
- ➕ Aggiungere nuova offerta
- ❌ Rimuovere offerta
- ☑️ Abilitare/disabilitare
- 🔄 Rigenerare Budget

**Nota:** La sidebar è più lenta della dialog rapida, usala solo se devi modificare descrizioni.

### Rimuovere un'Offerta

1. Sidebar: **Configura offerte...**
2. Click: **Rimuovi Offerta** (icona cestino)
3. Conferma l'eliminazione
4. Foglio Off_XX viene eliminato definitivamente

⚠️ **ATTENZIONE:**
- Non puoi rimuovere l'ultima offerta (deve esisterne almeno una)
- Operazione **irreversibile** (fai backup prima!)

### Ripristinare Master da Budget

**Scenario:** Hai modificato manualmente il Budget e vuoi usarlo come nuovo template per le future offerte.

**Come fare:**
1. Menu: **Automate → Opzioni Offerta → Ripristina Master da Budget**
2. Leggi l'avviso
3. Conferma
4. Master viene sovrascritto con il Budget corrente

⚠️ **ATTENZIONE:**
- Le offerte esistenti (Off_XX) **non cambiano**
- Solo le **nuove** offerte create dopo useranno il nuovo Master
- Operazione irreversibile

### Celle Speciali di Consistenza

Alcune celle richiedono **valori identici** in tutte le offerte (non somma o max).

**Esempio:** Cella **S452** (Tipo di assistenza)
- Se Off_01 = "Standard" e Off_02 = "Premium"
- Budget mostrerà: ⚠️ **"Tipo di assistenza incoerente tra le diverse offerte"**

**Soluzione:** Assicurati che tutte le offerte abilitate abbiano lo stesso valore in S452.

---

## 🛠️ Risoluzione Problemi

### Menu "Opzioni Offerta" non appare

**Causa:** Script non caricato
**Soluzione:**
1. Ricarica pagina (F5)
2. Se persiste: **Estensioni → Apps Script**
3. Cerca funzione `onOpen` nella lista
4. Click ▶ **Esegui**
5. Autorizza permessi se richiesto
6. Torna al foglio e ricarica (F5)

### "Sistema non inizializzato"

**Causa:** Non hai eseguito l'inizializzazione
**Soluzione:**
1. Menu: **Automate → Opzioni Offerta → Inizializza sistema multi-offerta**
2. Attendi completamento
3. Riprova l'operazione

### Budget non si aggiorna

**Causa:** Devi rigenerare manualmente
**Soluzione:**
1. ⚡ Gestione rapida offerte
2. Click: **Rigenera Budget**
3. Attendi 5-15 secondi
4. Verifica cella M62 per conferma

### Cella M62 non mostra nulla

**Causa:** Problema nella rigenerazione
**Soluzione:**
1. Apri **Estensioni → Apps Script**
2. Vai su **Visualizza → Log di esecuzione**
3. Cerca messaggi con "aggiornaEtichettaSintesi"
4. Se vedi errori, annota il messaggio e contatta supporto

### "Non puoi rimuovere l'ultima offerta"

**Causa:** Il sistema richiede almeno 1 offerta
**Soluzione:**
- Crea una nuova offerta prima di rimuovere l'ultima

### Offerta Off_XX ha tab grigio ma dovrebbe essere abilitata

**Causa:** Configurazione non sincronizzata
**Soluzione:**
1. ⚡ Gestione rapida offerte
2. Disabilita e riabilita l'offerta
3. Click: **Rigenera Budget**
4. Tab dovrebbe tornare verde

### Formule non si aggiornano dopo cambio configurazione

**Causa:** Bug risolto nella versione corrente
**Soluzione:**
- Verifica di avere l'ultima versione deployata
- Rigenera Budget due volte consecutive

### Dialog/Sidebar non si carica

**Causa:** Errore JavaScript o timeout
**Soluzione:**
1. Apri console browser (F12)
2. Cerca errori in rosso nella tab "Console"
3. Chiudi dialog/sidebar
4. Ricarica pagina (F5)
5. Riprova

### Celle verdi/gialle non trovate

**Causa:** Colori non riconosciuti dal sistema
**Soluzione:**
1. Verifica che le celle siano effettivamente:
   - **Verdi** in colonne S, T
   - **Gialle** in colonna U
2. Range valido: righe 69-495
3. Se usi colori personalizzati, potrebbero non essere riconosciuti

---

## ❓ Domande Frequenti

### Quante offerte posso creare?

Teoricamente illimitate (Off_01...Off_99). Praticamente, consigliamo **massimo 10 offerte** per performance.

### Posso rinominare i fogli Off_XX?

**NO.** Il sistema riconosce le offerte dal nome (Off_01, Off_02...). Se rinomini, l'offerta non funzionerà più.

Puoi però usare il campo **"Descrizione"** nella sidebar per identificare meglio le offerte.

### Posso modificare il Budget manualmente?

**Sì**, ma le modifiche verranno **sovrascritte** alla prossima rigenerazione.

Se vuoi salvare modifiche manuali come nuovo template: **Ripristina Master da Budget**.

### Posso lavorare su più offerte contemporaneamente?

**Sì!** Puoi aprire Off_01, Off_02, Off_03 in tab diverse e modificarli in parallelo. Ricorda solo di rigenerare il Budget quando hai finito.

### Il Budget si aggiorna automaticamente?

**NO.** Devi rigenerare manualmente per evitare rallentamenti. Clicca "Rigenera Budget" quando hai finito le modifiche.

### Posso annullare una rigenerazione?

**NO.** La rigenerazione sovrascrive il Budget. Se hai dubbi, fai una copia del file prima di rigenerare.

### Cosa succede se elimino accidentalmente Off_01?

Il foglio viene eliminato definitivamente. Se avevi dati importanti, devi recuperarli da un backup.

**Prevenzione:** Fai backup regolari (File → Crea una copia).

### Posso usare il sistema su file già esistenti?

**Sì!** Basta che il file abbia un foglio chiamato **"Budget"**. L'inizializzazione creerà tutto il necessario senza danneggiare i dati esistenti.

### Posso condividere il file con colleghi?

**Sì!** Il sistema funziona per tutti gli utenti con accesso al file. Ogni utente può:
- Modificare offerte
- Rigenerare Budget
- Aggiungere/rimuovere offerte

⚠️ Attenzione: Non lavorate sulla stessa offerta contemporaneamente!

### Posso esportare il Budget in PDF/Excel?

**Sì!** Il Budget è un normale foglio Google Sheets:
- **PDF:** File → Scarica → Documento PDF
- **Excel:** File → Scarica → Microsoft Excel (.xlsx)

### Le formule nel Budget sono modificabili?

Tecnicamente **sì**, ma verranno **sovrascritte** alla prossima rigenerazione. Se vuoi personalizzare formule, modificale **dopo** l'ultima rigenerazione finale.

### Cosa succede se aggiungo righe/colonne?

Il sistema lavora su un **range fisso** (righe 69-495, colonne S-T-U). Se aggiungi righe/colonne:
- Dentro il range: vengono processate normalmente
- Fuori dal range: ignorate dalla rigenerazione

---

## 📞 Supporto

### File di Log

Per diagnosticare problemi:
1. **Estensioni → Apps Script**
2. **Visualizza → Log di esecuzione**
3. Cerca messaggi di errore o warning

### Informazioni da Fornire

Se hai problemi, fornisci:
- Nome del file
- Operazione che hai tentato
- Messaggio di errore esatto
- Screenshot (se possibile)
- Log di esecuzione (se disponibile)

---

## 🎯 Workflow Consigliato

### Scenario: Preventivo con 3 Varianti

**Obiettivo:** Cliente vuole vedere configurazione Base, Premium e Custom.

#### Fase 1: Setup (Una Tantum)
1. ✅ Fai backup del file
2. ✅ Inizializza sistema multi-offerta
3. ✅ Verifica creazione Off_01

#### Fase 2: Creazione Varianti
1. ✅ Aggiungi Off_02, nome "Premium"
2. ✅ Aggiungi Off_03, nome "Custom"
3. ✅ Popola Off_01 con dati Base
4. ✅ Popola Off_02 con dati Premium
5. ✅ Popola Off_03 con dati Custom

#### Fase 3: Generazione Preventivi
**Solo Base:**
1. ⚡ Gestione rapida → Abilita solo Off_01
2. Rigenera Budget
3. M62 mostra: "Sintesi BOM (Offerta 01)"
4. Esporta PDF → "Preventivo_Base.pdf"

**Solo Premium:**
1. ⚡ Gestione rapida → Abilita solo Off_02
2. Rigenera Budget
3. M62 mostra: "Sintesi BOM (Offerta 02)"
4. Esporta PDF → "Preventivo_Premium.pdf"

**Base + Custom:**
1. ⚡ Gestione rapida → Abilita Off_01 e Off_03
2. Rigenera Budget
3. M62 mostra: "Sintesi BOM (Offerta 01+03)"
4. Esporta PDF → "Preventivo_Base_Custom.pdf"

**Tutte insieme:**
1. ⚡ Gestione rapida → Abilita tutte
2. Rigenera Budget
3. M62 mostra: "Sintesi BOM (Offerta 01+02+03)"
4. Esporta PDF → "Preventivo_Completo.pdf"

#### Fase 4: Presentazione Cliente
- Invia i 4 PDF al cliente
- Cliente sceglie la configurazione preferita
- Tu sai esattamente cosa hai quotato per ogni variante

---

## ✅ Checklist Rapida

### Prima di Iniziare
- [ ] Ho fatto backup del file
- [ ] Il foglio "Budget" esiste ed è compilato
- [ ] Ho letto il manuale (almeno "Primo Avvio" e "Uso Quotidiano")

### Dopo Inizializzazione
- [ ] Foglio Off_01 presente
- [ ] Cella M62 mostra "Sintesi BOM (Offerta 01)"
- [ ] Menu "Opzioni Offerta" visibile
- [ ] ⚡ Gestione rapida offerte funziona

### Prima di Esportare Budget
- [ ] Ho rigenerato il Budget con le offerte desiderate
- [ ] M62 mostra le offerte corrette
- [ ] I valori nel Budget sono corretti
- [ ] I tab Off_XX hanno i colori giusti (verde=abilitato, grigio=disabilitato)

---

## 📖 Glossario

**Budget:** Foglio sintesi che mostra la somma/massimo delle offerte abilitate. Quello che esporti e mostri al cliente.

**Offerta (Off_XX):** Singola variante del preventivo. Ogni offerta è un foglio indipendente.

**Master:** Template nascosto da cui vengono create tutte le nuove offerte.

**Rigenerazione:** Processo che aggiorna il Budget sommando/maximizzando i valori delle offerte abilitate.

**Abilitata/Disabilitata:** Stato di un'offerta. Solo le offerte abilitate vengono incluse nel Budget.

**Celle Verdi (S, T):** Celle che vengono sommate (SUM) durante la rigenerazione.

**Celle Gialle (U):** Celle che mostrano il valore massimo (MAX) durante la rigenerazione.

**Celle Blu:** Celle già processate dal sistema (erano verdi o gialle, ora blu con testo bianco).

**M62:** Cella etichetta nel Budget che mostra quali offerte sono incluse.

**Dialog Rapida:** Finestra modale veloce per gestire offerte e rigenerare Budget.

**Sidebar:** Pannello laterale più completo per configurazione avanzata offerte.

---

**Fine Manuale**

Versione 08 - 26 Novembre 2025
Sistema BOM Master con Multi-Offerta

Per assistenza, consulta la sezione "Risoluzione Problemi" o contatta il supporto tecnico.
