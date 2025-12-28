# Quick Start - Sistema Multi-Offerta

## ⚡ Setup Immediato

### Stato Attuale
✅ Codice deployato
❌ Sistema NON ancora inizializzato nel foglio

---

## 🚀 Primi Passi (5 minuti)

### Passo 1: Ricarica Foglio
1. Apri: https://docs.google.com/spreadsheets/d/18sXHc0oiIcpumQrgACtTSHPdO17t3JrJLQXLqsrmg3I
2. Premi **F5** per ricaricare
3. Attendi caricamento menu

### Passo 2: Verifica Menu
Cerca in alto il menu **"Automate"**

Dovrebbe contenere:
```
Automate
└─ Opzioni Offerta
   ├─ Configura offerte...
   ├─ Inizializza sistema multi-offerta  ← QUESTO!
   ├─ Rigenera Budget
   └─ Ripristina Master da Budget
```

**Se il menu non appare:**
1. Vai su **Estensioni → Apps Script**
2. Cerca la funzione `onOpen` nella lista
3. Click **Esegui** (icona ▶)
4. Autorizza se richiesto
5. Torna al foglio e ricarica (F5)

### Passo 3: Inizializza Sistema

**IMPORTANTE**: Fai questo SU UNA COPIA, non sul file originale!

1. **File → Crea una copia**
2. Apri la copia
3. Menu: **Automate → Opzioni Offerta → Inizializza sistema multi-offerta**
4. Attendi ~10 secondi
5. Dovresti vedere un alert di conferma

**Cosa succede:**
- ✓ Crea foglio "Master" (copia di Budget, nascosto)
- ✓ Crea foglio "Configurazione" (nascosto)
- ✓ Crea foglio "Off_01" (prima offerta)
- ✓ Rigenera Budget

### Passo 4: Verifica Creazione

**Controlla che esistano questi fogli:**
1. Guarda i tab in basso
2. Dovresti vedere:
   - Budget (originale)
   - **Off_01** (nuovo!)
   - Formule (se c'era già)
   - Label (se c'era già)

**Fogli nascosti** (Master e Configurazione):
- Non li vedi nei tab perché sono nascosti
- Per verificare: **Visualizza → Fogli nascosti**

### Passo 5: Apri Configurazione

1. Menu: **Automate → Opzioni Offerta → Configura offerte...**
2. Si apre sidebar a destra
3. Dovresti vedere Off_01 con checkbox selezionata

**Se funziona:** ✅ Sistema pronto!

---

## 🎨 Test Base

### Test 1: Aggiungi Offerta

1. Sidebar aperta
2. Click **+ Aggiungi Nuova Offerta**
3. Nome: "Test Premium"
4. Descrizione: "Variante di test"
5. Attendi creazione
6. Verifica apparizione **Off_02** nei tab

### Test 2: Modifica Valori

1. Apri foglio **Off_02**
2. Cerca celle VERDI nelle righe 69-495
3. Modifica alcuni valori
4. Torna alla sidebar
5. Click **Rigenera Budget**
6. Attendi ~5 secondi
7. Apri foglio **Budget**
8. Verifica che cella **M62** dica: "Sintesi BOM (Offerta 01+02)"

### Test 3: Disabilita Offerta

1. Sidebar aperta
2. Togli spunta da Off_02
3. Click **Rigenera Budget**
4. Verifica M62 → "Sintesi BOM (Offerta 01)"
5. I valori di Budget dovrebbero cambiare

---

## ❌ Troubleshooting

### Problema: Menu non appare

**Soluzione**:
```
1. Estensioni → Apps Script
2. Seleziona funzione "onOpen"
3. Click ▶ Esegui
4. Autorizza permessi
5. Ricarica foglio (F5)
```

### Problema: "Sistema già inizializzato"

**Causa**: Hai già eseguito inizializzazione

**Verifica**:
- Visualizza → Fogli nascosti
- Cerca "Configurazione"
- Se c'è: sistema già inizializzato ✓

### Problema: "Foglio Budget non trovato"

**Causa**: Foglio si chiama diversamente

**Soluzione**:
- Rinomina foglio in "Budget" (esatto)
- Riprova inizializzazione

### Problema: Errore durante inizializzazione

**Soluzione**:
1. Apri **Estensioni → Apps Script**
2. Vai su **Visualizza → Log di esecuzione**
3. Cerca errori in rosso
4. Copia messaggio errore
5. Potrebbe essere problema autorizzazioni

---

## 📞 Se Qualcosa Non Funziona

### Verifica 1: Script Deployato

```bash
# Sul tuo computer, nella cartella del progetto:
clasp status
```

Dovresti vedere:
```
Tracked files:
└─ OffertaManager.js
└─ OffertaConfigUI.html
└─ ...
```

### Verifica 2: File Corretti su Apps Script

1. Estensioni → Apps Script
2. Nella lista file a sinistra dovresti vedere:
   - OffertaManager.js
   - OffertaConfigUI.html
   - Config.js (con SHEETS e OFFERTE)
   - On Open.js

### Verifica 3: Errori Console

1. Nel foglio Google Sheets
2. F12 (apri console browser)
3. Cerca errori in rosso
4. Se ci sono, copia il messaggio

---

## ✅ Checklist Primo Avvio

Prima di inizializzare sul file VERO:

- [ ] Hai fatto copia del file originale
- [ ] Menu "Automate" appare
- [ ] Sottomenu "Opzioni Offerta" presente
- [ ] Hai testato su copia e funziona
- [ ] Hai letto GUIDA_SISTEMA_MULTI_OFFERTA.md
- [ ] Hai capito workflow base

---

## 🎯 Workflow Rapido

**Setup una volta sola:**
```
1. Copia file
2. Inizializza sistema multi-offerta
3. Verifica creazione fogli
```

**Uso quotidiano:**
```
1. Lavora su Off_XX
2. Rigenera Budget quando serve
3. Verifica M62 per conferma
```

**Gestione varianti:**
```
1. Aggiungi offerta (sidebar)
2. Popola valori
3. Abilita/disabilita combinazioni
4. Rigenera Budget per ogni combo
5. Esporta PDF/Excel
```

---

**Durata totale setup**: ~5 minuti
**Pronto?** Vai su Passo 1! 🚀
