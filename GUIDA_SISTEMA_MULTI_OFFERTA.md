# Guida Sistema Multi-Offerta - v08

**Data**: 25 Novembre 2025
**Versione**: 08 - Sistema Multi-Offerta

---

## 🎯 Panoramica

Il sistema Multi-Offerta permette di gestire **varianti multiple** di un'offerta BOM nello stesso file.

### Vantaggi

✅ **Varianti Parallele**: Crea più configurazioni (base, premium, custom)
✅ **Budget Dinamico**: Somma automatica delle offerte abilitate
✅ **Gestione Flessibile**: Aggiungi, rimuovi, abilita/disabilita offerte
✅ **Template Condiviso**: Tutte le offerte partono dallo stesso Master

---

## 📐 Struttura Fogli

```
Spreadsheet BOM v08
├─ Master (nascosto)        → Template protetto
├─ Budget                   → Sintesi automatica offerte abilitate
├─ Off_01                   → Offerta 1
├─ Off_02                   → Offerta 2
├─ Off_03                   → Offerta 3
├─ ...
├─ Off_NN                   → Offerta N
├─ Configurazione (nascosto) → Dati configurazione offerte
├─ Formule (nascosto)       → Formule di riferimento
└─ Label (nascosto)         → Label di riferimento
```

### Foglio Master

- **Scopo**: Template condiviso per tutte le offerte
- **Visibilità**: Nascosto e protetto
- **Origine**: Copia del Budget originale
- **Uso**: Quando crei nuova offerta, viene duplicato il Master

### Foglio Budget

- **Scopo**: Sintesi delle offerte abilitate
- **Contenuto**: Somma automatica delle celle verdi (righe 69-495)
- **Aggiornamento**: Manuale tramite "Rigenera Budget"
- **Etichetta**: Cella M62 mostra quali offerte sono incluse

### Fogli Off_XX

- **Scopo**: Varianti singole dell'offerta
- **Modificabili**: Sì, dall'utente
- **Struttura**: Identica al Master
- **Naming**: Off_01, Off_02, ..., Off_NN (numerazione automatica)

### Foglio Configurazione

- **Scopo**: Metadati offerte
- **Visibilità**: Nascosto
- **Struttura**:
  ```
  | ID     | Nome         | Descrizione    | Abilitata | Ordine |
  |--------|--------------|----------------|-----------|--------|
  | Off_01 | Offerta Base | Config std     | TRUE      | 1      |
  | Off_02 | Variante A   | Con opzione X  | TRUE      | 2      |
  | Off_03 | Variante B   | Con opzione Y  | FALSE     | 3      |
  ```

---

## 🚀 Setup Iniziale

### 1. Inizializzazione (Prima Volta)

**Prerequisiti**:
- File esistente con foglio "Budget" già popolato

**Procedura**:
1. Apri il file BOM
2. Menu: **Automate → Opzioni Offerta → Inizializza sistema multi-offerta**
3. Conferma

**Cosa succede**:
- ✓ Crea foglio "Master" (copia di Budget, nascosto)
- ✓ Crea foglio "Configurazione" (nascosto)
- ✓ Crea offerta "Off_01" (copia di Master)
- ✓ Rigenera Budget (copia da Off_01)

**Tempo**: ~10 secondi

---

## 📋 Gestione Offerte

### 2. Aprire Configurazione

Menu: **Automate → Opzioni Offerta → Configura offerte...**

Si apre una **sidebar** sulla destra con:
- Lista offerte esistenti
- Checkbox per abilitare/disabilitare
- Campo descrizione modificabile
- Pulsanti azione

### 3. Aggiungere Nuova Offerta

**Metodo A - Da Sidebar**:
1. Sidebar aperta
2. Click **+ Aggiungi Nuova Offerta**
3. Inserisci nome (es: "Variante Premium")
4. Inserisci descrizione (opzionale)
5. Offerta creata e abilitata

**Metodo B - Da Menu**:
1. Menu: **Automate → Opzioni Offerta → Aggiungi nuova offerta**
2. Prompt nome
3. Prompt descrizione

**Risultato**:
- Nuovo foglio Off_XX creato (duplica Master)
- Aggiunta a Configurazione
- Default: abilitata = TRUE

### 4. Modificare Offerta

**Nome/Descrizione**:
1. Sidebar aperta
2. Modifica campo descrizione
3. Auto-salva al cambio focus

**Abilitare/Disabilitare**:
1. Sidebar aperta
2. Togli/metti spunta su checkbox
3. Offerta esclusa/inclusa dal Budget

### 5. Rimuovere Offerta

**Procedura**:
1. Sidebar aperta
2. Click **Rimuovi Offerta** sull'offerta da eliminare
3. Conferma
4. Foglio Off_XX eliminato

**Attenzione**:
- Non puoi rimuovere l'ultima offerta
- Operazione irreversibile

---

## 🔄 Rigenerazione Budget

### Come Funziona

Il Budget è la **somma** delle offerte abilitate:
- Identifica celle verdi nel range righe 69-495
- Per ogni cella verde: somma valori da Off_XX abilitati
- Aggiorna etichetta M62

### Quando Rigenerare

**SEMPRE dopo**:
- ✓ Modificato valori in Off_XX
- ✓ Abilitato/disabilitato offerta
- ✓ Aggiunto/rimosso offerta

**NON automatico** per evitare rallentamenti!

### Procedura Manuale

**Metodo A - Da Sidebar**:
1. Sidebar aperta
2. Click **Rigenera Budget**
3. Conferma
4. Attendi completamento (5-15 secondi)

**Metodo B - Da Menu**:
1. Menu: **Automate → Opzioni Offerta → Rigenera Budget**
2. Conferma
3. Attendi completamento

### Etichetta Sintesi (M62)

Esempio output:
```
Sintesi BOM (Offerta 01+03+05)
```

Mostra quali offerte sono **attualmente** incluse nel Budget.

---

## 🎨 Workflow Tipico

### Scenario 1: Offerta con 2 Varianti

**Obiettivo**: Presentare cliente con opzione Base e Premium

**Step**:
1. **Inizializza** sistema (crea Off_01 = Base)
2. **Aggiungi** Off_02, nome "Premium"
3. Modifica Off_02: aggiungi righe specifiche Premium
4. **Configura offerte**: disabilita Off_02
5. **Rigenera Budget** → mostra solo Base
6. Salva/Esporta PDF Budget (solo Base)
7. **Configura offerte**: abilita Off_02, disabilita Off_01
8. **Rigenera Budget** → mostra solo Premium
9. Salva/Esporta PDF Budget (solo Premium)
10. **Configura offerte**: abilita entrambe
11. **Rigenera Budget** → mostra Base+Premium
12. Salva/Esporta PDF Budget (completo)

### Scenario 2: Offerta Modulare

**Obiettivo**: Cliente sceglie moduli (A+B+C)

**Step**:
1. **Inizializza** sistema
2. **Aggiungi** Off_02 (Modulo A), Off_03 (Modulo B), Off_04 (Modulo C)
3. Popola ogni modulo con costi specifici
4. **Configura**: abilita solo Off_01 (base)
5. Cliente vuole Base+A+C: abilita Off_01+Off_02+Off_04
6. **Rigenera Budget**
7. Etichetta M62: "Sintesi BOM (Offerta 01+02+04)"

### Scenario 3: Comparazione Fornitori

**Obiettivo**: Stessi servizi, 3 fornitori diversi

**Step**:
1. Off_01 = Fornitore A
2. Off_02 = Fornitore B
3. Off_03 = Fornitore C
4. Popola con prezzi specifici per fornitore
5. Abilita solo Fornitore A → Rigenera → Esporta
6. Abilita solo Fornitore B → Rigenera → Esporta
7. Abilita solo Fornitore C → Rigenera → Esporta
8. Compara i 3 Budget

---

## ⚙️ Funzioni Avanzate

### Ripristina Master da Budget

**Scopo**: Aggiornare il template Master con modifiche dal Budget

**Quando usare**:
- Hai modificato manualmente il Budget
- Vuoi che le nuove offerte partano da questa versione

**Procedura**:
1. Menu: **Automate → Opzioni Offerta → Ripristina Master da Budget**
2. Conferma (ATTENZIONE: sovrascrive Master)
3. Master aggiornato

**ATTENZIONE**: Le offerte esistenti (Off_XX) **non vengono modificate**!

### Celle Verdi da Sommare

**Definizione**: Celle con sfondo verde nel range 69-495

**Colori riconosciuti**:
- `#00ff00` - Verde puro
- `#ccffcc` - Verde chiaro
- `#d9ead3` - Verde pastello (Google Sheets standard)
- `#93c47d` - Verde medio
- `#6aa84f` - Verde scuro

**Altre celle**: Mantengono valori/formule originali del Master

---

## 🐛 Troubleshooting

### Problema: Menu "Opzioni Offerta" non appare

**Causa**: Script non caricato

**Soluzione**:
1. Ricarica foglio (F5)
2. Vai su **Estensioni → Apps Script**
3. Esegui `onOpen()`
4. Ricarica foglio

### Problema: "Sistema non inizializzato"

**Causa**: Non hai eseguito inizializzazione

**Soluzione**:
1. Menu: **Automate → Opzioni Offerta → Inizializza sistema multi-offerta**
2. Attendi completamento
3. Riprova operazione

### Problema: Budget non si aggiorna

**Causa**: Devi rigenerare manualmente

**Soluzione**:
1. Menu: **Automate → Opzioni Offerta → Rigenera Budget**
2. Attendi (5-15 secondi)
3. Verifica cella M62 per conferma

### Problema: "Non puoi rimuovere l'ultima offerta"

**Causa**: Deve esistere almeno 1 offerta

**Soluzione**:
- Crea nuova offerta prima di rimuovere l'ultima

### Problema: Sidebar non si carica

**Causa**: Errore JavaScript o timeout

**Soluzione**:
1. Apri console browser (F12)
2. Cerca errori in rosso
3. Chiudi sidebar
4. Riprova aprire

### Problema: Offerta Off_XX non esiste ma è in configurazione

**Causa**: Eliminazione manuale del foglio

**Soluzione**:
1. Sidebar: Rimuovi offerta dalla configurazione
2. O crea manualmente foglio con nome corretto

---

## 📊 Performance

### Tempo Rigenerazione Budget

**Dipende da**:
- Numero offerte abilitate
- Numero celle verdi
- Dimensione foglio

**Stima**:
- 1 offerta, 100 celle verdi: ~2 secondi
- 3 offerte, 300 celle verdi: ~6 secondi
- 10 offerte, 1000 celle verdi: ~20 secondi

**Ottimizzazioni**:
- Lettura batch (getValues invece di getValue)
- Scrittura batch (setValues invece di setValue)
- Caching mappa celle verdi

---

## 📝 Limitazioni

1. **Non auto-sync**: Modifiche Off_XX non rigenerano automaticamente Budget
2. **Solo celle verdi**: Solo righe 69-495 con sfondo verde vengono sommate
3. **No formule nelle celle verdi**: Somma solo VALORI, non formule
4. **Naming fisso**: Off_01, Off_02, ... (non personalizzabile)
5. **No ordinamento**: Offerte sempre in ordine numerico

---

## 🔐 Sicurezza e Protezione

- **Master**: Nascosto e protetto (warning only)
- **Budget**: NON protetto (può essere modificato ma verrà sovrascritto alla rigenerazione)
- **Off_XX**: NON protetti (modificabili dagli utenti)
- **Configurazione**: Nascosto (modificabile solo via API)

**Raccomandazione**: Se vuoi proteggere Budget, fallo manualmente dopo rigenerazione.

---

## ✅ Checklist Post-Inizializzazione

Dopo aver inizializzato il sistema:

- [ ] Foglio Master nascosto presente
- [ ] Foglio Configurazione nascosto presente
- [ ] Foglio Off_01 presente
- [ ] Budget aggiornato
- [ ] Cella M62 mostra "Sintesi BOM (Offerta 01)"
- [ ] Menu "Opzioni Offerta" visibile
- [ ] Sidebar "Configura offerte" funzionante

---

## 📚 Comandi Menu Completo

```
Automate
├─ Installa Attivatori
├─ Controlla BOM
├─ Azzera BOM
├─ Opzioni Offerta
│  ├─ Configura offerte...              → Apri sidebar
│  ├─ Inizializza sistema multi-offerta → Setup iniziale
│  ├─ Rigenera Budget                   → Aggiorna Budget
│  └─ Ripristina Master da Budget       → Salva Budget come Master
└─ Gestione Formule e Label (solo master)
   └─ ...
```

---

## 🎓 Best Practices

1. **Inizializza subito**: Prima operazione in un nuovo file
2. **Nomi descrittivi**: Usa descrizioni chiare per ogni offerta
3. **Rigenera sempre**: Dopo ogni modifica, rigenera Budget
4. **Backup prima di inizializzare**: Fai copia file prima del setup
5. **Testa su copia**: Prova funzioni su copia del file
6. **Disabilita invece di eliminare**: Meglio disabilitare che eliminare offerte
7. **Esporta Budget varianti**: Salva PDF/Excel per ogni combinazione
8. **Commenta in descrizione**: Annota cosa cambia in ogni offerta

---

## 📞 Supporto

**Problemi?**
1. Controlla questa guida
2. Guarda sezione Troubleshooting
3. Verifica log esecuzione in Apps Script
4. Controlla che foglio "Configurazione" esista

**File di riferimento**:
- `OffertaManager.js` - Logica principale
- `OffertaConfigUI.html` - Interfaccia sidebar
- `Config.js` - Configurazioni (SHEETS, OFFERTE)

---

**Versione**: 08 - Sistema Multi-Offerta
**Data**: 25 Novembre 2025
**Autore**: Claude Code
