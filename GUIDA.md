# MGS5 Tracker — Guida Installazione Completa

---

## STEP 1 — Installa Node.js

1. Vai su **nodejs.org**
2. Scarica il pulsante verde **LTS**
3. Installa normalmente (avanti, avanti, fine)
4. Verifica: apri il Prompt dei comandi e scrivi:
   ```
   node --version
   ```
   Se vedi un numero tipo `v20.x.x` è installato correttamente.

---

## STEP 2 — Prepara la cartella

1. Scarica il file **mgs5-app.zip** che ti ho dato
2. Spostalo sul **Desktop**
3. Fai click destro → **Estrai tutto** → estrai sul Desktop
4. Si crea la cartella `mgs5-app` sul Desktop

---

## STEP 3 — Apri il Prompt dei comandi

1. Premi **Windows + R**
2. Scrivi `cmd` e premi Invio
3. Scrivi questi comandi uno alla volta, premi Invio dopo ognuno:
   ```
   cd Desktop\mgs5-app
   npm install
   npm run dev
   ```
4. Aspetta che finisce (scarica i pacchetti, circa 1-2 minuti)
5. Quando vedi `Local: http://localhost:5173` è pronto

---

## STEP 4 — Testa sul PC

1. Apri **Chrome**
2. Vai su `http://localhost:5173`
3. L'app si apre — prova ad aggiungere una ricarica
4. I dati rimangono anche se ricarichi la pagina ✅

---

## STEP 5 — Costruisci la versione finale

Nel Prompt dei comandi scrivi:
```
npm run build
```
Si crea una cartella `dist` dentro `mgs5-app`

---

## STEP 6 — Installa sul telefono Samsung (PWA)

**Opzione A — Via rete locale (più semplice):**

1. Nel Prompt dei comandi scrivi:
   ```
   npm run preview -- --host
   ```
2. Ti dà un indirizzo tipo `http://192.168.1.X:4173`
3. Sul Samsung apri **Chrome**
4. Vai a quell'indirizzo
5. Chrome mostra in basso "Aggiungi a schermata Home" — tocca
6. L'app si installa con icona come un'app normale ✅

**Opzione B — Via GitHub Pages (accesso da ovunque, gratis):**

1. Crea account su **github.com**
2. Crea nuovo repository chiamato `mgs5`
3. Carica la cartella `dist` su GitHub
4. Vai su Settings → Pages → Branch: main → Save
5. GitHub ti dà un URL tipo `https://tuonome.github.io/mgs5`
6. Apri quell'URL sul Samsung con Chrome
7. **Aggiungi a schermata Home** ✅

---

## STEP 7 — Google Apps Script (sincronizzazione Sheet)

1. Apri il tuo **Google Sheet**
2. Vai su **Estensioni → Apps Script**
3. Cancella tutto il codice che c'è
4. Incolla il contenuto del file `google-apps-script.js` che trovi nella cartella
5. Clicca **Salva** (icona floppy)
6. Clicca **Distribuisci → Nuova distribuzione**
7. Tipo: **App web**
8. Esegui come: **Me**
9. Chi ha accesso: **Chiunque**
10. Clicca **Distribuisci**
11. **Copia l'URL** che ti dà — tipo `https://script.google.com/macros/s/XXXXX/exec`

---

## STEP 8 — Collega l'app allo Sheet

1. Apri l'app sul telefono
2. Vai su **Config** (icona ⚙️)
3. Trova il campo **Google Sheet URL**
4. Incolla l'URL copiato al punto 7
5. Da questo momento ogni ricarica salvata si sincronizza automaticamente su Google Sheet ✅

---

## RIEPILOGO FINALE

| Cosa | Dove |
|------|------|
| Dati salvati | Telefono (localStorage) |
| Backup automatico | Google Sheet |
| Aggiornamento app | Sostituisci App.jsx + `npm run build` |
| Accesso da ovunque | URL GitHub Pages |

---

## AGGIORNARE L'APP IN FUTURO

Quando vuoi aggiornare l'app con nuove funzioni:
1. Sostituisci il file `mgs5-app/src/App.jsx` con la versione nuova
2. Nel Prompt dei comandi scrivi:
   ```
   cd Desktop\mgs5-app
   npm run build
   ```
3. Carica di nuovo la cartella `dist` su GitHub
4. L'app sul telefono si aggiorna automaticamente entro pochi minuti
