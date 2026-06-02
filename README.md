<div align="center">

# 🎓 Stage Match — ITIS Pietro Paleocapa

**Portale di abbinamento studenti-aziende per lo stage PCTO estivo**  
Sviluppato dagli studenti del progetto SitLab · ITIS P. Paleocapa, Bergamo

---

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![Flask](https://img.shields.io/badge/Flask-3.x-000000?style=flat-square&logo=flask&logoColor=white)](https://flask.palletsprojects.com)
[![SQLite](https://img.shields.io/badge/SQLite-3-003B57?style=flat-square&logo=sqlite&logoColor=white)](https://sqlite.org)
[![License](https://img.shields.io/badge/License-Didattico-yellow?style=flat-square)](.)

</div>

---

## 📖 Cos'è Stage Match?

Stage Match è un'applicazione web che automatizza e ottimizza l'abbinamento tra studenti dell'ITIS Paleocapa e le aziende partner per gli stage PCTO estivi. Il sistema legge le risposte ai questionari Google Forms, le sincronizza in un database locale e calcola la compatibilità tra ogni studente e le aziende disponibili tramite un algoritmo a punteggio ponderato, con un'analisi testuale finale generata dall'AI.

---

## 🏗️ Architettura del sistema

```
PaleoMatch/
├── app.py                    # Backend Flask: route, modelli DB, sync, match, AI
├── shared_modules/
│   ├── sso_middleware.py     # Autenticazione SSO / JWT / Rate Limiting
│   └── __init__.py
├── static/
│   ├── css/style.css         # Foglio di stile completo (tema chiaro/scuro)
│   └── js/app.js             # Logica frontend (login, tema, cookie)
├── templates/
│   ├── base.html             # Layout base con navbar, modale login, footer
│   ├── index.html            # Home: hero, info, servizi, tech
│   ├── profilo.html          # Profilo studente con grafici e pesi
│   ├── profilo_azienda.html  # Profilo azienda con requisiti PCTO
│   ├── match.html            # Risultati abbinamento + analisi AI
│   ├── aziende.html          # Lista pubblica delle aziende
│   ├── privacy.html          # Informativa GDPR (studente + azienda)
│   ├── admin.html            # Pannello amministrazione
│   └── ...                   # Template modifica/profilo admin
├── instance/
│   └── database.db           # Database SQLite (auto-generato)
├── Credenziali.json          # Service Account Google (da fornire)
├── .env                      # Variabili di ambiente (da configurare)
├── requirements.txt          # Dipendenze Python
├── README.md                 # Questo file
└── TODO_SETUP.txt            # Lista file e chiavi ancora da inserire
```

---

## ⚙️ Come funziona

### 1. Questionari Google Forms → Sheets
Gli studenti e le aziende compilano separatamente due Google Forms. Le risposte vengono salvate automaticamente in due Google Sheets distinti.

### 2. Sincronizzazione automatica
Al primo accesso (o ogni 15 minuti), `app.py` legge le righe dai Sheets tramite le API Google e le inserisce nel database SQLite locale. La sincronizzazione usa una strategia **gap-filling**: se gli ID 1 e 2 sono stati eliminati, i nuovi record li reutilizzano prima di procedere con ID superiori.

### 3. Algoritmo di match
Per ogni coppia (studente, azienda):
- Si scorrono le 19 attività PCTO
- Per ogni attività **prevista dall'azienda**, si confronta il livello richiesto (1–3, normalizzato 1–5) con la media tra interesse e competenza dello studente (1–5)
- La compatibilità = `1 - (distanza / 5)`, moltiplicata per il **peso** configurabile
- Il punteggio finale è la media ponderata × 100

I pesi sono configurabili dall'admin (da 0.1 a 5.0) e influenzano in tempo reale tutti i calcoli di match.

### 4. Analisi AI
Dopo il calcolo, i top-5 candidati vengono inviati a un modello LLM (Ollama/OpenAI-compatible) che genera un'analisi testuale di 3–4 frasi con consigli pratici.

### 5. Autenticazione SSO
L'accesso avviene tramite email verificata nel database:
- **Studente**: inserisce la propria email scolastica → viene cercata nel DB → login
- **Azienda**: inserisce l'email con cui ha compilato il Form → login con ID azienda
- **Admin**: username + password configurati in `.env`

In modalità `dev`, il sistema bypassa JWT e consente login diretti tramite query string.

---

## 🚀 Avvio rapido

### Prerequisiti
- Python 3.10+
- Un Google Cloud Project con le API Sheets/Drive abilitate
- Un file `Credenziali.json` (Service Account)

## 🔑 Ruoli e accessi

| Ruolo | Come accede | Cosa può fare |
|-------|------------|---------------|
| **Studente** | Email scolastica nel login | Vede il proprio profilo, grafici ponderati, calcola il match con le aziende |
| **Azienda** | Email aziendale nel login | Vede il proprio profilo con requisiti PCTO, calcola il match con gli studenti |
| **Admin** | Username + password | Gestisce tutti i profili, modifica dati, configura i pesi dell'algoritmo, forzza le sync |

---

## 🎨 Design e tema

Il sito usa sempre il **tema chiaro** all'avvio (nessun salvataggio della preferenza). L'utente può passare al tema scuro tramite il pulsante in navbar, ma alla prossima sessione si riparte sempre dal tema chiaro.

I banner dei profili cambiano colore automaticamente in base al punteggio globale:
- 🟦 **Eccellente** (≥80%) — gradiente blu/viola
- 🟢 **Ottimo** (≥60%) — gradiente verde
- 🟡 **Buono** (≥40%) — gradiente ambra
- 🟠 **Base** (≥20%) — gradiente arancione
- ⚫ **Minimo** (<20%) — grigio neutro

---

## 📊 Database

Il DB SQLite viene creato automaticamente al primo avvio con tre tabelle:

- `questionari` — dati studenti (email, nome, cognome, classe, interesse/competenza 1–19)
- `questionari_aziende` — dati aziende (nome, email, descrizione, prevista/livello/formazione 1–19)
- `pesi_risposta` — pesi per ciascuna delle 19 domande (configurabili da admin)

Le migrazioni dello schema vengono applicate automaticamente ad ogni avvio.

---

## 🧠 AI Match

L'analisi AI utilizza un'istanza Ollama esposta tramite API compatibile OpenAI. Il modello riceve i top-5 candidati e produce un testo di orientamento di 3–4 frasi. Se Ollama non è raggiungibile, il match funziona comunque normalmente (l'analisi viene semplicemente omessa).

---

## 📄 Licenza

Progetto a scopo **esclusivamente didattico** — ITIS Pietro Paleocapa, Bergamo.  
Non destinato a uso commerciale o produzione.

---

<div align="center">
  <sub>Stage Match · SitLab · ITIS P. Paleocapa · 2026</sub>
</div>
