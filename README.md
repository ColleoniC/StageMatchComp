Stage Match — ITIS P. Paleocapa
 
Applicazione web per l'abbinamento studenti–aziende nello stage FSL estivo.
 
---
 
Cos'è
 
Gli studenti compilano un questionario su competenze e interessi. Un algoritmo abbina ogni studente alle aziende partner più compatibili, considerando affinità tecniche e vicinanza geografica.
 
Stack
 
- Backend — Python / Flask + SQLAlchemy (SQLite)
- Auth — SSO con JWT
- Dati — Sincronizzazione automatica da Google Sheets via Service Account
- Frontend — HTML/CSS/JS 
 
Avvio rapido

pip install -r requirements.txt
python app.py
→ http://127.0.0.1:5000
---
