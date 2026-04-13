import os
import sys
import secrets
from datetime import timedelta

import gspread
from flask import Flask, render_template, request, session, redirect, url_for, jsonify, flash
from flask_sqlalchemy import SQLAlchemy
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
from werkzeug.middleware.proxy_fix import ProxyFix

load_dotenv()

# ── SSO MIDDLEWARE (Path Login) ───────────────────────────────────────────────
# Importa esattamente il modulo condiviso di Path Login
sys.path.insert(0, os.path.dirname(__file__))
try:
    from shared_modules.sso_middleware import SSOMiddleware, WhitelistManager, RateLimiter, render_sso_error
except ImportError:
    from sso_middleware import SSOMiddleware, WhitelistManager, RateLimiter, render_sso_error

# ── FLASK APP ─────────────────────────────────────────────────────────────────

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)

app.secret_key = os.getenv('SERVER_SECRET_KEY', 'dev-secret-change-in-production')
app.permanent_session_lifetime = timedelta(hours=8)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# ── VARIABILI DI CONFIGURAZIONE ───────────────────────────────────────────────

SSO_MODE        = os.getenv('SSO_MODE', 'dev').lower()
DEV_USER_EMAIL  = os.getenv('DEV_USER_EMAIL', 'mario.rossi@itispaleocapa.it')

SSO_CONFIG = {
    'jwt_secret':   os.getenv('JWT_SECRET'),
    'jwt_algorithm': 'HS256',
    'jwt_issuer':   'sso-portal',
    'jwt_audience': os.getenv('APP_AUDIENCE', 'stage-match'),
    'session_timeout': 28800,
    'portal_url':   os.getenv('PORTAL_URL', 'http://localhost:5000'),
}

# Google Sheets — stesso schema del app.py condiviso
NOME_SHEET       = os.getenv('SHEET_NAME', 'Questionario PCTO - Studente')
FILE_CREDENZIALI = os.getenv('CREDENTIALS_FILE', 'Credenziali.json')
SCOPES_GOOGLE    = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
]

if SSO_MODE == 'production' and not SSO_CONFIG['jwt_secret']:
    raise ValueError("JWT_SECRET non configurato! Aggiungilo al file .env")

if SSO_MODE == 'production':
    app.config.update(
        SESSION_COOKIE_SECURE=True,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE='Lax',
    )

# ── DATABASE ──────────────────────────────────────────────────────────────────

db = SQLAlchemy(app)


class Questionario(db.Model):
    """
    Specchio locale del foglio Google 'Questionario PCTO - Studente'.
    Ogni riga del foglio corrisponde a un record qui.
    """
    __tablename__ = 'questionari'

    id                 = db.Column(db.Integer, primary_key=True, autoincrement=True)
    Indirizzo_email    = db.Column(db.String(255), unique=True, nullable=False)
    Nome               = db.Column(db.String(100), nullable=False)
    Cognome            = db.Column(db.String(100), nullable=False)
    Luogo_Di_Residenza = db.Column(db.String(100))

    Interesse_1  = db.Column(db.Integer);  Competenza_1  = db.Column(db.Integer)
    Interesse_2  = db.Column(db.Integer);  Competenza_2  = db.Column(db.Integer)
    Interesse_3  = db.Column(db.Integer);  Competenza_3  = db.Column(db.Integer)
    Interesse_4  = db.Column(db.Integer);  Competenza_4  = db.Column(db.Integer)
    Interesse_5  = db.Column(db.Integer);  Competenza_5  = db.Column(db.Integer)
    Interesse_6  = db.Column(db.Integer);  Competenza_6  = db.Column(db.Integer)
    Interesse_7  = db.Column(db.Integer);  Competenza_7  = db.Column(db.Integer)
    Interesse_8  = db.Column(db.Integer);  Competenza_8  = db.Column(db.Integer)
    Interesse_9  = db.Column(db.Integer);  Competenza_9  = db.Column(db.Integer)
    Interesse_10 = db.Column(db.Integer);  Competenza_10 = db.Column(db.Integer)
    Interesse_11 = db.Column(db.Integer);  Competenza_11 = db.Column(db.Integer)
    Interesse_12 = db.Column(db.Integer);  Competenza_12 = db.Column(db.Integer)
    Interesse_13 = db.Column(db.Integer);  Competenza_13 = db.Column(db.Integer)
    Interesse_14 = db.Column(db.Integer);  Competenza_14 = db.Column(db.Integer)
    Interesse_15 = db.Column(db.Integer);  Competenza_15 = db.Column(db.Integer)
    Interesse_16 = db.Column(db.Integer);  Competenza_16 = db.Column(db.Integer)
    Interesse_17 = db.Column(db.Integer);  Competenza_17 = db.Column(db.Integer)
    Interesse_18 = db.Column(db.Integer);  Competenza_18 = db.Column(db.Integer)
    Interesse_19 = db.Column(db.Integer);  Competenza_19 = db.Column(db.Integer)

    def __getitem__(self, key):
        """Permette accesso tipo dizionario: questionario['Interesse_1']"""
        return getattr(self, key)


with app.app_context():
    db.create_all()

# ── GOOGLE SHEETS SYNC ────────────────────────────────────────────────────────
# Logica identica all'app.py condiviso, adattata a Flask-SQLAlchemy ORM

def _parse_intero(valore):
    """Converte in int o ritorna None se vuoto/non valido."""
    try:
        return int(valore) if valore not in ('', None) else None
    except (ValueError, TypeError):
        return None


def sincronizza_google_sheets():
    """
    Legge TUTTE le righe del foglio Google e le inserisce/aggiorna nel DB locale.
    Viene chiamata sulle route che richiedono dati aggiornati (profilo studente).
    Il foglio è la fonte di verità: righe aggiunte mesi fa vengono comunque importate.
    """
    try:
        credenziali = ServiceAccountCredentials.from_json_keyfile_name(FILE_CREDENZIALI, SCOPES_GOOGLE)
        client = gspread.authorize(credenziali)
        foglio = client.open(NOME_SHEET).sheet1
        righe  = foglio.get_all_records()
    except Exception as e:
        app.logger.error(f"Errore Google Sheets: {e}")
        return

    for riga in righe:
        email = riga.get('Indirizzo email', '').strip()
        if not email:
            continue

        dati = {
            'Indirizzo_email':    email,
            'Nome':               riga.get('Nome', '').strip(),
            'Cognome':            riga.get('Cognome', '').strip(),
            'Luogo_Di_Residenza': riga.get('Luogo di residenza', '').strip(),
        }
        for i in range(1, 20):
            dati[f'Interesse_{i}']  = _parse_intero(riga.get(f'Interesse {i}'))
            dati[f'Competenza_{i}'] = _parse_intero(riga.get(f'Competenza {i}'))

        esistente = Questionario.query.filter_by(Indirizzo_email=email).first()
        if esistente:
            for chiave, valore in dati.items():
                setattr(esistente, chiave, valore)
        else:
            db.session.add(Questionario(**dati))

    try:
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Errore commit DB: {e}")


# ── RATE LIMITER & SSO (Path Login) ──────────────────────────────────────────
# Costruiti esattamente come in Path Login/app.py

# DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
# os.makedirs(DATA_DIR, exist_ok=True)
# whitelist_manager = WhitelistManager(whitelist_path=os.path.join(DATA_DIR, 'whitelist.json'))

rate_limiter = RateLimiter(
    max_sessions_per_user=int(os.getenv('MAX_SESSIONS_PER_USER', 3)),
    max_sessions_global=int(os.getenv('MAX_SESSIONS_GLOBAL', 100)),
    session_ttl_seconds=28800,
)

sso_middleware = SSOMiddleware(
    **SSO_CONFIG,
    # whitelist_manager=whitelist_manager,
    rate_limiter=rate_limiter,
)


# ── HELPER ────────────────────────────────────────────────────────────────────

def get_username(email: str) -> str:
    """Estrae la parte locale dell'email (prima della @)."""
    return email.split('@')[0] if '@' in email else email


def _complete_login(user_data: dict):
    """
    Logica post-validazione JWT (copiata da Path Login):
    1. Whitelist (commentata)
    2. Rate limit
    3. Crea sessione Flask e redirect
    """
    email = user_data.get('email', '')

    # 1. Controllo whitelist
    # if not whitelist_manager.is_authorized(email):
    #     return render_sso_error(
    #         f"Account non autorizzato: {email}",
    #         SSO_CONFIG['portal_url'], 403,
    #         "Account Non Autorizzato", "🚫"
    #     )

    # 2. Rate limit
    session_id = secrets.token_hex(32)
    allowed, reason = rate_limiter.register_session(session_id, email)
    if not allowed:
        app.logger.warning(f"Rate limit per: {email}")
        return render_sso_error(
            reason, SSO_CONFIG['portal_url'], 429,
            "Troppe Sessioni Attive", "⏱️"
        )

    # 3. Sessione Flask
    sso_middleware.create_session(user_data, session, session_id=session_id)
    session['tipo_utente'] = user_data.get('tipo', 'studente')

    return redirect(url_for('home'))


# ── ROUTE SSO ─────────────────────────────────────────────────────────────────

@app.route('/sso/login')
def sso_login():
    """
    Unico punto di ingresso autenticato.
    Il portale SSO chiama questa URL passando ?token=JWT[&tipo=studente|azienda].
    In modalità dev il token è opzionale: si usa DEV_USER_EMAIL.
    """
    tipo  = request.args.get('tipo', 'studente')
    token = request.args.get('token')

    if tipo == 'azienda':
        # Le aziende non hanno questionario — redirect diretto
        dummy = {'email': 'azienda@demo.it', 'name': 'Azienda Demo', 'tipo': 'azienda'}
        return _complete_login(dummy)

    # Modalità dev: login simulato senza portale reale
    if SSO_MODE == 'dev' and not token:
        dev_email = request.args.get('email') or DEV_USER_EMAIL
        app.logger.info(f"DEV MODE: login simulato per {dev_email}")
        user_data = {
            'email': dev_email,
            'name':  get_username(dev_email).replace('.', ' ').title(),
            'tipo':  tipo,
        }
        return _complete_login(user_data)

    if not token:
        return render_sso_error(
            "Token SSO mancante. Accedi tramite il portale.",
            SSO_CONFIG['portal_url']
        )

    try:
        user_data = sso_middleware.validate_jwt(token)
        user_data['tipo'] = tipo
        return _complete_login(user_data)
    except Exception as e:
        app.logger.error(f"Errore JWT: {e}")
        return render_sso_error(
            "Token SSO non valido o scaduto. Effettua nuovamente il login.",
            SSO_CONFIG['portal_url']
        )


@app.route('/logout')
def logout():
    """Rimuove sessione dal rate limiter e pulisce il cookie."""
    sid = session.get('session_id')
    if sid:
        rate_limiter.remove_session(sid)
    session.clear()
    return redirect(SSO_CONFIG['portal_url'])


# ── ROUTE APPLICAZIONE ────────────────────────────────────────────────────────

@app.route('/')
def home():
    return render_template('index.html',
                           utente=session.get('user'),
                           tipo=session.get('tipo_utente'))


@app.route('/aziende')
def lista_aziende():
    return render_template('aziende.html',
                           utente=session.get('user'),
                           tipo=session.get('tipo_utente'))


@app.route('/profilo')
@sso_middleware.sso_login_required
def profilo():
    """
    Profilo studente: protetta da SSO.
    Sincronizza Google Sheets prima di mostrare i dati,
    così il profilo è sempre aggiornato anche se i dati
    sono stati inseriti mesi fa.
    """
    if session.get('tipo_utente') != 'studente':
        return redirect(url_for('home'))

    utente = session.get('user')
    email  = utente.get('email', '').strip().lower()

    sincronizza_google_sheets()

    questionario = Questionario.query.filter(
        db.func.lower(Questionario.Indirizzo_email) == email
    ).first()

    return render_template('profilo.html',
                           utente=utente,
                           tipo='studente',
                           questionario=questionario)


@app.route('/mappa')
def mappa():
    return render_template('mappa.html',
                           utente=session.get('user'),
                           tipo=session.get('tipo_utente'))


@app.route('/api/studenti')
def api_studenti():
    """
    Ritorna la lista degli studenti nel DB per il selettore di login in modalità dev.
    Usato dal frontend per popolare la lista nella modale di login.
    """
    sincronizza_google_sheets()
    studenti = Questionario.query.order_by(Questionario.Cognome, Questionario.Nome).all()
    return jsonify([
        {
            'email': s.Indirizzo_email,
            'nome':  s.Nome,
            'cognome': s.Cognome,
        }
        for s in studenti
    ])


@app.route('/api/sessione')
def api_sessione():
    """Endpoint JSON per il frontend: stato login corrente."""
    utente = session.get('user')
    if utente:
        return jsonify({'loggato': True, 'utente': utente, 'tipo': session.get('tipo_utente')})
    return jsonify({'loggato': False})


# ── ENTRYPOINT ────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    port  = int(os.getenv('PORT', 5000))
    debug = os.getenv('DEBUG', 'False').lower() == 'true'
    print(f'Server avviato su http://127.0.0.1:{port}')
    app.run(host='127.0.0.1', port=port, debug=debug)

# ── ADMIN ─────────────────────────────────────────────────────────────────────
# Credenziali admin hardcoded (uso interno/didattico)
ADMIN_USERNAME = os.getenv('ADMIN_USERNAME', 'admin')
ADMIN_PASSWORD = os.getenv('ADMIN_PASSWORD', '1234')


@app.route('/admin/login', methods=['POST'])
def admin_login():
    """Verifica credenziali admin e crea sessione con tipo 'admin'."""
    nome     = request.form.get('nome', '').strip()
    password = request.form.get('password', '').strip()

    if nome == ADMIN_USERNAME and password == ADMIN_PASSWORD:
        session.permanent = True
        session['user'] = {
            'email': 'admin@itispaleocapa.it',
            'name':  'Administrator',
            'authenticated_at': __import__('datetime').datetime.utcnow().isoformat()
        }
        session['tipo_utente'] = 'admin'
        return redirect(url_for('admin_panel'))

    return jsonify({'ok': False, 'errore': 'Credenziali non valide'}), 401


@app.route('/admin')
def admin_panel():
    """Pannello admin: lista studenti con possibilità di eliminazione."""
    if session.get('tipo_utente') != 'admin':
        return redirect(url_for('home'))
    sincronizza_google_sheets()
    studenti = Questionario.query.order_by(Questionario.Cognome, Questionario.Nome).all()
    return render_template('admin.html',
                           utente=session.get('user'),
                           tipo='admin',
                           studenti=studenti)


@app.route('/admin/elimina/<int:id>', methods=['POST'])
def admin_elimina(id):
    """Elimina un record Questionario dal DB."""
    if session.get('tipo_utente') != 'admin':
        return redirect(url_for('home'))
    q = Questionario.query.get(id)
    if q:
        nome = f"{q.Cognome} {q.Nome}"
        db.session.delete(q)
        db.session.commit()
        flash(f'Record di {nome} eliminato correttamente.', 'success')
    else:
        flash('Record non trovato.', 'error')
    return redirect(url_for('admin_panel'))
