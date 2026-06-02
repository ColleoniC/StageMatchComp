import os, sys, secrets, time
from datetime import timedelta, datetime

import gspread
from flask import Flask, render_template, request, session, redirect, url_for, jsonify, flash
from flask_sqlalchemy import SQLAlchemy
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
from werkzeug.middleware.proxy_fix import ProxyFix

load_dotenv()

sys.path.insert(0, os.path.dirname(__file__))
try:
    from shared_modules.sso_middleware import SSOMiddleware, RateLimiter, render_sso_error
except ImportError:
    from sso_middleware import SSOMiddleware, RateLimiter, render_sso_error

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)
app.secret_key = os.getenv('SERVER_SECRET_KEY', 'dev-secret')
app.permanent_session_lifetime = timedelta(hours=8)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

SSO_MODE       = os.getenv('SSO_MODE', 'dev').lower()
DEV_USER_EMAIL = os.getenv('DEV_USER_EMAIL', 'mario.rossi@itispaleocapa.it')
SSO_CONFIG = {
    'jwt_secret':   os.getenv('JWT_SECRET'),
    'jwt_algorithm':'HS256',
    'jwt_issuer':   'sso-portal',
    'jwt_audience': os.getenv('APP_AUDIENCE', 'stage-match'),
    'session_timeout': 28800,
    'portal_url':   os.getenv('PORTAL_URL', 'http://localhost:5000'),
}

NOME_SHEET          = os.getenv('SHEET_NAME',          'Questionario PCTO - Studente')
NOME_SHEET_AZIENDE  = os.getenv('SHEET_NAME_AZIENDE',  'Questionario PCTO - Azienda')
FILE_CREDENZIALI    = os.getenv('CREDENTIALS_FILE',    'Credenziali.json')
SCOPES_GOOGLE       = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
ADMIN_USERNAME      = os.getenv('ADMIN_USERNAME', 'admin')
ADMIN_PASSWORD      = os.getenv('ADMIN_PASSWORD', '1234')

SYNC_INTERVAL_SECONDS  = 15 * 60
_ultimo_sync: float    = 0.0
_ultimo_sync_aziende: float = 0.0

if SSO_MODE == 'production' and not SSO_CONFIG['jwt_secret']:
    raise ValueError("JWT_SECRET non configurato!")
if SSO_MODE == 'production':
    app.config.update(
        SESSION_COOKIE_SECURE=True,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE='Lax',
    )

# ── DATABASE ──────────────────────────────────────────────────────────────────

db = SQLAlchemy(app)


class Questionario(db.Model):
    __tablename__ = 'questionari'
    id                = db.Column(db.Integer, primary_key=True, autoincrement=True)
    Indirizzo_email   = db.Column(db.String(255), unique=True, nullable=False)
    Nome              = db.Column(db.String(100), nullable=False)
    Cognome           = db.Column(db.String(100), nullable=False)
    Classe            = db.Column(db.String(10))
    admin_modified    = db.Column(db.Boolean, default=False, nullable=False, server_default='0')

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
        return getattr(self, key)


class QuestionarioAzienda(db.Model):
    __tablename__ = 'questionari_aziende'
    id            = db.Column(db.Integer, primary_key=True, autoincrement=True)
    Azienda       = db.Column(db.String(200), nullable=False)
    Email         = db.Column(db.String(255))
    Descrizione   = db.Column(db.String(500))
    admin_modified= db.Column(db.Boolean, default=False, nullable=False, server_default='0')

    Prevista_1  = db.Column(db.Boolean); Livello_1  = db.Column(db.Integer); Formazione_1  = db.Column(db.String(10))
    Prevista_2  = db.Column(db.Boolean); Livello_2  = db.Column(db.Integer); Formazione_2  = db.Column(db.String(10))
    Prevista_3  = db.Column(db.Boolean); Livello_3  = db.Column(db.Integer); Formazione_3  = db.Column(db.String(10))
    Prevista_4  = db.Column(db.Boolean); Livello_4  = db.Column(db.Integer); Formazione_4  = db.Column(db.String(10))
    Prevista_5  = db.Column(db.Boolean); Livello_5  = db.Column(db.Integer); Formazione_5  = db.Column(db.String(10))
    Prevista_6  = db.Column(db.Boolean); Livello_6  = db.Column(db.Integer); Formazione_6  = db.Column(db.String(10))
    Prevista_7  = db.Column(db.Boolean); Livello_7  = db.Column(db.Integer); Formazione_7  = db.Column(db.String(10))
    Prevista_8  = db.Column(db.Boolean); Livello_8  = db.Column(db.Integer); Formazione_8  = db.Column(db.String(10))
    Prevista_9  = db.Column(db.Boolean); Livello_9  = db.Column(db.Integer); Formazione_9  = db.Column(db.String(10))
    Prevista_10 = db.Column(db.Boolean); Livello_10 = db.Column(db.Integer); Formazione_10 = db.Column(db.String(10))
    Prevista_11 = db.Column(db.Boolean); Livello_11 = db.Column(db.Integer); Formazione_11 = db.Column(db.String(10))
    Prevista_12 = db.Column(db.Boolean); Livello_12 = db.Column(db.Integer); Formazione_12 = db.Column(db.String(10))
    Prevista_13 = db.Column(db.Boolean); Livello_13 = db.Column(db.Integer); Formazione_13 = db.Column(db.String(10))
    Prevista_14 = db.Column(db.Boolean); Livello_14 = db.Column(db.Integer); Formazione_14 = db.Column(db.String(10))
    Prevista_15 = db.Column(db.Boolean); Livello_15 = db.Column(db.Integer); Formazione_15 = db.Column(db.String(10))
    Prevista_16 = db.Column(db.Boolean); Livello_16 = db.Column(db.Integer); Formazione_16 = db.Column(db.String(10))
    Prevista_17 = db.Column(db.Boolean); Livello_17 = db.Column(db.Integer); Formazione_17 = db.Column(db.String(10))
    Prevista_18 = db.Column(db.Boolean); Livello_18 = db.Column(db.Integer); Formazione_18 = db.Column(db.String(10))
    Prevista_19 = db.Column(db.Boolean); Livello_19 = db.Column(db.Integer); Formazione_19 = db.Column(db.String(10))

    def __getitem__(self, key):
        return getattr(self, key)


class PesiRisposta(db.Model):
    __tablename__ = 'pesi_risposta'
    domanda  = db.Column(db.Integer, primary_key=True)
    peso     = db.Column(db.Float,   nullable=False, default=1.0)
    etichetta= db.Column(db.String(30), nullable=False, default='Normale')


def _migra_db():
    with db.engine.connect() as conn:
        from sqlalchemy import inspect, text
        cols = {c['name'] for c in inspect(db.engine).get_columns('questionari')}
        for colonna, sql in {
            'admin_modified': 'ALTER TABLE questionari ADD COLUMN admin_modified BOOLEAN DEFAULT 0 NOT NULL',
            'Classe':         'ALTER TABLE questionari ADD COLUMN Classe VARCHAR(10)',
        }.items():
            if colonna not in cols:
                try:
                    conn.execute(text(sql)); conn.commit()
                except Exception as e:
                    app.logger.warning(f"Migrazione '{colonna}': {e}")
        try:
            cols_az = {c['name'] for c in inspect(db.engine).get_columns('questionari_aziende')}
            for col, sql in {
                'Descrizione': 'ALTER TABLE questionari_aziende ADD COLUMN Descrizione VARCHAR(500)',
                'Email':       'ALTER TABLE questionari_aziende ADD COLUMN Email VARCHAR(255)',
            }.items():
                if col not in cols_az:
                    conn.execute(text(sql)); conn.commit()
        except Exception as e:
            app.logger.warning(f"Migrazione azienda: {e}")
        try:
            cols_p = {c['name'] for c in inspect(db.engine).get_columns('pesi_risposta')}
            if 'etichetta' not in cols_p:
                conn.execute(text("ALTER TABLE pesi_risposta ADD COLUMN etichetta VARCHAR(30) DEFAULT 'Normale' NOT NULL"))
                conn.commit()
        except Exception:
            pass


with app.app_context():
    db.create_all()
    _migra_db()
    for i in range(1, 20):
        if not db.session.get(PesiRisposta, i):
            db.session.add(PesiRisposta(domanda=i, peso=1.0, etichetta='Normale'))
    db.session.commit()


ETICHETTE_PESO = {0.5:'Bassa priorità', 1.0:'Normale', 1.5:'Importante', 2.0:'Molto importante', 3.0:'Prioritario'}


def _etichetta_peso(peso: float) -> str:
    return ETICHETTE_PESO[min(ETICHETTE_PESO, key=lambda v: abs(v - peso))]


def get_pesi() -> dict:
    return {p.domanda: p.peso for p in PesiRisposta.query.all()}


# ── ID GAP-FILLING ────────────────────────────────────────────────────────────

def _prossimo_id_libero(model):
    ids = {r.id for r in db.session.query(model.id).all()}
    n = 1
    while n in ids:
        n += 1
    return n


def _inserisci_con_id_minimo(obj, model):
    obj.id = _prossimo_id_libero(model)
    db.session.add(obj)


# ── SYNC STUDENTI ─────────────────────────────────────────────────────────────

def _parse_intero(v):
    if v in ('', None): return None
    try: return int(v)
    except (ValueError, TypeError):
        try: return int(float(v))
        except: return None


def _parse_bool(v):
    if isinstance(v, bool): return v
    if v in ('', None): return False
    return str(v).strip().lower() in ('si', 'sì', 'yes', 'true', 'vero', '1')


def _parse_formazione(v):
    if v in ('', None): return None
    val = str(v).strip().upper()
    return val if val in ('NP', 'C', 'F', 'B') else None


def sincronizza_google_sheets(forza: bool = False):
    global _ultimo_sync
    ora = time.time()
    if not forza and (ora - _ultimo_sync) < SYNC_INTERVAL_SECONDS: return
    if not os.path.exists(FILE_CREDENZIALI):
        app.logger.warning(f"Credenziali '{FILE_CREDENZIALI}' non trovate.")
        _ultimo_sync = ora; return
    try:
        cred   = ServiceAccountCredentials.from_json_keyfile_name(FILE_CREDENZIALI, SCOPES_GOOGLE)
        client = gspread.authorize(cred)
        righe  = client.open(NOME_SHEET).sheet1.get_all_records()
    except Exception as e:
        app.logger.error(f"Errore Google Sheets studenti: {e}"); return
    _ultimo_sync = ora
    email_set = {r.Indirizzo_email.lower() for r in Questionario.query.with_entities(Questionario.Indirizzo_email).all()}
    nuovi = []
    for riga in righe:
        email = riga.get('Indirizzo email', '').strip()
        if not email or email.lower() in email_set: continue
        dati = {'Indirizzo_email': email, 'Nome': riga.get('Nome','').strip(),
                'Cognome': riga.get('Cognome','').strip(), 'Classe': riga.get('Classe','').strip(), 'admin_modified': False}
        for i in range(1, 20):
            dati[f'Interesse_{i}']  = _parse_intero(riga.get(f'Interesse {i}'))
            dati[f'Competenza_{i}'] = _parse_intero(riga.get(f'Competenza {i}'))
        nuovi.append(Questionario(**dati))
        email_set.add(email.lower())
    if nuovi:
        try:
            for obj in nuovi: _inserisci_con_id_minimo(obj, Questionario)
            db.session.commit()
            app.logger.info(f"Sync studenti: +{len(nuovi)} profili.")
        except Exception as e:
            db.session.rollback(); app.logger.error(f"Commit sync studenti: {e}")


# ── SYNC AZIENDE ──────────────────────────────────────────────────────────────

def sincronizza_google_sheets_aziende(forza: bool = False):
    global _ultimo_sync_aziende
    ora = time.time()
    if not forza and (ora - _ultimo_sync_aziende) < SYNC_INTERVAL_SECONDS: return
    if not os.path.exists(FILE_CREDENZIALI):
        app.logger.warning(f"Credenziali '{FILE_CREDENZIALI}' non trovate.")
        _ultimo_sync_aziende = ora; return
    try:
        cred   = ServiceAccountCredentials.from_json_keyfile_name(FILE_CREDENZIALI, SCOPES_GOOGLE)
        client = gspread.authorize(cred)
        righe  = client.open(NOME_SHEET_AZIENDE).sheet1.get_all_records()
    except Exception as e:
        app.logger.error(f"Errore Google Sheets aziende: {e}"); return
    _ultimo_sync_aziende = ora
    esistenti = {r.Azienda.strip().lower(): r for r in QuestionarioAzienda.query.all()}
    nuove = aggiornate = 0
    for riga in righe:
        nome = riga.get('Azienda', '').strip()
        if not nome: continue
        dati = {'Azienda': nome, 'Email': riga.get('Email','').strip(), 'Descrizione': riga.get('Descrizione','').strip()}
        for i in range(1, 20):
            dati[f'Prevista_{i}']   = _parse_bool(riga.get(f'Prevista {i}', ''))
            lv = _parse_intero(riga.get(f'Livello {i}', ''))
            dati[f'Livello_{i}']    = lv if lv in (1, 2, 3) else None
            dati[f'Formazione_{i}'] = _parse_formazione(riga.get(f'Formazione {i}', ''))
        chiave = nome.lower()
        record = esistenti.get(chiave)
        if record is None:
            obj = QuestionarioAzienda(admin_modified=False, **dati)
            _inserisci_con_id_minimo(obj, QuestionarioAzienda)
            esistenti[chiave] = obj; nuove += 1
        elif not record.admin_modified:
            for campo, valore in dati.items(): setattr(record, campo, valore)
            aggiornate += 1
    try:
        db.session.commit()
        if nuove or aggiornate:
            app.logger.info(f"Sync aziende: +{nuove} nuove, {aggiornate} aggiornate.")
    except Exception as e:
        db.session.rollback(); app.logger.error(f"Commit sync aziende: {e}")


# ── ALGORITMO MATCH ───────────────────────────────────────────────────────────

def _calcola_match_studente(studente: Questionario, aziende: list, pesi: dict) -> list:
    risultati = []
    for az in aziende:
        punteggio = totale_peso = 0.0
        for i in range(1, 20):
            if not az[f'Prevista_{i}']: continue
            livello_az     = az[f'Livello_{i}']
            interesse_st   = studente[f'Interesse_{i}']
            competenza_st  = studente[f'Competenza_{i}']
            if any(v is None for v in (livello_az, interesse_st, competenza_st)): continue
            peso = pesi.get(i, 1.0)
            media_st    = (interesse_st + competenza_st) / 2.0
            livello_norm = (livello_az / 3.0) * 5.0
            compat = max(0, 1.0 - abs(media_st - livello_norm) / 5.0)
            punteggio += compat * peso
            totale_peso  += peso
        pct = round((punteggio / totale_peso) * 100) if totale_peso > 0 else 0
        risultati.append({'id': az.id, 'azienda': az.Azienda, 'email': az.Email or '',
                          'descrizione': (az.Descrizione or '')[:80], 'punteggio': pct})
    risultati.sort(key=lambda x: x['punteggio'], reverse=True)
    return risultati


def _calcola_match_azienda(az: QuestionarioAzienda, studenti: list, pesi: dict) -> list:
    risultati = []
    for st in studenti:
        punteggio = totale_peso = 0.0
        for i in range(1, 20):
            if not az[f'Prevista_{i}']: continue
            livello_az    = az[f'Livello_{i}']
            interesse_st  = st[f'Interesse_{i}']
            competenza_st = st[f'Competenza_{i}']
            if any(v is None for v in (livello_az, interesse_st, competenza_st)): continue
            peso = pesi.get(i, 1.0)
            media_st    = (interesse_st + competenza_st) / 2.0
            livello_norm = (livello_az / 3.0) * 5.0
            compat = max(0, 1.0 - abs(media_st - livello_norm) / 5.0)
            punteggio += compat * peso
            totale_peso  += peso
        pct = round((punteggio / totale_peso) * 100) if totale_peso > 0 else 0
        risultati.append({'id': st.id, 'nome': st.Nome, 'cognome': st.Cognome,
                          'email': st.Indirizzo_email, 'classe': st.Classe or '—', 'punteggio': pct})
    risultati.sort(key=lambda x: x['punteggio'], reverse=True)
    return risultati


# ── SSO ───────────────────────────────────────────────────────────────────────

rate_limiter = RateLimiter(
    max_sessions_per_user=int(os.getenv('MAX_SESSIONS_PER_USER', 3)),
    max_sessions_global=int(os.getenv('MAX_SESSIONS_GLOBAL', 100)),
    session_ttl_seconds=28800,
)
sso_middleware = SSOMiddleware(**SSO_CONFIG, rate_limiter=rate_limiter)


def _complete_login(user_data):
    sid = secrets.token_hex(32)
    allowed, reason = rate_limiter.register_session(sid, user_data.get('email', ''))
    if not allowed:
        return render_sso_error(reason, SSO_CONFIG['portal_url'], 429, "Troppe sessioni", "⏱️")
    sso_middleware.create_session(user_data, session, session_id=sid)
    session['tipo_utente'] = user_data.get('tipo', 'studente')
    if user_data.get('tipo') == 'azienda' and 'azienda_id' in user_data:
        session['azienda_id'] = user_data['azienda_id']
    return redirect(url_for('home'))


@app.route('/sso/login')
def sso_login():
    tipo  = request.args.get('tipo', 'studente')
    token = request.args.get('token')
    if SSO_MODE == 'dev' and not token:
        if tipo == 'azienda':
            az_id = request.args.get('id')
            if az_id:
                az = db.session.get(QuestionarioAzienda, int(az_id))
                if az:
                    return _complete_login({'email': az.Email or f'azienda_{az.id}@pcto.it',
                                            'name': az.Azienda, 'tipo': 'azienda', 'azienda_id': az.id})
            return redirect(url_for('home'))
        dev_email = request.args.get('email') or DEV_USER_EMAIL
        nome = dev_email.split('@')[0].replace('.', ' ').title()
        return _complete_login({'email': dev_email, 'name': nome, 'tipo': tipo})
    if not token:
        return render_sso_error("Token SSO mancante.", SSO_CONFIG['portal_url'])
    try:
        user_data = sso_middleware.validate_jwt(token)
        user_data['tipo'] = tipo
        return _complete_login(user_data)
    except Exception:
        return render_sso_error("Token non valido o scaduto.", SSO_CONFIG['portal_url'])


@app.route('/logout')
def logout():
    sid = session.get('session_id')
    if sid: rate_limiter.remove_session(sid)
    session.clear()
    return redirect(SSO_CONFIG['portal_url'])


# ── ROUTE PUBBLICHE ───────────────────────────────────────────────────────────

@app.route('/')
def home():
    return render_template('index.html', utente=session.get('user'), tipo=session.get('tipo_utente'))


@app.route('/aziende')
def lista_aziende():
    sincronizza_google_sheets_aziende()
    aziende = QuestionarioAzienda.query.order_by(QuestionarioAzienda.Azienda).all()
    return render_template('aziende.html', utente=session.get('user'), tipo=session.get('tipo_utente'), aziende=aziende)


@app.route('/aziende/<int:id>')
def profilo_azienda_pubblico(id):
    sincronizza_google_sheets_aziende()
    az = db.session.get(QuestionarioAzienda, id)
    if not az:
        flash('Azienda non trovata.', 'error')
        return redirect(url_for('lista_aziende'))
    pesi = PesiRisposta.query.order_by(PesiRisposta.domanda).all()
    return render_template('profilo_azienda.html', utente=session.get('user'),
                           tipo=session.get('tipo_utente'), az=az, pesi=pesi, vista_propria=False)


@app.route('/profilo')
@sso_middleware.sso_login_required
def profilo():
    tipo = session.get('tipo_utente')
    if tipo == 'azienda':
        az_id = session.get('azienda_id')
        sincronizza_google_sheets_aziende()
        az   = db.session.get(QuestionarioAzienda, az_id) if az_id else None
        pesi = PesiRisposta.query.order_by(PesiRisposta.domanda).all()
        return render_template('profilo_azienda.html', utente=session.get('user'),
                               tipo='azienda', az=az, pesi=pesi, vista_propria=True)
    if tipo != 'studente':
        return redirect(url_for('home'))
    utente = session.get('user')
    email  = utente.get('email', '').strip().lower()
    sincronizza_google_sheets()
    q    = Questionario.query.filter(db.func.lower(Questionario.Indirizzo_email) == email).first()
    pesi = PesiRisposta.query.order_by(PesiRisposta.domanda).all()
    return render_template('profilo.html', utente=utente, tipo='studente', questionario=q, pesi=pesi)


# ── MATCH ─────────────────────────────────────────────────────────────────────

@app.route('/match')
@sso_middleware.sso_login_required
def match():
    tipo = session.get('tipo_utente')
    pesi = get_pesi()
    if tipo == 'studente':
        utente = session.get('user')
        email  = utente.get('email', '').strip().lower()
        sincronizza_google_sheets(); sincronizza_google_sheets_aziende()
        q = Questionario.query.filter(db.func.lower(Questionario.Indirizzo_email) == email).first()
        if not q:
            flash('Profilo studente non trovato. Compila prima il questionario PCTO.', 'error')
            return redirect(url_for('profilo'))
        aziende   = QuestionarioAzienda.query.all()
        risultati = _calcola_match_studente(q, aziende, pesi)
        return render_template('match.html', utente=utente, tipo='studente',
                               risultati=risultati, profilo=q)
    elif tipo == 'azienda':
        az_id = session.get('azienda_id')
        sincronizza_google_sheets(); sincronizza_google_sheets_aziende()
        az = db.session.get(QuestionarioAzienda, az_id) if az_id else None
        if not az:
            flash('Profilo azienda non trovato.', 'error')
            return redirect(url_for('profilo'))
        studenti  = Questionario.query.all()
        risultati = _calcola_match_azienda(az, studenti, pesi)
        return render_template('match.html', utente=session.get('user'), tipo='azienda',
                               risultati=risultati, profilo=az)
    return redirect(url_for('home'))


# ── API ───────────────────────────────────────────────────────────────────────

@app.route('/api/studenti')
def api_studenti():
    sincronizza_google_sheets()
    studenti = Questionario.query.order_by(Questionario.Cognome, Questionario.Nome).all()
    return jsonify([{'email': s.Indirizzo_email, 'nome': s.Nome, 'cognome': s.Cognome} for s in studenti])


@app.route('/api/aziende')
def api_aziende():
    sincronizza_google_sheets_aziende()
    aziende = QuestionarioAzienda.query.order_by(QuestionarioAzienda.Azienda).all()
    return jsonify([{'id': a.id, 'azienda': a.Azienda,
                     'email': (a.Email or '').lower(),
                     'descrizione': (a.Descrizione or '')[:60]} for a in aziende])


@app.route('/api/sessione')
def api_sessione():
    utente = session.get('user')
    if utente:
        return jsonify({'loggato': True, 'utente': utente, 'tipo': session.get('tipo_utente')})
    return jsonify({'loggato': False})


# ── ADMIN ─────────────────────────────────────────────────────────────────────

def _richiedi_admin():
    return session.get('tipo_utente') != 'admin'


@app.route('/admin/login', methods=['POST'])
def admin_login():
    nome     = request.form.get('nome', '').strip()
    password = request.form.get('password', '').strip()
    if nome == ADMIN_USERNAME and password == ADMIN_PASSWORD:
        session.permanent = True
        session['user']        = {'email': 'admin@itispaleocapa.it', 'name': 'Administrator',
                                   'authenticated_at': datetime.utcnow().isoformat()}
        session['tipo_utente'] = 'admin'
        return redirect(url_for('admin_panel'))
    return jsonify({'ok': False, 'errore': 'Credenziali non valide'}), 401


@app.route('/admin')
def admin_panel():
    if _richiedi_admin(): return redirect(url_for('home'))
    sincronizza_google_sheets(); sincronizza_google_sheets_aziende()
    studenti = Questionario.query.order_by(Questionario.Cognome, Questionario.Nome).all()
    aziende  = QuestionarioAzienda.query.order_by(QuestionarioAzienda.Azienda).all()
    pesi     = PesiRisposta.query.order_by(PesiRisposta.domanda).all()
    return render_template('admin.html', utente=session.get('user'), tipo='admin',
                           studenti=studenti, aziende=aziende, pesi=pesi, etichette=ETICHETTE_PESO)


@app.route('/admin/profilo/<int:id>')
def admin_profilo(id):
    if _richiedi_admin(): return redirect(url_for('home'))
    q = db.session.get(Questionario, id)
    if not q:
        flash('Studente non trovato.', 'error'); return redirect(url_for('admin_panel'))
    pesi = PesiRisposta.query.order_by(PesiRisposta.domanda).all()
    return render_template('admin_profilo.html', utente=session.get('user'), tipo='admin', q=q, pesi=pesi)


@app.route('/admin/modifica/<int:id>', methods=['GET', 'POST'])
def admin_modifica(id):
    if _richiedi_admin(): return redirect(url_for('home'))
    q = db.session.get(Questionario, id)
    if not q:
        flash('Studente non trovato.', 'error'); return redirect(url_for('admin_panel'))
    if request.method == 'POST':
        q.Nome              = request.form.get('Nome', '').strip()
        q.Cognome           = request.form.get('Cognome', '').strip()
        q.Indirizzo_email   = request.form.get('Indirizzo_email', '').strip()
        q.Classe            = request.form.get('Classe', '').strip()
        for i in range(1, 20):
            setattr(q, f'Interesse_{i}',  _parse_intero(request.form.get(f'Interesse_{i}')))
            setattr(q, f'Competenza_{i}', _parse_intero(request.form.get(f'Competenza_{i}')))
        q.admin_modified = True
        try:
            db.session.commit(); flash(f'Profilo di {q.Cognome} {q.Nome} salvato.', 'success')
        except Exception as e:
            db.session.rollback(); flash(f'Errore: {e}', 'error')
        return redirect(url_for('admin_profilo', id=id))
    return render_template('admin_modifica.html', utente=session.get('user'), tipo='admin', q=q)


@app.route('/admin/elimina/<int:id>', methods=['POST'])
def admin_elimina(id):
    if _richiedi_admin(): return redirect(url_for('home'))
    q = db.session.get(Questionario, id)
    if q:
        nome = f"{q.Cognome} {q.Nome}"
        db.session.delete(q); db.session.commit()
        flash(f'Record di {nome} eliminato.', 'success')
    else:
        flash('Record non trovato.', 'error')
    return redirect(url_for('admin_panel'))


@app.route('/admin/azienda/profilo/<int:id>')
def admin_azienda_profilo(id):
    if _richiedi_admin(): return redirect(url_for('home'))
    az = db.session.get(QuestionarioAzienda, id)
    if not az:
        flash('Azienda non trovata.', 'error'); return redirect(url_for('admin_panel'))
    pesi = PesiRisposta.query.order_by(PesiRisposta.domanda).all()
    return render_template('admin_azienda_profilo.html', utente=session.get('user'), tipo='admin', az=az, pesi=pesi)


@app.route('/admin/azienda/modifica/<int:id>', methods=['GET', 'POST'])
def admin_azienda_modifica(id):
    if _richiedi_admin(): return redirect(url_for('home'))
    az = db.session.get(QuestionarioAzienda, id)
    if not az:
        flash('Azienda non trovata.', 'error'); return redirect(url_for('admin_panel'))
    if request.method == 'POST':
        az.Azienda      = request.form.get('Azienda', '').strip()
        az.Descrizione  = request.form.get('Descrizione', '').strip()
        for i in range(1, 20):
            setattr(az, f'Prevista_{i}',   request.form.get(f'Prevista_{i}') == 'on')
            setattr(az, f'Livello_{i}',    _parse_intero(request.form.get(f'Livello_{i}')))
            setattr(az, f'Formazione_{i}', _parse_formazione(request.form.get(f'Formazione_{i}')))
        az.admin_modified = True
        try:
            db.session.commit(); flash(f'Profilo di {az.Azienda} salvato.', 'success')
        except Exception as e:
            db.session.rollback(); flash(f'Errore: {e}', 'error')
        return redirect(url_for('admin_azienda_profilo', id=id))
    return render_template('admin_azienda_modifica.html', utente=session.get('user'), tipo='admin', az=az)


@app.route('/admin/azienda/elimina/<int:id>', methods=['POST'])
def admin_azienda_elimina(id):
    if _richiedi_admin(): return redirect(url_for('home'))
    az = db.session.get(QuestionarioAzienda, id)
    if az:
        nome = az.Azienda
        db.session.delete(az); db.session.commit()
        flash(f'Record di {nome} eliminato.', 'success')
    else:
        flash('Record non trovato.', 'error')
    return redirect(url_for('admin_panel'))


@app.route('/admin/pesi', methods=['POST'])
def admin_pesi_salva():
    if _richiedi_admin(): return redirect(url_for('home'))
    for i in range(1, 20):
        try:
            val = max(0.1, min(5.0, float(request.form.get(f'peso_{i}', 1.0))))
        except (TypeError, ValueError):
            val = 1.0
        p = db.session.get(PesiRisposta, i)
        if p:
            p.peso = val; p.etichetta = _etichetta_peso(val)
        else:
            db.session.add(PesiRisposta(domanda=i, peso=val, etichetta=_etichetta_peso(val)))
    db.session.commit()
    flash('Pesi aggiornati.', 'success')
    return redirect(url_for('admin_panel'))


@app.route('/admin/sync-aziende', methods=['POST'])
def admin_sync_aziende():
    if _richiedi_admin(): return redirect(url_for('home'))
    sincronizza_google_sheets_aziende(forza=True)
    flash('Sincronizzazione aziende completata.', 'success')
    return redirect(url_for('admin_panel'))


@app.route('/admin/sync-studenti', methods=['POST'])
def admin_sync_studenti():
    if _richiedi_admin(): return redirect(url_for('home'))
    sincronizza_google_sheets(forza=True)
    flash('Sincronizzazione studenti completata.', 'success')
    return redirect(url_for('admin_panel'))


@app.route('/admin/debug-sheet-aziende')
def admin_debug_sheet():
    if _richiedi_admin(): return redirect(url_for('home'))
    if not os.path.exists(FILE_CREDENZIALI):
        return jsonify({'errore': 'File credenziali non trovato'})
    try:
        cred   = ServiceAccountCredentials.from_json_keyfile_name(FILE_CREDENZIALI, SCOPES_GOOGLE)
        client = gspread.authorize(cred)
        ws     = client.open(NOME_SHEET_AZIENDE).sheet1
        raw    = ws.get_all_values()
        parsed = ws.get_all_records()
        intestazioni = raw[0] if raw else []
        prima_raw    = raw[1] if len(raw) > 1 else []
        prima_parsed = parsed[0] if parsed else {}
        confronto = {k: {'raw': prima_raw[intestazioni.index(k)] if k in intestazioni else '??',
                         'parsed': v, 'tipo': type(v).__name__}
                     for k, v in prima_parsed.items()}
        return jsonify({'righe': len(parsed), 'intestazioni': intestazioni, 'prima_riga': confronto})
    except Exception as e:
        import traceback
        return jsonify({'errore': str(e), 'traceback': traceback.format_exc()})


# ── MODIFICA PROFILO UTENTE ───────────────────────────────────────────────────

@app.route('/profilo/modifica', methods=['GET', 'POST'])
@sso_middleware.sso_login_required
def profilo_modifica_studente():
    if session.get('tipo_utente') != 'studente': return redirect(url_for('home'))
    utente = session.get('user')
    email  = utente.get('email', '').strip().lower()
    q      = Questionario.query.filter(db.func.lower(Questionario.Indirizzo_email) == email).first()
    if not q:
        flash('Profilo non trovato.', 'error'); return redirect(url_for('profilo'))
    if request.method == 'POST':
        q.Classe = request.form.get('Classe', '').strip()
        q.admin_modified = True
        try:
            db.session.commit(); flash('Profilo aggiornato.', 'success')
        except Exception as e:
            db.session.rollback(); flash(f'Errore: {e}', 'error')
        return redirect(url_for('profilo'))
    pesi = PesiRisposta.query.order_by(PesiRisposta.domanda).all()
    return render_template('profilo_modifica_studente.html', utente=utente, tipo='studente', q=q, pesi=pesi)


@app.route('/profilo/modifica-azienda', methods=['GET', 'POST'])
@sso_middleware.sso_login_required
def profilo_modifica_azienda():
    if session.get('tipo_utente') != 'azienda': return redirect(url_for('home'))
    az_id = session.get('azienda_id')
    az    = db.session.get(QuestionarioAzienda, az_id) if az_id else None
    if not az:
        flash('Profilo azienda non trovato.', 'error'); return redirect(url_for('profilo'))
    if request.method == 'POST':
        az.Email        = request.form.get('Email', '').strip()
        az.Descrizione  = request.form.get('Descrizione', '').strip()
        az.admin_modified = True
        try:
            db.session.commit(); flash('Profilo aggiornato.', 'success')
        except Exception as e:
            db.session.rollback(); flash(f'Errore: {e}', 'error')
        return redirect(url_for('profilo'))
    return render_template('profilo_modifica_azienda.html', utente=session.get('user'), tipo='azienda', az=az)


@app.route('/privacy')
def privacy():
    return render_template('privacy.html', utente=session.get('user'), tipo=session.get('tipo_utente'))


if __name__ == '__main__':
    port  = int(os.getenv('PORT', 5000))
    debug = os.getenv('DEBUG', 'False').lower() == 'true'
    print(f'Server su http://127.0.0.1:{port}')
    app.run(host='127.0.0.1', port=port, debug=debug)
