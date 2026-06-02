// ── THEME (localStorage persisted) ───────────────────────────────────────────
function _iconaLuna() {
  return '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1111.21 3a7 7 0 009.79 9.79z"/></svg>';
}
function _iconaSole() {
  return '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>';
}

function applicaTema() {
  const tema = localStorage.getItem('pm_tema') || 'chiaro';
  document.documentElement.dataset.tema = tema;
  const btn = document.getElementById('btnTema');
  if (btn) btn.innerHTML = tema === 'scuro' ? _iconaSole() : _iconaLuna();
}

function toggleTema() {
  const corrente = document.documentElement.dataset.tema;
  const nuovo = corrente === 'scuro' ? 'chiaro' : 'scuro';
  document.documentElement.dataset.tema = nuovo;
  localStorage.setItem('pm_tema', nuovo);
  const btn = document.getElementById('btnTema');
  if (btn) btn.innerHTML = nuovo === 'scuro' ? _iconaSole() : _iconaLuna();
}

// ── COOKIE BANNER ─────────────────────────────────────────────────────────────
(function () {
  if (!localStorage.getItem('pm_cookie_ok')) {
    const banner = document.getElementById('cookieBanner');
    if (banner) {
      banner.style.display = 'flex';
      requestAnimationFrame(() => banner.classList.add('cookie-banner--visibile'));
    }
  }
})();

function accettaCookie() {
  localStorage.setItem('pm_cookie_ok', '1');
  const banner = document.getElementById('cookieBanner');
  if (banner) {
    banner.classList.remove('cookie-banner--visibile');
    banner.classList.add('cookie-banner--uscita');
    setTimeout(() => (banner.style.display = 'none'), 400);
  }
}

// ── SESSIONE (localStorage cache per navbar) ──────────────────────────────────
function _salvaSessione(dati) {
  try {
    localStorage.setItem('pm_sessione', JSON.stringify({ ...dati, _ts: Date.now() }));
  } catch (_) {}
}

function _leggiSessioneCache() {
  try {
    const raw = localStorage.getItem('pm_sessione');
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (Date.now() - obj._ts > 30 * 60 * 1000) { localStorage.removeItem('pm_sessione'); return null; }
    return obj;
  } catch (_) { return null; }
}

// ── MODALE LOGIN ──────────────────────────────────────────────────────────────
function apriModalLogin() {
  const m = document.getElementById('modalLogin');
  if (m) { m.classList.add('aperto'); document.body.style.overflow = 'hidden'; }
}
function chiudiModalLogin() {
  const m = document.getElementById('modalLogin');
  if (m) { m.classList.remove('aperto'); document.body.style.overflow = ''; }
  tornaScelta();
}
function gestisciClickOverlay(e) {
  if (e.target === document.getElementById('modalLogin')) chiudiModalLogin();
}
function tornaScelta()         { _mostraStep('stepScelta'); }
function mostraSceltaStudente(){ _mostraStep('stepStudente'); _resetStep('studenteEmail', 'studenteErrore'); }
function mostraSceltaAzienda() { _mostraStep('stepAzienda');  _resetStep('aziendaEmail',   'aziendaErrore');  }
function mostraLoginAdmin()    {
  _mostraStep('stepAdmin');
  ['adminNome','adminPassword'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  const e = document.getElementById('adminErrore'); if(e) e.style.display='none';
}
function _mostraStep(id) {
  ['stepScelta','stepStudente','stepAzienda','stepAdmin'].forEach(s => {
    const el = document.getElementById(s);
    if (el) el.style.display = s === id ? '' : 'none';
  });
}
function _resetStep(inputId, errId) {
  const inp = document.getElementById(inputId);
  const err = document.getElementById(errId);
  if (inp) { inp.value = ''; setTimeout(() => inp.focus(), 60); }
  if (err) { err.style.display = 'none'; err.textContent = ''; }
}

// ── LOGIN STUDENTE ────────────────────────────────────────────────────────────
function loginStudenteEmail() {
  const inp = document.getElementById('studenteEmail');
  const err = document.getElementById('studenteErrore');
  const email = (inp ? inp.value : '').trim().toLowerCase();
  err.style.display = 'none';
  if (!email) { err.textContent = 'Inserisci la tua email scolastica.'; err.style.display = 'block'; return; }
  _setLoginLoading(inp, true);
  fetch('/api/studenti')
    .then(r => r.json())
    .then(studenti => {
      const trovato = studenti.find(s => s.email.toLowerCase() === email);
      if (trovato) {
        window.location.href = '/sso/login?tipo=studente&email=' + encodeURIComponent(trovato.email);
      } else {
        err.innerHTML = 'Email non trovata nel database PCTO.' +
          '<br><span style="font-size:.78rem;color:var(--testo-chiaro);">Hai compilato il questionario con questo indirizzo?</span>' +
          '<br><button onclick="apriSondaggio()" class="err-cta-btn">Compila il sondaggio PCTO</button>';
        err.style.display = 'block';
        _setLoginLoading(inp, false);
      }
    })
    .catch(() => { err.textContent = 'Errore di connessione. Riprova.'; err.style.display = 'block'; _setLoginLoading(inp, false); });
}

// ── LOGIN AZIENDA ─────────────────────────────────────────────────────────────
function loginAziendaEmail() {
  const inp = document.getElementById('aziendaEmail');
  const err = document.getElementById('aziendaErrore');
  const email = (inp ? inp.value : '').trim().toLowerCase();
  err.style.display = 'none';
  if (!email) { err.textContent = 'Inserisci la tua email aziendale.'; err.style.display = 'block'; return; }
  _setLoginLoading(inp, true);
  fetch('/api/aziende')
    .then(r => r.json())
    .then(aziende => {
      const trovata = aziende.find(a => a.email && a.email.toLowerCase() === email);
      if (trovata) {
        window.location.href = '/sso/login?tipo=azienda&id=' + trovata.id;
      } else {
        err.innerHTML = 'Email non trovata nel database PCTO.' +
          '<br><span style="font-size:.78rem;color:var(--testo-chiaro);">Hai compilato il questionario con questo indirizzo?</span>' +
          '<br><button onclick="apriSondaggioAzienda()" class="err-cta-btn">Compila il sondaggio PCTO Azienda</button>';
        err.style.display = 'block';
        _setLoginLoading(inp, false);
      }
    })
    .catch(() => { err.textContent = 'Errore di connessione. Riprova.'; err.style.display = 'block'; _setLoginLoading(inp, false); });
}

function _setLoginLoading(inp, loading) {
  if (!inp) return;
  inp.disabled = loading;
  inp.style.opacity = loading ? '.5' : '';
}

function apriSondaggio()        { window.open('https://forms.gle/WHEiPQb9FtRc2A4D8',  '_blank'); }
function apriSondaggioAzienda() { window.open('https://forms.gle/UgdAjvKemy7nU1MJA', '_blank'); }

// ── LOGIN ADMIN ───────────────────────────────────────────────────────────────
function loginAdmin() {
  const nome     = document.getElementById('adminNome').value.trim();
  const password = document.getElementById('adminPassword').value;
  const errEl    = document.getElementById('adminErrore');
  errEl.style.display = 'none';
  if (!nome || !password) { errEl.textContent = 'Compila tutti i campi.'; errEl.style.display = 'block'; return; }
  const form = new FormData();
  form.append('nome', nome); form.append('password', password);
  fetch('/admin/login', { method: 'POST', body: form })
    .then(r => { if (r.ok || r.redirected) { window.location.href = '/admin'; } else { return r.json().then(d => { throw new Error(d.errore || 'Errore'); }); } })
    .catch(e => { errEl.textContent = e.message || 'Credenziali non valide.'; errEl.style.display = 'block'; });
}

// ── NAVBAR ────────────────────────────────────────────────────────────────────
function aggiornaNavbar() {
  const cache = _leggiSessioneCache();
  if (cache && cache.loggato) { _applicaStatoLogin(cache); }
  fetch('/api/sessione')
    .then(r => r.json())
    .then(dati => { _salvaSessione(dati); _applicaStatoLogin(dati); })
    .catch(() => {});
}

function _applicaStatoLogin(dati) {
  if (dati.loggato && (dati.tipo === 'studente' || dati.tipo === 'azienda')) {
    const lp = document.getElementById('linkProfilo');
    if (lp) { lp.classList.remove('disabilitato'); lp.href = '/profilo'; }
  }
}

// ── HERO ──────────────────────────────────────────────────────────────────────
function inizializzaHero() {
  const img = document.getElementById('heroImg');
  if (img) img.src = 'https://lh3.googleusercontent.com/gps-cs-s/CIHM0ogKEICAgIDU2cCp0QE=w1814-h1360-k-no';
}

// ── INIT ──────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  applicaTema();
  inizializzaHero();
  aggiornaNavbar();
  document.addEventListener('keydown', e => { if (e.key === 'Escape') chiudiModalLogin(); });
  ['adminNome','adminPassword'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('keydown', e => { if (e.key === 'Enter') loginAdmin(); });
  });
  const sE = document.getElementById('studenteEmail');
  if (sE) sE.addEventListener('keydown', e => { if (e.key === 'Enter') loginStudenteEmail(); });
  const aE = document.getElementById('aziendaEmail');
  if (aE) aE.addEventListener('keydown', e => { if (e.key === 'Enter') loginAziendaEmail(); });
});
