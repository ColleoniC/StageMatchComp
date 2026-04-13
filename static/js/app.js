const SCHOOL_IMGS = [
  "https://lh3.googleusercontent.com/gps-cs-s/AHVAweoPIid29XZDiC_rvddu2kvuFfDQ7o9oJ4mijVN4uRswnUXLf_ryjHZlYqGo081shvGmJAiLgmfKjYN5e_I0SWXfiVzFEriTbQA5sm_vzwx5_CfnRPtQwpo5R7B3jcIh9pgIOeCA=w400-h300-k-no"
];

// ── MODALE LOGIN ──────────────────────────────────────────────────────────────

function apriModalLogin() {
  document.getElementById('modalLogin').classList.add('aperto');
  document.body.style.overflow = 'hidden';
}

function chiudiModalLogin() {
  document.getElementById('modalLogin').classList.remove('aperto');
  document.body.style.overflow = '';
  tornaScelta();
}

function gestisciClickOverlay(e) {
  if (e.target === document.getElementById('modalLogin')) chiudiModalLogin();
}

function tornaScelta() {
  _mostraStep('stepScelta');
}

function mostraSceltaStudente() {
  _mostraStep('stepStudente');
  caricaStudenti();
}

function mostraLoginAdmin() {
  _mostraStep('stepAdmin');
  document.getElementById('adminNome').value = '';
  document.getElementById('adminPassword').value = '';
  document.getElementById('adminErrore').style.display = 'none';
}

function _mostraStep(id) {
  ['stepScelta', 'stepStudente', 'stepAdmin'].forEach(s => {
    const el = document.getElementById(s);
    if (el) el.style.display = (s === id) ? '' : 'none';
  });
}

// ── STUDENTI ──────────────────────────────────────────────────────────────────

function caricaStudenti() {
  const lista = document.getElementById('listaStudenti');
  lista.innerHTML = '<div class="lista-studenti-login__loading">Caricamento…</div>';

  fetch('/api/studenti')
    .then(r => r.json())
    .then(studenti => {
      if (!studenti.length) {
        lista.innerHTML = '<div class="lista-studenti-login__vuoto">Nessuno studente registrato nel sistema.</div>';
        return;
      }
      lista.innerHTML = studenti.map(s => `
        <a href="/sso/login?tipo=studente&email=${encodeURIComponent(s.email)}"
           class="lista-studenti-login__item">
          <div class="lista-studenti-login__avatar">${(s.cognome || s.email)[0].toUpperCase()}</div>
          <div class="lista-studenti-login__dati">
            <span class="lista-studenti-login__nome">${s.cognome} ${s.nome}</span>
            <span class="lista-studenti-login__email">${s.email}</span>
          </div>
        </a>
      `).join('');
    })
    .catch(() => {
      lista.innerHTML = '<div class="lista-studenti-login__vuoto">Errore nel caricamento.</div>';
    });
}

// ── ADMIN LOGIN ───────────────────────────────────────────────────────────────

function loginAdmin() {
  const nome     = document.getElementById('adminNome').value.trim();
  const password = document.getElementById('adminPassword').value;
  const errEl    = document.getElementById('adminErrore');
  errEl.style.display = 'none';

  if (!nome || !password) {
    errEl.textContent = 'Compila tutti i campi.';
    errEl.style.display = 'block';
    return;
  }

  const form = new FormData();
  form.append('nome', nome);
  form.append('password', password);

  fetch('/admin/login', { method: 'POST', body: form })
    .then(r => {
      if (r.ok || r.redirected) {
        window.location.href = '/admin';
      } else {
        return r.json().then(d => { throw new Error(d.errore || 'Errore'); });
      }
    })
    .catch(e => {
      errEl.textContent = e.message || 'Credenziali non valide.';
      errEl.style.display = 'block';
    });
}

// ── NAVBAR ────────────────────────────────────────────────────────────────────

function aggiornaNavbar() {
  fetch('/api/sessione')
    .then(r => r.json())
    .then(dati => {
      if (dati.loggato && dati.tipo === 'studente') {
        const linkProfilo = document.getElementById('linkProfilo');
        if (linkProfilo) {
          linkProfilo.classList.remove('disabilitato');
          linkProfilo.href = '/profilo';
        }
      }
    })
    .catch(() => {});
}

// ── HERO ──────────────────────────────────────────────────────────────────────

function inizializzaHero() {
  const heroImg = document.getElementById('heroImg');
  if (heroImg) heroImg.src = SCHOOL_IMGS[0];
}

// ── ENTER KEY per admin form ──────────────────────────────────────────────────

function _adminEnter(e) {
  if (e.key === 'Enter') loginAdmin();
}

// ── INIT ──────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  inizializzaHero();
  aggiornaNavbar();
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') chiudiModalLogin();
  });
  ['adminNome','adminPassword'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('keydown', _adminEnter);
  });
});
