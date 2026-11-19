/* Acceso con Microsoft 365 · Tableros de Supervisión COMINVI
   - Inicio de sesión con la cuenta de Microsoft 365 (MSAL).
   - Los datos NO están en esta página: se leen de OneDrive con los permisos de quien inicia sesión.
   - La sesión se cierra tras 10 minutos sin uso.
   - Panel de administrador: alta/baja de usuarios, privilegios y tableros (da y quita permisos en OneDrive). */
(function () {
'use strict';
const CONF = JSON.parse(document.getElementById('conf').textContent);
const A = CONF.acceso || {};
const IDLE = (A.inactividadMin || 10) * 60e3, AVISO = 60e3;
const SC_LEER = ['User.Read', 'Files.Read.All'];
const SC_ADMIN = ['User.Read', 'Files.ReadWrite.All'];
const G = 'https://graph.microsoft.com/v1.0';
const TAB = { mina: 'Tablero Supervisión Mina', mantenimiento: 'Tablero Supervisión Mantenimiento' };
const TABS = ['mina', 'mantenimiento'];
const K = { act: 'tsup.act', abrir: 'tsup.abrirAdm' };
const BASE = (A.urls && A.urls[A.tablero]) || (location.origin + location.pathname);
let pca = null, acct = null, esAdmin = false, cerrando = false, avisoVisible = false;

const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const low = s => String(s || '').trim().toLowerCase();
const correoValido = s => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
const hoyTxt = () => new Date().toLocaleString('es-MX', { timeZone: 'America/Mexico_City', dateStyle: 'short', timeStyle: 'short' });

/* ---------- Estilos ---------- */
const css = `
.gate{position:fixed;inset:0;z-index:50;display:flex;align-items:center;justify-content:center;padding:20px;background:var(--bg)}
.gcard{width:100%;max-width:420px;background:var(--panel);border-radius:14px;border-top:5px solid var(--brand);box-shadow:0 4px 24px rgba(0,0,0,.08);padding:28px 24px 22px;text-align:center}
.gcard img{height:40px;width:auto;border-radius:4px;margin-bottom:12px}
.gcard h1{margin:0;font-family:var(--cond);font-weight:700;font-size:24px;text-transform:uppercase;letter-spacing:.02em}
.gcard .gs{margin:4px 0 18px;font-size:13px;color:var(--ink3)}
.gmsg{margin:0 0 16px;padding:10px 12px;border-radius:10px;font-size:14px;line-height:1.45;text-align:left;background:var(--soft);color:var(--ink2)}
.gmsg:empty{display:none}
.gmsg.warn{background:var(--brand-bg);color:var(--brand-ink);border:1px solid var(--brand-line)}
.gmsg b{color:var(--ink)}
.gbtn{display:flex;align-items:center;justify-content:center;gap:10px;width:100%;min-height:50px;padding:12px 16px;border:0;border-radius:10px;background:var(--brand);color:#fff;font-weight:600;font-size:16px;cursor:pointer}
.gbtn:hover{filter:brightness(1.05)}
.gbtn svg{width:20px;height:20px;flex:none}
.gbtn.alt{background:var(--soft);color:var(--ink);border:1px solid var(--line);margin-top:10px;font-weight:500;font-size:15px;min-height:44px}
.gnote{margin:14px 0 0;font-size:12.5px;color:var(--ink3);line-height:1.45}
.gload{display:flex;align-items:center;justify-content:center;gap:10px;color:var(--ink2);font-size:15px;padding:14px 0}
.spin{width:18px;height:18px;border:2.5px solid var(--line);border-top-color:var(--brand);border-radius:50%;animation:sp 0.8s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
.sesbar{display:flex;align-items:center;flex-wrap:wrap;gap:8px 12px;margin:-8px 0 16px;padding:8px 12px;background:var(--panel);border:1px solid var(--line);border-radius:10px;font-size:13px;color:var(--ink3)}
.sesbar .who{display:flex;align-items:center;gap:8px;min-width:0;flex:1}
.sesbar .who b{color:var(--ink2);font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sesbar .dot{width:8px;height:8px;border-radius:50%;background:var(--ok);flex:none}
.sesbar .lk{display:flex;gap:8px;flex-wrap:wrap}
.sbtn{padding:7px 12px;border-radius:8px;border:1px solid var(--line);background:var(--soft);font-size:13px;font-weight:600;cursor:pointer;color:var(--ink2);text-decoration:none;display:inline-flex;align-items:center;gap:6px}
.sbtn.sadm{background:var(--brand-bg);border-color:var(--brand-line);color:var(--brand-ink)}
.ovl{position:fixed;inset:0;z-index:60;background:rgba(10,12,16,.55);display:flex;align-items:center;justify-content:center;padding:16px}
.warnbox{max-width:380px;width:100%;background:var(--panel);border-radius:14px;padding:22px;text-align:center;border-top:5px solid var(--brand)}
.warnbox h2{margin:0 0 6px;font-family:var(--cond);font-size:22px;text-transform:uppercase}
.warnbox p{margin:0 0 16px;color:var(--ink2)}
.warnbox p b{font-size:20px;color:var(--brand-ink)}
.apanel{position:fixed;inset:0;z-index:55;background:var(--bg);overflow:auto;-webkit-overflow-scrolling:touch}
.apanel .in{max-width:1000px;margin:0 auto;padding:16px var(--pad) 48px}
.apanel .ah{display:flex;align-items:center;gap:12px;padding:6px 0 14px;border-bottom:1px solid var(--line);margin-bottom:16px}
.apanel .ah h2{margin:0;font-family:var(--cond);font-size:24px;text-transform:uppercase;flex:1}
.apanel .ah small{display:block;font-family:var(--font);text-transform:none;font-size:12.5px;color:var(--ink3);font-weight:400}
.abox{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:16px}
.abox h3{margin:0 0 12px;font-family:var(--cond);font-size:18px;text-transform:uppercase;letter-spacing:.03em}
.af{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.af label{display:flex;flex-direction:column;gap:5px;font-size:12.5px;font-weight:600;color:var(--ink3);text-transform:uppercase;letter-spacing:.04em}
.af input[type=email],.af input[type=text],.af select{font:inherit;font-size:15px;padding:10px 12px;border:1px solid var(--hair);border-radius:8px;background:var(--panel);color:var(--ink);text-transform:none;letter-spacing:0;font-weight:400}
.af .full{grid-column:1/-1}
.af .chk{display:flex;gap:10px;flex-wrap:wrap}
.af .chk label{flex-direction:row;align-items:center;gap:8px;padding:9px 12px;border:1px solid var(--hair);border-radius:8px;text-transform:none;letter-spacing:0;font-size:14.5px;color:var(--ink);font-weight:500;cursor:pointer}
.af .chk input{width:18px;height:18px;accent-color:var(--brand)}
.af .hint{font-size:12.5px;color:var(--ink3);font-weight:400;text-transform:none;letter-spacing:0}
.aact{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px}
.abtn{padding:10px 16px;border-radius:9px;border:0;background:var(--brand);color:#fff;font-weight:600;font-size:14.5px;cursor:pointer}
.abtn.alt{background:var(--soft);color:var(--ink);border:1px solid var(--line)}
.abtn.del{background:transparent;color:#c0392b;border:1px solid #e7b3ac}
.abtn:disabled{opacity:.55;cursor:wait}
.amsg{margin:0 0 14px;padding:10px 12px;border-radius:9px;font-size:14px;line-height:1.45}
.amsg:empty{display:none}
.amsg.ok{background:var(--okbg);color:var(--ok)}
.amsg.err{background:#fdecea;color:#a8322a}
.amsg.inf{background:var(--soft);color:var(--ink2)}
.aul{list-style:none;margin:0;padding:0}
.aul li{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:12px 0;border-top:1px solid var(--line)}
.aul li:first-child{border-top:0}
.aul .aup{flex:1;min-width:200px}
.aul .aup b{display:block;font-size:15px}
.aul .aup span{font-size:13px;color:var(--ink3);word-break:break-all}
.atg{display:inline-block;padding:3px 8px;border-radius:6px;font-size:12px;font-weight:600;margin:2px 4px 2px 0;background:var(--soft);color:var(--ink2);border:1px solid var(--line)}
.atg.a{background:var(--brand-bg);color:var(--brand-ink);border-color:var(--brand-line)}
.atg.ok{background:var(--okbg);color:var(--ok);border-color:transparent}
.atg.w{background:#fdecea;color:#a8322a;border-color:transparent}
.aul .bt{display:flex;gap:6px;flex-wrap:wrap}
.aul .bt button,.aul .bt a{padding:7px 10px;border-radius:8px;border:1px solid var(--line);background:var(--soft);font-size:13px;font-weight:600;cursor:pointer;color:var(--ink2);text-decoration:none}
.aul .bt .x{color:#c0392b;border-color:#e7b3ac;background:transparent}
.ahist2{font-size:13px;color:var(--ink2)}
.ahist2 div{padding:6px 0;border-top:1px solid var(--line)}
.ahist2 div:first-child{border-top:0}
.ahist2 small{color:var(--ink3)}
@media (max-width:640px){.af{grid-template-columns:1fr}.apanel .ah h2{font-size:21px}}
@media (prefers-color-scheme:dark){.amsg.err,.atg.w{background:#3a1a17;color:#f3a59b}.abtn.del,.aul .bt .x{color:#f3a59b;border-color:#6b2e28}}
`;
const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);

const MS = '<svg viewBox="0 0 21 21" aria-hidden="true"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>';
const logo = (document.querySelector('.top img') || {}).src || '';

/* ---------- Pantalla de acceso ---------- */
const gate = document.createElement('div');
gate.id = 'gate'; gate.className = 'gate';
gate.innerHTML = `<div class="gcard">${logo ? `<img src="${logo}" alt="COMINVI">` : ''}<h1>${esc(CONF.title)}</h1><p class="gs">${esc(CONF.sub || '')}</p><div id="gbody"></div><p class="gnote">Acceso solo para personal autorizado con cuenta de Microsoft 365 de la empresa.<br>La sesión se cierra tras ${IDLE / 60e3} minutos sin uso.</p></div>`;
document.body.prepend(gate);

function pantalla(tipo, msg) {
  gate.hidden = false;
  const app = $('app'); if (app) app.hidden = true;
  const b = $('gbody');
  if (tipo === 'cargando') { b.innerHTML = `<div class="gload"><span class="spin"></span>${esc(msg || 'Verificando acceso…')}</div>`; return; }
  if (tipo === 'config') { b.innerHTML = `<div class="gmsg warn">El acceso con Microsoft 365 se está configurando. Vuelve a intentarlo más tarde.</div>`; return; }
  if (tipo === 'sinacceso') {
    b.innerHTML = `<div class="gmsg warn">La cuenta <b>${esc(correo())}</b> no tiene acceso a este tablero.<br>Pide acceso al administrador del tablero.</div><button class="gbtn alt" id="gotra" type="button">Usar otra cuenta</button>`;
    $('gotra').onclick = () => salir('otra');
    return;
  }
  b.innerHTML = (msg ? `<div class="gmsg${tipo === 'error' ? ' warn' : ''}">${msg}</div>` : '') + `<button class="gbtn" id="glogin" type="button">${MS}<span>Iniciar sesión con Microsoft 365</span></button>`;
  $('glogin').onclick = entrar;
}

/* ---------- MSAL ---------- */
function correo() {
  if (!acct) return '';
  const c = acct.idTokenClaims || {};
  return low(acct.username || c.preferred_username || c.email || '');
}
async function entrar() {
  const b = $('glogin'); if (b) { b.disabled = true; b.querySelector('span').textContent = 'Abriendo Microsoft 365…'; }
  try {
    await pca.loginRedirect({ scopes: SC_LEER, prompt: A.pedirContrasena === false ? 'select_account' : 'login' });
  } catch (e) { pantalla('login', textoError(e)); }
}
async function token(scopes, estado) {
  try {
    return (await pca.acquireTokenSilent({ scopes, account: acct })).accessToken;
  } catch (e) {
    const code = (e && (e.errorCode || '')) + ' ' + (e && e.message || '');
    if ((window.msal && e instanceof msal.InteractionRequiredAuthError) || /interaction_required|consent_required|login_required|invalid_grant|no_tokens_found/.test(code)) {
      await pca.acquireTokenRedirect({ scopes, account: acct, state: estado || '' });
      return new Promise(() => {});
    }
    throw e;
  }
}
async function graph(method, path, scopes, body, extra) {
  const t = await token(scopes, scopes === SC_ADMIN ? 'admin' : '');
  const h = Object.assign({ Authorization: 'Bearer ' + t }, extra || {});
  let data;
  if (body !== undefined) {
    if (typeof body === 'string') data = body;
    else { data = JSON.stringify(body); h['Content-Type'] = h['Content-Type'] || 'application/json'; }
  }
  const r = await fetch(path.startsWith('http') ? path : G + path, { method, headers: h, body: data, cache: 'no-store' });
  if (!r.ok) {
    const e = new Error('Graph ' + r.status); e.status = r.status;
    try { e.body = await r.json(); e.message = (e.body.error && e.body.error.message) || e.message; } catch (_) {}
    throw e;
  }
  if (r.status === 204) return null;
  const ct = r.headers.get('content-type') || '';
  return ct.includes('json') ? r.json() : r.text();
}
async function leerJson(itemId, scopes) {
  const it = await graph('GET', `/drives/${A.drive}/items/${itemId}`, scopes || SC_LEER);
  const url = it['@microsoft.graph.downloadUrl'];
  if (!url) { const e = new Error('Sin liga de descarga'); e.status = 403; throw e; }
  const r = await fetch(url, { cache: 'no-store' });
  if (!r.ok) { const e = new Error('Descarga ' + r.status); e.status = r.status; throw e; }
  return { data: JSON.parse(await r.text()), eTag: it.eTag, mod: it.lastModifiedDateTime };
}
function textoError(e) {
  const m = String((e && (e.errorMessage || e.message)) || e || '');
  if (/AADSTS50105|AADSTS50020/.test(m)) return 'Tu cuenta no está autorizada para esta aplicación. Pide acceso al administrador.';
  if (/AADSTS65001|consent/i.test(m)) return 'Falta que el administrador de Microsoft 365 de la empresa apruebe la aplicación del tablero.';
  if (/user_cancelled|access_denied|AADSTS65004/.test(m)) return 'Se canceló el inicio de sesión.';
  if (/AADSTS700016|AADSTS90002/.test(m)) return 'La aplicación del tablero aún no está registrada en Microsoft 365.';
  const code = (m.match(/AADSTS\d+/) || [])[0];
  return 'No se pudo iniciar sesión' + (code ? ` (${code})` : '') + '. Intenta de nuevo.';
}

/* ---------- Sesión: 10 minutos sin uso ---------- */
function marcar() { try { sessionStorage.setItem(K.act, String(Date.now())); } catch (_) {} }
function ultima() { try { return +sessionStorage.getItem(K.act) || 0; } catch (_) { return 0; } }
let ultMarca = 0;
function actividad() {
  if (cerrando || !acct) return;
  const n = Date.now();
  if (n - ultima() >= IDLE) { expirar(); return; }
  if (avisoVisible) { ocultarAviso(); }
  if (n - ultMarca < 1500) return;
  ultMarca = n; marcar();
}
function revisar() {
  if (cerrando || !acct) return;
  const falta = IDLE - (Date.now() - ultima());
  if (falta <= 0) { expirar(); return; }
  if (falta <= AVISO) mostrarAviso(Math.ceil(falta / 1000));
  else if (avisoVisible) ocultarAviso();
}
function mostrarAviso(seg) {
  let o = $('aviso');
  if (!o) {
    o = document.createElement('div'); o.id = 'aviso'; o.className = 'ovl';
    o.innerHTML = `<div class="warnbox" role="alertdialog" aria-live="assertive"><h2>¿Sigues ahí?</h2><p>Por seguridad, la sesión se cerrará en <b id="aseg"></b> segundos por inactividad.</p><button class="gbtn" id="aseguir" type="button">Seguir en el tablero</button><button class="gbtn alt" id="asalir" type="button">Cerrar sesión</button></div>`;
    document.body.appendChild(o);
    $('aseguir').onclick = () => { marcar(); ocultarAviso(); };
    $('asalir').onclick = () => salir('salir');
  }
  avisoVisible = true; o.hidden = false; $('aseg').textContent = seg;
}
function ocultarAviso() { avisoVisible = false; const o = $('aviso'); if (o) o.hidden = true; }
function vigilar() {
  marcar();
  ['pointerdown', 'keydown', 'wheel', 'touchstart', 'mousemove'].forEach(ev => document.addEventListener(ev, actividad, { passive: true, capture: true }));
  window.addEventListener('scroll', actividad, { passive: true, capture: true });
  document.addEventListener('visibilitychange', revisar);
  window.addEventListener('pageshow', revisar);
  window.addEventListener('focus', revisar);
  setInterval(revisar, 1000);
}
async function cerrarLocal() {
  try { if (pca && acct) await pca.logoutRedirect({ account: acct, onRedirectNavigate: () => false }); } catch (_) {}
  try { sessionStorage.removeItem(K.act); sessionStorage.removeItem(K.abrir); } catch (_) {}
}
function borrarPantalla() {
  // Quita de inmediato los datos de la pantalla
  ['app', 'adm', 'aviso'].forEach(id => { const el = $(id); if (el) el.remove(); });
  pantalla('cargando', 'Cerrando sesión…');
}
async function expirar() {
  if (cerrando) return; cerrando = true;
  borrarPantalla();
  await cerrarLocal();
  location.replace(BASE + '?fin=inactividad');
}
async function salir(motivo) {
  if (cerrando) return; cerrando = true;
  borrarPantalla();
  await cerrarLocal();
  location.replace(BASE + '?fin=' + (motivo || 'salir'));
}

/* ---------- Mostrar el tablero ---------- */
function barraSesion() {
  const top = document.querySelector('.top'); if (!top || $('sesbar')) return;
  const otro = TABS.find(t => t !== A.tablero);
  const bar = document.createElement('div'); bar.id = 'sesbar'; bar.className = 'sesbar';
  bar.innerHTML = `<div class="who"><span class="dot"></span><span>Sesión de <b>${esc((acct.name || correo()))}</b></span></div><div class="lk">${A.urls && A.urls[otro] ? `<a class="sbtn" href="${esc(A.urls[otro])}">Ir a ${esc(otro === 'mina' ? 'Mina' : 'Mantenimiento')}</a>` : ''}<button class="sbtn sadm" id="admbtn" type="button" hidden>Administrar accesos</button><button class="sbtn" id="salir" type="button">Cerrar sesión</button></div>`;
  top.after(bar);
  $('salir').onclick = () => salir('salir');
  $('admbtn').onclick = abrirAdmin;
}
async function cargar() {
  pantalla('cargando', 'Cargando tablero…');
  let r;
  try { r = await leerJson(A.items[A.tablero]); }
  catch (e) {
    if (e.status === 403 || e.status === 404 || e.status === 401) { pantalla('sinacceso'); return; }
    pantalla('error', 'No se pudo cargar el tablero. Revisa tu conexión e intenta de nuevo.'); return;
  }
  gate.hidden = true;
  const app = $('app'); app.hidden = false;
  barraSesion();
  try { window.iniciarTablero(r.data); } catch (e) { console.error(e); }
  // ¿Es administrador? Solo los administradores pueden leer la lista de accesos.
  try {
    const ac = await leerJson(A.items.accesos);
    const u = (ac.data.usuarios || []).find(x => low(x.correo) === correo());
    if (u && u.rol === 'admin' && u.activo !== false) {
      esAdmin = true; $('admbtn').hidden = false;
      if (sessionStorage.getItem(K.abrir)) { sessionStorage.removeItem(K.abrir); abrirAdmin(); }
    }
  } catch (_) {}
}

/* ---------- Panel de administrador ---------- */
const IT = () => ({ carpeta: A.items.carpeta, mina: A.items.mina, mantenimiento: A.items.mantenimiento });
let ADM = null; // {lista, eTag, perms:{carpeta:[],mina:[],mantenimiento:[]}, dueno}

function correosDe(p) {
  const out = new Set();
  const add = x => { if (x) out.add(low(x)); };
  const fromId = g => {
    if (!g) return;
    if (g.user) add(g.user.email);
    if (g.siteUser) { add(g.siteUser.email); if (g.siteUser.loginName) add(String(g.siteUser.loginName).split('|').pop()); }
  };
  fromId(p.grantedToV2); fromId(p.grantedTo);
  (p.grantedToIdentitiesV2 || []).forEach(fromId); (p.grantedToIdentities || []).forEach(fromId);
  if (p.invitation) add(p.invitation.email);
  out.delete(''); return [...out].filter(correoValido);
}
const nombreDe = p => { const g = p.grantedToV2 || p.grantedTo || {}; return (g.user && g.user.displayName) || (g.siteUser && g.siteUser.displayName) || ''; };
const directos = (lista, email) => lista.filter(p => !p.inheritedFrom && !(p.roles || []).includes('owner') && correosDe(p).includes(email));

async function cargarAdmin() {
  const ac = await leerJson(A.items.accesos, SC_ADMIN);
  const perms = {};
  for (const [k, id] of Object.entries(IT())) {
    perms[k] = ((await graph('GET', `/drives/${A.drive}/items/${id}/permissions`, SC_ADMIN)) || {}).value || [];
  }
  const ownerP = perms.carpeta.find(p => (p.roles || []).includes('owner'));
  const lista = ac.data && Array.isArray(ac.data.usuarios) ? ac.data : { version: 1, usuarios: [] };
  if (!Array.isArray(lista.historial)) lista.historial = [];
  ADM = { lista, eTag: ac.eTag, perms, dueno: (ownerP && correosDe(ownerP)[0]) || low(A.dueno) };
}
function deseado(u) {
  const d = { carpeta: null, mina: null, mantenimiento: null };
  if (!u || u.activo === false) return d;
  if (u.rol === 'admin') { d.carpeta = 'write'; return d; }
  (u.tableros || []).forEach(t => { if (t in d) d[t] = 'read'; });
  return d;
}
function estadoReal(email) {
  // ¿Qué puede ver hoy en OneDrive? (directo o heredado de la carpeta)
  if (email === ADM.dueno) return { mina: true, mantenimiento: true, admin: true, dueno: true };
  const tiene = k => ADM.perms[k].some(p => correosDe(p).includes(email));
  const editor = ADM.perms.carpeta.some(p => correosDe(p).includes(email) && (p.roles || []).some(r => r === 'write' || r === 'owner'));
  return { mina: tiene('mina'), mantenimiento: tiene('mantenimiento'), admin: editor };
}
async function aplicar(u) {
  const email = low(u.correo);
  if (email === ADM.dueno) return; // el dueño siempre tiene acceso
  const want = deseado(u);
  for (const k of ['carpeta', 'mina', 'mantenimiento']) {
    const id = IT()[k];
    const tiene = directos(ADM.perms[k], email);
    const rol = want[k];
    const correcto = rol && tiene.some(p => (p.roles || []).includes(rol));
    for (const p of tiene) {
      if (correcto && (p.roles || []).includes(rol)) continue;
      await graph('DELETE', `/drives/${A.drive}/items/${id}/permissions/${p.id}`, SC_ADMIN);
    }
    if (rol && !correcto) {
      await graph('POST', `/drives/${A.drive}/items/${id}/invite`, SC_ADMIN, {
        requireSignIn: true, sendInvitation: false, roles: [rol], recipients: [{ email }]
      });
    }
  }
}
async function guardarLista(accion, email) {
  const L = ADM.lista;
  L.actualizado = new Date().toISOString();
  L.historial.unshift({ fecha: L.actualizado, por: correo(), accion, correo: email });
  L.historial = L.historial.slice(0, 200);
  const r = await graph('PUT', `/drives/${A.drive}/items/${A.items.accesos}/content`, SC_ADMIN, JSON.stringify(L, null, 1), { 'Content-Type': 'application/json', 'If-Match': ADM.eTag });
  ADM.eTag = r && r.eTag || ADM.eTag;
}

function abrirAdmin() {
  if (!esAdmin) return;
  let o = $('adm');
  if (!o) {
    o = document.createElement('div'); o.id = 'adm'; o.className = 'apanel';
    o.innerHTML = `<div class="in">
<div class="ah"><h2>Administrar accesos<small>Usuarios, privilegios y correos de Microsoft 365 · Mina y Mantenimiento</small></h2><button class="abtn alt" id="acerrar" type="button">Volver al tablero</button></div>
<div id="amsg" class="amsg"></div>
<section class="abox"><h3 id="aftit">Agregar usuario</h3>
<form class="af" id="aform" autocomplete="off">
 <label>Correo de Microsoft 365<input type="email" id="fcorreo" required placeholder="nombre@cominvi.com" inputmode="email" autocapitalize="off"></label>
 <label>Nombre<input type="text" id="fnombre" placeholder="Nombre y apellido"></label>
 <label>Privilegio<select id="frol"><option value="sup">Supervisor · solo ve tableros</option><option value="admin">Administrador · ve todo y administra accesos</option></select></label>
 <label>Tableros que puede ver<div class="chk"><label><input type="checkbox" id="tmina" value="mina">Mina</label><label><input type="checkbox" id="tmantenimiento" value="mantenimiento">Mantenimiento</label></div></label>
 <div class="full hint" id="fhint">El supervisor entra con su cuenta de Microsoft 365. No se envía ningún correo automático: usa "Enviar liga" para avisarle.</div>
 <div class="full aact"><button class="abtn" id="fguardar" type="submit">Guardar y dar acceso</button><button class="abtn alt" id="flimpiar" type="button">Limpiar</button></div>
</form></section>
<section class="abox"><h3>Usuarios con acceso <span id="acnt" class="atg"></span></h3><ul class="aul" id="alist"></ul></section>
<section class="abox" id="aotros" hidden><h3>Otras personas con acceso en OneDrive</h3><p class="hint" style="margin:0 0 8px;font-size:13px;color:var(--ink3)">Tienen permiso en los archivos pero no están en la lista. Si no deben ver los tableros, quítales el acceso.</p><ul class="aul" id="aolist"></ul></section>
<section class="abox"><h3>Historial</h3><div class="ahist2" id="ahist"></div></section>
</div>`;
    document.body.appendChild(o);
    $('acerrar').onclick = () => { o.hidden = true; $('app').hidden = false; window.scrollTo(0, 0); };
    $('frol').onchange = rolCambio;
    $('flimpiar').onclick = limpiarForm;
    $('aform').onsubmit = e => { e.preventDefault(); guardarUsuario(); };
    $('alist').onclick = accionLista;
    $('aolist').onclick = accionOtros;
  }
  o.hidden = false; $('app').hidden = true; window.scrollTo(0, 0);
  try { sessionStorage.setItem(K.abrir, '1'); } catch (_) {}
  refrescarAdmin().finally(() => { try { sessionStorage.removeItem(K.abrir); } catch (_) {} });
}
function msgAdm(t, cls) { const m = $('amsg'); m.className = 'amsg ' + (cls || 'inf'); m.innerHTML = t || ''; if (t) m.scrollIntoView({ block: 'nearest' }); }
function ocupado(on) { ['fguardar'].forEach(id => { const b = $(id); if (b) b.disabled = on; }); document.querySelectorAll('#alist button,#aolist button').forEach(b => b.disabled = on); }
async function refrescarAdmin(texto, cls) {
  msgAdm('<span class="spin" style="display:inline-block;vertical-align:-3px;margin-right:8px;width:14px;height:14px"></span>Leyendo accesos…', 'inf');
  try { await cargarAdmin(); }
  catch (e) { msgAdm('No se pudo leer la lista de accesos: ' + esc(e.message || e), 'err'); return; }
  msgAdm(texto || '', cls);
  pintarAdmin();
}
function rolCambio() {
  const adm = $('frol').value === 'admin';
  ['tmina', 'tmantenimiento'].forEach(id => { $(id).disabled = adm; if (adm) $(id).checked = true; });
}
function limpiarForm() {
  $('aform').reset(); $('fcorreo').readOnly = false; $('aftit').textContent = 'Agregar usuario';
  $('fguardar').textContent = 'Guardar y dar acceso'; rolCambio();
}
function ligaCorreo(u) {
  const ts = u.rol === 'admin' ? TABS : (u.tableros || []);
  const ligas = ts.map(t => `${TAB[t]}: ${A.urls[t]}`).join('\n');
  const body = `Hola ${u.nombre || ''},\n\nYa tienes acceso ${ts.length > 1 ? 'a los tableros' : 'al tablero'} de supervisión:\n\n${ligas}\n\nEntra con tu cuenta de Microsoft 365 (${u.correo}).\nPor seguridad, la sesión se cierra tras ${IDLE / 60e3} minutos sin uso.\n\nSaludos.`;
  return `mailto:${encodeURIComponent(u.correo)}?subject=${encodeURIComponent('Acceso a ' + (ts.length > 1 ? 'los Tableros de Supervisión' : TAB[ts[0]] || 'Tablero de Supervisión'))}&body=${encodeURIComponent(body)}`;
}
function pintarAdmin() {
  const L = ADM.lista.usuarios;
  // El dueño de la carpeta siempre aparece como administrador
  if (ADM.dueno && !L.some(u => low(u.correo) === ADM.dueno)) L.unshift({ correo: ADM.dueno, nombre: '', rol: 'admin', tableros: TABS, activo: true, alta: '' });
  $('acnt').textContent = L.length;
  $('alist').innerHTML = L.map((u, i) => {
    const e = low(u.correo), r = estadoReal(e), d = deseado(u);
    const ts = u.rol === 'admin' ? TABS : (u.tableros || []);
    const okMina = !d.mina && !d.carpeta ? !r.mina : r.mina, okMan = !d.mantenimiento && !d.carpeta ? !r.mantenimiento : r.mantenimiento;
    const okAdm = u.rol === 'admin' ? r.admin : true;
    const sync = okMina && okMan && okAdm;
    return `<li><div class="aup"><b>${esc(u.nombre || e)}</b><span>${esc(e)}</span></div><div>${r.dueno ? '<span class="atg a">Dueño</span>' : ''}<span class="atg${u.rol === 'admin' ? ' a' : ''}">${u.rol === 'admin' ? 'Administrador' : 'Supervisor'}</span>${ts.map(t => `<span class="atg">${t === 'mina' ? 'Mina' : 'Mantenimiento'}</span>`).join('')}${sync ? '<span class="atg ok">✓ Acceso activo</span>' : '<span class="atg w">⚠ Permiso pendiente</span>'}</div><div class="bt">${r.dueno ? '' : `<button type="button" data-a="ed" data-i="${i}">Editar</button>`}<a href="${esc(ligaCorreo(u))}">Enviar liga</a>${!sync ? `<button type="button" data-a="sy" data-i="${i}">Reparar</button>` : ''}${r.dueno || e === correo() ? '' : `<button type="button" class="x" data-a="del" data-i="${i}">Quitar acceso</button>`}</div></li>`;
  }).join('') || '<li>No hay usuarios.</li>';
  // Personas con permiso en OneDrive que no están en la lista
  const enLista = new Set(L.map(u => low(u.correo)));
  const otros = new Map();
  for (const k of ['carpeta', 'mina', 'mantenimiento']) {
    ADM.perms[k].forEach(p => {
      if (p.inheritedFrom || (p.roles || []).includes('owner')) return;
      if (p.link && !correosDe(p).length) { const key = 'liga:' + p.id; otros.set(key, { key, nombre: 'Liga para compartir (' + (p.link.scope || '') + ')', k: [k], perm: [[k, p.id]] }); return; }
      correosDe(p).forEach(e => { if (enLista.has(e)) return; const o = otros.get(e) || { key: e, nombre: nombreDe(p), correo: e, k: [], perm: [] }; o.k.push(k); o.perm.push([k, p.id]); otros.set(e, o); });
    });
  }
  $('aotros').hidden = !otros.size;
  ADM.otros = [...otros.values()];
  $('aolist').innerHTML = ADM.otros.map((o, i) => `<li><div class="aup"><b>${esc(o.nombre || o.correo)}</b><span>${esc(o.correo || '')}</span></div><div>${o.k.map(k => `<span class="atg">${k === 'carpeta' ? 'Carpeta completa' : k === 'mina' ? 'Mina' : 'Mantenimiento'}</span>`).join('')}</div><div class="bt"><button type="button" class="x" data-a="qo" data-i="${i}">Quitar acceso</button></div></li>`).join('');
  const H = ADM.lista.historial || [];
  $('ahist').innerHTML = H.slice(0, 25).map(h => `<div>${esc(h.accion)} · <b>${esc(h.correo || '')}</b><br><small>${esc(new Date(h.fecha).toLocaleString('es-MX', { timeZone: 'America/Mexico_City', dateStyle: 'medium', timeStyle: 'short' }))} · por ${esc(h.por || '')}</small></div>`).join('') || '<div><small>Sin movimientos todavía.</small></div>';
}
async function guardarUsuario() {
  const email = low($('fcorreo').value);
  if (!correoValido(email)) { msgAdm('Escribe un correo válido.', 'err'); return; }
  const dom = email.split('@')[1], doms = A.dominios || [];
  if (doms.length && !doms.includes(dom) && !confirm(`El correo ${email} no es de ${doms.join(' / ')}. Solo funciona si esa persona tiene cuenta en el Microsoft 365 de la empresa. ¿Continuar?`)) return;
  const rol = $('frol').value === 'admin' ? 'admin' : 'sup';
  const tableros = rol === 'admin' ? TABS.slice() : TABS.filter(t => $('t' + t).checked);
  if (!tableros.length) { msgAdm('Selecciona al menos un tablero.', 'err'); return; }
  ocupado(true); msgAdm('Guardando y dando permisos en OneDrive…', 'inf');
  try {
    await cargarAdmin();
    const L = ADM.lista.usuarios;
    let u = L.find(x => low(x.correo) === email), nuevo = !u;
    if (email === ADM.dueno) { msgAdm('El dueño de los archivos siempre tiene acceso completo.', 'inf'); ocupado(false); return; }
    if (nuevo) { u = { correo: email, alta: new Date().toISOString().slice(0, 10), por: correo() }; L.push(u); }
    if (!nuevo && u.rol === 'admin' && rol !== 'admin' && email === correo()) { msgAdm('No puedes quitarte tu propio privilegio de administrador.', 'err'); ocupado(false); return; }
    const antes = JSON.stringify([u.rol, u.tableros]);
    u.nombre = $('fnombre').value.trim() || u.nombre || ''; u.rol = rol; u.tableros = tableros; u.activo = true;
    await aplicar(u);
    await guardarLista(nuevo ? `Alta (${rol === 'admin' ? 'Administrador' : 'Supervisor: ' + tableros.join(', ')})` : (antes === JSON.stringify([rol, tableros]) ? 'Datos actualizados' : `Cambio a ${rol === 'admin' ? 'Administrador' : 'Supervisor: ' + tableros.join(', ')}`), email);
    limpiarForm();
    await refrescarAdmin(`Listo: <b>${esc(email)}</b> ${nuevo ? 'ya tiene acceso' : 'quedó actualizado'}. Usa "Enviar liga" para avisarle.`, 'ok');
  } catch (e) {
    await refrescarAdmin('No se pudo completar: ' + esc(explicaGraph(e)), 'err');
  } finally { ocupado(false); }
}
function explicaGraph(e) {
  const m = String(e && e.message || e || '');
  if (e && e.status === 412) return 'otra persona modificó la lista al mismo tiempo. Intenta de nuevo.';
  if (/invalidRequest|not found|notFound|does not exist|resolve/i.test(m) && e.status !== 403) return 'no se encontró esa cuenta en Microsoft 365. Revisa el correo. (' + m + ')';
  if (e && e.status === 403) return 'no tienes permiso para cambiar accesos en OneDrive. (' + m + ')';
  return m;
}
async function accionLista(ev) {
  const b = ev.target.closest('button'); if (!b) return;
  const u = ADM.lista.usuarios[+b.dataset.i]; if (!u) return;
  if (b.dataset.a === 'ed') {
    $('fcorreo').value = u.correo; $('fcorreo').readOnly = true; $('fnombre').value = u.nombre || '';
    $('frol').value = u.rol === 'admin' ? 'admin' : 'sup';
    TABS.forEach(t => { $('t' + t).checked = (u.tableros || []).includes(t); }); rolCambio();
    $('aftit').textContent = 'Editar usuario'; $('fguardar').textContent = 'Guardar cambios';
    $('aform').scrollIntoView({ behavior: 'smooth', block: 'start' }); return;
  }
  const email = low(u.correo);
  if (b.dataset.a === 'sy') {
    ocupado(true); msgAdm('Reparando permisos…', 'inf');
    try { await cargarAdmin(); const v = ADM.lista.usuarios.find(x => low(x.correo) === email); if (v) { await aplicar(v); await guardarLista('Permisos reparados', email); } await refrescarAdmin('Permisos de <b>' + esc(email) + '</b> reparados.', 'ok'); }
    catch (e) { await refrescarAdmin('No se pudo reparar: ' + esc(explicaGraph(e)), 'err'); } finally { ocupado(false); }
    return;
  }
  if (b.dataset.a === 'del') {
    if (!confirm(`¿Quitar el acceso de ${u.nombre || email}? Ya no podrá ver los tableros.`)) return;
    ocupado(true); msgAdm('Quitando acceso…', 'inf');
    try {
      await cargarAdmin();
      const L = ADM.lista.usuarios, k = L.findIndex(x => low(x.correo) === email);
      await aplicar({ correo: email, activo: false });
      if (k >= 0) L.splice(k, 1);
      await guardarLista('Baja (acceso retirado)', email);
      await refrescarAdmin('Se quitó el acceso de <b>' + esc(email) + '</b>.', 'ok');
    } catch (e) { await refrescarAdmin('No se pudo quitar: ' + esc(explicaGraph(e)), 'err'); } finally { ocupado(false); }
  }
}
async function accionOtros(ev) {
  const b = ev.target.closest('button'); if (!b) return;
  const o = ADM.otros[+b.dataset.i]; if (!o) return;
  if (!confirm(`¿Quitar el acceso en OneDrive de ${o.nombre || o.correo}?`)) return;
  ocupado(true); msgAdm('Quitando acceso…', 'inf');
  try {
    for (const [k, pid] of o.perm) await graph('DELETE', `/drives/${A.drive}/items/${IT()[k]}/permissions/${pid}`, SC_ADMIN);
    await cargarAdmin(); await guardarLista('Permiso externo retirado', o.correo || o.nombre);
    await refrescarAdmin('Acceso retirado.', 'ok');
  } catch (e) { await refrescarAdmin('No se pudo quitar: ' + esc(explicaGraph(e)), 'err'); } finally { ocupado(false); }
}

/* ---------- Inicio ---------- */
async function main() {
  const q = new URLSearchParams(location.search);
  const fin = q.get('fin');
  if (fin) history.replaceState(null, '', BASE + location.hash);
  if (!A.clientId) { pantalla('config'); return; }
  if (!window.msal) { pantalla('error', 'No se pudo cargar el inicio de sesión de Microsoft. Revisa tu conexión.'); return; }
  pantalla('cargando');
  try {
    pca = new msal.PublicClientApplication({
      auth: { clientId: A.clientId, authority: 'https://login.microsoftonline.com/' + A.tenant, redirectUri: BASE, postLogoutRedirectUri: BASE, navigateToLoginRequestUrl: true },
      cache: { cacheLocation: 'sessionStorage' },
      system: { allowRedirectInIframe: false }
    });
    await pca.initialize();
  } catch (e) { pantalla('error', 'No se pudo iniciar el acceso con Microsoft 365.'); return; }
  let res = null;
  try { res = await pca.handleRedirectPromise(); }
  catch (e) { pantalla('login', textoError(e)); return; }
  if (res && res.account) {
    acct = res.account; marcar();
    if (res.state === 'admin') { try { sessionStorage.setItem(K.abrir, '1'); } catch (_) {} }
  } else {
    const all = pca.getAllAccounts();
    if (all.length && Date.now() - ultima() < IDLE) acct = all[0];
    else {
      if (all.length) { acct = all[0]; await cerrarLocal(); acct = null; }
      const msg = fin === 'inactividad' ? `Tu sesión se cerró por inactividad (${IDLE / 60e3} minutos sin uso). Vuelve a iniciar sesión.`
        : fin === 'salir' ? 'Cerraste sesión.' : fin === 'otra' ? 'Inicia sesión con otra cuenta.' : '';
      pantalla('login', msg); return;
    }
  }
  pca.setActiveAccount(acct);
  vigilar();
  await cargar();
}
main();
})();
