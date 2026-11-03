/**
 * My Biblioteca · Servidor de salas (Google Apps Script)
 * Guarda en una Hoja de cálculo: salas (contraseña con hash), lienzos pintados (mapa de bits) y la música compartida.
 * Pasos: ver LEEME-SALAS.md
 */
const HOJA_SALAS = 'salas';    // sala | sal | hash | creado | musica(JSON) | mseq
const HOJA_ESTADO = 'estado';  // sala | lienzo | total | bits(base64) | v | fecha

function doGet() {
  return ContentService.createTextOutput('My Biblioteca · servidor de salas activo');
}

function doPost(e) {
  let r;
  try {
    r = manejar(JSON.parse(e.postData.contents));
  } catch (err) {
    r = { ok: false, error: 'peticion_invalida' };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function hoja(nombre, encabezado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(nombre);
  if (!h) { h = ss.insertSheet(nombre); h.appendRow(encabezado); }
  return h;
}

function hash(sal, clave) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + clave));
}

function manejar(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sala = String(d.sala || '').trim().toLowerCase();
    const clave = String(d.clave || '');
    if (!/^[a-z0-9ñ _-]{3,30}$/.test(sala) || clave.length < 4 || clave.length > 40) return { ok: false, error: 'datos' };
    const hs = hoja(HOJA_SALAS, ['sala', 'sal', 'hash', 'creado', 'musica', 'mseq']);
    const datos = hs.getDataRange().getValues();
    let fila = -1;
    for (let i = 1; i < datos.length; i++) if (datos[i][0] === sala) { fila = i + 1; break; }

    if (d.a === 'crear') {
      if (fila > 0) return { ok: false, error: 'existe' };
      const sal = Utilities.getUuid();
      hs.appendRow([sala, sal, hash(sal, clave), new Date(), '', 0]);
      return { ok: true };
    }
    if (fila < 0) return { ok: false, error: 'no_existe' };

    // Límite de intentos fallidos: 8 cada 10 minutos por sala
    const cache = CacheService.getScriptCache();
    const kf = 'f:' + sala;
    const fallos = +(cache.get(kf) || 0);
    if (fallos >= 8) return { ok: false, error: 'bloqueada' };
    const row = datos[fila - 1];
    if (hash(row[1], clave) !== row[2]) { cache.put(kf, String(fallos + 1), 600); return { ok: false, error: 'clave' }; }
    if (fallos) cache.remove(kf);

    if (d.a === 'unir') return { ok: true };
    if (d.a === 'sync') return sincronizar(d, sala, hs, fila, row);
    return { ok: false, error: 'accion' };
  } finally {
    lock.releaseLock();
  }
}

function sincronizar(d, sala, hs, fila, row) {
  const he = hoja(HOJA_ESTADO, ['sala', 'lienzo', 'total', 'bits', 'v', 'fecha']);
  const L = String(d.lienzo || '');
  const total = +d.total;
  if (!/^[\w.-]{1,80}@\d{1,6}$/.test(L) || !(total > 0 && total <= 200000)) return { ok: false, error: 'lienzo' };

  const ed = he.getDataRange().getValues();
  let f = -1;
  for (let i = 1; i < ed.length; i++) if (ed[i][0] === sala && ed[i][1] === L) { f = i + 1; break; }
  const nb = Math.ceil(total / 8);
  let bytes, v = 0;
  if (f < 0) {
    bytes = new Array(nb).fill(0);
    he.appendRow([sala, L, total, '', 0, new Date()]);
    f = he.getLastRow();
  } else {
    if (+ed[f - 1][2] !== total) return { ok: false, error: 'total' };
    const b64 = ed[f - 1][3];
    bytes = b64 ? Utilities.base64Decode(b64) : new Array(nb).fill(0);
    v = +ed[f - 1][4] || 0;
  }

  // Pintar solo AÑADE círculos: se unen (OR) los de todas las personas
  let cambio = false;
  const add = Array.isArray(d.add) ? d.add : [];
  for (const i of add) {
    if (Number.isInteger(i) && i >= 0 && i < total) {
      const k = i >> 3, m = 1 << (i & 7), cur = bytes[k] & 255;
      if (!(cur & m)) { bytes[k] = ((cur | m) << 24) >> 24; cambio = true; }
    }
  }
  if (cambio) {
    v++;
    he.getRange(f, 4, 1, 3).setValues([[Utilities.base64Encode(bytes), v, new Date()]]);
  }
  const out = { ok: true, v: v, now: Date.now() };
  if (cambio || (+d.v) !== v) out.bits = Utilities.base64Encode(bytes);

  // Música compartida
  let mseq = +row[5] || 0;
  let mus = row[4] ? JSON.parse(row[4]) : null;
  if (d.musica && typeof d.musica === 'object') {
    mus = { id: String(d.musica.id || '').slice(0, 100), play: !!d.musica.play, pos: Math.max(0, +d.musica.pos || 0), at: Date.now(), by: String(d.cid || '').slice(0, 40) };
    mseq++;
    hs.getRange(fila, 5, 1, 2).setValues([[JSON.stringify(mus), mseq]]);
  }
  out.mseq = mseq;
  if (mus && mseq !== (+d.mseq || 0)) out.musica = mus;
  return out;
}
