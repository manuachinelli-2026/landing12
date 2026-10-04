/**
 * Casamiento 12·12 — Apps Script de la reconfirmación.
 *
 * Va pegado en el Sheet "Casamiento 12·12 — Reconfirmación" (Extensiones → Apps Script).
 * doPost → reconfirmacion.html. Crea y llena solas estas pestañas:
 *   · Reconfirmación  una fila por invitado (si vuelve a confirmar, se actualiza su fila)
 *   · Visitas         una fila por visita: historias vistas, tiempo, clicks, hasta dónde llegó
 *   · Eventos         una fila por cada cosa que pasó (vio historia, tocó algo, llegó a una pregunta)
 *   · Resumen         números listos (se arma sola la primera vez)
 *
 * Columnas inteligentes: si la página manda un dato nuevo (ej. una pregunta nueva),
 * se agrega la columna al final automáticamente. No hay que tocar este script.
 */

const TABS = { rsvp: 'Reconfirmación', visitas: 'Visitas', eventos: 'Eventos', resumen: 'Resumen' };

/* ───────────── chequeo: abrir la URL /exec en el navegador muestra {"ok":true} ───────────── */
function doGet() {
  return json({ ok: true, sheet: SpreadsheetApp.getActiveSpreadsheet().getName() });
}

/* ───────────── reconfirmación + tracking ───────────── */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const s = data.session || {};

    if (data.events && data.events.length) {
      appendRows(tab(ss, TABS.eventos), data.events.map(ev => Object.assign(
        { recibido: now, sesion: s.sesion, visitante: s.visitante, grupo: s.grupo, test: s.test }, ev)));
    }
    if (s.sesion) {
      upsert(tab(ss, TABS.visitas), 'sesion', Object.assign({ actualizado: now }, s));
    }
    if (data.type === 'rsvp' && data.respuesta) {
      const r = data.respuesta;
      const invitado = normalize(r.nombre + ' ' + r.apellido);
      const sh = tab(ss, TABS.rsvp);
      const prev = findRow(sh, 'invitado', invitado);
      upsert(sh, 'invitado', Object.assign(
        { invitado: invitado, actualizado: now, envios: prev ? (Number(prev.envios) || 1) + 1 : 1 },
        prev ? {} : { primera_vez: now },
        r));
    }
    if (!ss.getSheetByName(TABS.resumen)) buildResumen(ss);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ───────────── helpers de hoja ───────────── */
function tab(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name, ss.getNumSheets());
}

// agrega al encabezado las columnas que falten y devuelve el encabezado completo
function ensureHeaders(sh, keys) {
  const lastCol = sh.getLastColumn();
  const headers = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  const missing = keys.filter(k => headers.indexOf(k) === -1);
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    if (!sh.getFrozenRows()) sh.setFrozenRows(1);
  }
  return headers.concat(missing);
}

function appendRows(sh, objs) {
  const keys = [];
  objs.forEach(o => Object.keys(o).forEach(k => { if (keys.indexOf(k) === -1) keys.push(k); }));
  const headers = ensureHeaders(sh, keys);
  const rows = objs.map(o => headers.map(h => clean(o[h])));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function findRow(sh, keyName, keyValue) {
  if (sh.getLastRow() < 2) return null;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = headers.indexOf(keyName);
  if (col === -1) return null;
  const cell = sh.getRange(2, col + 1, sh.getLastRow() - 1, 1)
    .createTextFinder(String(keyValue)).matchEntireCell(true).findNext();
  if (!cell) return null;
  const values = sh.getRange(cell.getRow(), 1, 1, headers.length).getValues()[0];
  const obj = { _row: cell.getRow() };
  headers.forEach((h, i) => obj[h] = values[i]);
  return obj;
}

// actualiza la fila con esa clave (o la crea), sin borrar columnas que no vinieron
function upsert(sh, keyName, obj) {
  const headers = ensureHeaders(sh, Object.keys(obj));
  const prev = findRow(sh, keyName, obj[keyName]);
  const row = headers.map(h => (h in obj) ? clean(obj[h]) : (prev ? prev[h] : ''));
  const r = prev ? prev._row : sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, headers.length).setValues([row]);
}

function clean(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string') {
    if (/^\d{4}-\d{2}-\d{2}T[\d:.]+Z$/.test(v)) return new Date(v);   // fechas ISO → fecha real
    if (/^[=+\-@]/.test(v)) return "'" + v;                            // evita fórmulas inyectadas
  }
  return v;
}

function normalize(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────────── pestaña Resumen ───────────── */
// Se arma sola. Para rearmarla: borrar la pestaña "Resumen" o correr buildResumen() a mano.
function buildResumen(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(TABS.resumen);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(TABS.resumen, 0);
  // saca la "Hoja 1" vacía que trae el Sheet nuevo
  ss.getSheets().forEach(x => { if (Object.values(TABS).indexOf(x.getName()) === -1 && x.getLastRow() === 0) ss.deleteSheet(x); });

  // columna por nombre de encabezado, así no importa en qué orden estén
  const col = (t, name) => `INDEX('${t}'!A2:ZZ,0,MATCH("${name}",'${t}'!A1:ZZ1,0))`;
  const R = n => col(TABS.rsvp, n), V = n => col(TABS.visitas, n), E = n => col(TABS.eventos, n);
  const real = t => `${col(t, 'test')},"<>Sí"`;    // criterio extra: ignora filas de prueba

  const rows = [
    ['CONFIRMACIONES', ''],
    ['Respondieron', `=IFERROR(COUNTIFS(${R('invitado')},"<>",${real(TABS.rsvp)}),0)`],
    ['Vienen', `=IFERROR(COUNTIFS(${R('asistencia')},"Sí, voy",${real(TABS.rsvp)}),0)`],
    ['No vienen', `=IFERROR(COUNTIFS(${R('asistencia')},"No puedo",${real(TABS.rsvp)}),0)`],
    ['Viernes', `=IFERROR(COUNTIFS(${R('dias')},"*Viernes*",${real(TABS.rsvp)}),0)`],
    ['Sábado', `=IFERROR(COUNTIFS(${R('dias')},"*Sábado*",${real(TABS.rsvp)}),0)`],
    ['Domingo', `=IFERROR(COUNTIFS(${R('dias')},"*Domingo*",${real(TABS.rsvp)}),0)`],
    ['', ''],
    ['LOGÍSTICA', ''],
    ['Bus: sí', `=IFERROR(COUNTIFS(${R('bus')},"Sí",${real(TABS.rsvp)}),0)`],
    ['Bus: por su cuenta', `=IFERROR(COUNTIFS(${R('bus')},"Voy por mi cuenta",${real(TABS.rsvp)}),0)`],
    ['Hospedaje: ya tiene', `=IFERROR(COUNTIFS(${R('hospedaje')},"Sí, ya tengo",${real(TABS.rsvp)}),0)`],
    ['Hospedaje: todavía no', `=IFERROR(COUNTIFS(${R('hospedaje')},"Todavía no",${real(TABS.rsvp)}),0)`],
    ['Hospedaje: no necesita', `=IFERROR(COUNTIFS(${R('hospedaje')},"No necesito",${real(TABS.rsvp)}),0)`],
    ['Vegetariano', `=IFERROR(COUNTIFS(${R('menu')},"Vegetariano",${real(TABS.rsvp)}),0)`],
    ['Vegano', `=IFERROR(COUNTIFS(${R('menu')},"Vegano",${real(TABS.rsvp)}),0)`],
    ['Celíaco', `=IFERROR(COUNTIFS(${R('menu')},"Celíaco",${real(TABS.rsvp)}),0)`],
    ['', ''],
    ['VISITAS', ''],
    ['Personas que abrieron el link', `=IFERROR(COUNTUNIQUEIFS(${V('visitante')},${real(TABS.visitas)}),0)`],
    ['Visitas totales', `=IFERROR(COUNTIFS(${V('sesion')},"<>",${real(TABS.visitas)}),0)`],
    ['Visitas que confirmaron', `=IFERROR(COUNTIFS(${V('confirmo')},"Sí",${real(TABS.visitas)}),0)`],
    ['Historias vistas (promedio)', `=IFERROR(ROUND(AVERAGEIFS(${V('historias_vistas')},${real(TABS.visitas)}),1),0)`],
    ['Segundos por visita (promedio)', `=IFERROR(ROUND(AVERAGEIFS(${V('segundos')},${real(TABS.visitas)}),0),0)`],
    ['Desde Instagram / WhatsApp / navegador', `=IFERROR(COUNTIFS(${V('app')},"Instagram",${real(TABS.visitas)})&" / "&COUNTIFS(${V('app')},"WhatsApp",${real(TABS.visitas)})&" / "&COUNTIFS(${V('app')},"Navegador",${real(TABS.visitas)}),"")`],
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);

  // alcance por historia (cuántas visitas llegaron a cada una)
  const top = 1;
  sh.getRange(top, 4, 1, 3).setValues([['HISTORIA', '', 'VISITAS QUE LA VIERON']]);
  const reach = [];
  for (let n = 1; n <= 25; n++) {
    const r = top + n;
    reach.push([
      n,
      `=IFERROR(INDEX(FILTER(${E('historia')},${E('historia_n')}=D${r},${E('evento')}="historia"),1),"")`,
      `=IF(E${r}="","",IFERROR(COUNTUNIQUEIFS(${E('sesion')},${E('evento')},"historia",${E('historia_n')},D${r},${real(TABS.eventos)}),0))`,
    ]);
  }
  sh.getRange(top + 1, 4, reach.length, 3).setValues(reach);

  // lo más tocado
  sh.getRange(top, 8).setValue('MÁS TOCADO');
  sh.getRange(top + 1, 8).setFormula(
    `=IFERROR(QUERY({${E('evento')},${E('detalle')},${E('test')}},"select Col2, count(Col1) where Col1='click' and Col3<>'Sí' group by Col2 order by count(Col1) desc limit 25 label Col2 'qué', count(Col1) 'veces'",0),"")`);

  [1, 4, 8].forEach(c => sh.getRange(top, c).setFontWeight('bold'));
  rows.forEach((r, i) => { if (r[0] && !r[1]) sh.getRange(i + 1, 1).setFontWeight('bold'); });
  sh.setColumnWidth(1, 280); sh.setColumnWidth(5, 200); sh.setColumnWidth(8, 260);
}
