/** Team Building – Presupuestos · Adventure Co.
 *  Recibe las solicitudes del formulario de la web y lleva el seguimiento.
 *  1) Ejecutar setup() una sola vez.  2) Implementar como aplicación web.
 */
// Token compartido entre el formulario de la web y este script, para validar las peticiones.
// Configuralo en el proyecto de Apps Script: Configuracion del proyecto > Propiedades del script > TOKEN.
const TOKEN = PropertiesService.getScriptProperties().getProperty('TOKEN') || 'YOUR_SECRET_TOKEN';
const SH = 'Solicitudes';
const RS = 'Resumen';
const HEAD = ['Fecha solicitud','Referencia','Estado','Empresa','Contacto','Email','Teléfono',
  'Fecha jornada','Fecha alternativa','Personas','Actividad','Presupuesto actividades (€)',
  'Catering restaurante (€)','Total jornada (€)','Importe final actividades (€)','Comisión 10 % sin IVA (€)',
  'Factura','Detalle del presupuesto','Comentarios del cliente','Notas internas'];
const VERDE = '#2d4a1e', NARANJA = '#e8704e', CREMA = '#f7f4ee';

function setup() {
  const ss = SpreadsheetApp.getActive();
  const ya = ss.getSheetByName(SH);
  if (ya && ya.getLastRow() > 1) throw new Error('La hoja ya tiene solicitudes: setup() no se vuelve a ejecutar para no borrarlas.');
  ss.setSpreadsheetLocale('es_ES');
  ss.setSpreadsheetTimeZone('Europe/Madrid');

  // ---------- Solicitudes ----------
  let s = ss.getSheetByName(SH) || ss.getSheets()[0];
  s.setName(SH);
  s.clear(); s.clearConditionalFormatRules();
  if (s.getMaxColumns() > HEAD.length) s.deleteColumns(HEAD.length + 1, s.getMaxColumns() - HEAD.length);
  s.getRange(1, 1, 1, HEAD.length).setValues([HEAD])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE)
    .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
  s.setFrozenRows(1); s.setFrozenColumns(4); s.setRowHeight(1, 46);
  const w = [130,125,105,180,150,200,115,110,110,80,140,120,120,110,130,110,70,380,260,220];
  w.forEach((px, i) => s.setColumnWidth(i + 1, px));

  s.getRange('A2:A').setNumberFormat('dd/mm/yyyy hh:mm');
  s.getRange('L2:P').setNumberFormat('#,##0.00 €');
  s.getRange('J2:J').setNumberFormat('0');
  s.getRange('A2:T').setVerticalAlignment('top');
  s.getRange('R2:S').setWrap(true);
  s.getRange('O2:O').setBackground('#fff8e1');   // a rellenar si el importe final cambia
  s.getRange('T2:T').setBackground('#fff8e1');

  s.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente', 'Aceptado', 'Rechazado'], true).setAllowInvalid(false).build());
  const C = s.getRange('C2:C');
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Aceptado').setBackground('#d9ead3').setFontColor('#1e5631').setBold(true).setRanges([C]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Rechazado').setBackground('#f4cccc').setFontColor('#8a1c1c').setRanges([C]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pendiente').setBackground('#fff2cc').setFontColor('#7a5b00').setRanges([C]).build()
  ]);
  s.getRange('B1').setNote('Referencia TB-… que recibe el cliente en su presupuesto');
  s.getRange('C1').setNote('Cambiar a "Aceptado" cuando el cliente confirme la reserva, o "Rechazado" si no.');
  s.getRange('O1').setNote('Solo si el importe final de actividades es distinto del presupuesto. Si se deja vacío, la comisión se calcula sobre el presupuesto.');
  s.getRange('P1').setNote('Se calcula solo cuando el estado es "Aceptado".');

  // ---------- Resumen ----------
  let r = ss.getSheetByName(RS) || ss.insertSheet(RS, 0);
  ss.setActiveSheet(r); ss.moveActiveSheet(1);
  r.clear();
  r.setHiddenGridlines(true);
  r.setColumnWidth(1, 24); r.setColumnWidth(2, 270); r.setColumnWidth(3, 150);
  r.setColumnWidth(4, 40); r.setColumnWidths(5, 4, 120);

  r.getRange('B2').setValue('Team Building · Presupuestos').setFontSize(18).setFontWeight('bold').setFontColor(VERDE);
  r.getRange('B3').setValue('Adventure Co. — se actualiza solo con cada solicitud de la web').setFontColor('#6d6a64');

  const S = "Solicitudes!";
  const kpis = [
    ['Solicitudes recibidas', `=COUNTA(${S}B2:B)`, '0'],
    ['Aceptadas', `=COUNTIF(${S}C2:C;"Aceptado")`, '0'],
    ['Pendientes', `=COUNTIF(${S}C2:C;"Pendiente")`, '0'],
    ['Rechazadas', `=COUNTIF(${S}C2:C;"Rechazado")`, '0'],
    ['Conversión (aceptadas / decididas)', '=IFERROR(C6/(C6+C8);0)', '0 %'],
    ['Importe presupuestado (actividades)', `=SUM(${S}L2:L)`, '#,##0.00 €'],
    ['Facturación aceptada (actividades)', `=SUM(${S}P2:P)*10`, '#,##0.00 €'],
    ['Comisión 10 %', `=SUM(${S}P2:P)`, '#,##0.00 €']
  ];
  kpis.forEach((k, i) => {
    const row = 5 + i;
    r.getRange(row, 2).setValue(k[0]).setFontColor('#3d3a35');
    r.getRange(row, 3).setFormula(k[1]).setNumberFormat(k[2]).setFontWeight('bold').setHorizontalAlignment('right');
  });
  r.getRange('B5:C12').setBackground(CREMA).setBorder(true, true, true, true, false, true, '#e3ded4', SpreadsheetApp.BorderStyle.SOLID);
  r.getRange('B12:C12').setBackground(VERDE).setFontColor('#ffffff');
  r.getRange('C12').setFontColor('#8dc63f').setFontSize(13);
  r.setRowHeights(5, 8, 30);
  r.getRange('B5:C12').setVerticalAlignment('middle');

  // por mes
  r.getRange('E4:H4').setValues([['Mes', 'Solicitudes', 'Aceptadas', 'Comisión']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(NARANJA).setHorizontalAlignment('center');
  r.getRange('E5').setFormula(`=IFERROR(SORT(UNIQUE(FILTER(TEXT(${S}A2:A;"yyyy-mm");${S}A2:A<>""));1;FALSE);"")`);
  for (let i = 5; i <= 28; i++) {
    const d = `${S}$A$2:$A;">="&DATEVALUE(E${i}&"-01");${S}$A$2:$A;"<"&EDATE(DATEVALUE(E${i}&"-01");1)`;
    r.getRange(i, 6).setFormula(`=IF(E${i}="";"";COUNTIFS(${d}))`);
    r.getRange(i, 7).setFormula(`=IF(E${i}="";"";COUNTIFS(${d};${S}$C$2:$C;"Aceptado"))`);
    r.getRange(i, 8).setFormula(`=IF(E${i}="";"";SUMIFS(${S}$P$2:$P;${d}))`);
  }
  r.getRange('E5:G28').setHorizontalAlignment('center');
  r.getRange('H5:H28').setNumberFormat('#,##0.00 €');
  r.getRange('E4:H28').setBorder(true, true, true, true, false, false, '#e3ded4', SpreadsheetApp.BorderStyle.SOLID);
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.token !== TOKEN) return out_('forbidden');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = SpreadsheetApp.getActive().getSheetByName(SH);
    const ref = txt_(p.ref);
    if (!ref) return out_('noref');
    asegurarColumnas_(s);
    if (p.mod === '1') return modificar_(s, p);
    const last = s.getLastRow();
    if (last > 1 && s.getRange(2, 2, last - 1, 1).getValues().some(v => v[0] === ref)) return out_('dup');
    s.appendRow([new Date(), ref, 'Pendiente', txt_(p.empresa), txt_(p.contacto), txt_(p.email), txt_(p.tel),
      txt_(p.fecha), txt_(p.fecha2), num_(p.personas), txt_(p.actividad), num_(p.total), num_(p.rest),
      num_(p.jornada), '', '', txt_(p.factura), txt_(p.resumen), txt_(p.coment), '',
      '', txt_(p.hora), num_(p.dur), txt_(p.k), txt_(p.lang), txt_(p.cat), '', txt_(p.snap), '', txt_(p.comida), lugarTxt_(p.lugar, p.comepax)]);
    const row = s.getLastRow();
    s.getRange(row, 16).setFormulaR1C1(FCOMACT_);
    s.getRange(row, 22).setNumberFormat('@').setValue(txt_(p.hora));
    s.getRange(row, C_.comcat).setFormulaR1C1(FCOMCAT_);
    s.getRange(row, C_.comida).setNumberFormat('@').setValue(txt_(p.comida));
    try { sincronizarFila_(s, row); } catch (err) { console.error(err); }
    try { tgPreReserva_(s, row); } catch (err) { console.error(err); }
    if (txt_(p.cat)) try { const v = s.getRange(row, 1, 1, C_.pagado).getValues()[0], dv = s.getRange(row, 1, 1, C_.pagado).getDisplayValues()[0];
      avisoRest_(v, 'PRE-RESERVA (pendiente de confirmar)', cuando_({ s: s, row: row, v: v, d: dv }), 'El grupo ha hecho una pre-reserva. Todavía no está confirmada: os avisaremos en cuanto la confirme.', 'PRE-RESERVA'); } catch (err) { console.error(err); }
    return out_('ok');
  } finally {
    lock.releaseLock();
    if (p.mod === '1') syncMA_();
  }
}



function txt_(v) {
  v = String(v == null ? '' : v).slice(0, 5000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;   // evita fórmulas inyectadas
}
function num_(v) {
  const x = parseFloat(String(v == null ? '' : v).replace(/[^\d,.-]/g, '').replace(',', '.'));
  return isNaN(x) ? '' : x;
}
function out_(m) { return ContentService.createTextOutput(m); }

/** Calendario Team Building – sincroniza la hoja "Solicitudes" con Google Calendar.
 *  ⏳ amarillo = pendiente · ✅ verde = aceptado · Rechazado = se borra del calendario.
 *  Ejecutar instalarCalendario() una sola vez.
 */
// ID del calendario de Google donde se crean los eventos de cada reserva.
// Configuralo en Propiedades del script > CAL_ID (o pega aqui el ID de tu calendario).
const CAL_ID = PropertiesService.getScriptProperties().getProperty('CAL_ID') || 'YOUR_CALENDAR_ID@group.calendar.google.com';
const C_ = { ref: 2, estado: 3, empresa: 4, contacto: 5, email: 6, tel: 7, fecha: 8, fecha2: 9, personas: 10,
  actividad: 11, total: 12, rest: 13, jornada: 14, final: 15, detalle: 18, coment: 19, notas: 20, evento: 21, hora: 22, dur: 23, token: 24, lang: 25, cat: 26, comcat: 27, snap: 28, hist: 29, comida: 30, lugar: 31, pagado: 32 };

function instalarCalendario() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alEditar').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alEditar').forSpreadsheet(ss).onEdit().create();
  const s = ss.getSheetByName(SH);
  if (s.getMaxColumns() < C_.evento) s.insertColumnsAfter(s.getMaxColumns(), C_.evento - s.getMaxColumns());
  s.getRange(1, C_.evento).setValue('ID evento calendario').setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE);
  s.hideColumns(C_.evento);
  for (let r = 2; r <= s.getLastRow(); r++) sincronizarFila_(s, r);
}

function alEditar(e) {
  try { alEditarCore_(e); } finally { if (e && e.range && e.range.getSheet().getName() === SH) syncMA_(); }
}
function alEditarCore_(e) {
  const r = e.range, s = r.getSheet();
  if (s.getName() !== SH || r.getLastRow() < 2) return;
  if (r.getLastColumn() < C_.estado || r.getColumn() > C_.dur || r.getColumn() === C_.evento) return;
  asegurarColumnas_(s);
  for (let row = Math.max(2, r.getRow()); row <= r.getLastRow(); row++) sincronizarFila_(s, row);
  // cambio de estado hecho a mano en la hoja → mismos avisos y PDF que si lo hace el cliente
  if (r.getNumRows() === 1 && r.getNumColumns() === 1 && r.getColumn() === C_.estado && r.getRow() > 1) {
    const nuevo = String(e.value || ''), viejo = String(e.oldValue || '');
    if (nuevo !== viejo && (nuevo === 'Aceptado' || nuevo === 'Cancelado')) avisoManual_(s, r.getRow(), nuevo, viejo);
  }
}

function fechaDe_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  let m = String(v).match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function eur_(n) { return (n === '' || n == null) ? '–' : Number(n).toLocaleString('es-ES', { minimumFractionDigits: Number(n) % 1 ? 2 : 0, maximumFractionDigits: 2 }) + ' €'; }

function sincronizarFila_(s, row) {
  const v = s.getRange(row, 1, 1, C_.lang).getDisplayValues()[0];
  const vr = s.getRange(row, 1, 1, C_.lang).getValues()[0];
  const g = k => vr[C_[k] - 1];
  if (!g('ref')) return;
  const cal = CalendarApp.getCalendarById(CAL_ID);
  if (!cal) throw new Error('No hay acceso al calendario Team Building');
  let ev = null;
  if (g('evento')) { try { ev = cal.getEventById(g('evento')); } catch (err) { ev = null; } }
  const dia = fechaDe_(vr[C_.fecha - 1]);
  const hm = horaDe_(v[C_.hora - 1]), dur = Number(g('dur')) || 3;
  const estado = String(g('estado'));
  if (estado === 'Rechazado' || estado === 'Cancelado' || !dia) {
    if (ev) ev.deleteEvent();
    s.getRange(row, C_.evento).setValue('');
    return;
  }
  const ok = estado === 'Aceptado';
  const titulo = (ok ? '✅ ' : '⏳ PENDIENTE · ') + (hm ? hm.txt + ' · ' : '') + g('empresa') + ' · ' + g('personas') + ' pers · ' + g('actividad');
  const desc = [
    'Referencia: ' + g('ref') + '  ·  Estado: ' + estado,
    hm ? 'Horario: ' + hm.txt + ' – ' + finTxt_(hm, dur) + ' (' + dur + ' h)' : '',
    'Contacto: ' + g('contacto') + ' · ' + g('tel') + ' · ' + g('email'),
    g('fecha2') ? 'Fecha alternativa: ' + (fechaDe_(g('fecha2')) ? Utilities.formatDate(fechaDe_(g('fecha2')), 'Europe/Madrid', 'dd/MM/yyyy') : g('fecha2')) : '',
    '',
    String(g('detalle') || ''),
    '',
    'Total actividades (Adventure Co.): ' + eur_(g('final') !== '' ? g('final') : g('total')),
    'Catering (Restaurante Partner): ' + eur_(g('rest')),
    'Total jornada: ' + eur_(g('jornada')),
    g('coment') ? '\nComentarios del cliente: ' + g('coment') : '',
    g('notas') ? 'Notas internas: ' + g('notas') : ''
  ].filter((x, i, a) => !(x === '' && a[i - 1] === '')).join('\n');
  if (ev && ev.isAllDayEvent() === !!hm) { ev.deleteEvent(); ev = null; }
  let ini = null, fin = null;
  if (hm) { ini = new Date(dia.getFullYear(), dia.getMonth(), dia.getDate(), hm.h, hm.m); fin = new Date(ini.getTime() + dur * 3600000); }
  if (!ev) {
    ev = hm ? cal.createEvent(titulo, ini, fin, { description: desc }) : cal.createAllDayEvent(titulo, dia, { description: desc });
    s.getRange(row, C_.evento).setValue(ev.getId());
  } else {
    ev.setTitle(titulo);
    ev.setDescription(desc);
    if (hm) ev.setTime(ini, fin); else ev.setAllDayDate(dia);
  }
  ev.setColor(ok ? CalendarApp.EventColor.GREEN : CalendarApp.EventColor.YELLOW);
}

/* ================= Reservas: disponibilidad, confirmar, modificar y cancelar ================= */
const CAPACIDAD = 2;              // grupos a la vez
const ABRE = 9, CIERRA = 15;      // hora de inicio: 9:00 – 15:00 (invierno) · 9:00 – 18:00 (verano, jun–sep)
function cierre_(f) { const m = +String(f).slice(5, 7); return (m >= 6 && m <= 9) ? 18 : CIERRA; }
const DIA_ENTERO = 12;            // reserva sin hora = ocupa todo el día
const BLOQ_TAB = 'Días bloqueados';
/** Días completos bloqueados: eventos de 'todo el día' del calendario + pestaña 'Días bloqueados' de la hoja de Empresa */
function bloqueados_(desde) {
  const c = CacheService.getScriptCache(), key = 'bloq_' + (desde || '');
  const hit = c.get(key); if (hit) return JSON.parse(hit);
  const set = {}, tz = 'Europe/Madrid', ini = desde ? fechaDe_(desde) : new Date(), fin = new Date(ini.getFullYear() + 2, ini.getMonth(), ini.getDate());
  try {
    const sh = main_().getSheetByName(SH), nr = sh.getLastRow() - 1, propios = {};
    if (nr > 0) sh.getRange(2, C_.evento, nr, 1).getValues().forEach(r => { if (r[0]) propios[String(r[0])] = 1; });
    CalendarApp.getCalendarById(CAL_ID).getEvents(ini, fin).filter(e => e.isAllDayEvent() && !propios[e.getId()] && !/(PENDIENTE|CONFIRMADA|CONFIRMADO|Team Building ·|TB-\d)/i.test(e.getTitle() + ' ' + e.getDescription())).forEach(e => {
      for (let d = new Date(e.getAllDayStartDate()); d < e.getAllDayEndDate(); d.setDate(d.getDate() + 1)) set[Utilities.formatDate(d, tz, 'yyyy-MM-dd')] = 1;
    });
  } catch (err) { console.error(err); }
  try {
    const id = PropertiesService.getScriptProperties().getProperty('MA_ID'), h = id ? SpreadsheetApp.openById(id).getSheetByName(BLOQ_TAB) : null;
    if (h && h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, 1).getValues().forEach(r => { const d = fechaDe_(r[0]); if (d) set[Utilities.formatDate(d, tz, 'yyyy-MM-dd')] = 1; });
  } catch (err) { console.error(err); }
  const out = Object.keys(set).filter(f => !desde || f >= desde).sort();
  c.put(key, JSON.stringify(out), 120);
  return out;
}
/** ¿Choca este horario? Un grupo a la vez: nunca dos grupos empiezan a la misma hora (el siguiente, mínimo 30 min después);
 *  un grupo de más de 50 personas cuenta como dos (dos subgrupos: a la hora y 30 min después). Máximo CAPACIDAD subgrupos a la vez. */
function choque_(f, ini, dur, pax, ref) {
  if (bloqueados_(f).indexOf(f) >= 0) return true;
  if (reservas_(f, f).some(x => x.tb && x.e === 'C' && x.r !== String(ref))) return true;   // team building: 1 por día
  const otras = reservas_(f, f).filter(x => x.e === 'C' && x.r !== String(ref));
  const W = x => (x.p > 50 ? 2 : 1), S = x => (x.p > 50 ? [x.i, x.i + 0.5] : [x.i]);
  const mias = pax > 50 ? [ini, ini + 0.5] : [ini], w = pax > 50 ? 2 : 1, fin = ini + dur;
  if (otras.some(x => S(x).some(a => mias.some(b => Math.abs(a - b) < 0.26)))) return true;
  const pts = [ini].concat(otras.map(x => x.i).filter(t => t > ini && t < fin));
  return pts.some(t => otras.filter(x => x.i <= t && x.i + x.d > t).reduce((s, x) => s + W(x), 0) + w > CAPACIDAD);
}
const AJ = 'Ajustes';
const WEB = 'https://www.empresa-ejemplo.com';

function horaDe_(t) {
  const m = String(t || '').match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  const h = +m[1], mi = +m[2];
  return { h: h, m: mi, dec: h + mi / 60, txt: ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2) };
}
function finTxt_(hm, dur) { const x = hm.dec + dur, h = Math.floor(x), mi = Math.round((x - h) * 60); return ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2); }

function asegurarColumnas_(s) {
  if (s.getMaxColumns() < C_.pagado) s.insertColumnsAfter(s.getMaxColumns(), C_.pagado - s.getMaxColumns());
  if (s.getRange(1, C_.hora).getValue() !== 'Hora inicio') {
    s.getRange(1, C_.hora, 1, 4).setValues([['Hora inicio', 'Duración (h)', 'Clave confirmación', 'Idioma']])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE).setHorizontalAlignment('center');
    s.setColumnWidth(C_.hora, 95); s.setColumnWidth(C_.dur, 95);
    s.getRange('V2:V').setNumberFormat('@');
    s.hideColumns(C_.token, 2);
  }
  if (s.getRange(1, C_.cat).getValue() !== 'Catering (restaurante)') {
    const ss = SpreadsheetApp.getActive();
    s.getRange(1, C_.cat, 1, 4).setValues([['Catering (restaurante)', 'Comisión catering (€)', 'Datos (JSON)', 'Historial']])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE).setHorizontalAlignment('center').setWrap(true);
    s.setColumnWidth(C_.cat, 300); s.setColumnWidth(C_.comcat, 120); s.setColumnWidth(C_.hist, 300);
    s.getRange('AA2:AA').setNumberFormat('#,##0.00 €');
    s.getRange('Z2:Z').setWrap(true); s.getRange('AC2:AC').setWrap(true);
    s.hideColumns(C_.snap);
    // estado "Cancelado"
    s.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente', 'Aceptado', 'Rechazado', 'Cancelado'], true).setAllowInvalid(false).build());
    const rules = s.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Cancelado').setBackground('#e6e3dd').setFontColor('#6d6a64').setStrikethrough(true).setRanges([s.getRange('C2:C')]).build());
    s.setConditionalFormatRules(rules);
    s.getRange('C1').setNote('Pendiente → el cliente aún no ha confirmado. Aceptado → confirmado (por el cliente o por ti). Rechazado → descartado por vosotros. Cancelado → cancelado por el cliente desde el correo.');
    // fórmulas de comisión catering en filas existentes
    for (let r = 2; r <= s.getLastRow(); r++) s.getRange(r, C_.comcat).setFormulaR1C1(FCOMCAT_);
    // hoja Ajustes
    let a = ss.getSheetByName(AJ) || ss.insertSheet(AJ);
    a.getRange('B2:C2').setValues([['% comisión catering (Restaurante Partner)', '']]);
    a.getRange('B3:C3').setValues([['Email del restaurante (avisos de reservas)', '']]);
    a.getRange('B5').setValue('Mientras no haya email del restaurante, los avisos "Para el restaurante" llegan a info@empresa-ejemplo.com para reenviarlos.').setFontColor('#6d6a64').setWrap(true);
    a.getRange('C2').setNumberFormat('0 %').setBackground('#fff8e1');
    a.getRange('C3').setBackground('#fff8e1');
    a.getRange('B2:B3').setFontWeight('bold').setFontColor(VERDE);
    a.setColumnWidth(1, 24); a.setColumnWidth(2, 380); a.setColumnWidth(3, 260);
    a.setHiddenGridlines(true);
    // Resumen: comisión catering y total
    const rs = ss.getSheetByName(RS);
    if (rs) {
      rs.getRange('B13').setValue('Comisión catering').setFontColor('#3d3a35');
      rs.getRange('C13').setFormula('=SUM(Solicitudes!AA2:AA)').setNumberFormat('#,##0.00 €').setFontWeight('bold').setHorizontalAlignment('right');
      rs.getRange('B14').setValue('Comisión total').setFontWeight('bold').setFontColor('#ffffff');
      rs.getRange('C14').setFormula('=C12+C13').setNumberFormat('#,##0.00 €').setFontWeight('bold').setHorizontalAlignment('right').setFontColor('#8dc63f').setFontSize(13);
      rs.getRange('B13:C13').setBackground(CREMA);
      rs.getRange('B14:C14').setBackground(VERDE);
      rs.setRowHeights(13, 2, 30);
      rs.getRange('B13:C14').setVerticalAlignment('middle');
    }
  }
  if (s.getRange(1, C_.comida).getValue() !== 'Hora comida') {
    s.getRange(1, C_.comida).setValue('Hora comida').setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE).setHorizontalAlignment('center');
    s.getRange(2, C_.comida, s.getMaxRows() - 1, 1).setNumberFormat('@');
    s.setColumnWidth(C_.comida, 95);
  }
  if (s.getRange(1, C_.lugar).getValue() !== 'Comida (lugar)') {
    s.getRange(1, C_.lugar).setValue('Comida (lugar)').setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE).setHorizontalAlignment('center');
    s.setColumnWidth(C_.lugar, 260);
    s.getRange(2, C_.lugar, s.getMaxRows() - 1, 1).setWrap(true);
  }
  if (s.getRange(1, C_.pagado).getValue() !== 'Pagado (Empresa)') {
    s.getRange(1, C_.pagado).setValue('Pagado (Empresa)').setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE).setHorizontalAlignment('center');
    s.setColumnWidth(C_.pagado, 190);
    const rules = s.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Pagado').setBackground('#d9ead3').setFontColor('#1e5631').setBold(true).setRanges([s.getRange(2, C_.pagado, s.getMaxRows() - 1, 1)]).build());
    s.setConditionalFormatRules(rules);
    const rs = SpreadsheetApp.getActive().getSheetByName(RS);
    if (rs) {
      rs.getRange('B16').setValue('Comisión de reservas pagadas').setFontColor('#3d3a35');
      rs.getRange('C16').setFormula('=SUMIFS(Solicitudes!P2:P;Solicitudes!AF2:AF;"Pagado*")').setNumberFormat('#,##0.00 €').setFontWeight('bold').setHorizontalAlignment('right').setFontColor('#1e5631');
      rs.getRange('B17').setValue('Comisión pendiente de cobro (confirmadas sin pagar)').setFontColor('#3d3a35');
      rs.getRange('C17').setFormula('=C12-C16').setNumberFormat('#,##0.00 €').setFontWeight('bold').setHorizontalAlignment('right').setFontColor('#b7791f');
      rs.getRange('B16:C17').setBackground(CREMA).setVerticalAlignment('middle');
      rs.setRowHeights(16, 2, 30);
    }
  }
}
/** Comisión de actividades: 10 % sobre la base SIN IVA (importe / 1,21) */
const IVA = 1.21;
const FCOMACT_ = '=IF(RC3="Aceptado";ROUND(IF(RC15<>"";RC15;RC12)/1,21*0,1;2);"")';
const FCOMCAT_ = '=IF(AND(RC3="Aceptado";Ajustes!R2C3<>"";RC13<>"");RC13*Ajustes!R2C3;"")';

function ajuste_(celda) {
  const a = SpreadsheetApp.getActive().getSheetByName(AJ);
  return a ? a.getRange(celda).getValue() : '';
}

function instalarReservas() {
  const s = SpreadsheetApp.getActive().getSheetByName(SH);
  asegurarColumnas_(s);
  MailApp.getRemainingDailyQuota();
  console.log('Zona horaria del script: ' + Session.getScriptTimeZone());
}

/** Lista de reservas (sin datos personales) para el calendario de la web */
function reservas_(desde, hasta) {
  const s = SpreadsheetApp.getActive().getSheetByName(SH);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  const V = s.getRange(2, 1, n, C_.lang).getValues();
  const D = s.getRange(2, 1, n, C_.lang).getDisplayValues();
  const out = [];
  for (let i = 0; i < n; i++) {
    const est = String(V[i][C_.estado - 1]);
    if (est === 'Rechazado' || est === 'Cancelado' || !V[i][C_.ref - 1]) continue;
    const d = fechaDe_(V[i][C_.fecha - 1]);
    if (!d) continue;
    const f = Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd');
    if ((desde && f < desde) || (hasta && f > hasta)) continue;
    const hm = horaDe_(D[i][C_.hora - 1]);
    out.push({ f: f, i: hm ? hm.dec : ABRE, d: hm ? (Number(V[i][C_.dur - 1]) || 3) : DIA_ENTERO,
      e: est === 'Aceptado' ? 'C' : 'P', r: String(V[i][C_.ref - 1]), p: Number(V[i][C_.personas - 1]) || 0, tb: 1 });
  }
  try { otrReservas_(desde, hasta).forEach(x => out.push(x)); } catch (err) { console.error(err); }
  return out;
}

function fila_(ref, k) {
  const s = SpreadsheetApp.getActive().getSheetByName(SH);
  const n = s.getLastRow() - 1;
  if (n < 1 || !ref || !k) return null;
  const V = s.getRange(2, 1, n, C_.pagado).getValues();
  const D = s.getRange(2, 1, n, C_.pagado).getDisplayValues();
  const i = V.findIndex(r => String(r[C_.ref - 1]) === String(ref));
  if (i < 0 || !V[i][C_.token - 1] || String(V[i][C_.token - 1]) !== String(k)) return null;
  return { s: s, row: i + 2, v: V[i], d: D[i] };
}
function cuando_(F) {
  const d = fechaDe_(F.v[C_.fecha - 1]), hm = horaDe_(F.d[C_.hora - 1]), dur = Number(F.v[C_.dur - 1]) || 3;
  return (d ? Utilities.formatDate(d, 'Europe/Madrid', 'dd/MM/yyyy') : '') + (hm ? ' · ' + hm.txt + '–' + finTxt_(hm, dur) : '');
}
function hist_(F, txt) {
  const prev = String(F.v[C_.hist - 1] || '');
  F.s.getRange(F.row, C_.hist).setValue((prev ? prev + '\n' : '') + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm') + ' · ' + txt);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.a === 'eva') return evaPagina_(p.k);
  if (p.a === 'disp') {
    const list = reservas_(p.desde, p.hasta).map(x => ({ f: x.f, i: x.i, d: x.d, e: x.e, p: x.p }));
    return json_({ ok: true, cap: CAPACIDAD, abre: ABRE, cierra: CIERRA, r: list, b: bloqueados_(p.desde).concat(tbDias_(p.desde)).filter((x, i, a) => a.indexOf(x) === i).sort() });
  }
  if (p.a === 'get') {
    const F = fila_(p.ref, p.k);
    if (!F) return json_({ ok: false, err: 'nf' });
    const est = String(F.v[C_.estado - 1]);
    if (est === 'Cancelado' || est === 'Rechazado') return json_({ ok: false, err: 'cancel' });
    const d = fechaDe_(F.v[C_.fecha - 1]), hm = horaDe_(F.d[C_.hora - 1]);
    let snap = {}; try { snap = JSON.parse(String(F.v[C_.snap - 1] || '{}')); } catch (err) { snap = {}; }
    return json_({ ok: true, ref: p.ref, estado: est, snap: snap,
      prev: d ? { f: Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd'), i: hm ? hm.dec : ABRE, d: hm ? (Number(F.v[C_.dur - 1]) || 3) : DIA_ENTERO, e: est === 'Aceptado' ? 'C' : 'P', p: Number(F.v[C_.personas - 1]) || 0 } : null });
  }
  if (p.a === 'confirmar' || p.a === 'cancelar') return paginaAccion_(p);
  if (p.a === 'pagar') return paginaPago_(p);
  if (p.a === 'restpago') return paginaRestPago_(p);
  if (p.a === 'fichaprev' && p.t === TOKEN) {
    const s = SpreadsheetApp.getActive().getSheetByName(SH), row = s.getLastRow();
    const v = s.getRange(row, 1, 1, C_.pagado).getValues()[0], d = s.getRange(row, 1, 1, C_.pagado).getDisplayValues()[0];
    const b = fichaPDF_(v, cuando_({ s: s, row: row, v: v, d: d }), p.tipo || 'ma', p.estado || 'CONFIRMADA');
    return json_({ ok: true, b64: Utilities.base64Encode(b.getBytes()) });
  }
  return out_('Team Building – Presupuestos: activo');
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- textos ES / CA ---------- */
function T_(ca) {
  return ca ? {
    ok: 'Reserva confirmada!', okt: 'Gràcies! Hem rebut la confirmació de la teva reserva. Uns dies abans de la jornada rebràs un correu de benvinguda amb tots els detalls.',
    ya: 'Aquesta reserva ja estava confirmada', yat: 'No cal fer res més. Ens posarem en contacte amb tu molt aviat.',
    full: 'Aquest horari ja s’ha completat', fullt: 'Un altre grup ha confirmat abans aquest horari. Pots modificar la teva pre-reserva i triar una altra hora o data, o escriure’ns a info@empresa-ejemplo.com i t’ajudem.',
    err: 'No hem trobat la reserva', errt: 'L’enllaç no és vàlid o ha caducat. Escriu-nos per WhatsApp al 600 000 000 i ho resolem.',
    canc: 'Reserva cancel·lada', canct: 'Hem cancel·lat la teva reserva i l’horari torna a quedar lliure. Si vols, pots tornar a demanar pressupost quan vulguis.',
    yacanc: 'Aquesta reserva ja està cancel·lada', yacanct: 'No cal fer res més. Si vols una nova data, pots demanar un pressupost nou.',
    qconf: 'Confirmes la reserva?', qconft: 'Revisa les dades i prem el botó per confirmar-la.',
    qcanc: 'Vols cancel·lar la reserva?', qcanct: 'Si la cancel·les, l’horari queda lliure per a altres grups. Aquesta acció no es pot desfer.',
    bconf: '✓ Sí, confirmar la reserva', bcanc: 'Sí, cancel·lar la reserva', bmod: 'Prefereixo modificar-la', wait: 'Un moment…',
    fail: 'No s’ha pogut completar. Torna-ho a provar o escriu-nos per WhatsApp al 600 000 000.',
    ref: 'Referència', when: 'Jornada', emp: 'Empresa', web: '/team-building-ca/' } : {
    ok: '¡Reserva confirmada!', okt: '¡Gracias! Hemos recibido la confirmación de tu reserva. Unos días antes de la jornada recibirás un correo de bienvenida con todos los detalles.',
    ya: 'Esta reserva ya estaba confirmada', yat: 'No hace falta hacer nada más. Nos pondremos en contacto contigo muy pronto.',
    full: 'Este horario ya se ha completado', fullt: 'Otro grupo ha confirmado antes este horario. Puedes modificar tu pre-reserva y elegir otra hora o fecha, o escribirnos a info@empresa-ejemplo.com y te ayudamos.',
    err: 'No hemos encontrado la reserva', errt: 'El enlace no es válido o ha caducado. Escríbenos por WhatsApp al 600 000 000 y lo resolvemos.',
    canc: 'Reserva cancelada', canct: 'Hemos cancelado tu reserva y el horario vuelve a quedar libre. Si quieres, puedes volver a pedir presupuesto cuando quieras.',
    yacanc: 'Esta reserva ya está cancelada', yacanct: 'No hace falta hacer nada más. Si quieres una nueva fecha, puedes pedir un presupuesto nuevo.',
    qconf: '¿Confirmas la reserva?', qconft: 'Revisa los datos y pulsa el botón para confirmarla.',
    qcanc: '¿Quieres cancelar la reserva?', qcanct: 'Si la cancelas, el horario queda libre para otros grupos. Esta acción no se puede deshacer.',
    bconf: '✓ Sí, confirmar la reserva', bcanc: 'Sí, cancelar la reserva', bmod: 'Prefiero modificarla', wait: 'Un momento…',
    fail: 'No se ha podido completar. Vuelve a intentarlo o escríbenos por WhatsApp al 600 000 000.',
    ref: 'Referencia', when: 'Jornada', emp: 'Empresa', web: '/es/team-building-es/' };
}

/** Página intermedia (evita que los antivirus de correo confirmen/cancelen al abrir el enlace) */
function paginaAccion_(p) {
  const ca = p.l === 'ca', T = T_(ca), tipo = p.a;
  const F = fila_(p.ref, p.k);
  if (!F) return pagina_(card_(T.err, T.errt, '', 'err', ca), T.err, ca);
  const est = String(F.v[C_.estado - 1]);
  const info = T.ref + ': <b>' + p.ref + '</b><br>' + T.emp + ': <b>' + esc_(F.v[C_.empresa - 1]) + '</b><br>' + T.when + ': <b>' + cuando_(F) + '</b>';
  if (est === 'Cancelado' || est === 'Rechazado') return pagina_(card_(T.yacanc, T.yacanct, info, 'full', ca), T.yacanc, ca);
  if (tipo === 'confirmar' && est === 'Aceptado') return pagina_(card_(T.ya, T.yat, info, 'ok', ca), T.ya, ca);
  const modUrl = WEB + T.web + '?mod=' + encodeURIComponent(p.ref) + '&k=' + encodeURIComponent(p.k) + '#calculadora';
  const conf = tipo === 'confirmar';
  const body = '<div class="ico" style="background:' + (conf ? '#2d4a1e' : '#b7791f') + '">' + (conf ? '✓' : '!') + '</div>' +
    '<h1 style="color:' + (conf ? '#2d4a1e' : '#b7791f') + '">' + (conf ? T.qconf : T.qcanc) + '</h1><p>' + (conf ? T.qconft : T.qcanct) + '</p>' +
    '<div class="i">' + info + '</div>' +
    '<button id="go" class="b' + (conf ? '' : ' r') + '">' + (conf ? T.bconf : T.bcanc) + '</button>' +
    (conf ? '' : '<p><a class="lk" href="' + modUrl + '">' + T.bmod + '</a></p>') +
    '<p class="er" id="er"></p>' +
    '<script>document.getElementById("go").onclick=function(){var b=this;b.disabled=true;b.textContent=' + JSON.stringify(T.wait) + ';' +
    'google.script.run.withSuccessHandler(function(h){document.getElementById("c").innerHTML=h;}).withFailureHandler(function(){b.disabled=false;b.textContent=' + JSON.stringify(conf ? T.bconf : T.bcanc) + ';document.getElementById("er").textContent=' + JSON.stringify(T.fail) + ';})' +
    '.ejecutarAccion(' + JSON.stringify(tipo) + ',' + JSON.stringify(String(p.ref)) + ',' + JSON.stringify(String(p.k)) + ',' + JSON.stringify(ca ? 'ca' : 'es') + ');};<\/script>';
  return pagina_(body, conf ? T.qconf : T.qcanc, ca);
}

/** Llamado desde la página (google.script.run). Devuelve el HTML de la tarjeta. */
function ejecutarAccion(tipo, ref, k, l) {
  const ca = l === 'ca', T = T_(ca);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const F = fila_(ref, k);
    if (!F) return card_(T.err, T.errt, '', 'err', ca);
    const est = String(F.v[C_.estado - 1]), cuando = cuando_(F);
    const info = T.ref + ': <b>' + ref + '</b><br>' + T.when + ': <b>' + cuando + '</b>';
    if (est === 'Cancelado' || est === 'Rechazado') return card_(T.yacanc, T.yacanct, info, 'full', ca);
    if (tipo === 'cancelar') {
      F.s.getRange(F.row, C_.estado).setValue('Cancelado');
      hist_(F, 'Cancelada por el cliente desde el correo (estado previo: ' + est + ')');
      try { sincronizarFila_(F.s, F.row); } catch (err) { console.error(err); }
      avisoMA_(F.v, '❌ Reserva CANCELADA por el cliente', 'El cliente ha cancelado la reserva desde el correo. El horario vuelve a estar libre y se ha borrado del calendario.', cuando, 'CANCELADA');
      if (F.v[C_.cat - 1] && avisadoRest_(F.v, est)) avisoRest_(F.v, 'CANCELADA', cuando, 'El grupo ha cancelado la reserva. Por favor, anuladla.', 'CANCELADA');
      return card_(T.canc, T.canct, info, 'full', ca);
    }
    // confirmar
    if (est === 'Aceptado') return card_(T.ya, T.yat, info, 'ok', ca);
    const d = fechaDe_(F.v[C_.fecha - 1]), hm = horaDe_(F.d[C_.hora - 1]), dur = Number(F.v[C_.dur - 1]) || 3;
    if (d) {
      const f = Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd');
      const ini = hm ? hm.dec : ABRE, dd = hm ? dur : DIA_ENTERO;
      if (choque_(f, ini, dd, Number(F.v[C_.personas - 1]) || 0, ref)) {
        avisoMA_(F.v, '⚠️ Horario completo', 'El cliente ha intentado confirmar, pero el horario ya está completo. Escríbele para proponer otra hora o fecha.', cuando);
        return card_(T.full, T.fullt, info, 'full', ca);
      }
    }
    const modificada = /Modificada/.test(String(F.v[C_.hist - 1] || ''));
    F.s.getRange(F.row, C_.estado).setValue('Aceptado');
    hist_(F, 'Confirmada por el cliente desde el correo');
    try { sincronizarFila_(F.s, F.row); } catch (err) { console.error(err); }
    avisarPendientes_(F.v);
    try { correoConfirmadoCliente_({ s: F.s, row: F.row, v: F.s.getRange(F.row, 1, 1, C_.pagado).getValues()[0], d: F.d }, cuando); } catch (err) { console.error(err); }
    avisoMA_(F.v, '✅ Reserva CONFIRMADA por el cliente', 'El cliente ha confirmado la reserva desde el correo. Ya está en verde en el calendario. ' + DIAS_BIENVENIDA + ' días antes de la jornada os llegará un aviso para enviarle el correo de bienvenida.', cuando, 'CONFIRMADA');
    if (F.v[C_.cat - 1]) avisoRest_(F.v, modificada ? 'CONFIRMADA (con cambios)' : 'CONFIRMADA', cuando, modificada ? 'Esta reserva sustituye a la versión anterior que os enviamos.' : '', 'CONFIRMADA');
    return card_(T.ok, T.okt, info, 'ok', ca);
  } finally {
    lock.releaseLock();
    syncMA_();
  }
}

/** Modificación desde la web (doPost con mod=1) */
function modificar_(s, p) {
  const F = fila_(p.ref, p.k);
  if (!F) return out_('forbidden');
  const est = String(F.v[C_.estado - 1]);
  if (est === 'Cancelado' || est === 'Rechazado') return out_('cancelled');
  const antes = cuando_(F), catAntes = String(F.v[C_.cat - 1] || '');
  const r = F.row;
  s.getRange(r, C_.estado).setValue('Pendiente');
  s.getRange(r, C_.empresa, 1, 11).setValues([[txt_(p.empresa), txt_(p.contacto), txt_(p.email), txt_(p.tel), txt_(p.fecha), '',
    num_(p.personas), txt_(p.actividad), num_(p.total), num_(p.rest), num_(p.jornada)]]);
  s.getRange(r, 17, 1, 3).setValues([[txt_(p.factura), txt_(p.resumen), txt_(p.coment)]]);
  s.getRange(r, C_.hora).setNumberFormat('@').setValue(txt_(p.hora));
  s.getRange(r, C_.dur).setValue(num_(p.dur));
  s.getRange(r, C_.lang).setValue(txt_(p.lang));
  s.getRange(r, C_.cat).setValue(txt_(p.cat));
  s.getRange(r, C_.snap).setValue(txt_(p.snap));
  s.getRange(r, C_.comida).setNumberFormat('@').setValue(txt_(p.comida));
  s.getRange(r, C_.lugar).setValue(lugarTxt_(p.lugar, p.comepax));
  hist_(F, 'Modificada por el cliente (antes: ' + antes + ', ' + est + ')');
  try { sincronizarFila_(s, r); } catch (err) { console.error(err); }
  const nv = s.getRange(r, 1, 1, C_.pagado).getValues()[0];
  const F2 = { s: s, row: r, v: nv, d: s.getRange(r, 1, 1, C_.pagado).getDisplayValues()[0] };
  const ahora = cuando_(F2);
  avisoMA_(nv, '✏️ Reserva MODIFICADA por el cliente', 'El cliente ha modificado la reserva desde la web. Queda como Pendiente hasta que la vuelva a confirmar desde el nuevo correo.\nAntes: ' + antes + ' (' + est + ')', ahora, 'MODIFICADA');
  if (catAntes || nv[C_.cat - 1]) avisoRest_(nv, 'MODIFICADA – pendiente de confirmar', ahora,
    'El grupo ha cambiado su reserva. Antes: ' + antes + (catAntes ? ' · ' + catAntes : '') + '. Os avisaremos cuando la confirme.', 'MODIFICADA');
  return out_('ok');
}

/* ---------- avisos por correo ---------- */
function envio_(to, subject, body, pdf, html) {
  try {
    const o = { to: to, subject: subject, body: body, name: 'Adventure Co. – Reservas' };
    if (pdf) o.attachments = [pdf];
    if (html) o.htmlBody = html;
    MailApp.sendEmail(o);
  } catch (err) { console.error(err); }
  try { tgEspejo_(to, subject, body); } catch (err) { console.error(err); }
}
function ficha_(r, cuando, tipo, estadoFicha) {
  if (!estadoFicha) return null;
  try { return fichaPDF_(r, cuando, tipo, estadoFicha); } catch (err) { console.error(err); return null; }
}

/** Aviso a Adventure Co. (info@) y, por separado, a colaborador (con comisiones). Cada uno con su ficha PDF. */
function avisoMA_(r, titulo, intro, cuando, estadoFicha) {
  const pct = ajuste_('C2');
  const rest = Number(r[C_.rest - 1]) || 0;
  const comAct = (Number(r[C_.final - 1]) || Number(r[C_.total - 1]) || 0) / IVA * 0.1;
  const comCat = (pct !== '' && rest) ? rest * Number(pct) : null;
  const datos = intro + '\n\n' +
    'Referencia: ' + r[C_.ref - 1] + '\nEmpresa: ' + r[C_.empresa - 1] + '\nContacto: ' + r[C_.contacto - 1] + ' · ' + r[C_.tel - 1] + ' · ' + r[C_.email - 1] +
    '\nJornada: ' + cuando + (comidaTxt_(r) ? '\nComida: ' + comidaTxt_(r) : '') + '\nPersonas: ' + r[C_.personas - 1] + ' · ' + r[C_.actividad - 1] +
    '\n\nActividades (Adventure Co.): ' + eur_(r[C_.total - 1]) +
    '\nCatering (Restaurante Partner): ' + (rest ? eur_(rest) + ' (base imponible ' + eur_(iv10_(rest)[0]) + ' · IVA 10 % ' + eur_(iv10_(rest)[1]) + ')' : '–') + (r[C_.cat - 1] ? '\n  ' + String(r[C_.cat - 1]).split(' · ').join('\n  ') : '') +
    '\nTotal jornada: ' + eur_(r[C_.jornada - 1]);
  const com = '\n\nComisión actividades (10 % sin IVA): ' + eur_(Math.round(comAct * 100) / 100) +
    '\nComisión catering: ' + (comCat === null ? (rest ? 'pendiente de definir el % en la hoja (pestaña Ajustes)' : '–') : eur_(Math.round(comCat * 100) / 100) + ' (' + Math.round(Number(pct) * 100) + ' %)');
  const coment = r[C_.coment - 1] ? '\n\nComentarios del cliente: ' + r[C_.coment - 1] : '';
  const pie = estadoFicha ? '\n\nAdjunto: ficha de la reserva en PDF.' : '';
  const asunto = titulo + ' – ' + r[C_.ref - 1] + ' – ' + r[C_.empresa - 1];
  const caC = String(r[C_.lang - 1]) === 'ca', tr = t => caC ? caMA_(t) : t;
  envio_('info@empresa-ejemplo.com', tr(asunto), tr(datos + coment + pie), ficha_(r, cuando, 'ma', estadoFicha), estadoFicha === 'CONFIRMADA' ? htmlMA_(r, tr(datos + coment + pie), caC) : null);
  envio_('colaborador@example.com', asunto, datos + com + coment + pie, ficha_(r, cuando, 'colab', estadoFicha));
}

/** Aviso al restaurante, solo con el catering y su ficha PDF */
function avisoRest_(r, estado, cuando, nota, estadoFicha) {
  const dest = String(ajuste_('C3') || '').trim();
  const lineas = String(r[C_.cat - 1] || '').split(' · ').filter(String);
  const cuerpo = 'Hola,\n\n' + 'Reserva de grupo de Adventure Co.: ' + estado + '\n' + (nota ? nota + '\n' : '') + '\n' +
    'Grupo: ' + r[C_.empresa - 1] + '\nContacto del grupo: ' + r[C_.contacto - 1] + ' · ' + r[C_.tel - 1] + ' · ' + r[C_.email - 1] + '\nPersonas: ' + r[C_.personas - 1] + (Number(r[C_.personas - 1]) > 50 ? ' (grupo grande: la actividad se hace en dos grupos, el segundo empieza 30 min después; lo más recomendable es comer en el porche)' : '') + '\nFecha y horario de la actividad: ' + cuando + (servicioRest_(r) ? '\n' + servicioRest_(r) : '') +
    '\n\nCatering:\n' + (lineas.length ? lineas.map(x => '  • ' + x).join('\n') : '  • Sin catering') +
    '\nImporte catering (IVA incl.): ' + eur_(r[C_.rest - 1]) + ' (base imponible ' + eur_(iv10_(r[C_.rest - 1])[0]) + ' · IVA 10 % ' + eur_(iv10_(r[C_.rest - 1])[1]) + ')' +
    (r[C_.coment - 1] ? '\n\nComentarios del grupo (alergias, etc.): ' + r[C_.coment - 1] : '') +
    '\n\nReferencia: ' + r[C_.ref - 1] + '\nContacto Adventure Co.: 600 000 000 · info@empresa-ejemplo.com' +
    (estadoFicha ? '\n\nAdjunto: ficha del catering en PDF.' : '') + '\n\nGracias,\nAdventure Co.';
  const asunto = '🍽️ ' + (dest ? '' : 'PARA EL RESTAURANTE · ') + 'Reserva ' + estado + ' – ' + r[C_.empresa - 1] + ' – ' + cuando + (r[C_.comida - 1] ? ' · comida ' + r[C_.comida - 1] + ' h' + (enInst_(r) ? ' en las instalaciones de Adventure Co.' : ' en el restaurante') : '');
  envio_(dest || 'info@empresa-ejemplo.com', asunto, cuerpo, ficha_(r, cuando, 'rest', estadoFicha));
}
function avisadoRest_(r, estPrevio) { return estPrevio === 'Aceptado' || /Modificada/.test(String(r[C_.hist - 1] || '')); }

/* ---------- HTML ---------- */
function esc_(t) { return String(t == null ? '' : t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function card_(tit, txt, info, tipo, ca) {
  const col = tipo === 'ok' ? '#2d4a1e' : tipo === 'full' ? '#b7791f' : '#9b2c1f';
  const ico = tipo === 'ok' ? '✓' : tipo === 'full' ? '!' : '×';
  return '<div class="ico" style="background:' + col + '">' + ico + '</div><h1 style="color:' + col + '">' + tit + '</h1><p>' + txt + '</p>' +
    (info ? '<div class="i">' + info + '</div>' : '') +
    '<p>WhatsApp: <b>600 000 000</b></p><a class="b" href="' + WEB + T_(ca).web + '" target="_top">empresa-ejemplo.com</a>';
}
function pagina_(inner, tit, ca) {
  const html = '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><base target="_top">' +
    '<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;800&display=swap" rel="stylesheet">' +
    '<style>body{margin:0;background:#f7f4ee;font-family:Poppins,Arial,sans-serif;color:#3d3a35}' +
    '.bar{height:90px;background:#ff6b35 url(https://www.empresa-ejemplo.com/wp-content/uploads/2017/02/banner-photo.png) center/cover}' +
    '.card{max-width:520px;margin:-45px auto 40px;background:#fff;border-radius:18px;padding:34px 28px;text-align:center;box-shadow:0 10px 30px rgba(0,0,0,.08)}' +
    '.ico{width:64px;height:64px;border-radius:50%;color:#fff;font-size:34px;line-height:64px;margin:0 auto 14px;font-weight:800}' +
    'h1{font-size:24px;margin:0 0 10px}p{line-height:1.6;margin:10px 0}.i{background:#f7f4ee;border-radius:12px;padding:12px;margin:18px 0;line-height:1.7}' +
    '.b{display:inline-block;border:0;cursor:pointer;background:#e8704e;color:#fff;text-decoration:none;padding:13px 26px;border-radius:30px;font:600 16px Poppins,Arial,sans-serif;margin-top:8px}' +
    'button.b{background:#2d4a1e}button.b.r{background:#9b2c1f}button.b:disabled{opacity:.6;cursor:default}' +
    '.lk{color:#2d4a1e;font-weight:600}.er{color:#9b2c1f;font-size:14px}' +
    'img.lg{height:54px;margin-bottom:14px}</style></head><body><div class="bar"></div><div class="card">' +
    '<img class="lg" src="https://www.empresa-ejemplo.com/wp-content/uploads/2026/10/logo-empresa.png" alt="Adventure Co.">' +
    '<div id="c">' + inner + '</div></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle(tit + ' · Adventure Co.').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================= Fichas PDF de reserva confirmada ================= */
function lineasResumen_(txt) {
  const act = [], cat = [];
  let sec = '';
  String(txt || '').split('\n').forEach(l => {
    if (/^ACTIVIDADES|^ACTIVITATS/.test(l)) sec = 'a';
    else if (/^CATERING/.test(l)) sec = 'c';
    const m = l.match(/^- (.+?): (\d+) × (.+?) = (.+)$/);
    if (m && sec) (sec === 'a' ? act : cat).push([m[1].replace(/ \(Restaurante Partner\)$/, ''), m[2], m[3], m[4]]);
  });
  return { act: act, cat: cat };
}
const FV_ = { verde: '#2d4a1e', or: '#e8704e', orbar: '#ff6b35', crema: '#f7f4ee', txt: '#3d3a35', gris: '#6d6a64', lima: '#5f8f22', marron: '#463428', borde: '#e3ded4' };
// Nota: el conversor HTML→PDF de Google no pinta fondos; el diseño usa texto de color y bordes.
function h2_(t, col) { return '<p style="color:' + (col || FV_.or) + ';font-size:8.5pt;font-weight:bold;letter-spacing:1.5px;margin:16px 0 4px">' + t + '</p>'; }
function filas_(L) {
  return L.map(x => '<tr><td style="padding:6px 8px;border-bottom:1px solid ' + FV_.borde + '">' + esc_(x[0]) + '</td><td style="padding:6px 8px;text-align:center;border-bottom:1px solid ' + FV_.borde + '">' + esc_(x[1]) + '</td><td style="padding:6px 8px;text-align:center;border-bottom:1px solid ' + FV_.borde + '">' + esc_(x[2]) + '</td><td style="padding:6px 8px;text-align:right;border-bottom:1px solid ' + FV_.borde + '">' + esc_(x[3]) + '</td></tr>').join('');
}
function tabla_(L, head, col) {
  const th = 'padding:6px 8px;color:' + (col || FV_.verde) + ';font-size:7.5pt;letter-spacing:.8px;border-bottom:2px solid ' + (col || FV_.verde);
  return '<table style="width:100%;border-collapse:collapse;font-size:9.5pt">' +
    '<tr><th style="' + th + ';text-align:left">CONCEPTO</th><th style="' + th + '">' + (head || 'CANTIDAD') + '</th><th style="' + th + '">PRECIO</th><th style="' + th + ';text-align:right">IMPORTE</th></tr>' +
    filas_(L) + '</table>';
}
function totRow_(lbl, val, col, grande) {
  return '<table style="width:100%;border-collapse:collapse"><tr><td style="padding:8px;border-top:2px solid ' + col + ';border-bottom:2px solid ' + col + ';color:' + col + ';font-weight:bold;font-size:' + (grande ? '12pt' : '10.5pt') + '">' + lbl + '</td><td style="padding:8px;border-top:2px solid ' + col + ';border-bottom:2px solid ' + col + ';color:' + col + ';font-weight:bold;text-align:right;font-size:' + (grande ? '14pt' : '12pt') + '">' + val + '</td></tr></table>';
}
function subRow_(lbl, val) {
  return '<tr><td style="padding:4px 8px;color:' + FV_.gris + ';font-size:9pt">' + lbl + '</td><td style="padding:4px 8px;color:' + FV_.gris + ';font-size:9pt;text-align:right">' + val + '</td></tr>';
}
function dato_(k, v) { return '<td style="padding:6px 8px;width:50%;vertical-align:top;border-bottom:1px solid ' + FV_.borde + '"><span style="font-size:7pt;color:' + FV_.gris + ';letter-spacing:.6px">' + k.toUpperCase() + '</span><br><span style="font-size:10.5pt">' + esc_(v || '–') + '</span></td>'; }

/** tipo: 'ma' (Adventure Co.), 'colab' (con comisiones) o 'rest' (restaurante, solo catering) */
function fichaPDF_(r, cuando, tipo, estado) {
  const L = lineasResumen_(r[C_.detalle - 1]);
  const rest = Number(r[C_.rest - 1]) || 0, tot = Number(r[C_.total - 1]) || 0;
  const base = Math.round(tot / 1.21 * 100) / 100, iva = Math.round((tot - base) * 100) / 100;
  const pct = ajuste_('C2');
  const hoy = Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm');
  const esRest = tipo === 'rest', esColab = tipo === 'colab';
  const est = String(estado).split(' ')[0];
  const EB = est === 'CANCELADA' ? ['#b3261e', '✕  RESERVA CANCELADA', 'El grupo ha cancelado la reserva.'] : est === 'MODIFICADA' ? ['#b7791f', '✎  RESERVA MODIFICADA', 'Pendiente de que el cliente la vuelva a confirmar.'] : ['#2e7d32', '✓  RESERVA CONFIRMADA', 'Confirmada · ' + hoy];
  const comida = String(r[C_.comida - 1] || '');
  const LOGO_MA = 'https://www.empresa-ejemplo.com/wp-content/uploads/2026/10/logo-empresa.png';
  const LOGO_V = 'https://www.empresa-ejemplo.com/wp-content/uploads/2026/10/logo-restaurante-partner.png';
  let h = '<html><head><meta charset="utf-8"></head><body style="font-family:Arial,Helvetica,sans-serif;color:' + FV_.txt + ';margin:0">' +
    '<table style="width:100%;border-collapse:collapse"><tr>' +
    '<td style="padding:6px 4px 10px;width:50%;border-bottom:5px solid ' + FV_.orbar + '"><img src="' + LOGO_MA + '" style="height:56px"></td>' +
    '<td style="padding:6px 4px 10px;width:50%;text-align:right;border-bottom:5px solid ' + FV_.orbar + '"><img src="' + LOGO_V + '" style="height:44px"></td></tr></table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:10px"><tr>' +
    '<td style="padding:4px"><span style="font-size:19pt;font-weight:bold;color:' + FV_.verde + '">' + (esRest ? 'Reserva de catering' : 'Ficha de reserva') + '</span><br>' +
    '<span style="font-size:9pt;color:' + FV_.gris + '">' + (esRest ? 'Grupo de Adventure Co. · Restaurante Partner' : esColab ? 'Team building · copia del colaborador (con comisiones)' : 'Team building · Adventure Co.') + '</span></td>' +
    '<td style="padding:4px;text-align:right;vertical-align:top"><span style="font-size:12pt;font-weight:bold;color:' + FV_.or + '">' + esc_(r[C_.ref - 1]) + '</span><br><span style="font-size:8pt;color:' + FV_.gris + '">Generada: ' + hoy + '</span></td></tr></table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:10px"><tr><td style="border-left:6px solid ' + EB[0] + ';border-top:1px solid ' + EB[0] + ';border-bottom:1px solid ' + EB[0] + ';border-right:1px solid ' + EB[0] + ';padding:8px 12px">' +
    '<span style="color:' + EB[0] + ';font-weight:bold;font-size:12pt">' + EB[1] + '</span><br><span style="color:' + FV_.gris + ';font-size:8.5pt">' + EB[2] + '</span></td></tr></table>' +
    (esRest ? (servicioRest_(r) ? '<table style="width:100%;border-collapse:collapse;margin-top:10px"><tr><td style="border:2px solid ' + FV_.or + ';padding:10px 12px;color:' + FV_.or + ';font-weight:bold;font-size:12.5pt">' + esc_(servicioRest_(r)) + '</td></tr></table>' : '')
      : (comidaTxt_(r) ? '<table style="width:100%;border-collapse:collapse;margin-top:10px"><tr><td style="border:2px solid ' + FV_.or + ';padding:9px 12px;color:' + FV_.or + ';font-weight:bold;font-size:11pt">COMIDA: ' + esc_(comidaTxt_(r)) + '</td></tr></table>' : '')) +
    h2_(esRest ? 'GRUPO' : 'DATOS DEL CLIENTE') +
    '<table style="width:100%;border-collapse:collapse">' +
    '<tr>' + dato_('Empresa', r[C_.empresa - 1]) + dato_('Actividad · horario', cuando) + '</tr>' +
    (comidaTxt_(r) ? '<tr>' + dato_('Comida', comidaTxt_(r)) + dato_('Personas · actividad', r[C_.personas - 1] + ' · ' + r[C_.actividad - 1]) + '</tr>'
      : '<tr>' + dato_('Personas', r[C_.personas - 1]) + dato_('Actividad', r[C_.actividad - 1]) + '</tr>') +
    '<tr>' + dato_('Contacto', r[C_.contacto - 1] + ' · ' + r[C_.tel - 1]) + dato_('Email', r[C_.email - 1]) + '</tr>' +
    (esRest ? '<tr>' + dato_('Referencia', r[C_.ref - 1]) + dato_('Contacto Adventure Co.', '600 000 000') + '</tr>' : '<tr>' + dato_('Factura a nombre de la empresa', r[C_.factura - 1]) + dato_('Referencia', r[C_.ref - 1]) + '</tr>') +
    '</table>' + (Number(r[C_.personas - 1]) > 50 ? '<table style="width:100%;border-collapse:collapse;margin-top:8px"><tr><td style="border-left:5px solid ' + FV_.or + ';background:#fff8e1;padding:7px 12px;font-size:9.5pt"><b>Grupo de más de 50 personas:</b> la actividad se hace en dos grupos (el segundo empieza 30 minutos después) y lo más recomendable es comer en el porche de las instalaciones.</td></tr></table>' : '');
  if (!esRest && L.act.length) {
    h += h2_('ACTIVIDADES · ADVENTURE CO.') + tabla_(L.act) +
      '<table style="width:100%;border-collapse:collapse">' + subRow_('Base imponible', eur_(base)) + subRow_('IVA (21 %)', eur_(iva)) + '</table>' +
      totRow_('Total actividades (IVA incluido)', eur_(tot), FV_.verde);
  }
  if (L.cat.length) {
    h += h2_('CATERING · RESTAURANTE PARTNER' + (comida ? (enInst_(r) ? ' · SERVIR EN LAS INSTALACIONES A LAS ' : ' · COMIDA A LAS ') + comida + ' H' : ''), FV_.marron) + tabla_(L.cat, 'PERSONAS', FV_.marron) +
      '<table style="width:100%;border-collapse:collapse">' + subRow_('Base imponible', eur_(iv10_(rest)[0])) + subRow_('IVA (10 %)', eur_(iv10_(rest)[1])) + '</table>' +
      totRow_('Total catering (IVA incluido)', eur_(rest), FV_.marron);
    if (/Calçotada/.test(String(r[C_.cat - 1]))) h += '<p style="font-size:8.5pt;color:' + FV_.gris + ';margin:6px 0 0">Calçotada (menú bufet): ensaladas, calçots a la llama con romesco (15 u.), parrillada de carne a la brasa con patatas al caliu, alcachofas y judías del ganxet, alioli, pan, postre y bebida (agua, cerveza, refresco o vino).</p>';
  } else if (esRest) {
    h += '<p>Sin catering.</p>';
  }
  if (!esRest && L.cat.length) h += '<div style="height:10px"></div>' + totRow_('Total jornada (dos facturas)', eur_(r[C_.jornada - 1]), FV_.or, true);
  if (tipo === 'ma' && est === 'CONFIRMADA' && !String(r[C_.pagado - 1] || '')) h += '<table style="width:100%;border-collapse:collapse;margin-top:14px"><tr><td style="border:2px solid #2e7d32;padding:12px 14px;text-align:center">' +
    '<a href="' + pagoUrl_(r[C_.ref - 1]) + '" style="color:#2e7d32;font-weight:bold;font-size:12.5pt;text-decoration:none">MARCAR COMO PAGADA</a><br><span style="font-size:8pt;color:' + FV_.gris + '">Pulsa aquí cuando el cliente haya pagado las actividades. Se te pedirá confirmación.</span></td></tr></table>';
  if (esColab) {
    const comAct = Math.round((Number(r[C_.final - 1]) || tot) / IVA * 0.1 * 100) / 100;
    const comCat = (pct !== '' && rest) ? Math.round(rest * Number(pct) * 100) / 100 : null;
    h += h2_('COMISIONES · USO INTERNO', FV_.lima) + '<table style="width:100%;border-collapse:collapse;border-top:1px solid ' + FV_.borde + '">' +
      subRow_('Comisión actividades (10 % sin IVA)', eur_(comAct)) +
      subRow_('Comisión catering' + (comCat === null ? '' : ' (' + Math.round(Number(pct) * 100) + ' %)'), comCat === null ? (rest ? 'pendiente de definir' : '–') : eur_(comCat)) + '</table>';
  }
  if (r[C_.coment - 1]) h += h2_(esRest ? 'ALERGIAS Y COMENTARIOS DEL GRUPO' : 'COMENTARIOS DEL CLIENTE') +
    '<table style="width:100%;border-collapse:collapse"><tr><td style="border-left:5px solid ' + FV_.or + ';padding:6px 12px;font-size:10pt">' + esc_(r[C_.coment - 1]) + '</td></tr></table>';
  h += '<table style="width:100%;border-collapse:collapse;margin-top:22px"><tr>' +
    '<td style="padding:10px 4px;border-top:3px solid ' + FV_.verde + ';font-size:8pt;color:' + FV_.gris + ';vertical-align:top"><b style="color:' + FV_.verde + '">Adventure Co. S.L.</b> · CIF B00000000<br>[Dirección de la empresa]<br>600 000 000 · info@empresa-ejemplo.com' + (L.cat.length ? '<br>Catering: Restaurante Partner · ' + REST_EMAIL : '') + '</td>' +
    '<td style="padding:10px 4px;border-top:3px solid ' + FV_.verde + ';font-size:8pt;color:' + FV_.gris + ';text-align:right;vertical-align:top">' + (esRest ? '<b style="color:' + FV_.verde + '">El catering lo factura el restaurante</b><br>directamente al cliente.' : '<b style="color:' + FV_.verde + '">Reservas y dudas</b><br>WhatsApp y teléfono: 600 000 000') + '<br><span style="color:' + FV_.or + '">empresa-ejemplo.com</span></td></tr></table>' +
    '</body></html>';
  if (tipo === 'ma' && String(r[C_.lang - 1]) === 'ca') h = caMA_(h);
  const nombre = (esRest ? 'Catering_' : 'Reserva_') + est + '_' + r[C_.ref - 1] + '_' + String(r[C_.empresa - 1]).replace(/[^\wÀ-ÿ]+/g, '_').slice(0, 30) + '.pdf';
  return Utilities.newBlob(h, MimeType.HTML, 'ficha.html').getAs(MimeType.PDF).setName(nombre);
}

/** Envía a colaborador las dos fichas de la última solicitud, para revisar el diseño */
function probarFicha() {
  const s = SpreadsheetApp.getActive().getSheetByName(SH);
  const row = s.getLastRow();
  const v = s.getRange(row, 1, 1, C_.pagado).getValues()[0], d = s.getRange(row, 1, 1, C_.pagado).getDisplayValues()[0];
  const cuando = cuando_({ s: s, row: row, v: v, d: d });
  const a = fichaPDF_(v, cuando, 'ma', 'CONFIRMADA (prueba)'), b = fichaPDF_(v, cuando, 'rest', 'CONFIRMADA (prueba)'), c = fichaPDF_(v, cuando, 'colab', 'CONFIRMADA (prueba)');
  MailApp.sendEmail({ to: 'colaborador@example.com', subject: 'Prueba de fichas PDF – ' + v[C_.ref - 1], body: 'Fichas PDF de prueba: Adventure Co., restaurante y la tuya (con comisiones).', name: 'Adventure Co. – Reservas', attachments: [a, b, c] });
}

/** Avisos cuando el estado se cambia a mano en la hoja */
function avisoManual_(s, row, nuevo, viejo) {
  try {
    const v = s.getRange(row, 1, 1, C_.pagado).getValues()[0], d = s.getRange(row, 1, 1, C_.pagado).getDisplayValues()[0];
    if (!v[C_.ref - 1]) return;
    const F = { s: s, row: row, v: v, d: d }, cuando = cuando_(F);
    if (nuevo === 'Aceptado') {
      hist_(F, 'Confirmada a mano en la hoja (estado previo: ' + (viejo || 'vacío') + ')');
      correoConfirmadoCliente_({ s: s, row: row, v: s.getRange(row, 1, 1, C_.pagado).getValues()[0], d: d }, cuando);
      avisarPendientes_(v);
      avisoMA_(v, '✅ Reserva CONFIRMADA (desde la hoja)', 'La reserva se ha marcado como Aceptado en la hoja. Ya está en verde en el calendario. ' + DIAS_BIENVENIDA + ' días antes de la jornada os llegará un aviso para enviarle el correo de bienvenida.', cuando, 'CONFIRMADA');
      if (v[C_.cat - 1]) avisoRest_(v, /Modificada/.test(String(v[C_.hist - 1] || '')) ? 'CONFIRMADA (con cambios)' : 'CONFIRMADA', cuando, '', 'CONFIRMADA');
    } else if (nuevo === 'Cancelado') {
      hist_(F, 'Cancelada a mano en la hoja (estado previo: ' + (viejo || 'vacío') + ')');
      avisoMA_(v, '❌ Reserva CANCELADA (desde la hoja)', 'La reserva se ha marcado como Cancelado en la hoja. Se ha borrado del calendario y el horario queda libre.', cuando, 'CANCELADA');
      if (v[C_.cat - 1] && avisadoRest_(v, viejo)) avisoRest_(v, 'CANCELADA', cuando, 'El grupo ha cancelado la reserva. Por favor, anuladla.', 'CANCELADA');
    }
  } catch (err) { console.error(err); }
}

/* ================= Aviso a solicitudes pendientes cuando otro grupo confirma ================= */
function avisarPendientes_(cv) {
  try {
    const d = fechaDe_(cv[C_.fecha - 1]);
    if (!d) return;
    const f = Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd');
    const todas = reservas_(f, f), conf = todas.filter(x => x.e === 'C');
    const nuevo = todas.find(x => x.r === String(cv[C_.ref - 1]));
    if (!nuevo) return;
    const s = SpreadsheetApp.getActive().getSheetByName(SH), n = s.getLastRow() - 1;
    const V = s.getRange(2, 1, n, C_.pagado).getValues(), D = s.getRange(2, 1, n, C_.pagado).getDisplayValues();
    todas.filter(x => x.e === 'P' && x.r !== nuevo.r).forEach(p => {
      const i = V.findIndex(r => String(r[C_.ref - 1]) === p.r);
      if (i < 0) return;
      const r = V[i];
      if (!r[C_.email - 1] || !r[C_.token - 1]) return;
      const lleno = choque_(f, p.i, p.d, p.p, p.r);
      const F = { s: s, row: i + 2, v: r, d: D[i] };
      correoUrgencia_(F, lleno);
      hist_(F, lleno ? 'Aviso al cliente: su horario se ha completado (otro grupo confirmó antes)' : 'Aviso al cliente: el horario se está llenando, se le pide que confirme');
    });
  } catch (err) { console.error(err); }
}

/* correoUrgencia_ → ver bloque v22 */

/* ================= Correo al cliente cuando se confirma a mano desde la hoja ================= */
function correoConfirmadoCliente_(F, cuando) {
  const r = F.v;
  if (!r[C_.email - 1]) return;
  const ca = String(r[C_.lang - 1]) === 'ca', T = T_(ca);
  const ref = String(r[C_.ref - 1]), k = String(r[C_.token - 1] || ''), l = ca ? 'ca' : 'es';
  const url = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
  const canc = url + '?a=cancelar&ref=' + encodeURIComponent(ref) + '&k=' + encodeURIComponent(k) + '&l=' + l;
  const mod = WEB + T.web + '?mod=' + encodeURIComponent(ref) + '&k=' + encodeURIComponent(k) + '#calculadora';
  const nom = esc_(String(r[C_.contacto - 1] || '').split(' ')[0]), comida = String(r[C_.comida - 1] || '');
  const X = ca ? {
    s: '✅ Reserva confirmada · ' + ref + ' · Adventure Co.',
    t: 'La teva reserva està confirmada!',
    p: 'Hem confirmat la teva jornada de team building. Aquestes són les dades:',
    k: ['Empresa', 'Jornada', 'Activitat', 'Persones', 'Dinar', 'Càtering'],
    n: 'Uns dies abans de la jornada rebràs un correu de benvinguda amb tots els detalls. Si necessites canviar alguna cosa, fes-ho des d’aquí:',
    mo: 'Modificar la reserva', ce: 'Cancel·lar la reserva',
    wa: 'Dubtes? WhatsApp i telèfon: 600 000 000' } : {
    s: '✅ Reserva confirmada · ' + ref + ' · Adventure Co.',
    t: '¡Tu reserva está confirmada!',
    p: 'Hemos confirmado tu jornada de team building. Estos son los datos:',
    k: ['Empresa', 'Jornada', 'Actividad', 'Personas', 'Comida', 'Catering'],
    n: 'Unos días antes de la jornada recibirás un correo de bienvenida con todos los detalles. Si necesitas cambiar algo, hazlo desde aquí:',
    mo: 'Modificar la reserva', ce: 'Cancelar la reserva',
    wa: '¿Dudas? WhatsApp y teléfono: 600 000 000' };
  const fila = (a, b) => b ? '<tr><td style="padding:7px 12px;color:#6d6a64;font-size:13px;border-bottom:1px solid #eee9df;width:42%">' + a + '</td><td style="padding:7px 12px;font-size:14px;border-bottom:1px solid #eee9df"><b>' + esc_(b) + '</b></td></tr>' : '';
  const html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#3d3a35;max-width:560px;font-size:15px;line-height:1.55">' +
    '<div style="border-bottom:4px solid #ff6b35;padding-bottom:10px;margin-bottom:16px"><img src="https://www.empresa-ejemplo.com/wp-content/uploads/2026/10/logo-empresa.png" alt="Adventure Co." style="height:52px"></div>' +
    '<p style="font-size:20px;font-weight:bold;color:#2d4a1e;margin:0 0 8px">' + X.t + '</p>' +
    '<p>Hola ' + nom + ',<br>' + X.p + '</p>' +
    '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:2px solid #2d4a1e;border-radius:10px">' +
    fila(T.ref, ref) + fila(X.k[0], r[C_.empresa - 1]) + fila(X.k[1], cuando) + fila(X.k[2], r[C_.actividad - 1]) + fila(X.k[3], String(r[C_.personas - 1] || '')) +
    fila(X.k[4], ca ? comidaCa_(comidaTxt_(r)) : comidaTxt_(r)) + fila(X.k[5], String(r[C_.cat - 1] || '')) + '</table>' +
    '<p style="margin-top:16px">' + X.n + '</p>' +
    (k ? '<p><a href="' + mod + '" style="color:#2d4a1e;font-weight:bold">' + X.mo + '</a> &nbsp;·&nbsp; <a href="' + canc + '" style="color:#b3261e">' + X.ce + '</a></p>' : '') +
    '<p style="font-size:13px;color:#6d6a64;border-top:1px solid #e3ded4;padding-top:10px;margin-top:20px">' + X.wa + ' · info@empresa-ejemplo.com' + (r[C_.cat - 1] ? '<br>' + (ca ? 'Càtering' : 'Catering') + ': Restaurante Partner · ' + REST_EMAIL : '') + '<br>Adventure Co. · [Dirección de la empresa]</p></div>';
  try {
    MailApp.sendEmail({ to: String(r[C_.email - 1]), subject: X.s, htmlBody: html, body: X.t + '\n' + ref + ' · ' + cuando, name: 'Adventure Co.', replyTo: 'info@empresa-ejemplo.com' });
    hist_(F, 'Correo de confirmación enviado al cliente');
  } catch (err) { console.error(err); }
}

/* ================= Lugar de la comida ================= */
function lugarTxt_(code, pax) {
  const t = { 'rest': 'Restaurante Partner', 'inst-rest': 'Instalaciones de Adventure Co. · comida del restaurante (la traen)',
    'inst-propia': 'Instalaciones de Adventure Co. · comida propia', 'inst-bbq': 'Instalaciones de Adventure Co. · barbacoa propia' }[String(code || '')];
  const n = Number(pax) || 0;
  return t ? t + (n ? ' · ' + n + ' pers.' : '') : '';
}
function enInst_(r) { return /^Instalaciones/.test(String(r[C_.lugar - 1] || '')); }
function paxDe_(r) { const m = String(r[C_.lugar - 1] || '').match(/(\d+) pers\./); return m ? m[1] : ''; }
function comidaTxt_(r) {
  const l = String(r[C_.lugar - 1] || ''), h = String(r[C_.comida - 1] || '');
  if (!l && !h) return '';
  return (l || 'Restaurante Partner') + (h ? ' · ' + h + ' h' : '');
}
function servicioRest_(r) {
  const h = String(r[C_.comida - 1] || ''), n = paxDe_(r);
  if (!h) return '';
  return (enInst_(r) ? 'LLEVAR LA COMIDA A LAS INSTALACIONES DE ADVENTURE CO. A LAS ' : 'COMEN EN EL RESTAURANTE A LAS ') + h + ' H' + (n ? ' · ' + n + ' PERSONAS' : '');
}
function comidaCa_(t) {
  return String(t).replace('Instalaciones de Adventure Co.', 'Instal·lacions de Adventure Co.').replace('comida del restaurante (la traen)', 'menjar del restaurant (us el porten)')
    .replace('comida propia', 'menjar propi').replace('barbacoa propia', 'barbacoa pròpia');
}
function prepararColumnas() { asegurarColumnas_(SpreadsheetApp.getActive().getSheetByName(SH)); }

/* ================= Pago de la reserva (Adventure Co.) ================= */
// URL de esta app web una vez implementada (Implementar > Nueva implementacion > Aplicacion web).
// Configurala en Propiedades del script > EXEC_URL, o pegala aqui directamente.
const EXEC_URL = PropertiesService.getScriptProperties().getProperty('EXEC_URL') || 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
function pagoKey_(ref) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(ref) + '|pago|' + TOKEN).slice(0, 8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function pagoUrl_(ref) { return EXEC_URL + '?a=pagar&ref=' + encodeURIComponent(ref) + '&p=' + pagoKey_(ref); }
function filaRef_(ref) {
  const s = SpreadsheetApp.getActive().getSheetByName(SH), n = s.getLastRow() - 1;
  if (n < 1 || !ref) return null;
  const V = s.getRange(2, 1, n, C_.pagado).getValues(), D = s.getRange(2, 1, n, C_.pagado).getDisplayValues();
  const i = V.findIndex(r => String(r[C_.ref - 1]) === String(ref));
  return i < 0 ? null : { s: s, row: i + 2, v: V[i], d: D[i] };
}
function infoPago_(F) {
  const r = F.v;
  return 'Referencia: <b>' + esc_(r[C_.ref - 1]) + '</b><br>Empresa: <b>' + esc_(r[C_.empresa - 1]) + '</b><br>Jornada: <b>' + cuando_(F) + '</b><br>Actividades: <b>' + eur_(Number(r[C_.final - 1]) || Number(r[C_.total - 1]) || 0) + '</b>';
}
function htmlMA_(r, txt, ca) {
  const pag = String(r[C_.pagado - 1] || '');
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#3d3a35;line-height:1.55;max-width:600px">' +
    (pag ? '<p style="color:#2e7d32;font-weight:bold">✓ ' + esc_(pag) + '</p>' :
      '<table cellpadding="0" cellspacing="0" border="0" style="margin:4px 0 6px"><tr><td style="background:#2e7d32;border-radius:30px"><a href="' + pagoUrl_(r[C_.ref - 1]) + '" style="display:inline-block;padding:13px 28px;color:#ffffff;text-decoration:none;font-weight:bold;font-size:15px;font-family:Arial,sans-serif">💶 ' + (ca ? 'Marcar com a PAGADA' : 'Marcar como PAGADA') + '</a></td></tr></table>' +
      '<p style="font-size:12.5px;color:#6d6a64;margin:0 0 16px">' + (ca ? 'Prem el botó quan el client hagi pagat les activitats. Se’t demanarà confirmació.' : 'Pulsa el botón cuando el cliente haya pagado las actividades. Se te pedirá confirmación.') + '</p>') +
    esc_(txt).replace(/\n/g, '<br>') + '</div>';
}
function paginaPago_(p) {
  const F = (p.p && p.p === pagoKey_(p.ref)) ? filaRef_(p.ref) : null;
  if (!F) return pagina_(card_('Enlace no válido', 'No hemos encontrado la reserva.', '', 'err', false), 'Pago', false);
  const est = String(F.v[C_.estado - 1]), pag = String(F.v[C_.pagado - 1] || ''), info = infoPago_(F);
  if (pag) return pagina_(card_('Ya estaba marcada como pagada', esc_(pag) + '. No hace falta hacer nada más.', info, 'ok', false), 'Pago', false);
  if (est !== 'Aceptado') return pagina_(card_('La reserva no está confirmada', 'Solo se pueden marcar como pagadas las reservas confirmadas (estado actual: ' + esc_(est) + ').', info, 'full', false), 'Pago', false);
  const body = '<div class="ico" style="background:#2e7d32">€</div><h1 style="color:#2e7d32">Marcar reserva como pagada</h1>' +
    '<p>Pulsa el botón cuando el cliente haya pagado las actividades a Adventure Co..</p><div class="i">' + info + '</div>' +
    '<button id="go" class="b" style="background:#2e7d32">💶 Marcar como pagada</button><p class="er" id="er"></p>' +
    '<div id="md" style="display:none;position:fixed;inset:0;background:rgba(30,40,25,.55);align-items:center;justify-content:center;padding:16px;z-index:9">' +
    '<div style="background:#fff;border-radius:18px;max-width:400px;width:100%;padding:26px 22px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.25)">' +
    '<img src="https://www.empresa-ejemplo.com/wp-content/uploads/2026/10/logo-empresa.png" alt="Adventure Co." style="height:46px;margin-bottom:10px">' +
    '<h2 style="color:#2d4a1e;margin:0 0 8px;font-size:20px">¿Estás seguro de confirmar el pago?</h2>' +
    '<p style="margin:0 0 16px;font-size:14px">' + esc_(F.v[C_.empresa - 1]) + ' · ' + esc_(F.v[C_.ref - 1]) + '<br>La reserva quedará marcada como <b>pagada</b>.</p>' +
    '<div style="display:flex;gap:10px;flex-wrap:wrap"><button id="no" class="b" style="flex:1 1 140px;background:#fff;color:#2d4a1e;border:1px solid #e3ded4">Cancelar</button>' +
    '<button id="si" class="b" style="flex:1 1 140px;background:#2e7d32">Sí, confirmar pago</button></div></div></div>' +
    '<script>var md=document.getElementById("md");document.getElementById("go").onclick=function(){md.style.display="flex"};' +
    'document.getElementById("no").onclick=function(){md.style.display="none"};' +
    'document.getElementById("si").onclick=function(){var b=this;b.disabled=true;b.textContent="Un momento…";' +
    'google.script.run.withSuccessHandler(function(h){md.style.display="none";document.getElementById("c").innerHTML=h;}).withFailureHandler(function(){md.style.display="none";b.disabled=false;b.textContent="Sí, confirmar pago";document.getElementById("er").textContent="No se ha podido registrar. Vuelve a intentarlo.";})' +
    '.ejecutarPago(' + JSON.stringify(String(p.ref)) + ',' + JSON.stringify(String(p.p)) + ');};<\/script>';
  return pagina_(body, 'Marcar como pagada', false);
}
/** Llamado desde la página de pago (google.script.run). */
function ejecutarPago(ref, key) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!key || key !== pagoKey_(ref)) return card_('Enlace no válido', 'No hemos encontrado la reserva.', '', 'err', false);
    const F = filaRef_(ref);
    if (!F) return card_('Enlace no válido', 'No hemos encontrado la reserva.', '', 'err', false);
    const info = infoPago_(F), pag = String(F.v[C_.pagado - 1] || '');
    if (pag) return card_('Ya estaba marcada como pagada', esc_(pag) + '.', info, 'ok', false);
    if (String(F.v[C_.estado - 1]) !== 'Aceptado') return card_('La reserva no está confirmada', 'Solo se pueden marcar como pagadas las reservas confirmadas.', info, 'full', false);
    marcarPagado_(F, 'desde el correo');
    return card_('¡Pago registrado!', 'La reserva ha quedado marcada como pagada y la comisión se ha actualizado en la hoja.', info, 'ok', false);
  } finally { lock.releaseLock(); syncMA_(); }
}
function marcarPagado_(F, via) {
    const txt = 'Pagado · ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm');
    F.s.getRange(F.row, C_.pagado).setValue(txt);
    hist_(F, 'Marcada como PAGADA por Adventure Co. ' + via);
    const r = F.v, com = Math.round((Number(r[C_.final - 1]) || Number(r[C_.total - 1]) || 0) / IVA * 0.1 * 100) / 100;
    envio_('colaborador@example.com', '💶 Reserva PAGADA – ' + r[C_.ref - 1] + ' – ' + r[C_.empresa - 1],
      'Adventure Co. ha marcado la reserva como pagada (' + via + ').\n\nReferencia: ' + r[C_.ref - 1] + '\nEmpresa: ' + r[C_.empresa - 1] + '\nJornada: ' + cuando_(F) +
      '\nActividades: ' + eur_(Number(r[C_.final - 1]) || Number(r[C_.total - 1]) || 0) + '\nComisión actividades (10 % sin IVA): ' + eur_(com) + '\n\nYa está sumada en la hoja (Resumen › Comisión de reservas pagadas).', null);
}


/* ================= Reenvío del presupuesto a Cliente Ejemplo S.L. (TB-240101-0001) ================= */
function prepararLaura_() {
  const F = filaRef_('TB-240101-0001');
  if (!F) throw new Error('No encuentro la reserva TB-240101-0001');
  let k = String(F.v[C_.token - 1] || '');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 24); F.s.getRange(F.row, C_.token).setValue(k); }
  if (!F.v[C_.lang - 1]) F.s.getRange(F.row, C_.lang).setValue('ca');
  if (!F.v[C_.cat - 1]) F.s.getRange(F.row, C_.cat).setValue('Esmorzar opció 2 × 27');
  if (!F.v[C_.snap - 1]) {
    const o = { ppl: '27', acomp: '0', porche: false, poN: '27', transfer: false, trN: '27', desay: true, comensD: '27', catering: false, comens: '27', calcot: false, comensC: '27', ninosC: '0',
      r_act: '22', r_trw: '1', r_des: '9.5', r_menu: '22', f_empresa: 'Cliente Ejemplo S.L.', f_contacto: 'Laura', f_email: 'cliente@empresa-ejemplo.com', f_tel: '600000000', f_coment: '',
      f_factura: false, f_copia: true, obj: [], fecha: '2026-12-17', hora: '', meal: '', loc: '', food: '', ed: [false, false, false, false, false] };
    F.s.getRange(F.row, C_.snap).setValue(JSON.stringify(o));
  }
  return { F: filaRef_('TB-240101-0001'), k: k };
}
function correoLaura_(to, ejemplo, nota) {
  const P = prepararLaura_(), r = P.F.v, ref = String(r[C_.ref - 1]);
  const mod = WEB + '/team-building-ca/?mod=' + encodeURIComponent(ref) + '&k=' + encodeURIComponent(P.k) + '#calculadora';
  const fila = (a, b) => '<tr><td style="padding:8px 12px;border-bottom:1px solid #eee9df;font-size:14px">' + a + '</td><td style="padding:8px 12px;border-bottom:1px solid #eee9df;font-size:14px;text-align:right;white-space:nowrap"><b>' + b + '</b></td></tr>';
  const subj = (ejemplo ? '[EJEMPLO] ' : '') + 'La vostra pre-reserva i pressupost orientatiu – Adventure Co. · ' + ref;
  const btn = (txt) => '<table cellpadding="0" cellspacing="0" border="0" style="margin:18px 0"><tr><td style="background:#2d4a1e;border-radius:30px"><a href="' + mod + '" style="display:inline-block;padding:14px 30px;color:#ffffff;text-decoration:none;font-weight:bold;font-size:16px;font-family:Arial,sans-serif">' + txt + '</a></td></tr></table>';
  const html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#3d3a35;max-width:560px;font-size:15px;line-height:1.55">' +
    '<div style="border-bottom:4px solid #ff6b35;padding-bottom:10px;margin-bottom:16px"><img src="https://www.empresa-ejemplo.com/wp-content/uploads/2026/10/logo-empresa.png" alt="Adventure Co." style="height:52px"></div>' +
    '<p>Hola Laura,</p>' +
    '<p>moltes gràcies per la vostra pre-reserva de team building. Et volem explicar per què no t’havia arribat el <b>pressupost orientatiu</b>: just quan la vau fer, estàvem fent millores al sistema de reserves de la web, i la sol·licitud ens va arribar sense enviar-te el pressupost automàticament. A més, hi falten dues dades: <b>l’hora de l’activitat</b> i <b>on voleu esmorzar</b>.</p>' +
    '<p>Perdona les molèsties. Aquests són els detalls de la vostra pre-reserva:</p>' +
    '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #e3ded4">' +
    fila('Empresa', 'Cliente Ejemplo S.L.') + fila('Data', 'Dijous, 17 de desembre de 2026') + fila('Hora de l’activitat', '<span style="color:#e8704e">Pendent</span>') +
    fila('Orientació · 27 × 22 €', '594 €') + fila('Esmorzar opció 2 (Restaurante Partner) · 27 × 9,50 €', '256,50 €') +
    fila('Lloc de l’esmorzar', '<span style="color:#e8704e">Pendent</span>') +
    '<tr><td style="padding:10px 12px;font-size:15px;color:#2d4a1e"><b>Total orientatiu</b></td><td style="padding:10px 12px;text-align:right;font-size:16px;color:#e8704e"><b>850,50 €</b></td></tr></table>' +
    '<p style="font-size:13px;color:#6d6a64;margin-top:8px">IVA inclòs. Adventure Co. i el Restaurante Partner són dues empreses independents: el pressupost és conjunt, però cada una emet la seva factura (activitats 594 € · esmorzar 256,50 €).</p>' +
    '<p style="margin-bottom:6px"><b>On voleu esmorzar?</b></p>' +
    '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #e3ded4;margin-bottom:14px">' +
    fila('🍽️ Al Restaurante Partner<br><span style="font-size:12.5px;color:#6d6a64">Sense cost d’espai · opció de transfer: 1 € per persona i trajecte</span>', '0 €') +
    fila('🏕️ A les nostres instal·lacions<br><span style="font-size:12.5px;color:#6d6a64">Espai + porxo (sala de 160 m² + porxo)</span>', '+6 €/pers.') + '</table>' +
    '<p>Amb el nou sistema <b>no cal que truqueu al restaurant</b>: quan confirmeu la reserva, l’encàrrec de l’esmorzar li arriba automàticament.</p>' +
    '<p>Per completar-la només has de prémer el botó: tens totes les dades ja omplertes i només cal que triïs <b>l’hora de l’activitat</b> i on mengeu. Si marqueu <b>«A les nostres instal·lacions»</b>, s’hi afegirà automàticament el cost de l’espai + porxo (6 €/pers.). Si esmorzeu al restaurant, marqueu «Sense dinar» i, si el necessiteu, el «Transfer» als extres. En acabar, rebràs per correu el pressupost final actualitzat, amb el botó per confirmar la reserva.</p>' +
    btn('Completar la meva pre-reserva') +
    '<p style="background:#fff8e1;border-left:4px solid #e8704e;padding:10px 12px;font-size:14px"><b>Important:</b> la pre-reserva encara no és una reserva confirmada. La reserva queda feta quan premeu <b>«Confirmar reserva»</b> des del correu o del PDF del pressupost.</p>' +
    '<p style="font-size:13.5px">És sense compromís. Si vols canviar alguna cosa (persones, activitat, esmorzar…), també ho pots fer des del mateix enllaç.</p>' +
    '<p style="font-size:13px;color:#6d6a64;border-top:1px solid #e3ded4;padding-top:10px;margin-top:20px">Referència: ' + ref + '<br>Dubtes? WhatsApp i telèfon: 600 000 000 · info@empresa-ejemplo.com<br>Adventure Co. · [Dirección de la empresa]</p></div>';
  if (nota) {
    const pre = '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;background:#fff8e1;border:1px solid #f0d9a8;border-radius:10px;padding:14px 16px;margin-bottom:22px;font-size:14px;color:#3d3a35;line-height:1.5">' + nota + '</div>';
    MailApp.sendEmail({ to: to, subject: 'PROPUESTA de correo para Laura (Cliente Ejemplo S.L.) – revisar y enviar', htmlBody: pre + html, body: 'Propuesta de correo para Laura. Asunto sugerido: ' + subj + '\n' + mod, name: 'Sistema de reservas · Team Building' });
    return P.F;
  }
  MailApp.sendEmail({ to: to, subject: subj, htmlBody: html, body: subj + '\n' + mod, name: 'Adventure Co.', replyTo: 'info@empresa-ejemplo.com' });
  return P.F;
}
function propuestaLaura() {
  correoLaura_('info@empresa-ejemplo.com', false, '<b>Hola equipo,</b><br>esta es una <b>propuesta de correo para Laura</b> (Cliente Ejemplo S.L. · TB-240101-0001). Su pre-reserva llegó mientras se estaba cambiando el sistema de la web y no le llegó el presupuesto orientativo.<br><br><b>Cómo enviarlo:</b> pulsa <b>Reenviar</b>, borra este recuadro amarillo, cambia lo que quieras y envíalo a <b>cliente@empresa-ejemplo.com</b>.<br>Asunto sugerido: <i>La vostra pre-reserva i pressupost orientatiu – Adventure Co. · TB-240101-0001</i><br><br>El botón verde ya lleva el enlace personal de Laura: si lo abres para probarlo, no envíes el formulario (cambiaría su reserva).');
}
function ejemploLaura() { correoLaura_('colaborador@example.com', true); }
function enviarLaura() { const F = correoLaura_('cliente@empresa-ejemplo.com', false); hist_(F, 'Presupuesto reenviado al cliente (con enlace para elegir la hora y confirmar)'); }


/* ================= Hoja para Adventure Co. (reservas confirmadas, en tiempo real) =================
 *  1) Ejecutar crearHojaEmpresa() una sola vez → crea la hoja, la conecta y guarda el enlace en Ajustes (C7).
 *  2) Compartir esa hoja con Adventure Co. (Editor). Solo tienen que cambiar la columna "Pago".
 */
const MA_TAB = 'Team Building';
const MA_HEAD = ['Referencia', 'Fecha jornada', 'Horario', 'Empresa', 'Contacto', 'Teléfono', 'Email', 'Personas',
  'Actividad', 'Comida', 'Importe actividades (€)', 'Pago', 'Fecha de pago', 'Detalle'];
const MA_PAGO = 12;
function main_() {
  const id = PropertiesService.getScriptProperties().getProperty('MAIN_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}
function hojaMA_() {
  const id = PropertiesService.getScriptProperties().getProperty('MA_ID');
  if (!id) return null;
  const ss = SpreadsheetApp.openById(id);
  return ss.getSheetByName(MA_TAB);
}
function crearHojaEmpresa() {
  const main = SpreadsheetApp.getActive(), P = PropertiesService.getScriptProperties();
  P.setProperty('MAIN_ID', main.getId());
  let ss = P.getProperty('MA_ID') ? SpreadsheetApp.openById(P.getProperty('MA_ID')) : null;
  if (!ss) { ss = SpreadsheetApp.create('Adventure Co. · Reservas confirmadas'); P.setProperty('MA_ID', ss.getId()); }
  ss.setSpreadsheetTimeZone('Europe/Madrid');
  let s = ss.getSheetByName(MA_TAB);
  if (!s) { s = ss.getSheets()[0]; s.setName(MA_TAB); }
  s.clear(); s.clearConditionalFormatRules();
  s.getRange(1, 1, 1, MA_HEAD.length).setValues([MA_HEAD]).setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  s.setRowHeight(1, 42); s.setFrozenRows(1); s.setFrozenColumns(1);
  [120, 110, 110, 190, 150, 115, 200, 80, 220, 260, 120, 120, 135, 360].forEach((w, i) => s.setColumnWidth(i + 1, w));
  if (s.getMaxColumns() > MA_HEAD.length) s.deleteColumns(MA_HEAD.length + 1, s.getMaxColumns() - MA_HEAD.length);
  s.getRange(1, MA_PAGO).setNote('Cambiad a "Pagada" cuando el cliente haya pagado las actividades. Es lo mismo que el botón "Marcar como PAGADA" del correo.');
  s.getRange(1, MA_PAGO).setBackground(NARANJA);
  const a = main.getSheetByName(AJ) || main.insertSheet(AJ);
  a.getRange('B7').setValue('Hoja de Adventure Co. (reservas confirmadas)').setFontWeight('bold').setFontColor(VERDE);
  a.getRange('C7').setValue(ss.getUrl());
  // triggers: edición en la hoja de Empresa + sincronización periódica
  ScriptApp.getProjectTriggers().filter(t => ['alEditarMA', 'sincronizarMA'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alEditarMA').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('sincronizarMA').timeBased().everyMinutes(5).create();
  sincronizarMA();
  console.log('Hoja creada: ' + ss.getUrl());
}
function sincronizarMA() {
  const PR = PropertiesService.getScriptProperties();
  if (!PR.getProperty('COM_SIN_IVA')) { PR.setProperty('COM_SIN_IVA', '1'); comisionSinIVA(); return; }
  if (!PR.getProperty('V18')) { PR.setProperty('V18', '1'); actualizarV18(); }
  if (!PR.getProperty('V22')) { PR.setProperty('V22', '1'); instalarTareasDiarias(); }
  if (!PR.getProperty('V25')) { PR.setProperty('V25', '1'); crearHojaRestaurante(); }
  if (!PR.getProperty('V32')) { PR.setProperty('V30', '1'); PR.setProperty('V32', '1'); PR.setProperty('REST_MAIL', '1'); crearHojaRestaurante(); }
  if (!PR.getProperty('V30')) { PR.setProperty('V28', '1'); PR.setProperty('V29', '1'); PR.setProperty('V30', '1'); PR.setProperty('REST_MAIL', '1'); crearHojaRestaurante(); }
  if (!PR.getProperty('V29')) { PR.setProperty('V28', '1'); PR.setProperty('V29', '1'); PR.setProperty('REST_MAIL', '1'); crearHojaRestaurante(); }
  if (!PR.getProperty('V28')) { PR.setProperty('V26', '1'); PR.setProperty('V27', '1'); PR.setProperty('V28', '1'); PR.setProperty('REST_MAIL', '1'); crearHojaRestaurante(); }
  try { sincronizarRest_(); } catch (err) { console.error(err); }
  const h = hojaMA_();
  if (!h) return;
  const s = main_().getSheetByName(SH), n = s.getLastRow() - 1;
  const V = n > 0 ? s.getRange(2, 1, n, C_.pagado).getValues() : [], D = n > 0 ? s.getRange(2, 1, n, C_.pagado).getDisplayValues() : [];
  const rows = [];
  V.forEach((r, i) => {
    if (String(r[C_.estado - 1]) !== 'Aceptado' || !r[C_.ref - 1]) return;
    const d = fechaDe_(r[C_.fecha - 1]), hm = horaDe_(D[i][C_.hora - 1]), dur = Number(r[C_.dur - 1]) || 3;
    const pag = String(r[C_.pagado - 1] || ''), lug = String(r[C_.lugar - 1] || ''), hc = String(D[i][C_.comida - 1] || '');
    rows.push([String(r[C_.ref - 1]), d || '', hm ? hm.txt + '–' + finTxt_(hm, dur) : '', r[C_.empresa - 1], r[C_.contacto - 1], D[i][C_.tel - 1], r[C_.email - 1],
      r[C_.personas - 1], r[C_.actividad - 1], lug ? lug + (hc ? ' · ' + hc + ' h' : '') : 'Sin comida',
      Number(r[C_.final - 1]) || Number(r[C_.total - 1]) || 0, pag ? 'Pagada' : 'Pendiente', pag ? pag.replace(/^Pagado · /, '') : '',
      String(r[C_.detalle - 1] || '')]);
  });
  rows.sort((x, y) => (x[1] ? x[1].getTime() : 0) - (y[1] ? y[1].getTime() : 0));
  const last = h.getLastRow();
  if (last > 1) { h.getRange(2, 1, last - 1, MA_HEAD.length).clearContent().clearDataValidations(); }
  if (rows.length) {
    h.getRange(2, 1, rows.length, MA_HEAD.length).setValues(rows).setVerticalAlignment('top').setWrap(true);
    h.getRange(2, 2, rows.length, 1).setNumberFormat('dd/MM/yyyy');
    h.getRange(2, 6, rows.length, 1).setNumberFormat('@');
    h.getRange(2, 11, rows.length, 1).setNumberFormat('#,##0.00 €');
    h.getRange(2, MA_PAGO, rows.length, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente', 'Pagada'], true).setAllowInvalid(false).build()).setHorizontalAlignment('center').setFontWeight('bold');
  }
  const rg = h.getRange(2, 1, Math.max(rows.length, 1), MA_HEAD.length), pg = h.getRange(2, MA_PAGO, Math.max(rows.length, 1), 1);
  h.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pagada').setBackground('#d9ead3').setFontColor('#1e5631').setRanges([pg]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pendiente').setBackground('#fff4e5').setFontColor('#b7791f').setRanges([pg]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=ISEVEN(ROW())').setBackground(CREMA).setRanges([rg]).build()]);
  SpreadsheetApp.flush();
}
function syncMA_() { try { sincronizarMA(); } catch (err) { console.error(err); } }

/** Cambio en la hoja de Adventure Co.: columna "Pago" */
function alEditarMA(e) {
  const r = e.range, h = r.getSheet();
  if (h.getName() !== MA_TAB || r.getRow() < 2) return;
  if (r.getNumRows() !== 1 || r.getNumColumns() !== 1 || r.getColumn() !== MA_PAGO) { syncMA_(); return; }
  const ref = String(h.getRange(r.getRow(), 1).getValue()), nuevo = String(e.value || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = main_().getSheetByName(SH), n = s.getLastRow() - 1;
    const V = s.getRange(2, 1, n, C_.pagado).getValues(), D = s.getRange(2, 1, n, C_.pagado).getDisplayValues();
    const i = V.findIndex(x => String(x[C_.ref - 1]) === ref);
    if (i < 0) return;
    const F = { s: s, row: i + 2, v: V[i], d: D[i] }, pag = String(V[i][C_.pagado - 1] || '');
    if (nuevo === 'Pagada' && !pag) marcarPagado_(F, 'desde su hoja');
    else if (nuevo === 'Pendiente' && pag) {
      s.getRange(F.row, C_.pagado).setValue('');
      hist_(F, 'Pago DESMARCADO por Adventure Co. desde su hoja');
      envio_('colaborador@example.com', '↩️ Pago desmarcado – ' + V[i][C_.ref - 1] + ' – ' + V[i][C_.empresa - 1],
        'Adventure Co. ha vuelto a poner la reserva como PENDIENTE de pago desde su hoja.\n\nReferencia: ' + V[i][C_.ref - 1] + '\nEmpresa: ' + V[i][C_.empresa - 1] + '\nJornada: ' + cuando_(F), null);
    }
  } finally { lock.releaseLock(); syncMA_(); }
}

/** Ejecutar una vez: pasa la comisión de actividades a 10 % sobre la base sin IVA en todas las filas. */
function comisionSinIVA() {
  const ss = main_(), s = ss.getSheetByName(SH);
  s.getRange(1, 16).setValue('Comisión 10 % sin IVA (€)');
  for (let r = 2; r <= s.getLastRow(); r++) if (s.getRange(r, C_.ref).getValue()) s.getRange(r, 16).setFormulaR1C1(FCOMACT_);
  const rs = ss.getSheetByName(RS);
  if (rs) rs.createTextFinder('Comisión 10 %').matchEntireCell(true).findAll().forEach(c => c.setValue('Comisión actividades (10 % sin IVA)'));
  s.getRange('P1').setNote('10 % sobre el importe de actividades SIN IVA (importe ÷ 1,21). Usa "Importe final actividades" si está relleno; si no, el presupuesto.');
  syncMA_();
}

/** Ejecutada una vez (automáticamente): email del restaurante y pestaña 'Días bloqueados'. */
const REST_EMAIL = 'restaurante@example.com';
function actualizarV18() {
  const a = main_().getSheetByName(AJ);
  if (a && !String(a.getRange('C3').getValue() || '').trim()) a.getRange('C3').setValue(REST_EMAIL);
  const id = PropertiesService.getScriptProperties().getProperty('MA_ID'); if (!id) return;
  const ss = SpreadsheetApp.openById(id); let h = ss.getSheetByName(BLOQ_TAB);
  if (!h) {
    h = ss.insertSheet(BLOQ_TAB);
    h.getRange(1, 1, 1, 2).setValues([['Fecha', 'Motivo (casal, escuelas, evento…)']]).setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE);
    h.setFrozenRows(1); h.setColumnWidth(1, 130); h.setColumnWidth(2, 360);
    h.getRange('A2:A').setNumberFormat('dd/MM/yyyy').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
    h.getRange('A1').setNote('Escribe aquí los días que estén ocupados todo el día. En la calculadora de la web aparecerán como completos. También vale crear un evento de "todo el día" en el calendario Team Building.');
  }
}

/* ================= Correos automáticos al cliente y avisos a Empresa (v22) ================= */
function mailUrl_() { return 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec'; }
function fechaLarga_(d, ca) {
  if (!d) return '';
  const dias = ca ? ['diumenge', 'dilluns', 'dimarts', 'dimecres', 'dijous', 'divendres', 'dissabte'] : ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
  const mes = ca ? ['de gener', 'de febrer', 'de març', 'd’abril', 'de maig', 'de juny', 'de juliol', 'd’agost', 'de setembre', 'd’octubre', 'de novembre', 'de desembre']
    : ['de enero', 'de febrero', 'de marzo', 'de abril', 'de mayo', 'de junio', 'de julio', 'de agosto', 'de septiembre', 'de octubre', 'de noviembre', 'de diciembre'];
  return dias[d.getDay()] + ' ' + d.getDate() + ' ' + mes[d.getMonth()];
}
function mailDatos_(F, ca) {
  const r = F.v;
  const fila = (a, b) => b ? '<tr><td style="padding:7px 12px;color:#6d6a64;font-size:13px;border-bottom:1px solid #eee9df;width:42%">' + a + '</td><td style="padding:7px 12px;font-size:14px;border-bottom:1px solid #eee9df"><b>' + esc_(b) + '</b></td></tr>' : '';
  const com = comidaTxt_(r), rest = Number(r[C_.rest - 1]) || 0, act = Number(r[C_.final - 1]) || Number(r[C_.total - 1]) || 0;
  return '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:2px solid #2d4a1e;border-radius:10px">' +
    fila(ca ? 'Referència' : 'Referencia', r[C_.ref - 1]) + fila('Empresa', r[C_.empresa - 1]) + fila(ca ? 'Jornada' : 'Jornada', cuando_(F)) +
    fila(ca ? 'Activitat' : 'Actividad', r[C_.actividad - 1] + ' · ' + r[C_.personas - 1] + (ca ? ' persones' : ' personas')) +
    fila(ca ? 'Dinar' : 'Comida', com ? (ca ? comidaCa_(com) : com) : '') + fila(ca ? 'Càtering' : 'Catering', String(r[C_.cat - 1] || '')) +
    fila(ca ? 'Total orientatiu' : 'Total orientativo', eur_(act) + (ca ? ' activitats' : ' actividades') + (rest ? ' · ' + eur_(rest) + (ca ? ' càtering' : ' catering') : '')) +
    fila(ca ? 'IVA inclòs' : 'IVA incluido', (ca ? 'activitats 21 %: ' : 'actividades 21 %: ') + eur_(Math.round((act - act / IVA) * 100) / 100) + (rest ? (ca ? ' · càtering 10 %: ' : ' · catering 10 %: ') + eur_(iv10_(rest)[1]) : '')) + '</table>';
}
function mailMarco_(titulo, cuerpo, F, ca) {
  const r = F.v;
  return '<div style="font-family:Arial,Helvetica,sans-serif;color:#3d3a35;max-width:560px;font-size:15px;line-height:1.55">' +
    '<div style="border-bottom:4px solid #ff6b35;padding-bottom:10px;margin-bottom:16px"><img src="https://www.empresa-ejemplo.com/wp-content/uploads/2026/10/logo-empresa.png" alt="Adventure Co." style="height:52px"></div>' +
    '<p style="font-size:20px;font-weight:bold;color:#2d4a1e;margin:0 0 8px">' + titulo + '</p>' + cuerpo +
    '<p style="font-size:13px;color:#6d6a64;border-top:1px solid #e3ded4;padding-top:10px;margin-top:20px">' + (ca ? 'Referència' : 'Referencia') + ': ' + r[C_.ref - 1] + '<br>' +
    (ca ? 'Dubtes? WhatsApp i telèfon: 600 000 000' : '¿Dudas? WhatsApp y teléfono: 600 000 000') + ' · info@empresa-ejemplo.com' +
    (r[C_.cat - 1] ? '<br>' + (ca ? 'Càtering' : 'Catering') + ': Restaurante Partner · ' + REST_EMAIL : '') +
    '<br>Adventure Co. · [Dirección de la empresa]</p></div>';
}
function mailBtn_(href, txt, col) {
  return '<table cellpadding="0" cellspacing="0" border="0" style="margin:18px 0 10px"><tr><td style="background:' + (col || '#2d4a1e') + ';border-radius:30px"><a href="' + href + '" style="display:inline-block;padding:14px 30px;color:#ffffff;text-decoration:none;font-weight:bold;font-size:16px;font-family:Arial,sans-serif">' + txt + '</a></td></tr></table>';
}
function mailLinks_(F, ca) {
  const r = F.v, ref = String(r[C_.ref - 1]), k = String(r[C_.token - 1] || ''), T = T_(ca);
  return { conf: mailUrl_() + '?a=confirmar&ref=' + encodeURIComponent(ref) + '&k=' + encodeURIComponent(k) + '&l=' + (ca ? 'ca' : 'es'),
    canc: mailUrl_() + '?a=cancelar&ref=' + encodeURIComponent(ref) + '&k=' + encodeURIComponent(k) + '&l=' + (ca ? 'ca' : 'es'),
    mod: WEB + T.web + '?mod=' + encodeURIComponent(ref) + '&k=' + encodeURIComponent(k) + '#calculadora' };
}
function mailAviso_(ca) {
  return '<p style="background:#fff8e1;border-left:4px solid #e8704e;padding:10px 12px;font-size:14px">' + (ca
    ? '<b>Recorda:</b> la pre-reserva encara no és una reserva confirmada. La data i l’hora només queden reservades quan prems «Confirmar reserva».'
    : '<b>Recuerda:</b> la pre-reserva todavía no es una reserva confirmada. La fecha y la hora solo quedan reservadas cuando pulsas «Confirmar reserva».') + '</p>' +
    '<p style="font-size:14px">' + (ca ? 'Tens dubtes o vols que t’ajudem amb algun canvi? Escriu-nos a ' : '¿Tienes dudas o quieres que te ayudemos con algún cambio? Escríbenos a ') +
    '<a href="mailto:info@empresa-ejemplo.com" style="color:#2d4a1e;font-weight:bold">info@empresa-ejemplo.com</a>' + (ca ? ' i ho resolem amb tu.' : ' y lo resolvemos contigo.') + '</p>';
}
function enviarCliente_(F, asunto, html) {
  const r = F.v;
  try { MailApp.sendEmail({ to: String(r[C_.email - 1]), subject: asunto, htmlBody: html, body: asunto, name: 'Adventure Co.', replyTo: 'info@empresa-ejemplo.com' }); return true; }
  catch (err) { console.error(err); return false; }
}
/** Otro grupo ha confirmado el mismo día: se está llenando (lleno=false) o su hora ya está ocupada (lleno=true) */
function correoUrgencia_(F, lleno) {
  const r = F.v, ca = String(r[C_.lang - 1]) === 'ca', ref = String(r[C_.ref - 1]), L = mailLinks_(F, ca);
  const nom = esc_(String(r[C_.contacto - 1] || '').split(' ')[0]), d = fechaDe_(r[C_.fecha - 1]), hm = horaDe_(F.d[C_.hora - 1]);
  const dia = fechaLarga_(d, ca), hora = hm ? hm.txt : '';
  let asunto, titulo, cuerpo;
  if (!lleno) {
    asunto = ca ? '⏳ El dia de la teva pre-reserva s’està omplint – confirma o modifica · ' + ref : '⏳ El día de tu pre-reserva se está llenando – confirma o modifica · ' + ref;
    titulo = ca ? 'El dia de la teva pre-reserva s’està omplint' : 'El día de tu pre-reserva se está llenando';
    cuerpo = (ca
      ? '<p>Hola ' + nom + ',</p><p>un altre grup acaba de <b>confirmar la seva reserva el mateix dia</b> que la teva pre-reserva (<b>' + dia + '</b>). Atenem un sol grup alhora i les hores s’assignen per ordre de confirmació.</p>' +
        '<p>La teva hora' + (hora ? ' (' + hora + ')' : '') + ' encara està lliure. Pots <b>confirmar-la ara</b> per no quedar-te sense plaça o, si et va millor, <b>modificar la reserva</b> i triar una altra hora o un altre dia.</p>'
      : '<p>Hola ' + nom + ',</p><p>otro grupo acaba de <b>confirmar su reserva el mismo día</b> que tu pre-reserva (<b>' + dia + '</b>). Atendemos a un solo grupo a la vez y las horas se asignan por orden de confirmación.</p>' +
        '<p>Tu hora' + (hora ? ' (' + hora + ')' : '') + ' todavía está libre. Puedes <b>confirmarla ahora</b> para no quedarte sin plaza o, si te va mejor, <b>modificar la reserva</b> y elegir otra hora u otro día.</p>') +
      mailDatos_(F, ca) + mailBtn_(L.conf, ca ? '✓ Confirmar la meva reserva' : '✓ Confirmar mi reserva') +
      '<p style="font-size:14px;margin-top:0">' + (ca ? 'Prefereixes una altra hora o data? ' : '¿Prefieres otra hora o fecha? ') + '<a href="' + L.mod + '" style="color:#2d4a1e;font-weight:bold">' + (ca ? 'Modificar la reserva' : 'Modificar la reserva') + '</a></p>' + mailAviso_(ca);
  } else {
    asunto = ca ? '📅 La teva hora ja està ocupada – tria una altra hora o data · ' + ref : '📅 Tu hora ya está ocupada – elige otra hora o fecha · ' + ref;
    titulo = ca ? 'L’hora de la teva pre-reserva ja està ocupada' : 'La hora de tu pre-reserva ya está ocupada';
    cuerpo = (ca
      ? '<p>Hola ' + nom + ',</p><p>ho sentim: un altre grup ha <b>confirmat abans</b> l’hora que vas demanar (<b>' + dia + (hora ? ' · ' + hora : '') + '</b>), i només atenem un grup alhora.</p>' +
        '<p>La teva pre-reserva continua guardada amb totes les dades: només has de <b>modificar-la i triar una altra hora o un altre dia</b> al calendari. Et portarà un minut.</p>'
      : '<p>Hola ' + nom + ',</p><p>lo sentimos: otro grupo ha <b>confirmado antes</b> la hora que pediste (<b>' + dia + (hora ? ' · ' + hora : '') + '</b>), y solo atendemos a un grupo a la vez.</p>' +
        '<p>Tu pre-reserva sigue guardada con todos los datos: solo tienes que <b>modificarla y elegir otra hora u otro día</b> en el calendario. Te llevará un minuto.</p>') +
      mailDatos_(F, ca) + mailBtn_(L.mod, ca ? 'Triar una altra hora o data' : 'Elegir otra hora o fecha', '#e8704e') + mailAviso_(ca);
  }
  enviarCliente_(F, asunto, mailMarco_(titulo, cuerpo, F, ca));
}
/** Recordatorio 3 días antes de que caduque la pre-reserva (15 días) */
function correoCaduca_(F, limite) {
  const r = F.v, ca = String(r[C_.lang - 1]) === 'ca', ref = String(r[C_.ref - 1]), L = mailLinks_(F, ca);
  const nom = esc_(String(r[C_.contacto - 1] || '').split(' ')[0]), lim = fechaLarga_(limite, ca);
  const asunto = ca ? '⏰ La teva pre-reserva caduca el ' + lim + ' – confirma-la · ' + ref : '⏰ Tu pre-reserva caduca el ' + lim + ' – confírmala · ' + ref;
  const cuerpo = (ca
    ? '<p>Hola ' + nom + ',</p><p>fa uns dies ens vas demanar el pressupost per a la vostra jornada de team building. El pressupost orientatiu és vàlid <b>15 dies</b>, així que la teva pre-reserva <b>caduca el ' + lim + '</b>.</p>' +
      '<p>Recorda que la pre-reserva encara <b>no és una reserva confirmada</b>: la data i l’hora només queden reservades quan prems «Confirmar reserva».</p>'
    : '<p>Hola ' + nom + ',</p><p>hace unos días nos pediste el presupuesto para vuestra jornada de team building. El presupuesto orientativo es válido <b>15 días</b>, así que tu pre-reserva <b>caduca el ' + lim + '</b>.</p>' +
      '<p>Recuerda que la pre-reserva todavía <b>no es una reserva confirmada</b>: la fecha y la hora solo quedan reservadas cuando pulsas «Confirmar reserva».</p>') +
    mailDatos_(F, ca) + mailBtn_(L.conf, ca ? '✓ Confirmar la meva reserva' : '✓ Confirmar mi reserva') +
    '<p style="background:#fff8e1;border-left:4px solid #e8704e;padding:10px 12px;font-size:14px">' + (ca
      ? 'Si no la confirmes abans del <b>' + lim + '</b>, la pre-reserva s’anul·larà i l’hora quedarà lliure per a altres grups.'
      : 'Si no la confirmas antes del <b>' + lim + '</b>, la pre-reserva se anulará y la hora quedará libre para otros grupos.') + '</p>' +
    '<p style="font-size:13.5px">' + (ca ? 'Prefereixes canviar alguna cosa? ' : '¿Prefieres cambiar algo? ') + '<a href="' + L.mod + '" style="color:#2d4a1e;font-weight:bold">Modificar la reserva</a> &nbsp;·&nbsp; <a href="' + L.canc + '" style="color:#b3261e">' + (ca ? 'Cancel·lar la pre-reserva' : 'Cancelar la pre-reserva') + '</a></p>' +
    '<p style="font-size:14px">' + (ca ? 'Tens dubtes? Escriu-nos a ' : '¿Tienes dudas? Escríbenos a ') + '<a href="mailto:info@empresa-ejemplo.com" style="color:#2d4a1e;font-weight:bold">info@empresa-ejemplo.com</a>.</p>';
  return enviarCliente_(F, asunto, mailMarco_(ca ? 'La teva pre-reserva caduca aviat' : 'Tu pre-reserva caduca pronto', cuerpo, F, ca));
}
/** Aviso a Empresa (info@) unos días antes: botón para escribir el correo de bienvenida (en el idioma del cliente) */
function avisoBienvenida_(F, dias) {
  const r = F.v, ca = String(r[C_.lang - 1]) === 'ca';
  const emp = String(r[C_.empresa - 1] || ''), cont = String(r[C_.contacto - 1] || ''), nom = cont.split(' ')[0];
  const mail = String(r[C_.email - 1] || ''), d = fechaDe_(r[C_.fecha - 1]), hm = horaDe_(F.d[C_.hora - 1]);
  const dia = fechaLarga_(d, ca), hora = hm ? hm.txt : '';
  const X = ca ? {
    cu: dias === 0 ? 'avui' : dias === 1 ? 'demà' : 'd’aquí a ' + dias + ' dies', av: 'AVÍS INTERN', toca: 'Toca escriure a ',
    es: 'La seva jornada és ', env: '. Envia’ls el correu de benvinguda.', jo: 'Jornada', co: 'Contacte', gr: 'Grup', pe: ' persones · ', di: 'Dinar', ca: 'Càtering', cm: 'Comentaris',
    pa: 'Només 2 passos:', p1: '1. Prem el botó verd. S’obre un correu nou per a ', p1b: ', amb l’assumpte ja posat.', p2: '2. Escriu el teu missatge i prem <b>Enviar</b>.',
    bt: 'Escriure a ', no: 'No s’obre? <a href="{mt}" style="color:#2d4a1e">Prova amb aquest altre enllaç</a> o escriu a ', nr: 'No responguis aquest avís: li arribaria a Empresa, no al client.',
    id: '<b>Idees per al correu:</b> hora d’arribada · com arribar / ubicació · roba i calçat còmodes · el vostre telèfon 600 000 000.', pie: ' · Aquest avís només arriba un cop per grup.',
    su: 'Benvinguda a Adventure Co. – la vostra jornada del ', as: 'Toca escriure a ', ja: ' – jornada ', cc: 'Correu del client: ', ass: 'Assumpte suggerit: ' } : {
    cu: dias === 0 ? 'hoy' : dias === 1 ? 'mañana' : 'dentro de ' + dias + ' días', av: 'AVISO INTERNO', toca: 'Toca escribir a ',
    es: 'Su jornada es ', env: '. Envíales el correo de bienvenida.', jo: 'Jornada', co: 'Contacto', gr: 'Grupo', pe: ' personas · ', di: 'Comida', ca: 'Catering', cm: 'Comentarios',
    pa: 'Solo 2 pasos:', p1: '1. Pulsa el botón verde. Se abre un correo nuevo para ', p1b: ', con el asunto ya puesto.', p2: '2. Escribe tu mensaje y pulsa <b>Enviar</b>.',
    bt: 'Escribir a ', no: '¿No se abre? <a href="{mt}" style="color:#2d4a1e">Prueba con este otro enlace</a> o escribe a ', nr: 'No respondas a este aviso: le llegaría a Empresa, no al cliente.',
    id: '<b>Ideas para el correo:</b> hora de llegada · cómo llegar / ubicación · ropa y calzado cómodos · vuestro teléfono 600 000 000.', pie: ' · Este aviso solo llega una vez por grupo.',
    su: 'Bienvenida a Adventure Co. – vuestra jornada del ', as: 'Toca escribir a ', ja: ' – jornada ', cc: 'Correo del cliente: ', ass: 'Asunto sugerido: ' };
  const su = X.su + dia, bo = 'Hola ' + nom + ',\n\n';
  const gm = 'https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(mail) + '&su=' + encodeURIComponent(su) + '&body=' + encodeURIComponent(bo);
  const mt = 'mailto:' + mail + '?subject=' + encodeURIComponent(su) + '&body=' + encodeURIComponent(bo);
  const com = comidaTxt_(r), cat = String(r[C_.cat - 1] || '');
  const fila = (a, b) => b ? '<tr><td style="padding:7px 12px;color:#6d6a64;border-bottom:1px solid #eee9df;width:38%">' + a + '</td><td style="padding:7px 12px;border-bottom:1px solid #eee9df"><b>' + esc_(b) + '</b></td></tr>' : '';
  const html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#3d3a35;max-width:560px;font-size:15px;line-height:1.55">' +
    '<div style="border-bottom:4px solid #ff6b35;padding-bottom:10px;margin-bottom:16px;font-weight:bold;color:#2d4a1e;font-size:13px;letter-spacing:.5px">ADVENTURE CO. · ' + X.av + '</div>' +
    '<p style="font-size:22px;font-weight:bold;color:#2d4a1e;margin:0 0 6px">📩 ' + X.toca + esc_(emp) + '</p>' +
    '<p style="margin:0 0 16px">' + X.es + '<b>' + X.cu + '</b>' + X.env + '</p>' +
    '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:2px solid #2d4a1e;font-size:14px">' +
    fila(X.jo, dia + (hora ? ' · ' + hora : '')) + fila(X.co, cont + (r[C_.tel - 1] ? ' · ' + r[C_.tel - 1] : '')) +
    fila(X.gr, r[C_.personas - 1] + X.pe + r[C_.actividad - 1]) + fila(X.di, ca ? comidaCa_(com) : com) + fila(X.ca, cat) +
    fila(X.cm, String(r[C_.coment - 1] || '')) + '</table>' +
    '<p style="margin:20px 0 6px;font-weight:bold">' + X.pa + '</p>' +
    '<p style="margin:0 0 4px">' + X.p1 + esc_(nom) + X.p1b + '</p>' +
    '<p style="margin:0 0 4px">' + X.p2 + '</p>' +
    '<table cellpadding="0" cellspacing="0" style="margin:18px 0 8px"><tr><td style="background:#2d4a1e;border-radius:30px"><a href="' + gm + '" style="display:inline-block;padding:16px 34px;color:#ffffff;text-decoration:none;font-weight:bold;font-size:17px">✉️ ' + X.bt + esc_(nom) + ' (' + esc_(emp) + ')</a></td></tr></table>' +
    '<p style="font-size:13px;color:#6d6a64;margin:0 0 4px">' + X.no.replace('{mt}', mt) + '<b>' + esc_(mail) + '</b>.</p>' +
    '<p style="font-size:13px;color:#6d6a64;margin:0 0 16px">' + X.nr + '</p>' +
    '<p style="background:#fff8e1;border-left:4px solid #e8704e;padding:10px 12px;font-size:14px;margin:0">' + X.id + '</p>' +
    '<p style="font-size:12px;color:#9a968e;border-top:1px solid #e3ded4;padding-top:10px;margin-top:20px">Ref. ' + r[C_.ref - 1] + X.pie + '</p></div>';
  const txt = X.toca + emp + ' (' + X.cu + ': ' + dia + (hora ? ' ' + hora : '') + ').\n' + X.cc + mail + '\n' + X.ass + su;
  envio_('info@empresa-ejemplo.com', '📩 ' + X.as + emp + X.ja + X.cu + ' (' + dia + ')', txt, null, html);
}
const DIAS_CADUCA = 15, DIAS_RECORDATORIO = 12, DIAS_BIENVENIDA = 4;
/** Cada día a las 9: recordatorios de caducidad, pre-reservas caducadas y avisos de bienvenida */
function tareasDiarias() {
  try { vincularTelegram(); } catch (err) { console.error(err); }
  const s = main_().getSheetByName(SH), n = s.getLastRow() - 1;
  if (n < 1) return;
  const V = s.getRange(2, 1, n, C_.pagado).getValues(), D = s.getRange(2, 1, n, C_.pagado).getDisplayValues();
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const DIA = 86400000;
  V.forEach((r, i) => {
    if (!r[C_.ref - 1]) return;
    const F = { s: s, row: i + 2, v: r, d: D[i] }, est = String(r[C_.estado - 1]), hist = String(r[C_.hist - 1] || '');
    try {
      if (est === 'Pendiente' && r[C_.email - 1] && r[C_.token - 1]) {
        let base = r[0] instanceof Date ? new Date(r[0]) : null;
        const mods = hist.match(/(\d{2})\/(\d{2})\/(\d{4})[^\n]*Modificada/g);
        if (mods) { const m = mods[mods.length - 1].match(/(\d{2})\/(\d{2})\/(\d{4})/); base = new Date(+m[3], +m[2] - 1, +m[1]); }
        if (!base) return;
        base.setHours(0, 0, 0, 0);
        const edad = Math.round((hoy - base) / DIA), limite = new Date(base.getTime() + DIAS_CADUCA * DIA);
        const yaAvisado = hist.lastIndexOf('Recordatorio de caducidad') > hist.lastIndexOf('Modificada');
        if (edad >= DIAS_CADUCA) {
          s.getRange(F.row, C_.estado).setValue('Cancelado');
          hist_(F, 'Pre-reserva caducada (15 días sin confirmar)');
          try { sincronizarFila_(s, F.row); } catch (err) { console.error(err); }
          const caC = String(r[C_.lang - 1]) === 'ca';
          envio_('info@empresa-ejemplo.com', (caC ? '⌛ Pre-reserva CADUCADA – ' : '⌛ Pre-reserva CADUCADA – ') + r[C_.ref - 1] + ' – ' + r[C_.empresa - 1],
            (caC ? 'La pre-reserva ha caducat (15 dies sense confirmar) i s’ha anul·lat. L’horari queda lliure.\n\nReferència: ' : 'La pre-reserva ha caducado (15 días sin confirmar) y se ha anulado. El horario queda libre.\n\nReferencia: ') + r[C_.ref - 1] + '\nEmpresa: ' + r[C_.empresa - 1] + '\nJornada: ' + cuando_(F), null);
        } else if (edad >= DIAS_RECORDATORIO && !yaAvisado) {
          if (correoCaduca_(F, limite)) hist_(F, 'Recordatorio de caducidad enviado al cliente');
        }
      } else if (est === 'Aceptado') {
        const d = fechaDe_(r[C_.fecha - 1]);
        if (!d) return;
        const faltan = Math.round((d - hoy) / DIA);
        if (faltan >= 0 && faltan <= DIAS_BIENVENIDA && hist.indexOf('Aviso de bienvenida') < 0) {
          avisoBienvenida_(F, faltan);
          hist_(F, 'Aviso de bienvenida enviado a Adventure Co.');
        }
      }
    } catch (err) { console.error(err); }
  });
  syncMA_();
}
function instalarTareasDiarias() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tareasDiarias').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tareasDiarias').timeBased().everyDays(1).atHour(9).inTimezone('Europe/Madrid').create();
}

/* ================= Avisos a info@ en catalán (cliente catalán) ================= */
const CA_MA_ = [
  ['❌ Reserva CANCELADA por el cliente', '❌ Reserva CANCEL·LADA pel client'],
  ['El cliente ha cancelado la reserva desde el correo. El horario vuelve a estar libre y se ha borrado del calendario.', 'El client ha cancel·lat la reserva des del correu. L’horari torna a estar lliure i s’ha esborrat del calendari.'],
  ['⚠️ Horario completo', '⚠️ Horari complet'],
  ['El cliente ha intentado confirmar, pero el horario ya está completo. Escríbele para proponer otra hora o fecha.', 'El client ha intentat confirmar, però l’horari ja està complet. Escriu-li per proposar-li una altra hora o data.'],
  ['✅ Reserva CONFIRMADA por el cliente', '✅ Reserva CONFIRMADA pel client'],
  ['El cliente ha confirmado la reserva desde el correo. Ya está en verde en el calendario. ', 'El client ha confirmat la reserva des del correu. Ja surt en verd al calendari. '],
  [' días antes de la jornada os llegará un aviso para enviarle el correo de bienvenida.', ' dies abans de la jornada us arribarà un avís per enviar-li el correu de benvinguda.'],
  ['✏️ Reserva MODIFICADA por el cliente', '✏️ Reserva MODIFICADA pel client'],
  ['El cliente ha modificado la reserva desde la web. Queda como Pendiente hasta que la vuelva a confirmar desde el nuevo correo.\nAntes: ', 'El client ha modificat la reserva des del web. Queda com a Pendent fins que la torni a confirmar des del nou correu.\nAbans: '],
  ['✅ Reserva CONFIRMADA (desde la hoja)', '✅ Reserva CONFIRMADA (des del full)'],
  ['La reserva se ha marcado como Aceptado en la hoja. Ya está en verde en el calendario. ', 'La reserva s’ha marcat com a Acceptada al full. Ja surt en verd al calendari. '],
  ['❌ Reserva CANCELADA (desde la hoja)', '❌ Reserva CANCEL·LADA (des del full)'],
  ['La reserva se ha marcado como Cancelado en la hoja. Se ha borrado del calendario y el horario queda libre.', 'La reserva s’ha marcat com a Cancel·lada al full. S’ha esborrat del calendari i l’horari queda lliure.'],
  ['\nReferencia: ', '\nReferència: '], ['\nContacto: ', '\nContacte: '], ['\nComida: ', '\nDinar: '], ['\nPersonas: ', '\nPersones: '],
  ['Actividades (Adventure Co.): ', 'Activitats (Adventure Co.): '], ['\nCatering (Restaurante Partner): ', '\nCàtering (Restaurante Partner): '],
  ['\n\nComentarios del cliente: ', '\n\nComentaris del client: '], ['\n\nAdjunto: ficha de la reserva en PDF.', '\n\nAdjunt: fitxa de la reserva en PDF.'],
  // ficha PDF
  ['>Ficha de reserva<', '>Fitxa de reserva<'],
  ['✕  RESERVA CANCELADA', '✕  RESERVA CANCEL·LADA'], ['El grupo ha cancelado la reserva.', 'El grup ha cancel·lat la reserva.'],
  ['Pendiente de que el cliente la vuelva a confirmar.', 'Pendent que el client la torni a confirmar.'],
  ['>COMIDA: ', '>DINAR: '], ['DATOS DEL CLIENTE', 'DADES DEL CLIENT'], ['>ACTIVIDAD · HORARIO<', '>ACTIVITAT · HORARI<'], ['>PERSONAS · ACTIVIDAD<', '>PERSONES · ACTIVITAT<'],
  ['>COMIDA<', '>DINAR<'], ['>PERSONAS<', '>PERSONES<'], ['>ACTIVIDAD<', '>ACTIVITAT<'], ['>CONTACTO<', '>CONTACTE<'],
  ['>FACTURA A NOMBRE DE LA EMPRESA<', '>FACTURA A NOM DE L’EMPRESA<'], ['>REFERENCIA<', '>REFERÈNCIA<'],
  ['<b>Grupo de más de 50 personas:</b> la actividad se hace en dos grupos (el segundo empieza 30 minutos después) y lo más recomendable es comer en el porche de las instalaciones.',
    '<b>Grup de més de 50 persones:</b> l’activitat es fa en dos grups (el segon comença 30 minuts després) i el més recomanable és dinar al porxo de les instal·lacions.'],
  ['ACTIVIDADES · ADVENTURE CO.', 'ACTIVITATS · ADVENTURE CO.'], ['>CONCEPTO<', '>CONCEPTE<'], ['>CANTIDAD<', '>QUANTITAT<'], ['>PRECIO<', '>PREU<'], ['>IMPORTE<', '>IMPORT<'],
  ['Base imponible', 'Base imposable'], ['(base imponible ', '(base imposable '], ['Total actividades (IVA incluido)', 'Total activitats (IVA inclòs)'],
  [' · SERVIR EN LAS INSTALACIONES A LAS ', ' · SERVIR A LES INSTAL·LACIONS A LES '], [' · COMIDA A LAS ', ' · DINAR A LES '],
  ['Total catering (IVA incluido)', 'Total càtering (IVA inclòs)'], ['Total jornada (dos facturas)', 'Total jornada (dues factures)'],
  ['Calçotada (menú bufet): ensaladas, calçots a la llama con romesco (15 u.), parrillada de carne a la brasa con patatas al caliu, alcachofas y judías del ganxet, alioli, pan, postre y bebida (agua, cerveza, refresco o vino).',
    'Calçotada (menú bufet): amanides, calçots a la flama amb romesco (15 u.), graellada de carn a la brasa amb patates al caliu, carxofes i mongetes del ganxet, allioli, pa, postres i beguda (aigua, cervesa, refresc o vi).'],
  ['MARCAR COMO PAGADA', 'MARCAR COM A PAGADA'], ['Pulsa aquí cuando el cliente haya pagado las actividades. Se te pedirá confirmación.', 'Prem aquí quan el client hagi pagat les activitats. Se’t demanarà confirmació.'],
  ['COMENTARIOS DEL CLIENTE', 'COMENTARIS DEL CLIENT'], ['Reservas y dudas', 'Reserves i dubtes'], ['WhatsApp y teléfono: ', 'WhatsApp i telèfon: '], ['<br>Catering: Restaurant', '<br>Càtering: Restaurant']
];
function caMA_(t) {
  let x = comidaCa_(String(t));
  CA_MA_.forEach(p => { x = x.split(p[0]).join(p[1]); });
  return x;
}

/* ================= Hoja del restaurante (solo lectura + columna "Pago") (v25) ================= */
const R_PRE = 'Pre-reservas', R_CONF = 'Reservas confirmadas', C_RPAG = 33;
const R_HEAD = ['Referencia', 'Fecha', 'Hora comida', 'Servicio', 'Grupo (empresa)', 'Contacto', 'Teléfono', 'Email', 'Personas que comen', 'Menú / catering', 'Importe catering (€)', 'Alergias y comentarios'];
const R_HEAD_C = R_HEAD.concat(['Pago', 'Fecha de pago']), R_PAGO = R_HEAD_C.length - 1;
const R_MARRON = '#463428', R_FACT = 'Facturación mensual', R_RES = 'Resumen';
const R_FACT_HEAD = ['Mes', 'Referencia', 'Fecha', 'Grupo (empresa)', 'Personas que comen', 'Menú / catering', 'Base imponible (€)', 'IVA 10 % (€)', 'Total IVA incl. (€)', 'Pago'];
function iv10_(t) { t = Number(t) || 0; const b = Math.round(t / 1.1 * 100) / 100; return [b, Math.round((t - b) * 100) / 100]; }
function hojaRest_() {
  const id = PropertiesService.getScriptProperties().getProperty('REST_ID');
  return id ? SpreadsheetApp.openById(id) : null;
}
const R_ROW = 3; // fila 1: logos · fila 2: cabecera · datos desde la fila 3
const LOGO_MA_URL = 'https://www.empresa-ejemplo.com/wp-content/uploads/2026/10/logo-empresa.png';
const LOGO_VALL_URL = 'https://www.empresa-ejemplo.com/wp-content/uploads/2026/10/logo-restaurante-partner.png';
function cabRest_(h, head, nota) {
  h.setFrozenRows(0); h.setFrozenColumns(0);
  h.getRange(1, 1, h.getMaxRows(), h.getMaxColumns()).breakApart();
  h.clear(); h.clearConditionalFormatRules();
  if (h.getMaxColumns() > head.length) h.deleteColumns(head.length + 1, h.getMaxColumns() - head.length);
  const n = head.length;
  h.getRange(1, 1, 1, n).setBackground('#ffffff').setVerticalAlignment('middle');
  h.getRange(1, 1, 1, 2).merge().setFormula('=IMAGE("' + LOGO_MA_URL + '")');
  h.getRange(1, 3, 1, 2).merge().setFormula('=IMAGE("' + LOGO_VALL_URL + '")');
  h.getRange(1, 5, 1, n - 4).merge().setValue('Adventure Co.  ×  Restaurante Partner\nGrupos de team building que comen con vosotros')
    .setHorizontalAlignment('left').setWrap(true).setFontWeight('bold').setFontSize(13).setFontColor(VERDE);
  h.getRange(1, 1, 1, n).setBorder(null, null, true, null, null, null, '#ff6b35', SpreadsheetApp.BorderStyle.SOLID_THICK);
  h.setRowHeight(1, 80);
  h.getRange(2, 1, 1, n).setValues([head]).setFontWeight('bold').setFontColor('#ffffff').setBackground(R_MARRON)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  h.setRowHeight(2, 42); h.setFrozenRows(2);
  [120, 105, 90, 230, 200, 150, 110, 200, 95, 320, 115, 260, 110, 135].slice(0, head.length).forEach((w, i) => h.setColumnWidth(i + 1, w));
  h.getRange(2, 1).setNote(nota);
}
function crearHojaRestaurante() {
  const P = PropertiesService.getScriptProperties();
  let ss = hojaRest_();
  if (!ss) { ss = SpreadsheetApp.create('Restaurante Partner · Reservas de Adventure Co.'); P.setProperty('REST_ID', ss.getId()); }
  ss.setSpreadsheetTimeZone('Europe/Madrid');
  let a = ss.getSheetByName(R_PRE); if (!a) { a = ss.getSheets()[0]; a.setName(R_PRE); }
  let c = ss.getSheetByName(R_CONF); if (!c) c = ss.insertSheet(R_CONF);
  let fa = ss.getSheetByName(R_FACT); if (!fa) fa = ss.insertSheet(R_FACT);
  let rz = ss.getSheetByName(R_RES); if (!rz) rz = ss.insertSheet(R_RES, 1);
  cabRest_(fa, R_FACT_HEAD, 'Facturación de los grupos de Adventure Co. que comen con vosotros (solo reservas confirmadas), mes a mes. Se actualiza sola. Importes con IVA del 10 % desglosado.');
  [110, 120, 100, 200, 95, 300, 120, 105, 120, 100].forEach((w, i) => fa.setColumnWidth(i + 1, w));
  cabRest_(a, R_HEAD, 'Pre-reservas: el cliente todavía NO ha confirmado. Es solo para que lo tengáis previsto; no hace falta preparar nada hasta que pase a "Reservas confirmadas".');
  cabRest_(c, R_HEAD_C, 'CÓMO USAR ESTA HOJA\n\n• Aquí salen los grupos que ya han confirmado y comen con vosotros. Se actualiza sola cada pocos minutos: no hace falta apuntar nada.\n\n• Cuando el cliente os pague el catering: en la columna «Pago», pulsad la flecha ▼ de su fila y elegid «Pagada». La celda se pondrá en verde y se apuntará la fecha de pago.\n\n• Si os equivocáis, volved a elegir «Pendiente».\n\n• El resto de columnas está bloqueado para que no se borre nada sin querer.\n\n• Dudas: 600 000 000 · info@empresa-ejemplo.com');
  c.getRange(2, R_PAGO).setBackground(NARANJA).setNote('Cambiad a "Pagada" cuando el cliente os haya pagado el catering.');
  ss.setActiveSheet(c); ss.moveActiveSheet(1);
  // Protección: todo bloqueado salvo la columna "Pago" de confirmadas
  const own = ss.getOwner() ? ss.getOwner().getEmail() : '';
  [a, c, fa, rz].forEach(h => {
    h.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const p = h.protect().setDescription('Solo lectura (se actualiza sola desde la web)');
    if (h === c) p.setUnprotectedRanges([h.getRange(R_ROW, R_PAGO, h.getMaxRows() - R_ROW + 1, 1)]);
    p.removeEditors(p.getEditors().filter(u => u.getEmail() !== own));
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });
  const s = main_().getSheetByName(SH);
  if (s.getMaxColumns() < C_RPAG) s.insertColumnsAfter(s.getMaxColumns(), C_RPAG - s.getMaxColumns());
  s.getRange(1, C_RPAG).setValue('Pago catering (restaurante)').setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE).setHorizontalAlignment('center').setWrap(true);
  s.setColumnWidth(C_RPAG, 150);
  const aj = main_().getSheetByName(AJ);
  if (aj) { aj.getRange('B8').setValue('Hoja del restaurante (pre-reservas y confirmadas)').setFontWeight('bold').setFontColor(VERDE); aj.getRange('C8').setValue(ss.getUrl()); }
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alEditarRest').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alEditarRest').forSpreadsheet(ss).onEdit().create();
  sincronizarRest_();
  if (!P.getProperty('REST_MAIL')) { P.setProperty('REST_MAIL', '1'); envio_('colaborador@example.com', '🍽️ Hoja del restaurante creada', 'Hoja del restaurante (pre-reservas y confirmadas): ' + ss.getUrl() + '\n\nTodavía NO está compartida con el restaurante.', null); }
}
function sincronizarRest_() {
  const ss = hojaRest_(); if (!ss) return;
  const s = main_().getSheetByName(SH), n = s.getLastRow() - 1;
  const V = n > 0 ? s.getRange(2, 1, n, C_RPAG).getValues() : [], D = n > 0 ? s.getRange(2, 1, n, C_RPAG).getDisplayValues() : [];
  const pre = [], conf = [];
  V.forEach((r, i) => {
    const est = String(r[C_.estado - 1]);
    if (!r[C_.ref - 1] || !r[C_.cat - 1] || (est !== 'Pendiente' && est !== 'Aceptado')) return;
    const d = fechaDe_(r[C_.fecha - 1]), hc = String(D[i][C_.comida - 1] || '');
    const serv = !hc ? '' : enInst_(r) ? 'Llevar a las instalaciones de Adventure Co.' : 'Comen en el restaurante';
    const fila = [String(r[C_.ref - 1]), d || '', hc, serv, r[C_.empresa - 1], r[C_.contacto - 1], D[i][C_.tel - 1], r[C_.email - 1], paxDe_(r) || r[C_.personas - 1],
      String(r[C_.cat - 1]).split(' · ').join('\n'), Number(r[C_.rest - 1]) || 0, String(r[C_.coment - 1] || '')];
    if (est === 'Pendiente') pre.push(fila);
    else { const pg = String(r[C_RPAG - 1] || ''); conf.push(fila.concat([pg ? 'Pagada' : 'Pendiente', pg ? pg.replace(/^Pagado · /, '') : ''])); }
  });
  const ord = (x, y) => (x[1] ? x[1].getTime() : 0) - (y[1] ? y[1].getTime() : 0);
  pre.sort(ord); conf.sort(ord);
  try { factRest_(ss, conf); } catch (err) { console.error(err); }
  try { resumenRest_(ss, conf, pre); } catch (err) { console.error(err); }
  [[ss.getSheetByName(R_PRE), pre, R_HEAD.length], [ss.getSheetByName(R_CONF), conf, R_HEAD_C.length]].forEach(([h, rows, nc]) => {
    if (!h) return;
    const last = h.getLastRow();
    if (last >= R_ROW) h.getRange(R_ROW, 1, last - R_ROW + 1, nc).clearContent().clearDataValidations();
    if (rows.length) {
      h.getRange(R_ROW, 1, rows.length, nc).setValues(rows).setVerticalAlignment('top').setWrap(true);
      h.getRange(R_ROW, 2, rows.length, 1).setNumberFormat('dd/MM/yyyy');
      h.getRange(R_ROW, 3, rows.length, 1).setNumberFormat('@'); h.getRange(R_ROW, 7, rows.length, 1).setNumberFormat('@');
      h.getRange(R_ROW, 11, rows.length, 1).setNumberFormat('#,##0.00 €');
      if (nc === R_HEAD_C.length) h.getRange(R_ROW, R_PAGO, rows.length, 1).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(['Pendiente', 'Pagada'], true).setAllowInvalid(false).build()).setHorizontalAlignment('center').setFontWeight('bold');
    }
    const rg = h.getRange(R_ROW, 1, Math.max(rows.length, 1), nc), reglas = [SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=ISEVEN(ROW())').setBackground(CREMA).setRanges([rg]).build()];
    if (nc === R_HEAD_C.length) {
      const pg = h.getRange(R_ROW, R_PAGO, Math.max(rows.length, 1), 1);
      reglas.unshift(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pagada').setBackground('#d9ead3').setFontColor('#1e5631').setRanges([pg]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pendiente').setBackground('#fff4e5').setFontColor('#b7791f').setRanges([pg]).build());
    }
    h.setConditionalFormatRules(reglas);
  });
  SpreadsheetApp.flush();
}
/** El restaurante cambia la columna "Pago" */
function alEditarRest(e) {
  const r = e.range, h = r.getSheet();
  if (h.getName() !== R_CONF || r.getRow() < R_ROW) return;
  if (r.getNumRows() !== 1 || r.getNumColumns() !== 1 || r.getColumn() !== R_PAGO) { sincronizarRest_(); return; }
  const ref = String(h.getRange(r.getRow(), 1).getValue()), nuevo = String(e.value || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = main_().getSheetByName(SH), n = s.getLastRow() - 1;
    const V = s.getRange(2, 1, n, C_RPAG).getValues(), D = s.getRange(2, 1, n, C_RPAG).getDisplayValues();
    const i = V.findIndex(x => String(x[C_.ref - 1]) === ref);
    if (i < 0) return;
    const F = { s: s, row: i + 2, v: V[i], d: D[i] }, pag = String(V[i][C_RPAG - 1] || ''), x = V[i];
    if (nuevo === 'Pagada' && !pag) {
      s.getRange(F.row, C_RPAG).setValue('Pagado · ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm'));
      hist_(F, 'Catering marcado como PAGADO por el restaurante');
      envio_('colaborador@example.com', '🍽️💶 Catering PAGADO – ' + x[C_.ref - 1] + ' – ' + x[C_.empresa - 1],
        'El restaurante ha marcado el catering como pagado.\n\nReferencia: ' + x[C_.ref - 1] + '\nEmpresa: ' + x[C_.empresa - 1] + '\nJornada: ' + cuando_(F) + '\nCatering: ' + eur_(Number(x[C_.rest - 1]) || 0), null);
    } else if (nuevo === 'Pendiente' && pag) {
      s.getRange(F.row, C_RPAG).setValue('');
      hist_(F, 'Pago del catering DESMARCADO por el restaurante');
      envio_('colaborador@example.com', '↩️ Catering desmarcado – ' + x[C_.ref - 1] + ' – ' + x[C_.empresa - 1],
        'El restaurante ha vuelto a poner el catering como PENDIENTE de pago.\n\nReferencia: ' + x[C_.ref - 1] + '\nEmpresa: ' + x[C_.empresa - 1] + '\nJornada: ' + cuando_(F), null);
    }
  } finally { lock.releaseLock(); sincronizarRest_(); }
}

/** Pestaña "Facturación mensual" del restaurante: detalle por mes con IVA 10 % desglosado */
function factRest_(ss, conf) {
  const h = ss.getSheetByName(R_FACT); if (!h) return;
  const MES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const out = [], subs = [], r2 = x => Math.round(x * 100) / 100;
  let cur = null, sb = 0, si = 0, st = 0, gb = 0, gi = 0, gt = 0;
  const cierra = () => { if (cur !== null) { out.push(['Total ' + cur, '', '', '', '', '', r2(sb), r2(si), r2(st), '']); subs.push(out.length - 1); } };
  conf.forEach(f => {
    const d = f[1], k = d ? MES[d.getMonth()] + ' ' + d.getFullYear() : 'Sin fecha';
    if (k !== cur) { cierra(); cur = k; sb = si = st = 0; }
    const t = Number(f[10]) || 0, v = iv10_(t);
    sb += v[0]; si += v[1]; st += t; gb += v[0]; gi += v[1]; gt += t;
    out.push([k, f[0], d || '', f[4], f[8], f[9], v[0], v[1], t, f[12]]);
  });
  cierra();
  if (out.length) { out.push(['TOTAL GENERAL', '', '', '', '', '', r2(gb), r2(gi), r2(gt), '']); subs.push(out.length - 1); }
  const nc = R_FACT_HEAD.length, last = h.getLastRow();
  if (last >= R_ROW) h.getRange(R_ROW, 1, last - R_ROW + 1, nc).clearContent().setBackground(null).setFontWeight('normal').setFontColor(null);
  h.setConditionalFormatRules([]);
  if (!out.length) return;
  h.getRange(R_ROW, 1, out.length, nc).setValues(out).setVerticalAlignment('top').setWrap(true);
  h.getRange(R_ROW, 3, out.length, 1).setNumberFormat('dd/MM/yyyy');
  h.getRange(R_ROW, 7, out.length, 3).setNumberFormat('#,##0.00 €');
  subs.forEach((i, j) => h.getRange(R_ROW + i, 1, 1, nc).setFontWeight('bold').setBackground(j === subs.length - 1 ? '#ffe0cc' : CREMA).setFontColor(j === subs.length - 1 ? '#b34a1a' : VERDE));
  out.forEach((x, i) => { if (x[9] === 'Pagada') h.getRange(R_ROW + i, 10).setFontColor('#1e5631').setFontWeight('bold'); else if (x[9] === 'Pendiente') h.getRange(R_ROW + i, 10).setFontColor('#b7791f').setFontWeight('bold'); });
}

/* ================= Botón de pago del restaurante (Pagada / Cancelar pago) (v31) ================= */
function restKey_(ref) { return pagoKey_(String(ref) + '|rest'); }
function restPagoUrl_(ref) { return EXEC_URL + '?a=restpago&ref=' + encodeURIComponent(ref) + '&p=' + restKey_(ref); }
function filaRest_(ref) {
  const s = main_().getSheetByName(SH), n = s.getLastRow() - 1;
  if (n < 1 || !ref) return null;
  const V = s.getRange(2, 1, n, C_RPAG).getValues(), D = s.getRange(2, 1, n, C_RPAG).getDisplayValues();
  const i = V.findIndex(r => String(r[C_.ref - 1]) === String(ref));
  return i < 0 ? null : { s: s, row: i + 2, v: V[i], d: D[i] };
}
function restPagoCard_(F) {
  const r = F.v, est = String(r[C_.estado - 1]), pg = String(r[C_RPAG - 1] || ''), t = Number(r[C_.rest - 1]) || 0, v = iv10_(t);
  const info = 'Grupo: <b>' + esc_(r[C_.empresa - 1]) + '</b><br>Referencia: <b>' + esc_(r[C_.ref - 1]) + '</b><br>Jornada: <b>' + cuando_(F) + '</b>' +
    (comidaTxt_(r) ? '<br>Comida: <b>' + esc_(comidaTxt_(r)) + '</b>' : '') + '<br>Catering: <b>' + eur_(t) + '</b> <span style="font-size:13px">(base ' + eur_(v[0]) + ' · IVA 10 % ' + eur_(v[1]) + ')</span>';
  if (est === 'Cancelado' || est === 'Rechazado') return '<div class="ico" style="background:#9b2c1f">✕</div><h1 style="color:#9b2c1f">Reserva cancelada</h1><div class="i">' + info + '</div>';
  if (est !== 'Aceptado') return '<div class="ico" style="background:#b7791f">⏳</div><h1 style="color:#b7791f">Pre-reserva</h1><p>El grupo todavía no ha confirmado. Cuando confirme, desde este mismo enlace podréis marcar el pago.</p><div class="i">' + info + '</div>';
  if (pg) return '<div class="ico" style="background:#2e7d32">✓</div><h1 style="color:#2e7d32">PAGADA</h1><p>' + esc_(pg) + '</p><div class="i">' + info + '</div>' +
    '<button id="go" class="b r" data-a="cancelar">Cancelar pago</button><p class="er" id="er"></p>';
  return '<div class="ico" style="background:#b7791f">€</div><h1 style="color:#b7791f">Pendiente de pago</h1><p>Pulsad el botón cuando el grupo os haya pagado el catering.</p><div class="i">' + info + '</div>' +
    '<button id="go" class="b" style="background:#2e7d32" data-a="pagar">💶 Marcar como PAGADA</button><p class="er" id="er"></p>';
}
function paginaRestPago_(p) {
  const F = (p.p && p.p === restKey_(p.ref)) ? filaRest_(p.ref) : null;
  if (!F) return pagina_(card_('Enlace no válido', 'No hemos encontrado la reserva.', '', 'err', false), 'Pago catering', false);
  const js = '<script>function bindPago(){var g=document.getElementById("go");if(!g)return;g.onclick=function(){var a=g.getAttribute("data-a");g.disabled=true;g.textContent="Un momento…";' +
    'google.script.run.withSuccessHandler(function(h){document.getElementById("c").innerHTML=h;bindPago();}).withFailureHandler(function(){g.disabled=false;g.textContent=(a==="pagar"?"💶 Marcar como PAGADA":"Cancelar pago");document.getElementById("er").textContent="No se ha podido registrar. Vuelve a intentarlo.";})' +
    '.ejecutarRestPago(' + JSON.stringify(String(p.ref)) + ',' + JSON.stringify(String(p.p)) + ',a);};}bindPago();<\/script>';
  return pagina_(restPagoCard_(F) + js, 'Pago catering', false);
}
/** Llamado desde la página de pago del restaurante (google.script.run). a = 'pagar' | 'cancelar' */
function ejecutarRestPago(ref, key, a) {
  if (key !== restKey_(ref)) throw new Error('Enlace no válido');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let F;
  try {
    F = filaRest_(ref);
    if (!F) throw new Error('No encontrada');
    const x = F.v, pg = String(x[C_RPAG - 1] || ''), est = String(x[C_.estado - 1]);
    if (a === 'pagar' && !pg && est === 'Aceptado') {
      F.s.getRange(F.row, C_RPAG).setValue('Pagado · ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm'));
      hist_(F, 'Catering marcado como PAGADO por el restaurante (botón)');
      envio_('colaborador@example.com', '🍽️💶 Catering PAGADO – ' + x[C_.ref - 1] + ' – ' + x[C_.empresa - 1],
        'El restaurante ha marcado el catering como pagado (botón).\n\nReferencia: ' + x[C_.ref - 1] + '\nEmpresa: ' + x[C_.empresa - 1] + '\nJornada: ' + cuando_(F) + '\nCatering: ' + eur_(Number(x[C_.rest - 1]) || 0), null);
    } else if (a === 'cancelar' && pg) {
      F.s.getRange(F.row, C_RPAG).setValue('');
      hist_(F, 'Pago del catering CANCELADO por el restaurante (botón)');
      envio_('colaborador@example.com', '↩️ Catering desmarcado – ' + x[C_.ref - 1] + ' – ' + x[C_.empresa - 1],
        'El restaurante ha cancelado el pago del catering (vuelve a pendiente).\n\nReferencia: ' + x[C_.ref - 1] + '\nEmpresa: ' + x[C_.empresa - 1] + '\nJornada: ' + cuando_(F), null);
    }
    SpreadsheetApp.flush();
    F = filaRest_(ref);
  } finally { lock.releaseLock(); }
  try { sincronizarRest_(); } catch (err) { console.error(err); }
  return restPagoCard_(F);
}

/** Pestaña "Resumen" del restaurante: panel de lo facturado con Adventure Co. (mismo estilo que el Resumen de Empresa) */
function resumenRest_(ss, conf, pre) {
  const h = ss.getSheetByName(R_RES); if (!h) return;
  const MES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const r2 = x => Math.round(x * 100) / 100, LIMA = '#8dc63f';
  let tot = 0, pag = 0, pax = 0, prev = 0;
  const meses = {}, orden = [];
  conf.forEach(f => {
    const t = Number(f[10]) || 0, p = Number(f[8]) || 0, ok = f[12] === 'Pagada', d = f[1];
    const k = d ? d.getFullYear() * 100 + d.getMonth() : 0;
    if (!meses[k]) { meses[k] = { l: d ? MES[d.getMonth()] + ' ' + d.getFullYear() : 'Sin fecha', g: 0, p: 0, t: 0, c: 0 }; orden.push(k); }
    const m = meses[k]; m.g++; m.p += p; m.t += t; if (ok) m.c += t;
    tot += t; pax += p; if (ok) pag += t;
  });
  pre.forEach(f => { prev += Number(f[10]) || 0; });
  const iv = iv10_(tot);
  h.getRange(1, 1, h.getMaxRows(), h.getMaxColumns()).breakApart();
  h.clear(); h.clearConditionalFormatRules(); h.setHiddenGridlines(true);
  if (h.getMaxColumns() < 10) h.insertColumnsAfter(h.getMaxColumns(), 10 - h.getMaxColumns());
  [20, 330, 140, 30, 170, 80, 90, 150, 130, 130].forEach((w, i) => h.setColumnWidth(i + 1, w));
  h.setRowHeight(1, 70);
  h.getRange('B1').setFormula('=IMAGE("' + LOGO_MA_URL + '")');
  h.getRange('C1').setFormula('=IMAGE("' + LOGO_VALL_URL + '")');
  h.getRange('B2').setValue('Facturación · Restaurante Partner').setFontSize(20).setFontWeight('bold').setFontColor(VERDE);
  h.getRange('B3').setValue('Grupos de team building de Adventure Co. — se actualiza solo con cada reserva').setFontColor('#6d6a64');
  const K = [
    ['Grupos confirmados', conf.length, '0'],
    ['Personas que comen', pax, '0'],
    ['Base imponible', r2(iv[0]), '#,##0.00 €'],
    ['IVA (10 %)', r2(iv[1]), '#,##0.00 €'],
    ['Total facturado (IVA incluido)', r2(tot), '#,##0.00 €'],
    ['Cobrado (marcado como pagado)', r2(pag), '#,##0.00 €'],
    ['Pendiente de cobro', r2(tot - pag), '#,##0.00 €'],
    ['Pre-reservas en curso (aún sin confirmar)', pre.length, '0'],
    ['Importe previsto de las pre-reservas', r2(prev), '#,##0.00 €']];
  const k0 = 5;
  K.forEach((x, i) => {
    const rg = h.getRange(k0 + i, 2, 1, 2);
    rg.setValues([[x[0], x[1]]]).setBackground(CREMA).setFontColor('#3d3a35').setFontSize(11)
      .setBorder(true, true, true, true, false, false, '#e3ded4', SpreadsheetApp.BorderStyle.SOLID);
    h.getRange(k0 + i, 3).setNumberFormat(x[2]).setHorizontalAlignment('right').setFontWeight('bold');
    h.setRowHeight(k0 + i, 30);
  });
  [4, 5].forEach(i => h.getRange(k0 + i, 2, 1, 2).setBackground(VERDE).setFontColor('#ffffff').setFontWeight('bold'));
  [4, 5].forEach(i => h.getRange(k0 + i, 3).setFontColor(LIMA).setFontSize(13));
  h.getRange(k0 + 6, 3).setFontColor('#b7791f');
  h.getRange(k0 + 8, 3).setFontColor('#6d6a64'); h.getRange(k0 + 7, 3).setFontColor('#6d6a64');
  // tabla por mes
  const head = ['Mes', 'Grupos', 'Personas', 'Facturado (IVA incl.)', 'Cobrado', 'Pendiente'];
  h.getRange(4, 5, 1, 6).setValues([head]).setBackground(NARANJA).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  h.setRowHeight(4, 30);
  const rows = orden.map(k => { const m = meses[k]; return [m.l, m.g, m.p, r2(m.t), r2(m.c), r2(m.t - m.c)]; });
  if (!rows.length) rows.push(['Todavía sin reservas confirmadas', '', '', '', '', '']);
  h.getRange(5, 5, rows.length, 6).setValues(rows).setHorizontalAlignment('center').setVerticalAlignment('middle');
  h.getRange(5, 8, rows.length, 3).setNumberFormat('#,##0.00 €').setHorizontalAlignment('right');
  h.getRange(4, 5, rows.length + 1, 6).setBorder(true, true, true, true, null, null, '#e3ded4', SpreadsheetApp.BorderStyle.SOLID);
  h.getRange(5, 10, rows.length, 1).setFontColor('#b7791f');
  h.getRange(5, 9, rows.length, 1).setFontColor('#1e5631');
  h.setFrozenRows(0);
}


/* ================= Avisos por Telegram (bot @YourBookingBot) ================= */
// 1) Propiedad del proyecto TELEGRAM_TOKEN = token del bot.
// 2) Cada persona abre su enlace UNA vez y pulsa Iniciar:
//    colaborador:          https://t.me/YourBookingBot?start=colab
//    Empresa:     https://t.me/YourBookingBot?start=empresa
//    Restaurante:  https://t.me/YourBookingBot?start=restaurante
// 3) Ejecutar probarTelegramReservas().
// Quién recibe qué: colaborador y Empresa → todo (pre-reservas, confirmadas, cambios, cancelaciones y pagos).
// Restaurante → solo la parte de la comida (pre-reservas, confirmadas, cambios, cancelaciones y catering pagado).
const TG_ROLES_ = { colab: 'TG_COLAB', empresa: 'TG_EMPRESA', restaurante: 'TG_REST' };

function tgToken_() { return String(PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN') || '').replace(/\s+/g, ''); }

function tgEnviar_(rol, texto) {
  const P = PropertiesService.getScriptProperties(), token = tgToken_();
  if (!token || !texto || !TG_ROLES_[rol]) return;
  let chat = P.getProperty(TG_ROLES_[rol]);
  if (!chat) { try { vincularTelegram(); } catch (err) { console.error(err); } chat = P.getProperty(TG_ROLES_[rol]); }
  if (!chat) return;
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chat, text: String(texto).slice(0, 4000), disable_web_page_preview: true })
  });
}

/** Lee los "/start colab|empresa|restaurante" recibidos por el bot y guarda el chat de cada uno. */
function vincularTelegram() {
  const P = PropertiesService.getScriptProperties(), token = tgToken_();
  if (!token) throw new Error('Falta TELEGRAM_TOKEN en las propiedades del proyecto.');
  const r = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates', { muteHttpExceptions: true }).getContentText());
  if (!r.ok) throw new Error('Telegram: ' + (r.description || 'error'));
  const hechos = [];
  (r.result || []).forEach(u => {
    const msg = u.message; if (!msg || !msg.text) return;
    const k = (String(msg.text).match(/^\/start\s+(\w+)/) || [])[1];
    if (k && TG_ROLES_[k]) { P.setProperty(TG_ROLES_[k], String(msg.chat.id)); hechos.push(k); }
  });
  console.log('Vinculados ahora: ' + (hechos.join(', ') || 'ninguno') + ' · colaborador: ' + (P.getProperty('TG_COLAB') ? 'sí' : 'no') +
    ' · Empresa: ' + (P.getProperty('TG_EMPRESA') ? 'sí' : 'no') + ' · Restaurante: ' + (P.getProperty('TG_REST') ? 'sí' : 'no'));
  return hechos;
}

function probarTelegramReservas() {
  vincularTelegram();
  tgEnviar_('colab', '✅ Avisos de reservas de Adventure Co. conectados. Te llegará todo: pre-reservas, confirmadas, cambios y pagos.');
  tgEnviar_('empresa', '✅ Avisos de reservas conectados. Aquí os llegarán las pre-reservas, las reservas confirmadas, los cambios y los pagos.');
  tgEnviar_('restaurante', '✅ Avisos de Adventure Co. conectados. Aquí os llegarán las pre-reservas y reservas confirmadas con comida, y los pagos del catering.');
}

/** Copia por Telegram de cada correo que envía el sistema, según a quién va. */
function tgEspejo_(to, asunto, cuerpo) {
  asunto = String(asunto || ''); const t = asunto + '\n\n' + String(cuerpo || '');
  const para = String(to || '').toLowerCase().trim();
  const destRest = String(ajuste_('C3') || '').toLowerCase().trim();
  const sinComision = x => x.split('\n').filter(l => !/Comisi[oó]n/i.test(l)).join('\n');
  if (para === 'colaborador@example.com') {
    tgEnviar_('colab', t);
    if (/^(💶|🍽️💶|↩️)/.test(asunto)) tgEnviar_('empresa', sinComision(t));        // pagos y pagos desmarcados
    if (/^(🍽️💶|↩️ Catering)/.test(asunto)) tgEnviar_('restaurante', sinComision(t)); // catering pagado / desmarcado
  } else if (/^🍽️/.test(asunto) || (destRest && para === destRest)) {
    tgEnviar_('restaurante', t);                                                      // avisos de comida al restaurante
  } else if (para === 'info@empresa-ejemplo.com') {
    tgEnviar_('empresa', t);
  }
}

/** Aviso de pre-reserva nueva para colaborador y Empresa (el restaurante ya recibe el suyo si hay comida). */
function tgPreReserva_(s, row) {
  const v = s.getRange(row, 1, 1, C_.pagado).getValues()[0], dv = s.getRange(row, 1, 1, C_.pagado).getDisplayValues()[0];
  const cuando = cuando_({ s: s, row: row, v: v, d: dv });
  const txt = '🆕 Nueva PRE-RESERVA (pendiente de confirmar) – ' + v[C_.ref - 1] + ' – ' + v[C_.empresa - 1] + '\n\n' +
    'Empresa: ' + v[C_.empresa - 1] + '\nContacto: ' + v[C_.contacto - 1] + ' · ' + v[C_.tel - 1] + ' · ' + v[C_.email - 1] +
    '\nJornada: ' + cuando + (comidaTxt_(v) ? '\nComida: ' + comidaTxt_(v) : '') +
    '\nPersonas: ' + v[C_.personas - 1] + ' · ' + v[C_.actividad - 1] +
    '\n\nActividades: ' + eur_(v[C_.total - 1]) + '\nCatering: ' + (Number(v[C_.rest - 1]) ? eur_(v[C_.rest - 1]) : '–') + '\nTotal jornada: ' + eur_(v[C_.jornada - 1]) +
    (v[C_.coment - 1] ? '\n\nComentarios del cliente: ' + v[C_.coment - 1] : '');
  tgEnviar_('colab', txt);
  tgEnviar_('empresa', String(v[C_.lang - 1]) === 'ca' ? caMA_(txt) : txt);
}


/* ================= OTRAS RESERVAS (formulario del equipo) =================
 * Página: EXEC_URL + '?a=eva&k=' + clave (ver evaEnlace()).
 * El equipo apunta reservas de otros servicios (fecha + hora o día entero). Se guardan en la pestaña
 * "Otras reservas" de su hoja, se crean en el calendario Team Building y cuentan para la
 * disponibilidad de la web (team building y escuelas), así no se solapan. */
const OTR_TAB = 'Otras reservas';
const OTR_HEAD = ['Apuntada el', 'Fecha', 'Hora inicio', 'Duración (h)', 'Día entero', 'Servicio', 'Cliente / notas', 'Personas', 'Estado', 'Evento'];

function otrosKey_() {
  const P = PropertiesService.getScriptProperties();
  let k = P.getProperty('EVA_KEY');
  // Si no existe todavia, se genera una clave aleatoria una sola vez y se guarda.
  if (!k) { k = Utilities.getUuid().replace(/-/g, '').slice(0, 16); P.setProperty('EVA_KEY', k); }
  return k;
}

/** Ejecutar para ver el enlace del formulario del equipo. */
function evaEnlace() { const u = EXEC_URL + '?a=eva&k=' + otrosKey_(); console.log(u); return u; }

function hojaOtr_() {
  const id = PropertiesService.getScriptProperties().getProperty('MA_ID');
  if (!id) return null;
  const ss = SpreadsheetApp.openById(id);
  let s = ss.getSheetByName(OTR_TAB);
  if (s) return s;
  s = ss.insertSheet(OTR_TAB);
  s.getRange(1, 1, 1, OTR_HEAD.length).setValues([OTR_HEAD]).setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  s.setRowHeight(1, 40); s.setFrozenRows(1);
  [130, 105, 90, 90, 85, 170, 280, 80, 100, 60].forEach((w, i) => s.setColumnWidth(i + 1, w));
  s.hideColumns(OTR_HEAD.length);
  return s;
}

function otrReservas_(desde, hasta) {
  const s = hojaOtr_();
  if (!s || s.getLastRow() < 2) return [];
  const n = s.getLastRow() - 1, V = s.getRange(2, 1, n, OTR_HEAD.length).getValues(), D = s.getRange(2, 1, n, OTR_HEAD.length).getDisplayValues(), out = [];
  for (let i = 0; i < n; i++) {
    if (String(V[i][8]) === 'Cancelada') continue;
    const d = fechaDe_(V[i][1]);
    if (!d) continue;
    const f = Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd');
    if ((desde && f < desde) || (hasta && f > hasta)) continue;
    const ent = String(V[i][4]) === 'Sí', hm = horaDe_(D[i][2]);
    out.push({ f: f, i: ent || !hm ? ABRE : hm.dec, d: ent || !hm ? DIA_ENTERO : (Number(V[i][3]) || 2), e: 'C', r: 'OTR-' + (i + 2), p: Number(V[i][7]) || 0 });
  }
  return out;
}

function evaGuardar(o) {
  if (!o || o.k !== otrosKey_()) throw new Error('Enlace no válido');
  const dia = fechaDe_(o.fecha);
  if (!dia) throw new Error('Falta la fecha');
  const ent = !!o.entero, hm = ent ? null : horaDe_(o.hora);
  if (!ent && !hm) throw new Error('Falta la hora (o marca «Día entero»)');
  const dur = ent ? 0 : (Number(o.dur) || 2);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = hojaOtr_();
    const serv = String(o.serv || 'Otro servicio').slice(0, 80), notas = String(o.notas || '').slice(0, 500), pax = Number(o.pax) || '';
    s.appendRow([new Date(), dia, hm ? hm.txt : '', dur || '', ent ? 'Sí' : 'No', serv, notas, pax, 'Activa', '']);
    const row = s.getLastRow();
    s.getRange(row, 1).setNumberFormat('dd/MM/yyyy HH:mm'); s.getRange(row, 2).setNumberFormat('dd/MM/yyyy'); s.getRange(row, 3).setNumberFormat('@').setValue(hm ? hm.txt : '');
    const cal = CalendarApp.getCalendarById(CAL_ID), tit = '🔒 ' + serv + (notas ? ' · ' + notas.slice(0, 60) : '') + (pax ? ' · ' + pax + ' pers' : '');
    const desc = 'Reserva de otro servicio apuntada por Adventure Co. (formulario del equipo).\n' + (notas ? 'Notas: ' + notas : '');
    let ev;
    if (ent) ev = cal.createAllDayEvent(tit, dia, { description: desc });
    else { const ini = new Date(dia.getFullYear(), dia.getMonth(), dia.getDate(), hm.h, hm.m); ev = cal.createEvent(tit, ini, new Date(ini.getTime() + dur * 3600000), { description: desc }); }
    ev.setColor(CalendarApp.EventColor.GRAY);
    s.getRange(row, 10).setValue(ev.getId());
    CacheService.getScriptCache().removeAll(['bloq_', 'bloq_' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd')]);
    return evaLista(o.k);
  } finally { lock.releaseLock(); }
}

function evaCancelar(k, row) {
  if (k !== otrosKey_()) throw new Error('Enlace no válido');
  const s = hojaOtr_(); row = Number(row);
  if (!(row >= 2 && row <= s.getLastRow())) throw new Error('No encontrada');
  s.getRange(row, 9).setValue('Cancelada');
  const id = String(s.getRange(row, 10).getValue() || '');
  if (id) { try { const ev = CalendarApp.getCalendarById(CAL_ID).getEventById(id); if (ev) ev.deleteEvent(); } catch (err) { console.error(err); } }
  s.getRange(row, 10).setValue('');
  return evaLista(k);
}

function evaLista(k) {
  if (k !== otrosKey_()) throw new Error('Enlace no válido');
  const s = hojaOtr_();
  if (!s || s.getLastRow() < 2) return [];
  const hoy = Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd'), n = s.getLastRow() - 1;
  const V = s.getRange(2, 1, n, OTR_HEAD.length).getValues(), D = s.getRange(2, 1, n, OTR_HEAD.length).getDisplayValues(), out = [];
  for (let i = 0; i < n; i++) {
    if (String(V[i][8]) === 'Cancelada') continue;
    const d = fechaDe_(V[i][1]); if (!d) continue;
    const f = Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd'); if (f < hoy) continue;
    out.push({ row: i + 2, f: f, fecha: D[i][1], hora: String(V[i][4]) === 'Sí' ? 'Día entero' : D[i][2] + ' (' + D[i][3] + ' h)', serv: D[i][5], notas: D[i][6], pax: D[i][7] });
  }
  return out.sort((a, b) => a.f < b.f ? -1 : 1).slice(0, 60);
}

function evaPagina_(k) {
  if (k !== otrosKey_()) return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Enlace no válido.</p>');
  const h = `<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap" rel="stylesheet">
<style>body{margin:0;background:#f7f4ee;font-family:Poppins,sans-serif;color:#3d3a35}.w{max-width:560px;margin:0 auto;padding:20px 16px 40px}h1{color:#2d4a1e;font-size:22px;margin:6px 0 4px}p.s{color:#6d6a64;font-size:14px;margin:0 0 16px}
.card{background:#fff;border:1px solid #e3ded4;border-radius:16px;padding:18px;margin-bottom:16px}label{display:block;font-weight:600;font-size:14px;margin:12px 0 6px}input,select,textarea{width:100%;box-sizing:border-box;padding:12px;border:1.5px solid #e3ded4;border-radius:10px;font:15px Poppins,sans-serif;background:#fff}
.row{display:flex;gap:10px}.row>div{flex:1}.chk{display:flex;align-items:center;gap:10px;margin-top:14px;font-weight:600}.chk input{width:22px;height:22px}button{width:100%;margin-top:18px;padding:15px;border:0;border-radius:30px;background:#e8704e;color:#fff;font:700 16px Poppins,sans-serif;cursor:pointer}
button:disabled{opacity:.6}.msg{margin-top:12px;font-size:14px;font-weight:600}.ok{color:#1e5631}.er{color:#b4471f}.it{display:flex;justify-content:space-between;gap:10px;border-top:1px solid #eee;padding:10px 0;font-size:14px}.it b{color:#2d4a1e}.it small{color:#6d6a64;display:block}.x{width:auto;margin:0;padding:6px 12px;font-size:13px;background:#f1efea;color:#b4471f}</style></head><body><div class="w">
<h1>Apuntar reserva de otro servicio</h1><p class="s">Quedará en el calendario y la web no dejará reservar team building ni escuelas en ese horario.</p>
<div class="card"><label>Fecha</label><input type="date" id="f"><label class="chk"><input type="checkbox" id="e"> Día entero (bloquear todo el día)</label>
<div class="row" id="hh"><div><label>Hora de inicio</label><input type="time" id="h" value="10:00" step="900"></div><div><label>Duración</label><select id="d"><option value="1">1 h</option><option value="1.5">1 h 30</option><option value="2" selected>2 h</option><option value="2.5">2 h 30</option><option value="3">3 h</option><option value="4">4 h</option><option value="5">5 h</option><option value="6">6 h</option><option value="8">8 h</option></select></div></div>
<label>Servicio</label><select id="sv"><option>Paintball</option><option>Paintball júnior</option><option>Multiaventura</option><option>Tiro con arco</option><option>Orientación</option><option>Trekking</option><option>BTT</option><option>Pack</option><option>Cumpleaños</option><option>Alquiler de espacio</option><option>Otro servicio</option></select>
<div class="row"><div><label>Personas</label><input type="number" id="p" min="1" placeholder="Ej. 12"></div></div>
<label>Cliente / notas</label><textarea id="n" rows="2" placeholder="Nombre, teléfono, detalles…"></textarea>
<button id="b">Guardar en el calendario</button><div class="msg" id="m"></div></div>
<div class="card"><b style="color:#2d4a1e">Próximas reservas apuntadas</b><div id="l"><p class="s">Cargando…</p></div></div></div>
<script>var K=${JSON.stringify(k)};function $(i){return document.getElementById(i)}
$('f').valueAsDate=new Date();$('e').onchange=function(){$('hh').style.display=this.checked?'none':'flex'};
function pinta(L){if(!L.length){$('l').innerHTML='<p class="s">No hay ninguna.</p>';return}$('l').innerHTML=L.map(function(x){return '<div class="it"><div><b>'+x.fecha+' · '+x.hora+'</b><small>'+esc(x.serv)+(x.pax?' · '+x.pax+' pers':'')+(x.notas?' · '+esc(x.notas):'')+'</small></div><button class="x" onclick="can('+x.row+')">Anular</button></div>'}).join('')}
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
function msg(t,ok){$('m').className='msg '+(ok?'ok':'er');$('m').textContent=t}
$('b').onclick=function(){var o={k:K,fecha:$('f').value,entero:$('e').checked,hora:$('h').value,dur:$('d').value,serv:$('sv').value,pax:$('p').value,notas:$('n').value};if(!o.fecha){msg('Elige la fecha');return}$('b').disabled=true;msg('Guardando…',true);
google.script.run.withSuccessHandler(function(L){$('b').disabled=false;msg('✅ Guardada en el calendario',true);$('n').value='';$('p').value='';pinta(L)}).withFailureHandler(function(e){$('b').disabled=false;msg('Error: '+e.message,false)}).evaGuardar(o)};
function can(r){if(!confirm('¿Anular esta reserva?'))return;google.script.run.withSuccessHandler(pinta).withFailureHandler(function(e){msg('Error: '+e.message,false)}).evaCancelar(K,r)}
google.script.run.withSuccessHandler(pinta).evaLista(K);</script></body></html>`;
  return HtmlService.createHtmlOutput(h).setTitle('Adventure Co. · Otras reservas').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/** Días con un team building confirmado: el día queda completo (1 team building por día). */
function tbDias_(desde) { return reservas_(desde).filter(x => x.tb && x.e === 'C').map(x => x.f); }
