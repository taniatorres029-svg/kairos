/** @OnlyCurrentDoc */
// Guarda cada solicitud del formulario /factura en la hoja "Facturas".
// Instalar: en la hoja de Google, Extensiones > Apps Script, pegar este archivo,
// Implementar > Nueva implementación > Aplicación web
// (Ejecutar como: Yo · Quién tiene acceso: Cualquier persona) y copiar la URL /exec.

const HOJA = 'Facturas';
const COLUMNAS = [
  'Recibido', 'Completado', 'Folio', 'Atendido por',
  'RFC', 'Nombre / razón social', 'C.P. fiscal', 'Régimen fiscal', 'Uso CFDI', 'Correo', 'Teléfono',
  'Paciente', 'Concepto', 'Fecha de pago', 'Monto', 'Forma de pago', 'Referencia',
];

function hoja_() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const sh = libro.getSheetByName(HOJA) || libro.insertSheet(HOJA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNAS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold').setBackground('#e3efeb');
  }
  return sh;
}

// Evita que un texto que empieza con = + - @ se interprete como fórmula.
function texto_(v) {
  const s = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const f = d.filas || {};
  const fila = COLUMNAS.map(c => {
    if (c === 'Recibido') return new Date();
    if (c === 'Completado') return false;
    if (c === 'Folio') return texto_(d.folio);
    if (c === 'Monto') return Number(f[c]) || 0;
    return texto_(f[c]);
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = hoja_();
    sh.appendRow(fila);
    const r = sh.getLastRow();
    sh.getRange(r, 2).insertCheckboxes();
    sh.getRange(r, 15).setNumberFormat('$#,##0.00');
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Ejecutar una vez desde el editor para crear la hoja y dar permisos.
function preparar() {
  hoja_();
}
