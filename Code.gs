/**
 * Chaqueta de Ladera - backend de la encuesta (Google Apps Script)
 *
 * Pegar en: Hoja de cálculo > Extensiones > Apps Script > Code.gs
 * Implementar > Gestionar implementaciones > Editar (lápiz) > Versión: "Nueva versión"
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier persona
 * Así la URL /exec NO cambia y el sitio sigue funcionando.
 *
 * El sitio envía (POST, text/plain): {"calificacion": 1-5, "conocia": "Sí"|"No"}
 * y lee (GET): [{"r": 5, "c": "Sí"}, ...]
 */

var SHEET_NAME = 'Respuestas';
var HEADERS = ['Fecha', 'Calificación', '¿Conocía la chaqueta?'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita que dos envíos simultáneos se pisen
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch (err) { data = {}; }
    }
    // Respaldo: también acepta parámetros de formulario (?calificacion=5&conocia=Sí)
    if (e && e.parameter) {
      if (data.calificacion == null) data.calificacion = e.parameter.calificacion;
      if (data.conocia == null) data.conocia = e.parameter.conocia;
    }

    var r = parseInt(data.calificacion, 10);
    var c = String(data.conocia || '').trim();
    if (!(r >= 1 && r <= 5)) return json_({ ok: false, error: 'calificacion inválida' });
    if (/^s/i.test(c)) c = 'Sí'; else if (/^n/i.test(c)) c = 'No';
    else return json_({ ok: false, error: 'conocia inválido' });

    getSheet_().appendRow([new Date(), r, c]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return json_([]);
  var rows = sh.getRange(2, 1, last - 1, 3).getValues();
  var out = [];
  rows.forEach(function (row) {
    var r = parseInt(row[1], 10);
    if (r >= 1 && r <= 5) out.push({ r: r, c: String(row[2] || '') });
  });
  return json_(out);
}

/** Prueba manual: selecciona "probar" arriba y dale Ejecutar. Luego borra la fila de prueba. */
function probar() {
  var res = doPost({ postData: { contents: JSON.stringify({ calificacion: 5, conocia: 'Sí' }) } });
  Logger.log(res.getContent());
  Logger.log(doGet().getContent());
}
