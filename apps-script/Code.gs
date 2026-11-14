/**
 * MULTIPAR S.A. — Guardia v2 · Puente con Drive
 *
 * Entrega a la página el Excel de la guardia tal como está en Drive en ese momento.
 * La página lo lee con el mismo lector que el archivo local (assets/guardia-reader.js),
 * así que acá NO se interpreta nada: solo se busca el archivo y se envía.
 *
 * INSTALACIÓN
 * 1. script.google.com → Proyecto nuevo → pegar este archivo en Code.gs.
 * 2. Configuración del proyecto (⚙️) → Propiedades del script → agregar:
 *      ARCHIVO_ID   = id del archivo de la guardia en Drive (lo que va entre /d/ y /edit en el link)
 *      CLAVE_ACCESO = la clave que se escribe en la página
 * 3. Implementar → Nueva implementación → Aplicación web
 *      Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier usuario
 *    La primera vez pide permisos de Drive: aceptar.
 * 4. Copiar la URL (termina en /exec) en assets/config.js del repo.
 * 5. Si se cambia este código: Implementar → Gestionar implementaciones → editar → Nueva versión.
 *
 * Sirve tanto si el archivo es un .xlsx subido a Drive como si es un Google Sheet.
 *
 * HISTORIAL DE LECTURAS (estadísticas automáticas)
 * Cada 30 minutos guarda una "foto" de las hojas clave de la guardia en una planilla APARTE
 * ("Guardia MULTIPAR - Historial de lecturas", la crea el script solo). Solo guarda una hoja si cambió
 * desde la última foto. Para activarlo, una sola vez: ejecutar la función  instalarDisparador .
 * Requiere que el archivo de la guardia sea un Google Sheet (no un .xlsx subido).
 */
var TIPO_SHEET = 'application/vnd.google-apps.spreadsheet';

function doGet(e) {
  var props = PropertiesService.getScriptProperties();
  var claveEsperada = props.getProperty('CLAVE_ACCESO');
  if (claveEsperada && (e.parameter.clave || '') !== claveEsperada) return salida({ error: 'No autorizado' });

  var id = props.getProperty('ARCHIVO_ID');
  if (!id) return salida({ error: 'Falta la propiedad ARCHIVO_ID en el script' });

  // Historial de lecturas (para las estadísticas): /exec?clave=...&modo=historial&desde=2026-11-14T00:00:00
  if (e.parameter.modo === 'historial') return salida(leerHistorial(e.parameter.desde || ''));

  try {
    var archivo = DriveApp.getFileById(id);
    var blob;
    if (archivo.getMimeType() === TIPO_SHEET) {
      var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx', {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() !== 200) return salida({ error: 'No pude exportar la planilla (' + resp.getResponseCode() + ')' });
      blob = resp.getBlob();
    } else {
      blob = archivo.getBlob();
    }
    var notas = [];
    if (archivo.getMimeType() === TIPO_SHEET) {
      try { notas = leerNotas(id); } catch (e2) { notas = []; }
    }
    return salida({
      notas: notas,
      nombre: archivo.getName(),
      modificado: archivo.getLastUpdated().toISOString(),
      consultado: new Date().toISOString(),
      xlsx: Utilities.base64Encode(blob.getBytes())
    });
  } catch (err) {
    return salida({ error: 'No pude abrir el archivo: ' + err.message });
  }
}

/** Notas de celda (el triangulito negro) de todas las hojas: [{hoja, celda, fila, texto}] */
function leerNotas(id) {
  var out = [];
  SpreadsheetApp.openById(id).getSheets().forEach(function (hoja) {
    var rango = hoja.getDataRange();
    var notas = rango.getNotes();
    var f0 = rango.getRow(), c0 = rango.getColumn();
    for (var i = 0; i < notas.length; i++) {
      for (var j = 0; j < notas[i].length; j++) {
        if (notas[i][j]) out.push({ hoja: hoja.getName(), celda: hoja.getRange(f0 + i, c0 + j).getA1Notation(), fila: f0 + i, texto: notas[i][j] });
      }
    }
  });
  return out;
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


// ───────────────────────── HISTORIAL DE LECTURAS ─────────────────────────
var HOJAS_HISTORIAL = ['TURNO', 'FRANCOS', 'SUBIDAS', 'BAJADAS', 'CAMPANA'];

/** Ejecutar UNA vez a mano: crea el disparador cada 30 min y saca la primera foto. */
function instalarDisparador() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'guardarLectura') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('guardarLectura').timeBased().everyMinutes(30).create();
  guardarLectura();
}

/** Planilla de historial: la crea la primera vez y guarda su id en las propiedades del script. */
function planillaHistorial() {
  var props = PropertiesService.getScriptProperties();
  var hid = props.getProperty('HISTORIAL_ID');
  if (hid) { try { return SpreadsheetApp.openById(hid); } catch (e) { /* si se borró, se crea de nuevo */ } }
  var ss = SpreadsheetApp.create('Guardia MULTIPAR - Historial de lecturas');
  var h = ss.getSheets()[0]; h.setName('LECTURAS');
  h.getRange(1, 1, 1, 5).setValues([['LECTURA', 'HOJA', 'HASH', 'FILAS', 'DATOS (JSON)']]).setFontWeight('bold');
  h.setFrozenRows(1);
  props.setProperty('HISTORIAL_ID', ss.getId());
  return ss;
}

/** Saca la foto: por cada hoja clave, si cambió respecto de la última foto, agrega una fila con su contenido visible. */
function guardarLectura() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('ARCHIVO_ID');
    if (!id) return;
    var origen = SpreadsheetApp.openById(id);
    var hist = planillaHistorial().getSheetByName('LECTURAS');
    var ahora = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', "yyyy-MM-dd'T'HH:mm:ss");
    var nuevas = [];
    HOJAS_HISTORIAL.forEach(function (nombre) {
      var hoja = origen.getSheetByName(nombre);
      if (!hoja) return;
      var vals = hoja.getDataRange().getDisplayValues();
      while (vals.length && vals[vals.length - 1].every(function (c) { return c === ''; })) vals.pop();   // sin filas vacías al final
      var json = JSON.stringify(vals);
      var hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json));
      if (hash === props.getProperty('HASH_' + nombre)) return;
      if (json.length > 49000) json = json.slice(0, 49000);   // tope de una celda; las hojas de la guardia están muy por debajo
      nuevas.push([ahora, nombre, hash, vals.length, json]);
      props.setProperty('HASH_' + nombre, hash);
    });
    if (nuevas.length) hist.getRange(hist.getLastRow() + 1, 1, nuevas.length, 5).setValues(nuevas);
  } finally {
    lock.releaseLock();
  }
}

/** Devuelve las fotos desde una fecha (ISO, hora de Buenos Aires): [{lectura, hoja, filas, datos}] */
function leerHistorial(desde) {
  var hid = PropertiesService.getScriptProperties().getProperty('HISTORIAL_ID');
  if (!hid) return { error: 'Todavía no hay historial: ejecutar instalarDisparador en el script.' };
  var hist = SpreadsheetApp.openById(hid).getSheetByName('LECTURAS');
  var ult = hist.getLastRow();
  if (ult < 2) return { lecturas: [] };
  var filas = hist.getRange(2, 1, ult - 1, 5).getValues();
  var out = [];
  filas.forEach(function (f) {
    var l = String(f[0]);
    if (desde && l < desde) return;
    var datos = null; try { datos = JSON.parse(f[4]); } catch (e) { datos = null; }
    out.push({ lectura: l, hoja: f[1], filas: f[3], datos: datos });
  });
  return { lecturas: out };
}
