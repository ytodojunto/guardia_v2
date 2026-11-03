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
 */
var TIPO_SHEET = 'application/vnd.google-apps.spreadsheet';

function doGet(e) {
  var props = PropertiesService.getScriptProperties();
  var claveEsperada = props.getProperty('CLAVE_ACCESO');
  if (claveEsperada && (e.parameter.clave || '') !== claveEsperada) return salida({ error: 'No autorizado' });

  var id = props.getProperty('ARCHIVO_ID');
  if (!id) return salida({ error: 'Falta la propiedad ARCHIVO_ID en el script' });

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
