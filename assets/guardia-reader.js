/*
 * guardia-reader.js — lector del Excel de la guardia MULTIPAR (estructura v5)
 *
 * Funciona en el navegador y en Node. Necesita ExcelJS (se le pasa como parámetro).
 *   const datos = await GuardiaReader.leerGuardia(ExcelJS, arrayBufferOBuffer);
 *
 * Principios:
 *  - Lee por ENCABEZADO y por ANCLA (rótulos de la columna A), nunca por fila/columna fija.
 *  - No rompe: lo que no entiende lo devuelve crudo y lo anota en `avisos`.
 *  - Los colores se interpretan con la leyenda "REFERENCIAS DE COLOR" del propio archivo.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.GuardiaReader = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var VERSION = '0.1';

  // ───────────── utilidades de texto / valores ─────────────
  function norm(s) {
    return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toUpperCase().replace(/\s+/g, ' ').trim();
  }
  function limpio(s) { return String(s).replace(/\\/g, '').replace(/\s+/g, ' ').trim(); }
  function pad(n) { return String(n).padStart(2, '0'); }
  function ymd(d) { return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  var EPOCH_1899 = Date.UTC(1899, 11, 30);

  // Valor "plano" de una celda de ExcelJS (fórmulas, rich text, etc.)
  function raw(cell) {
    // Celda combinada: el valor vive en la celda de arriba a la izquierda; ExcelJS lo repite en el resto
    // del rango, así que esas se tratan como vacías (si no, un dato combinado se leería varias veces).
    if (cell && cell.isMerged && cell.master && cell.master !== cell && cell.master.address !== cell.address) return null;
    var v = cell ? cell.value : null;
    if (v == null) return null;
    if (typeof v === 'object' && !esDate(v)) {
      if ('result' in v) v = v.result;
      else if (v.richText) v = v.richText.map(function (t) { return t.text; }).join('');
      else if (v.text != null) v = v.text;
      else return null;
    }
    if (v == null) return null;
    if (typeof v === 'string') { v = limpio(v); if (v === '') return null; }
    return v;
  }

  // ───────────── fechas y horas ─────────────
  function esDate(v) { return Object.prototype.toString.call(v) === '[object Date]'; }
  function esFecha(v) { return esDate(v) && v.getUTCFullYear() >= 1950; }
  function esSoloHora(v) { return esDate(v) && v.getUTCFullYear() < 1950; }

  function aFecha(v) {
    if (esFecha(v)) return ymd(v);
    if (typeof v === 'number' && v >= 1) return ymd(new Date(EPOCH_1899 + Math.round(v) * 86400000));
    if (typeof v === 'string') {
      var m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
      if (m) {
        var y = +m[3]; if (y < 100) y += 2000;
        return y + '-' + pad(+m[2]) + '-' + pad(+m[1]);
      }
    }
    return null;
  }

  // → { hora:'HH:MM', franja:null } | { hora:null, franja:'AM'|'PM' } | null (no entendida)
  function aHora(v) {
    if (v == null) return { hora: null, franja: null };
    if (esSoloHora(v)) {
      var mins = Math.round((((v.getTime() - EPOCH_1899) % 86400000) + 86400000) % 86400000 / 60000);
      return { hora: pad(Math.floor(mins / 60) % 24) + ':' + pad(mins % 60), franja: null };
    }
    if (typeof v === 'number' && v >= 0 && v < 1) {
      var m2 = Math.round(v * 1440);
      return { hora: pad(Math.floor(m2 / 60) % 24) + ':' + pad(m2 % 60), franja: null };
    }
    if (typeof v === 'string') {
      var s = norm(v);
      if (s === 'AM' || s === 'PM') return { hora: null, franja: s };
      var m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
      if (m && +m[1] < 24 && +m[2] < 60) return { hora: pad(+m[1]) + ':' + pad(+m[2]), franja: null };
    }
    return null;
  }

  // Día "dd" + hora sin mes → fecha completa, eligiendo el mes más cercano a la fecha de referencia
  function resolverDia(dia, hh, mm, ref) {
    if (!ref) return null;
    var p = ref.split('-').map(Number), refT = Date.UTC(p[0], p[1] - 1, p[2]), best = null, bd = Infinity;
    [-1, 0, 1].forEach(function (dm) {
      var t = Date.UTC(p[0], p[1] - 1 + dm, dia), dt = new Date(t);
      if (dt.getUTCDate() !== dia) return;
      var diff = Math.abs(t - refT);
      if (diff < bd) { bd = diff; best = dt; }
    });
    return best ? ymd(best) + 'T' + pad(hh) + ':' + pad(mm) : null;
  }

  // ───────────── puertos (equivalencias) ─────────────
  function esPend(v) { return v === '?' || (v != null && /^A\s*CONFIRMAR$/i.test(String(v).trim())); }
  var PUERTOS_ALIAS = {
    SL: ['SL', 'SANLORENZO'], ROS: ['ROS', 'ROSARIO'], AS: ['AS', 'ARROYOSECO'],
    SN: ['SN', 'SANNICOLAS'], SP: ['SP', 'SANPEDRO'], TIM: ['TIM', 'TIMBUES'],
    CPNA: ['CPNA', 'CAMPANA'], ZTE: ['ZTE', 'ZARATE'], ESC: ['ESC', 'ESCOBAR'],
    RLLO: ['RLLO', 'RAMALLO'], NPAL: ['NPAL', 'NUEVAPALMIRA', 'NUEVAPALMIRA(UY)', 'NUEVAPALMIRAUY'], LP: ['LP', 'LASPALMAS'],
    K171: ['K171', 'KM171', 'KM171PG'], ZC: ['ZC', 'ZONACOMUN'], VCON: ['VCON', 'VILLACONSTITUCION'],
    BL: ['BL', 'BRAZOLARGO'], GUAZU: ['GUAZU', 'TERMINALGUAZU', 'DELGUAZU'], DD: ['DD', 'DELTADOCK'],
    PPER: ['PPER', 'PUNTAPEREIRA', 'PUNTAPEREIRA(UY)', 'PUNTAPEREIRAUY'], K238: ['K238', 'KM238PA', 'KM238'],
    IBI: ['IBI', 'IBICUY'], MAZ: ['MAZ', 'MAZARUCA'], GLAG: ['GLAG', 'GENERALLAGOS'], VGG: ['VGG'],
    SM: ['SM', 'SANMARTIN'], DIAM: ['DIAM', 'DIAMANTE'], SF: ['SF', 'SANTAFE'], PNA: ['PNA', 'PARANA'], ATU: ['ATU', 'ATUCHA']
  };
  var ALIAS_A_CODIGO = {};
  Object.keys(PUERTOS_ALIAS).forEach(function (k) {
    PUERTOS_ALIAS[k].forEach(function (a) { ALIAS_A_CODIGO[a] = k; });
  });
  function codigoPuerto(texto) {
    if (texto == null) return null;
    return ALIAS_A_CODIGO[norm(texto).replace(/[.\s]/g, '')] || null;
  }

  // ───────────── colores / leyenda ─────────────
  function argb(cell) {
    var f = cell && cell.fill;
    if (!f || f.type !== 'pattern' || f.pattern !== 'solid' || !f.fgColor || !f.fgColor.argb) return null;
    return f.fgColor.argb.slice(-6).toUpperCase();
  }
  function distancia(a, b) {
    var d = 0;
    for (var i = 0; i < 3; i++) {
      var x = parseInt(a.substr(i * 2, 2), 16) - parseInt(b.substr(i * 2, 2), 16);
      d += x * x;
    }
    return Math.sqrt(d);
  }
  function idLeyenda(etiqueta) {
    var n = norm(etiqueta);
    if (/^X\s*=/.test(n) || /DESPACHADO$/.test(n) && /^X/.test(n)) return 'despachado';
    if (/PROXIMO A DESPACHAR/.test(n)) return 'proximo_a_despachar';
    if (/DESPACHADO NUESTRO/.test(n)) return 'despachado_nuestro';
    if (/MANIOBRA/.test(n)) return 'maniobra';
    if (/DEMORA/.test(n)) return 'demora';
    if (/CUBIERTO EN FRANCO/.test(n)) return 'cubierto_en_franco';
    if (/CUBIERTO POR OTRA/.test(n)) return 'cubierto_por_otra_empresa'; // nos cubren
    if (/CUBIERTO A OTRA/.test(n)) return 'cubierto_a_otra_empresa';     // cubrimos a otra
    if (/CANCELADO/.test(n)) return 'cancelado';
    if (/CALADO DESDE/.test(n)) return 'calado_2_practicos';
    return null;
  }
  var TOLERANCIA_COLOR = 28; // el verde cambió de 75A67C a 7FB77E entre versiones: se tolera
  function colorId(hex, leyenda) {
    if (!hex || !leyenda.length) return null;
    var mejor = null, md = Infinity;
    leyenda.forEach(function (e) { var d = distancia(hex, e.hex); if (d < md) { md = d; mejor = e; } });
    return md <= TOLERANCIA_COLOR ? mejor.id : null;
  }

  // ───────────── lector principal ─────────────
  function parseGuardia(wb) {
    var avisos = [], vistos = {};
    function aviso(nivel, hoja, celda, tipo, mensaje) {
      var k = [nivel, hoja, celda, tipo].join('|');
      if (vistos[k]) return; vistos[k] = 1;
      avisos.push({ nivel: nivel, hoja: hoja, celda: celda || null, tipo: tipo, mensaje: mensaje });
    }
    var out = {
      version: VERSION, hojas: wb.worksheets.map(function (w) { return w.name; }),
      guardia: null, leyendaColores: [], catalogo: [],
      turno: [], guardiaRosario: [],
      viajes: { aRosario: [], deRosario: [], aCampana: [], deCampana: [], movimientos: [] },
      francos: { enFranco: [], fueraDeTurno: [], licencias: [] },
      programacion: {}, servicios: {}, avisos: avisos
    };

    function hoja(nombre) {
      var ws = wb.getWorksheet(nombre);
      if (!ws) aviso('error', nombre, null, 'hoja_faltante', 'No encuentro la hoja ' + nombre);
      return ws;
    }

    // ── catálogo de prácticos (hoja LISTAS, columna B) ──
    var catalogoSet = {};
    var wsL = hoja('LISTAS');
    if (wsL) {
      for (var r = 2; r <= wsL.rowCount; r++) {
        var n = raw(wsL.getCell(r, 2));
        if (n) { out.catalogo.push(limpio(n)); catalogoSet[norm(n)] = true; }
      }
    }
    function esPropio(nombre) { return !!catalogoSet[norm(nombre)]; }

    // ── hoja TURNO ──
    var wsT = hoja('TURNO');
    var leyenda = [];
    var refFecha = null;

    if (wsT) {
      var NC = Math.max(wsT.columnCount, 30), NR = wsT.rowCount;

      // encabezado de la guardia (filas 1-4)
      var fechas = [], horas = [], operador = null, coord = null;
      for (var rr = 1; rr <= 4; rr++) {
        for (var cc = 1; cc <= NC; cc++) {
          var v = raw(wsT.getCell(rr, cc));
          if (v == null) continue;
          if (rr === 2 && esFecha(v)) fechas.push(ymd(v));
          else if (rr === 2 && esSoloHora(v)) horas.push(aHora(v).hora);
          var t = typeof v === 'string' ? norm(v) : '';
          if (t === 'OPERADOR' || t === 'COORD.' || t === 'COORDINADOR') {
            for (var k = cc + 1; k <= NC; k++) {
              var vv = raw(wsT.getCell(rr, k));
              if (vv != null) { if (t === 'OPERADOR') operador = limpio(vv); else coord = limpio(vv); break; }
            }
          }
        }
      }
      out.guardia = {
        fechaInicio: fechas[0] || null, horaInicio: horas[0] || null,
        fechaFin: fechas[1] || null, horaFin: horas[1] || null,
        operador: operador, coordinador: coord
      };
      refFecha = out.guardia.fechaInicio;
      if (!refFecha) aviso('error', 'TURNO', 'D2', 'fecha_guardia', 'No encuentro la fecha de la guardia en el encabezado: no puedo completar los días de los puertos.');

      // anclas en columna A
      var ANCLAS = {
        'GUARDIA ROSARIO': 'rosario', 'EN VIAJE A ROSARIO': 'aRosario', 'EN VIAJE DE ROSARIO': 'deRosario',
        'EN VIAJE A CAMPANA': 'aCampana', 'EN VIAJE DE CAMPANA': 'deCampana', 'MOVIMIENTOS': 'movimientos',
        'REFERENCIAS DE COLOR': 'referencias'
      };
      var anclas = [];
      for (var ra = 1; ra <= NR; ra++) {
        var ta = norm(raw(wsT.getCell(ra, 1)) || '');
        if (!ta) continue;
        if (ta.indexOf('LISTADO DE PRACTICOS') === 0) anclas.push({ id: 'turno', fila: ra });
        else if (ANCLAS[ta]) anclas.push({ id: ANCLAS[ta], fila: ra });
      }
      Object.keys(ANCLAS).map(function (k) { return ANCLAS[k]; }).concat(['turno']).forEach(function (id) {
        if (!anclas.some(function (a) { return a.id === id; }))
          aviso('error', 'TURNO', null, 'ancla_faltante', 'No encuentro el bloque "' + id + '" en la columna A.');
      });
      anclas.sort(function (a, b) { return a.fila - b.fila; });
      anclas.forEach(function (a, i) { a.fin = (anclas[i + 1] ? anclas[i + 1].fila - 1 : NR); });
      function bloque(id) { return anclas.filter(function (a) { return a.id === id; })[0]; }

      // leyenda de colores (se lee del propio archivo)
      var bl = bloque('referencias');
      if (bl) {
        for (var rl = bl.fila + 1; rl <= bl.fin; rl++) {
          var etiqueta = null;
          for (var cl = 2; cl <= 5 && !etiqueta; cl++) { var e = raw(wsT.getCell(rl, cl)); if (typeof e === 'string') etiqueta = e; }
          if (!etiqueta) continue;
          var hexL = null;
          for (var ch = 1; ch <= 5 && !hexL; ch++) hexL = argb(wsT.getCell(rl, ch));
          if (!hexL) continue;
          var idL = idLeyenda(etiqueta);
          if (!idL) { aviso('aviso', 'TURNO', 'A' + rl, 'leyenda_desconocida', 'Color de la leyenda sin significado conocido: "' + etiqueta + '"'); continue; }
          leyenda.push({ id: idL, etiqueta: etiqueta, hex: hexL });
        }
      }
      if (!leyenda.length) aviso('error', 'TURNO', null, 'sin_leyenda', 'No pude leer la leyenda de colores: no voy a interpretar colores.');
      out.leyendaColores = leyenda;

      // lectura de encabezados de un bloque
      function encabezado(b, claves) {
        var hr = null;
        for (var r2 = b.fila + 1; r2 <= Math.min(b.fila + 3, b.fin); r2++) {
          var a = norm(raw(wsT.getCell(r2, 1)) || '');
          if (a === 'NRO' || a === 'SERVICIO') { hr = r2; break; }
        }
        if (!hr) { aviso('error', 'TURNO', null, 'encabezado_faltante', 'No encuentro la fila de encabezados bajo "' + b.id + '".'); return null; }
        var campos = [];
        for (var c = 1; c <= NC; c++) { var tx = norm(raw(wsT.getCell(hr, c)) || ''); if (tx) campos.push({ key: tx, col: c }); }
        campos.forEach(function (f, i) { f.fin = campos[i + 1] ? campos[i + 1].col - 1 : NC; });
        var mapa = {};
        campos.forEach(function (f) { if (!mapa[f.key]) mapa[f.key] = f; });
        var falta = claves.filter(function (k) { return !mapa[k]; });
        if (falta.length) { aviso('error', 'TURNO', 'fila ' + hr, 'encabezado_faltante', 'Faltan columnas en "' + b.id + '": ' + falta.join(', ')); return null; }
        return { fila: hr, f: mapa };
      }
      function leer(r2, f) {
        for (var c = f.col; c <= f.fin; c++) { var v2 = raw(wsT.getCell(r2, c)); if (v2 != null) return { v: v2, cell: wsT.getCell(r2, c) }; }
        return { v: null, cell: wsT.getCell(r2, f.col) };
      }
      function colorDe(r2, f) {
        for (var c = f.col; c <= f.fin; c++) { var id = colorId(argb(wsT.getCell(r2, c)), leyenda); if (id) return id; }
        return null;
      }
      function texto(x) { return x.v == null ? null : limpio(x.v); }

      // ── listas de prácticos (turno y guardia Rosario) ──
      function listaPracticos(id, destino) {
        var b = bloque(id); if (!b) return;
        var h = encabezado(b, ['PRACTICO', 'FECHA E/T', 'HORA']); if (!h) return;
        var pos = 0, repetidos = {};
        for (var r2 = h.fila + 1; r2 <= b.fin; r2++) {
          pos++;
          var p = leer(r2, h.f['PRACTICO']);
          if (p.v == null) continue;
          var nombre = limpio(p.v), fecha = leer(r2, h.f['FECHA E/T']), hora = leer(r2, h.f['HORA']);
          var fE = aFecha(fecha.v), hE = aHora(hora.v);
          if (fecha.v != null && !fE) aviso('aviso', 'TURNO', fecha.cell.address, 'fecha_invalida', 'Fecha que no entiendo: "' + fecha.v + '" (' + nombre + ')');
          if (hE === null) { aviso('aviso', 'TURNO', hora.cell.address, 'hora_invalida', 'Hora que no entiendo: "' + hora.v + '" (' + nombre + ')'); hE = { hora: null, franja: null }; }
          var marca = h.f['☎'] ? leer(r2, h.f['☎']) : { v: null };
          var m = marca.v == null ? null : limpio(marca.v).toUpperCase();
          if (m && m !== '*' && m !== 'X') { aviso('aviso', 'TURNO', marca.cell && marca.cell.address, 'marca_invalida', 'Marca ☎ rara: "' + m + '" (' + nombre + ')'); m = null; }
          if (repetidos[norm(nombre)]) aviso('aviso', 'TURNO', p.cell.address, 'repetido', nombre + ' aparece más de una vez en ' + id + '.');
          repetidos[norm(nombre)] = 1;
          if (!esPropio(nombre)) aviso('info', 'TURNO', p.cell.address, 'fuera_de_catalogo', nombre + ' no está en el catálogo de prácticos propios.');
          var obs = h.f['OBSERVACION'] ? texto(leer(r2, h.f['OBSERVACION'])) : null;
          destino.push({
            pos: pos, practico: nombre, marca: m,
            entrada: fE ? fE + (hE.hora ? 'T' + hE.hora : '') : null,
            fecha: fE, hora: hE.hora, observacion: obs,
            estado: colorDe(r2, h.f['PRACTICO']), fila: r2
          });
        }
      }
      listaPracticos('turno', out.turno);
      listaPracticos('rosario', out.guardiaRosario);

      // ── buques en viaje ──
      function punto(x, hoja2) {
        if (x.v == null) return null;
        var s = limpio(x.v);
        var m = s.match(/^([A-Za-zÁÉÍÓÚÑáéíóúñ.\s]+?)\s*:?\s*(\d{1,2})\/(\d{1,2}):(\d{2})$/);
        if (m) {
          var cod = codigoPuerto(m[1]);
          if (!cod) aviso('aviso', hoja2, x.cell.address, 'puerto_desconocido', 'Puerto que no conozco: "' + m[1].trim() + '"');
          var iso = resolverDia(+m[2], +m[3], +m[4], refFecha);
          return { raw: s, puerto: cod, puertoTexto: m[1].trim(), cuando: iso };
        }
        if (/\d/.test(s)) {
          aviso('aviso', hoja2, x.cell.address, 'formato_no_reconocido', 'No entiendo "' + s + '" (esperaba PUERTO dd/hh:mm).');
          return { raw: s, puerto: null, cuando: null };
        }
        return { raw: s, puerto: codigoPuerto(s), cuando: null }; // solo un destino, sin hora
      }
      var CANALES = ['CEM', 'CMG', 'RPG', 'RPB', 'RPP', 'RPA'];
      function viajes(id, destino) {
        var b = bloque(id); if (!b) return;
        var h = encabezado(b, ['SERVICIO', 'BUQUE', 'PRACTICO 1', 'PRACTICO 2', 'PUERTO E INICIO', 'DESEMBARCO 1', 'DESEMBARCO 2']);
        if (!h) return;
        // Dos formatos: todo en un texto ("SL 01/09:50") o en tres celdas (puerto | FECHA | HORA).
        function leerPunto(r2, f) {
          var s1 = norm(raw(wsT.getCell(h.fila, f.col + 1)) || ''), s2 = norm(raw(wsT.getCell(h.fila, f.col + 2)) || '');
          if (s1 !== 'FECHA' || s2 !== 'HORA') return punto(leer(r2, f), 'TURNO');
          var p = wsT.getCell(r2, f.col), d = wsT.getCell(r2, f.col + 1), t = wsT.getCell(r2, f.col + 2);
          var pv = raw(p), dv = raw(d), tv = raw(t);
          if (pv == null && dv == null && tv == null) return null;
          pv = pv == null ? null : limpio(pv);
          var fe = aFecha(dv), ho = aHora(tv);
          if (dv != null && !fe) aviso('aviso', 'TURNO', d.address, 'fecha_invalida', 'Fecha que no entiendo: "' + dv + '"');
          if (tv != null && ho === null) { aviso('aviso', 'TURNO', t.address, 'hora_invalida', 'Hora que no entiendo: "' + tv + '"'); }
          ho = ho || { hora: null, franja: null };
          var cod = pv ? codigoPuerto(pv) : null;
          if (pv && pv !== '?' && !cod) aviso('aviso', 'TURNO', p.address, 'puerto_desconocido', 'Puerto que no conozco: "' + pv + '"');
          if (ho.hora && !fe) aviso('aviso', 'TURNO', d.address, 'falta_fecha', 'Hay hora pero falta la fecha (' + (pv || p.address) + ').');
          if (fe && refFecha && Math.abs(Date.parse(fe) - Date.parse(refFecha)) > 20 * 86400000)
            aviso('aviso', 'TURNO', d.address, 'fecha_lejana', 'La fecha ' + fe + ' queda lejos de la guardia (' + refFecha + '): ¿día o mes mal puesto?');
          return { raw: [pv, fe, ho.hora].filter(Boolean).join(' '), puerto: cod, puertoTexto: pv, pendiente: pv === '?',
                   cuando: fe ? fe + (ho.hora ? 'T' + ho.hora : '') : null };
        }
        for (var r2 = h.fila + 1; r2 <= b.fin; r2++) {
          var sv = leer(r2, h.f['SERVICIO']), bq = leer(r2, h.f['BUQUE']);
          if (sv.v == null && bq.v == null) continue;
          var servicio = sv.v == null ? null : String(sv.v).replace(/\.0$/, '');
          var p1 = texto(leer(r2, h.f['PRACTICO 1'])), p2 = texto(leer(r2, h.f['PRACTICO 2']));
          if (bq.v == null && !p1 && !p2) {
            aviso('info', 'TURNO', 'fila ' + r2, 'fila_sin_buque', 'La fila tiene número de servicio (' + servicio + ') pero no buque ni prácticos: se ignora (restos de una fila anterior).');
            continue;
          }
          if (bq.v != null && servicio == null)
            aviso('aviso', 'TURNO', 'fila ' + r2, 'sin_servicio', 'Falta el número de servicio de ' + limpio(bq.v) + ': sin él no se cruza con la programación.');
          [p1, p2].forEach(function (p) { if (p && !esPropio(p)) aviso('info', 'TURNO', 'fila ' + r2, 'cobertura', p + ' no es de MULTIPAR (cobertura) — buque ' + (bq.v || servicio)); });
          var cal = h.f['CALADO'] ? leer(r2, h.f['CALADO']) : { v: null }, calado = null;
          if (cal.v != null) {
            calado = typeof cal.v === 'number' ? cal.v : Number(String(cal.v).replace(',', '.'));
            if (isNaN(calado)) { aviso('aviso', 'TURNO', cal.cell.address, 'calado_invalido', 'Calado que no entiendo: "' + cal.v + '"'); calado = null; }
          }
          var can = h.f['CANAL'] ? texto(leer(r2, h.f['CANAL'])) : null;
          if (can && CANALES.indexOf(can.toUpperCase()) < 0) aviso('aviso', 'TURNO', 'fila ' + r2, 'canal_invalido', 'Canal fuera de la lista: "' + can + '"');
          var reg = {
            bloque: id, servicio: servicio, buque: texto(bq),
            practico1: p1, practico2: p2, calado: calado, canal: can,
            inicio: leerPunto(r2, h.f['PUERTO E INICIO']),
            desembarco1: leerPunto(r2, h.f['DESEMBARCO 1']),
            desembarco2: leerPunto(r2, h.f['DESEMBARCO 2']),
            observaciones: h.f['OBSERVACIONES'] ? texto(leer(r2, h.f['OBSERVACIONES'])) : null,
            colores: {
              buque: colorDe(r2, h.f['BUQUE']), practico1: colorDe(r2, h.f['PRACTICO 1']),
              practico2: colorDe(r2, h.f['PRACTICO 2']), calado: h.f['CALADO'] ? colorDe(r2, h.f['CALADO']) : null
            },
            ejemplo: !!(servicio && /^99\d{3}$/.test(servicio)), fila: r2
          };
          destino.push(reg);
          if (servicio) {
            out.servicios[servicio] = out.servicios[servicio] || { viajes: [], programacion: [] };
            out.servicios[servicio].viajes.push({ bloque: id, fila: r2 });
          }
        }
      }
      viajes('aRosario', out.viajes.aRosario);
      viajes('deRosario', out.viajes.deRosario);
      viajes('aCampana', out.viajes.aCampana);
      viajes('deCampana', out.viajes.deCampana);
      viajes('movimientos', out.viajes.movimientos);
    }

    // ── FRANCOS ──
    var wsF = hoja('FRANCOS');
    if (wsF) {
      var seccion = 'enFranco', hdr = null;
      for (var rf = 1; rf <= wsF.rowCount; rf++) {
        var a1 = raw(wsF.getCell(rf, 1)), ta1 = typeof a1 === 'string' ? norm(a1) : '';
        if (ta1.indexOf('FUERA DE TURNO') === 0) { seccion = 'fueraDeTurno'; hdr = null; continue; }
        if (ta1.indexOf('LICENCIAS SOLICITADAS') === 0) { seccion = 'licencias'; hdr = null; continue; }
        if (ta1 === 'NRO') { hdr = rf; continue; }
        if (seccion === 'licencias') { if (typeof a1 === 'string' && ta1.indexOf('LISTADO') !== 0) out.francos.licencias.push(a1); continue; }
        if (!hdr) continue;
        var pr = raw(wsF.getCell(rf, 3)); if (pr == null) continue;
        var nombreF = limpio(pr);
        var d1 = aFecha(raw(wsF.getCell(rf, 4))), h1 = aHora(raw(wsF.getCell(rf, 5)));
        var d2 = aFecha(raw(wsF.getCell(rf, 6))), h2 = aHora(raw(wsF.getCell(rf, 7)));
        var obsF = raw(wsF.getCell(rf, 8));
        if (!d1) aviso('aviso', 'FRANCOS', 'D' + rf, 'fecha_invalida', 'Fecha de salida que no entiendo (' + nombreF + ')');
        if (seccion === 'enFranco' && !d2) aviso('aviso', 'FRANCOS', 'F' + rf, 'fecha_invalida', 'Falta fecha de ingreso (' + nombreF + ')');
        if (seccion === 'enFranco' && !esPropio(nombreF)) aviso('info', 'FRANCOS', 'C' + rf, 'fuera_de_catalogo', nombreF + ' no está en el catálogo.');
        out.francos[seccion].push({
          practico: nombreF,
          salida: d1 ? d1 + (h1 && h1.hora ? 'T' + h1.hora : '') : null,
          ingreso: d2 ? d2 + (h2 && h2.hora ? 'T' + h2.hora : '') : null,
          observacion: obsF == null ? null : limpio(obsF), fila: rf
        });
      }
      // cruces útiles
      var enTurno = {}; out.turno.concat(out.guardiaRosario).forEach(function (t) { enTurno[norm(t.practico)] = 1; });
      out.francos.enFranco.forEach(function (f) {
        if (enTurno[norm(f.practico)]) aviso('aviso', 'FRANCOS', 'C' + f.fila, 'franco_y_turno', f.practico + ' figura de franco y también en el turno.');
      });
      out.francos.fueraDeTurno.forEach(function (f) {
        if (enTurno[norm(f.practico)]) aviso('aviso', 'FRANCOS', 'C' + f.fila, 'fuera_de_turno_y_turno', f.practico + ' figura fuera de turno (enfermo/licencia) y también en la lista del turno' + (f.salida ? ' — fuera desde ' + f.salida.replace('T', ' ') : '') + '. Revisar cuál es el dato correcto.');
      });
    }

    // ── programación: SUBIDAS / BAJADAS / CAMPANA ──
    var MESES = { ENERO: 1, FEBRERO: 2, MARZO: 3, ABRIL: 4, MAYO: 5, JUNIO: 6, JULIO: 7, AGOSTO: 8, SEPTIEMBRE: 9, SETIEMBRE: 9, OCTUBRE: 10, NOVIEMBRE: 11, DICIEMBRE: 12 };
    function programacion(nombre, clave) {
      var ws = hoja(nombre); if (!ws) return;
      var NC2 = Math.max(ws.columnCount, 15), hr = null;
      for (var r2 = 1; r2 <= Math.min(10, ws.rowCount) && !hr; r2++)
        for (var c = 1; c <= NC2; c++) if (norm(raw(ws.getCell(r2, c)) || '') === 'AGENCIA') { hr = r2; break; }
      if (!hr) { aviso('error', nombre, null, 'encabezado_faltante', 'No encuentro el encabezado (columna AGENCIA).'); return; }
      var col = {};
      for (var c2 = 1; c2 <= NC2; c2++) {
        var tx = norm(raw(ws.getCell(hr, c2)) || '');
        if (tx && !col[tx]) col[tx] = c2;
      }
      var colETA = col['Z.C. // ETD'] || col['CONFIRMADO'];
      var obligatorias = { SERVICIO: col['SERVICIO'], BUQUE: col['BUQUE'], 'RECALADA / PREG.': col['RECALADA / PREG.'], ETA: colETA, DESDE: col['DESDE'], HASTA: col['HASTA'], AUTORIZADO: col['AUTORIZADO'] };
      var falt = Object.keys(obligatorias).filter(function (k) { return !obligatorias[k]; });
      if (falt.length) { aviso('error', nombre, 'fila ' + hr, 'encabezado_faltante', 'Faltan columnas: ' + falt.join(', ')); return; }
      var sub = norm(raw(ws.getCell(hr + 1, col['RECALADA / PREG.'])) || '') + '/' + norm(raw(ws.getCell(hr + 1, col['RECALADA / PREG.'] + 1)) || '');
      if (sub !== 'FECHA/HORA') aviso('aviso', nombre, 'fila ' + (hr + 1), 'subencabezado', 'Esperaba FECHA / HORA bajo RECALADA y veo "' + sub + '".');
      var colMarca = col['☎'] || 1;

      var info = { publicado: null, buques: [], coefPromedioPlanilla: null, coefPromedio: null };
      var tit = String(raw(ws.getCell(2, 1)) || '');
      var mt = norm(tit).match(/(\d{1,2}) DE ([A-Z]+) DE (\d{4})/), mh = tit.match(/(\d{1,2}):(\d{2})/);
      if (mt && MESES[mt[2]]) info.publicado = mt[3] + '-' + pad(MESES[mt[2]]) + '-' + pad(+mt[1]) + (mh ? 'T' + pad(+mh[1]) + ':' + mh[2] : '');
      else aviso('aviso', nombre, 'A2', 'titulo_sin_fecha', 'No pude leer la fecha de publicación del título.');

      function fechaHora(r3, cF, tag) {
        var fv = raw(ws.getCell(r3, cF)), hv = raw(ws.getCell(r3, cF + 1));
        var f = aFecha(fv), h = aHora(hv);
        if (fv != null && !f) aviso('aviso', nombre, ws.getCell(r3, cF).address, 'fecha_invalida', tag + ': fecha que no entiendo "' + fv + '"');
        if (h === null) { aviso('aviso', nombre, ws.getCell(r3, cF + 1).address, 'hora_invalida', tag + ': hora que no entiendo "' + hv + '"'); h = { hora: null, franja: null }; }
        return { fecha: f, hora: h.hora, franja: h.franja };
      }
      function txt(r3, c3) { var v3 = raw(ws.getCell(r3, c3)); return v3 == null ? null : limpio(v3); }
      function num(r3, c3, tag) {
        var v3 = raw(ws.getCell(r3, c3)); if (v3 == null) return null;
        var n3 = typeof v3 === 'number' ? v3 : Number(String(v3).replace(',', '.'));
        if (isNaN(n3)) { aviso('aviso', nombre, ws.getCell(r3, c3).address, 'numero_invalido', tag + ': no es un número "' + v3 + '"'); return null; }
        return n3;
      }
      for (var r3 = hr + 2; r3 <= ws.rowCount; r3++) {
        var fila = [];
        for (var c3 = 1; c3 <= 6; c3++) fila.push(norm(raw(ws.getCell(r3, c3)) || ''));
        if (fila.join(' ').indexOf('COEF. PROMEDIO') >= 0) {
          info.coefPromedioPlanilla = col['COEF.'] ? num(r3, col['COEF.'], 'coef. promedio') : null;
          break;
        }
        if (fila.some(function (t) { return t.indexOf('REFERENCIAS') === 0; })) break;
        var sv = raw(ws.getCell(r3, col['SERVICIO'])), bq = raw(ws.getCell(r3, col['BUQUE']));
        if (sv == null && bq == null) continue;
        var servicio = sv == null ? null : String(sv).replace(/\.0$/, '');
        var mk = raw(ws.getCell(r3, colMarca)), marca = mk == null ? null : limpio(mk).toUpperCase();
        if (marca && marca !== '*' && marca !== 'X') { aviso('aviso', nombre, ws.getCell(r3, colMarca).address, 'marca_invalida', 'Marca ☎ rara: "' + marca + '"'); marca = null; }
        var aut = txt(r3, col['AUTORIZADO']), autN = aut ? aut.toUpperCase() : null;
        if (autN && autN !== 'SI' && autN !== 'NO') aviso('info', nombre, ws.getCell(r3, col['AUTORIZADO']).address, 'autorizado_dudoso', 'AUTORIZADO = "' + aut + '" (' + (bq || servicio) + ')');
        var desde = txt(r3, col['DESDE']), hasta = txt(r3, col['HASTA']);
        [['DESDE', desde], ['HASTA', hasta]].forEach(function (p) {
          if (p[1] && p[1] !== '?' && !codigoPuerto(p[1])) aviso('aviso', nombre, ws.getCell(r3, col[p[0]]).address, 'puerto_desconocido', 'Puerto que no conozco en ' + p[0] + ': "' + p[1] + '"');
        });
        var colores = null;
        for (var cx = 1; cx <= NC2 && !colores; cx++) { var idc = colorId(argb(ws.getCell(r3, cx)), leyenda); if (idc) colores = idc; }
        var reg = {
          hoja: nombre, marca: marca, agencia: txt(r3, col['AGENCIA']), servicio: servicio, buque: bq == null ? null : limpio(bq),
          coef: col['COEF.'] ? num(r3, col['COEF.'], 'coef.') : null, calado: col['CALADO'] ? num(r3, col['CALADO'], 'calado') : null,
          recalada: fechaHora(r3, col['RECALADA / PREG.'], 'recalada'),
          eta: fechaHora(r3, colETA, clave === 'campana' ? 'confirmado' : 'ETD'),
          desde: desde, hasta: hasta, desdeCod: codigoPuerto(desde), hastaCod: codigoPuerto(hasta),
          desdePendiente: esPend(desde), hastaPendiente: esPend(hasta), // "?" = destino a confirmar
          autorizado: autN, terminal: txt(r3, col['TERMINAL']), observaciones: txt(r3, col['OBSERVACIONES']),
          estado: colores, fila: r3
        };
        info.buques.push(reg);
        if (servicio) {
          out.servicios[servicio] = out.servicios[servicio] || { viajes: [], programacion: [] };
          out.servicios[servicio].programacion.push({ hoja: nombre, fila: r3 });
        }
      }
      var cs = info.buques.map(function (b) { return b.coef; }).filter(function (x) { return typeof x === 'number'; });
      info.coefPromedio = cs.length ? Math.round(cs.reduce(function (a, b) { return a + b; }, 0) / cs.length * 10) / 10 : null;
      if (info.publicado && refFecha && info.publicado.slice(0, 10) < refFecha)
        aviso('aviso', nombre, 'A2', 'programacion_vieja', 'La programación está publicada el ' + info.publicado.replace('T', ' ') + ', anterior a la guardia del ' + refFecha + '.');
      out.programacion[clave] = info;
    }
    programacion('SUBIDAS', 'subidas');
    programacion('BAJADAS', 'bajadas');
    programacion('CAMPANA', 'campana');

    // ── NOVEDADES OPERATIVAS (fecha | informó | novedad; una novedad puede seguir en filas siguientes con solo el texto) ──
    out.novedades = [];
    var wsN = hoja('NOVEDADES OPERATIVAS');
    if (wsN) {
      var ultimaNov = null;
      for (var rn = 2; rn <= wsN.rowCount; rn++) {
        var fN = raw(wsN.getCell(rn, 1)), aN = raw(wsN.getCell(rn, 2)), tN = raw(wsN.getCell(rn, 3));
        if (fN == null && aN == null && tN == null) continue;
        if (fN == null && aN == null && ultimaNov) { ultimaNov.texto += '\n' + limpio(tN); continue; }
        ultimaNov = { fecha: esFecha(fN) ? ymd(fN) : (fN == null ? null : limpio(fN)), autor: aN == null ? null : limpio(aN), texto: tN == null ? '' : limpio(tN), fila: rn };
        out.novedades.push(ultimaNov);
      }
    }

    // ── GUARDIA REMISES: dos bloques lado a lado (07 a 19 desde la columna A, 19 a 07 desde la E) ──
    out.remises = null;
    var wsR = hoja('GUARDIA REMISES');
    if (wsR) {
      var bloqueRem = function (c0) {
        var b = { titulo: null, fecha: null, remiseros: [], otros: [], novedades: [] }, sec = 'rem';
        var t0 = raw(wsR.getCell(1, c0)); b.titulo = t0 == null ? null : limpio(t0);
        var f0 = raw(wsR.getCell(2, c0 + 1)); b.fecha = esFecha(f0) ? ymd(f0) : null;
        for (var rr2 = 4; rr2 <= wsR.rowCount; rr2++) {
          var a0 = raw(wsR.getCell(rr2, c0)), ta0 = typeof a0 === 'string' ? norm(a0) : '';
          if (ta0.indexOf('FRANCO') === 0) { sec = 'otros'; continue; }
          if (ta0 === 'NOVEDADES') { sec = 'nov'; continue; }
          if (sec === 'nov') { if (a0 != null) b.novedades.push(limpio(a0)); continue; }
          var nm = raw(wsR.getCell(rr2, c0 + 1)), ob = raw(wsR.getCell(rr2, c0 + 2));
          if (nm == null) continue;
          var reg2 = { ord: typeof a0 === 'number' ? a0 : null, nombre: limpio(nm), observacion: ob == null ? null : limpio(ob) };
          (sec === 'rem' ? b.remiseros : b.otros).push(reg2);
        }
        return b;
      };
      out.remises = { turno1: bloqueRem(1), turno2: bloqueRem(5) };
    }

    // ── TELEFONOS de las lanchas (puerto | teléfonos) y enlaces útiles ──
    out.telefonos = []; out.enlaces = [];
    var wsTel = hoja('TELEFONOS');
    if (wsTel) {
      for (var rt = 2; rt <= wsTel.rowCount; rt++) {
        var pT = raw(wsTel.getCell(rt, 1)); if (pT == null) continue;
        if (/^https?:/i.test(String(pT))) {
          var dsc = raw(wsTel.getCell(rt + 1, 1));
          out.enlaces.push({ url: String(pT).trim(), descripcion: dsc != null && !/^https?:/i.test(String(dsc)) ? limpio(dsc) : null });
          continue;
        }
        var partes = [];
        for (var ct = 2; ct <= Math.max(wsTel.columnCount, 12); ct++) { var vt = raw(wsTel.getCell(rt, ct)); if (vt != null) partes.push(limpio(vt)); }
        if (partes.length) out.telefonos.push({ puerto: limpio(pT), telefono: partes.join(' · ') });
      }
    }

    return out;
  }

  // Algunos programas guardan las imágenes (el logo del encabezado) con un formato interno
  // que ExcelJS no entiende y se cae. Las imágenes no sirven para leer la guardia:
  // si pasa, se vacían los dibujos y se reintenta (necesita JSZip).
  var DIBUJO_VACIO = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>';

  async function cargar(ExcelJS, datos) {
    var wb = new ExcelJS.Workbook();
    await wb.xlsx.load(datos);
    return wb;
  }

  async function leerGuardia(ExcelJS, datos, JSZipOpcional) {
    var wb, avisoDibujos = false;
    try {
      wb = await cargar(ExcelJS, datos);
    } catch (e) {
      var Z = JSZipOpcional || (typeof self !== 'undefined' && self.JSZip) || null;
      if (!Z) throw e;
      var zip = await Z.loadAsync(datos);
      zip.file(/^xl\/drawings\/drawing\d+\.xml$/).forEach(function (f) { zip.file(f.name, DIBUJO_VACIO); });
      var hojas = zip.file(/^xl\/worksheets\/sheet\d+\.xml$/);
      for (var i = 0; i < hojas.length; i++) { // las "tablas" de Excel (hoja de novedades) tampoco se usan
        var xml = await hojas[i].async('string');
        zip.file(hojas[i].name, xml.replace(/<tableParts[\s\S]*?<\/tableParts>/g, ''));
      }
      wb = await cargar(ExcelJS, await zip.generateAsync({ type: 'uint8array' }));
      avisoDibujos = true;
    }
    var out = parseGuardia(wb);
    if (avisoDibujos) out.avisos.push({ nivel: 'info', hoja: '(archivo)', celda: null, tipo: 'imagenes_ignoradas', mensaje: 'El archivo trae imágenes y tablas en un formato que no se lee; se ignoraron (no afectan los datos).' });
    return out;
  }

  return { VERSION: VERSION, parseGuardia: parseGuardia, leerGuardia: leerGuardia, codigoPuerto: codigoPuerto };
});
