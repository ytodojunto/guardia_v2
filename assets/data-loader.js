// ── CARGA DE DATOS EN VIVO DESDE GOOGLE SHEETS (vía Apps Script) ──
(function () {

  function esc(str) {
    return String(str || '').replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function soloDigitos(str) {
    return String(str || '').replace(/\D/g, '');
  }

  // ── ESTADO DE CARGA ──────────────────────────────────────────────
  // El HTML arranca con "loading-skel" (puntitos animados) en vez de datos
  // de ejemplo, para que nunca se confunda un placeholder con un dato real.
  function limpiarSkeletons() {
    document.querySelectorAll('.loading-skel').forEach(function (el) { el.remove(); });
  }

  function ocultarBanner() {
    var banner = document.getElementById('loadingBanner');
    if (banner) banner.classList.add('oculto');
    limpiarSkeletons();
  }

  function mostrarErrorBanner(msg) {
    var banner = document.getElementById('loadingBanner');
    if (banner) {
      banner.classList.remove('oculto');
      banner.classList.add('error');
      banner.textContent = '⚠️ ' + msg;
    }
  }

  // ── LECTURA: Apps Script entrega el Excel de Drive, GuardiaReader lo interpreta y se adapta al diseño ──
  function aBytes(b64) { var bin = atob(b64), u = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u.buffer; }
  function dmy(iso) { var m = /^(\d{4})-(\d\d)-(\d\d)/.exec(iso || ''); return m ? m[3] + '/' + m[2] + '/' + m[1] : ''; }
  function dm(iso) { var m = /^(\d{4})-(\d\d)-(\d\d)/.exec(iso || ''); return m ? m[3] + '/' + m[2] : ''; }
  function hh(iso) { var m = /T(\d\d:\d\d)/.exec(iso || ''); return m ? m[1] : ''; }
  function dmh(iso) { return (dm(iso) + ' ' + hh(iso)).trim(); }
  var ESTADO = { despachado_nuestro: 'Despachado', despachado: 'Despachado', maniobra: 'Maniobra', demora: 'Demora', cubierto_en_franco: 'Cubre franco', cubierto_por_otra_empresa: 'Nos cubren', cubierto_a_otra_empresa: 'Cubrimos', cancelado: 'Cancelado', proximo_a_despachar: 'Próximo' };

  function puntoTxt(etq, p) {
    if (!p) return null;
    return etq + ': ' + (p.puertoTexto || p.raw || '?') + (p.cuando ? ' ' + dmh(p.cuando) : '');
  }
  function tarjetaDeViaje(v) {
    var datos = [puntoTxt('Inicio', v.inicio), puntoTxt('Desembarco 1', v.desembarco1), puntoTxt('Desembarco 2', v.desembarco2)].filter(Boolean);
    return { buque: (v.servicio ? v.servicio + ' · ' : '') + (v.buque || '(sin nombre)'), fecha: '', canal: v.canal || '',
             calado: v.calado != null ? String(v.calado).replace('.', ',') + ' m' : '',
             practicos: [v.practico1, v.practico2].filter(Boolean), datos: datos, observaciones: v.observaciones || '', ejemplo: v.ejemplo };
  }

  // Notas de celda (las trae el Apps Script: [{hoja, celda, fila, texto}]). Se pegan al práctico / buque de esa fila.
  function aplicarNotas(g, notas) {
    (notas || []).forEach(function (n) {
      var txt = String(n.texto || '').trim(); if (!txt) return;
      var h = String(n.hoja || '').toUpperCase(), hit = false;
      if (h === 'TURNO') {
        g.turno.forEach(function (t) { if (t.fila === n.fila) { t.nota = (t.nota ? t.nota + ' · ' : '') + txt; hit = true; } });
        Object.keys(g.viajes || {}).forEach(function (k) {
          (g.viajes[k] || []).forEach(function (v) { if (v.fila === n.fila) { v.observaciones = (v.observaciones ? v.observaciones + ' · ' : '') + '📝 ' + txt; hit = true; } });
        });
      }
      if (!hit) (g.novedades = g.novedades || []).push({ fecha: '', autor: 'Nota ' + n.hoja + ' ' + n.celda, texto: txt });
    });
  }

  function adaptar(g, meta) {
    var fueraSet = {};
    g.francos.fueraDeTurno.forEach(function (f) { fueraSet[f.practico.toUpperCase()] = f; });
    // buque en el que figura cada práctico dentro de los bloques de viajes
    var buqueDe = {};
    Object.keys(g.viajes || {}).forEach(function (k) {
      (g.viajes[k] || []).forEach(function (v) {
        if (!v.buque) return;
        [v.practico1, v.practico2].forEach(function (n) { if (n) buqueDe[String(n).toUpperCase()] = v.buque; });
      });
    });
    var turno = {
      practicos: g.turno.map(function (t) {
        var bq = buqueDe[String(t.practico).toUpperCase()];
        var dest = t.observacion || '';
        if (bq && t.estado === 'despachado_nuestro' && dest.toUpperCase().indexOf(bq.toUpperCase()) < 0) dest = bq;
        return { pos: t.pos, nombre: t.practico, cambio: t.marca === '*', destino: dest,
                 detalle: 'Entró ' + dmh(t.entrada) + (t.nota ? ' · 📝 ' + t.nota : ''), tipo: fueraSet[t.practico.toUpperCase()] ? '⚠ Figura fuera de turno' : (ESTADO[t.estado] || '') };
      }),
      francos: g.francos.enFranco.map(function (f, i) {
        return { nro: i + 1, nombre: f.practico, cambio: false, fechaSalida: dm(f.salida), horaSalida: hh(f.salida),
                 fechaIngreso: dm(f.ingreso), horaIngreso: hh(f.ingreso), obs: f.observacion };
      }),
      enfermos: [],
      observaciones: g.francos.fueraDeTurno.map(function (f) { return f.practico + ' (fuera de turno)' + (f.salida ? ' desde ' + dmh(f.salida) : '') + (f.observacion ? ' — ' + f.observacion : ''); })
        .concat(g.francos.licencias || []),
      coordinador: g.guardia.coordinador, operador: g.guardia.operador, determinantes: [],
      telefonos: (g.telefonos || []).concat((g.enlaces || []).map(function (e) { return { puerto: '🔗', telefono: (e.descripcion ? e.descripcion + ' — ' : '') + e.url }; }))
    };

    var previstos = [];
    [['subidas', 'subida', 'ETD'], ['bajadas', 'bajada', 'ETD'], ['campana', 'campana', 'Confirmado']].forEach(function (k) {
      var p = g.programacion[k[0]]; if (!p) return;
      p.buques.forEach(function (b) {
        var obs = [];
        if (b.autorizado === 'SI') obs.push('AUTORIZADO'); else if (b.autorizado === 'NO') obs.push('NO AUT');
        if (b.terminal) obs.push('Terminal ' + b.terminal);
        if (b.observaciones) obs.push(b.observaciones);
        var rec = b.recalada || {}, eta = b.eta || {};
        previstos.push({ tipo: k[1], buque: b.buque || '', coef: b.coef == null ? '' : b.coef, agencia: b.agencia || '', marca: b.marca || '',
          fechaRecalada: dmy(rec.fecha), horaRecalada: rec.hora || rec.franja || '', labelFecha1: 'Recalada',
          fechaConfirmado: dmy(eta.fecha), horaConfirmado: eta.hora || eta.franja || '', labelFecha2: k[2],
          desde: b.desde || '?', hasta: b.hasta || '?', observaciones: obs.join(' · ') });
      });
    });

    var transitos = {
      aRosario: g.viajes.aRosario.map(tarjetaDeViaje),
      deRosario: g.viajes.deRosario.map(tarjetaDeViaje),
      aCampana: g.viajes.aCampana.map(tarjetaDeViaje),
      deCampana: g.viajes.deCampana.map(tarjetaDeViaje),
      movimientos: g.viajes.movimientos.map(tarjetaDeViaje),
      guardiaRosario: g.guardiaRosario.map(function (t) { return { practico: (t.marca === '*' ? '📞 ' : '') + t.practico, hora: 'Entró ' + dmh(t.entrada), tipo: t.observacion || '' }; })
    };

    var novedades = (g.novedades || []).slice().sort(function (a, b) { return String(b.fecha || '').localeCompare(String(a.fecha || '')); })
      .map(function (n) { return { autor: n.autor || 'Guardia', texto: n.texto, fecha: /^\d{4}-/.test(n.fecha || '') ? dmy(n.fecha) : (n.fecha || '') }; });

    var R = g.remises;
    var remises = R ? { turno1: R.turno1, turno2: R.turno2 } : null;

    var G = g.guardia;
    return { turno: turno, previstos: previstos, transitos: transitos, novedades: novedades, remises: remises,
             avisos: g.avisos.filter(function (a) { return a.nivel === 'error' || a.tipo === 'fuera_de_turno_y_turno' || a.tipo === 'franco_y_turno' || a.tipo === 'sin_servicio' || a.tipo === 'turno_viaje_distinto' || a.tipo === 'viaje_sin_despacho'; }).map(function (a) { return a.hoja + (a.celda ? ' ' + a.celda : '') + ': ' + a.mensaje; }),
             encabezado: 'Guardia ' + dmh(G.fechaInicio + 'T' + (G.horaInicio || '')) + ' → ' + dmh(G.fechaFin + 'T' + (G.horaFin || '')) + (meta ? ' · planilla guardada ' + meta : '') + ' · leída ' + (function (d) { return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); })(new Date()) + (window.GUARDIA_VERSION ? ' · ' + window.GUARDIA_VERSION : '') };
  }

  async function cargarDatos() {
    var url = window.GUARDIA_API_URL;
    if (!url) { mostrarErrorBanner('Falta configurar la conexión con Drive.'); return; }
    var clave = sessionStorage.getItem('mp_clave') || '';
    try {
      var resp = await fetch(url + (url.indexOf('?') === -1 ? '?' : '&') + 'clave=' + encodeURIComponent(clave), { cache: 'no-store' });
      var d = await resp.json();
      if (d.error) {
        console.error('Backend:', d.error);
        mostrarErrorBanner(/autorizado/i.test(d.error) ? 'No se pudo autenticar con Drive. Revisá la clave y recargá.' : d.error);
        return;
      }
      var g = await GuardiaReader.leerGuardia(ExcelJS, aBytes(d.xlsx), window.JSZip);
      aplicarNotas(g, d.notas);
      var m = d.modificado ? new Date(d.modificado) : null;
      var data = adaptar(g, m ? m.toLocaleDateString('es-AR') + ' ' + m.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' }) : '');
      window.__guardia = { leida: g, vista: data };
      renderTurno(data.turno);
      renderPrevistos(data.previstos);
      renderNovedades(data.novedades);
      renderRemises(data.remises);
      renderBuques(data.transitos);
      renderGuardiaRosarioTurno(data.transitos.guardiaRosario);
      var fechaEl = document.querySelector('.hdr-fecha');
      if (fechaEl) fechaEl.textContent = '📅 ' + data.encabezado;
      limpiarSkeletons();
      if (data.avisos.length) mostrarErrorBanner('Revisar planilla: ' + data.avisos.join(' · '));
      else { var banner = document.getElementById('loadingBanner'); if (banner) banner.classList.add('oculto'); }
    } catch (err) {
      console.error('Error cargando datos de guardia:', err);
      mostrarErrorBanner('No se pudieron cargar los datos. Revisá tu conexión y recargá la página.');
    }
  }

  // ── TURNO ──────────────────────────────────────────────────────
  function renderTurno(t) {
    if (!t) return;
    var listaTurno = document.getElementById('listaTurno');
    var listaFrancos = document.getElementById('listaFrancos');
    var cardVarios = document.getElementById('cardVarios');
    if (!listaTurno || !listaFrancos) return;

    listaTurno.innerHTML = t.practicos.map(function (p, i) {
      // El asterisco de la planilla significa "el operador se comunicó con el práctico"
      // (no un cambio de turno) — se muestra como una etiqueta discreta, sin recolorear el nombre.
      return '<div class="prac-row" data-tipo="turno' + (p.cambio ? ' cambio' : '') + '" data-nombre="' + esc(p.nombre) + '" data-pos="' + esc(p.pos) + '" data-hora="' + esc(soloDigitos(p.detalle)) + '">' +
        '<div class="prac-num">' + esc(p.pos) + '</div>' +
        '<div class="prac-info"><div class="prac-nombre">' + esc(p.nombre) + (p.cambio ? ' <span class="badge b-verde" style="font-size:9px;vertical-align:middle">📞 Contactado</span>' : '') + '</div>' +
        '<div class="prac-detalle">' + esc(p.destino) + (p.destino && p.detalle ? ' — ' : '') + esc(p.detalle) + '</div></div>' +
        '<div class="prac-badge">' + esc(p.tipo || '—') + '</div></div>';
    }).join('') || '<div class="empty">Sin prácticos en turno cargados</div>';

    listaFrancos.innerHTML = t.francos.map(function (f, i) {
      return '<div class="franco-row" data-tipo="franco' + (f.cambio ? ' cambio' : '') + '" data-nombre="' + esc(f.nombre) + '">' +
        '<div class="franco-top"><div class="franco-num">' + esc(f.nro || (i + 1)) + '</div>' +
        '<div class="franco-nombre">' + esc(f.nombre) + '</div><span class="badge b-naranja">FRANCO</span></div>' +
        '<div class="franco-fechas"><span>Salida: ' + esc(f.fechaSalida) + ' – ' + esc(f.horaSalida) + '</span>' +
        (f.fechaIngreso ? '<span>Regresa: ' + esc(f.fechaIngreso) + ' – ' + esc(f.horaIngreso) + '</span>' : '') +
        (f.obs && f.obs.toUpperCase() !== 'FRANCO' ? '<span>' + esc(f.obs) + '</span>' : '') + '</div></div>';
    }).join('') || '<div class="empty">Sin francos cargados</div>';

    if (cardVarios) {
      var varios = '';
      if (t.enfermos.length) {
        varios += '<div class="sec-label">Enfermos</div>';
        varios += t.enfermos.map(function (en) {
          return '<div class="varios-row" data-tipo="varios" data-nombre="' + esc(en.nombre) + '">' +
            '<div class="varios-icon">🏥</div><div><div class="varios-nombre">' + esc(en.nombre) + '</div>' +
            '<div class="varios-detalle">Desde ' + esc(en.fecha) + (en.hora ? ' – ' + esc(en.hora) : '') + '</div></div></div>';
        }).join('');
      }
      if (t.observaciones.length) {
        varios += '<div class="sec-label">Observaciones</div>';
        varios += t.observaciones.map(function (obs) {
          return '<div class="varios-row" data-tipo="varios"><div class="varios-icon">📋</div><div><div class="varios-detalle">' + esc(obs) + '</div></div></div>';
        }).join('');
      }
      cardVarios.innerHTML = '<div class="card-head"><span>🔵</span><span class="card-head-title">Varios</span><span class="card-head-count">' + (t.enfermos.length + t.observaciones.length) + '</span></div>' +
        (varios || '<div class="empty">Sin novedades de turno</div>');
    }

    // stats
    var stats = document.querySelectorAll('.stats-row .stat-num');
    if (stats.length >= 4) {
      stats[0].textContent = t.practicos.length;
      stats[1].textContent = t.francos.length;
      stats[2].textContent = t.enfermos.length + t.observaciones.length;
      stats[3].textContent = t.practicos.filter(function (p) { return p.cambio; }).length;
    }

    // coordinador en el header (se quitó Presidente por pedido)
    var hdrCoord = document.querySelector('.hdr-coord');
    if (hdrCoord && t.coordinador) hdrCoord.textContent = '👤 Coord: ' + t.coordinador;

    // determinantes reales (hoja GUARDIA, celda B71 en adelante)
    renderDeterminantes(t.determinantes);

    // teléfonos reales en la pestaña Info (reemplaza los de ejemplo)
    if (t.telefonos && t.telefonos.length) {
      var infoCards = document.querySelectorAll('#sec-info .card');
      var cardTel = infoCards[1]; // segunda tarjeta = "Teléfonos de servicio"
      if (cardTel) {
        cardTel.innerHTML = '<div class="card-head"><span>📞</span><span class="card-head-title">Teléfonos de servicio</span></div>' +
          t.telefonos.map(function (tf) {
            return '<div class="tel-row"><span class="tel-puerto">' + esc(tf.puerto) + '</span><span class="tel-num">' + esc(tf.telefono) + '</span></div>';
          }).join('');
      }
    }
    var cntTurno = document.getElementById('cntTurno');
    var cntFrancos = document.getElementById('cntFrancos');
    if (cntTurno) cntTurno.textContent = t.practicos.length;
    if (cntFrancos) cntFrancos.textContent = t.francos.length;
  }

  // ── GUARDIA ROSARIO EN PESTAÑA TURNO (bajada) ────────────────────
  // Se muestra en la página principal Turno, después del listado en
  // turno y antes de Francos, además de en la pestaña Buques.
  function renderGuardiaRosarioTurno(lista) {
    var card = document.getElementById('cardGuardiaRosario');
    var cont = document.getElementById('listaGuardiaRosario');
    var cnt = document.getElementById('cntGuardiaRosario');
    if (!card || !cont) return;
    if (!lista || !lista.length) {
      card.style.display = 'none';
      return;
    }
    card.style.display = '';
    if (cnt) cnt.textContent = lista.length;
    cont.innerHTML = lista.map(function (g) {
      // g.tipo trae el estado tal como está en la planilla (p.ej. nombre de
      // buque una vez asignado, o vacío mientras no se confirma la bajada).
      var estadoBadge = g.tipo
        ? '<span class="badge b-verde">🚢 ' + esc(g.tipo) + '</span>'
        : '<span class="badge b-naranja">⏳ En espera</span>';
      return '<div class="prac-row"><div class="prac-num">•</div><div class="prac-info"><div class="prac-nombre">' + esc(g.practico) + '</div>' +
        '<div class="prac-detalle">' + esc(g.hora) + '</div></div>' + estadoBadge + '</div>';
    }).join('');
  }

  // ── DETERMINANTES (canal/km/calado reales, hoja GUARDIA B71+) ────
  function renderDeterminantes(lista) {
    var cont = document.getElementById('listaDeterminantes');
    if (!cont) return;
    if (!lista || !lista.length) { var card = cont.closest ? cont.closest('.card') : null; if (card) card.style.display = 'none'; return; }
    cont.innerHTML = (lista || []).map(function (linea) {
      return '<div class="det-row"><span class="det-label">' + esc(linea) + '</span></div>';
    }).join('') || '<div class="empty">Sin determinantes cargadas</div>';
  }

  // ── PREVISTOS (subidas / bajadas / campana) ──────────────────
  // Convierte "dd/MM/yyyy" + "HH:mm" en una clave ordenable "YYYYMMDDHHmm".
  // Antes se concatenaban los dígitos tal cual venían en la planilla
  // (DDMMYYYYHHmm), lo que ordena por día antes que por año/mes y mezcla
  // mal fechas de distintos meses o años entre sí — con Subidas, Bajadas
  // y Campana en la misma lista, esto rompía el orden cronológico real.
  function claveOrdenFecha(fecha, hora) {
    var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(fecha || '');
    var partefecha = m
      ? (m[3] + m[2].padStart(2, '0') + m[1].padStart(2, '0'))
      : '99999999'; // fechas no reconocidas van al final, no rompen el orden
    var partehora = soloDigitos(hora).padStart(4, '0').slice(0, 4) || '0000';
    return partefecha + partehora;
  }

  function renderPrevistos(lista) {
    var cont = document.getElementById('listaPrevistos');
    if (!cont || !lista) return;
    cont.innerHTML = lista.map(function (b) {
      var obs = b.observaciones || '';
      var estado = 'sin';
      if (/AUTORIZADO/i.test(obs) && !/NO AUT/i.test(obs)) estado = 'autorizado';
      else if (/NO AUT/i.test(obs)) estado = 'posible';
      // Antes había un tercer criterio (si tenía fechaConfirmado -> autorizado) que
      // generaba falsos positivos: tener una fecha de Z.C./confirmación cargada no
      // significa que la agencia haya autorizado el despacho. Solo el texto explícito
      // ("AUTORIZADO" / "NO AUT") define el estado; si no dice nada, queda "sin confirmar".
      var tipoFiltro = b.tipo === 'campana' ? 'movimiento' : b.tipo;
      var dataHora = claveOrdenFecha(b.fechaRecalada, b.horaRecalada);
      var icon = b.tipo === 'subida' ? '⬆️' : (b.tipo === 'bajada' ? '⬇️' : '🔄');
      var marcaBadge = '';
      if (b.marca === 'X') marcaBadge = '<span class="badge b-verde">✔ Despachado</span>';
      else if (b.marca === '*') marcaBadge = '<span class="badge b-naranja">🔔 Novedad agencia</span>';
      return '<div class="buque-row ' + estado + '" data-tipo="' + tipoFiltro + '" data-nombre="' + esc(b.buque) + '" data-agencia="' + esc(b.agencia) + '" data-hora="' + esc(dataHora) + '" data-estado="' + estado + '">' +
        '<div class="buque-top"><div class="buque-nombre">' + icon + ' ' + esc(b.buque) + '</div><div class="buque-cf">CF ' + esc(b.coef) + '</div></div>' +
        (marcaBadge ? '<div style="padding:0 12px 6px">' + marcaBadge + '</div>' : '') +
        '<div class="buque-bottom">' +
        '<div class="buque-dato"><span>Agencia</span><strong>' + esc(b.agencia) + '</strong></div>' +
        '<div class="buque-dato"><span>' + esc(b.labelFecha1 || 'Recalada') + '</span><strong>' + esc(b.fechaRecalada) + ' ' + esc(b.horaRecalada) + '</strong></div>' +
        (b.fechaConfirmado ? '<div class="buque-dato"><span>' + esc(b.labelFecha2 || 'Confirmado') + '</span><strong>' + esc(b.fechaConfirmado) + ' ' + esc(b.horaConfirmado) + '</strong></div>' : '') +
        '<div class="buque-dato"><span>Ruta</span><strong>' + esc(b.desde) + ' → ' + esc(b.hasta) + '</strong></div>' +
        '</div>' +
        (obs ? '<div class="buque-obs">' + esc(obs) + '</div>' : '') +
        '</div>';
    }).join('') || '<div class="empty">Sin buques previstos cargados</div>';
  }

  // ── NOVEDADES ──────────────────────────────────────────────────
  var novedadesData = [];
  function renderNovedades(lista) {
    novedadesData = lista || [];
    pintarNovedades(novedadesData);
    var searchInput = document.getElementById('novSearch');
    if (searchInput && !searchInput.dataset.wired) {
      searchInput.dataset.wired = '1';
      searchInput.addEventListener('input', function () {
        var q = this.value.toLowerCase().trim();
        var filtradas = !q ? novedadesData : novedadesData.filter(function (n) {
          return (n.texto + ' ' + n.autor).toLowerCase().includes(q);
        });
        pintarNovedades(filtradas);
      });
    }
  }
  function pintarNovedades(lista) {
    var cont = document.getElementById('listaNov');
    if (!cont) return;
    cont.innerHTML = lista.map(function (n, i) {
      return '<div class="obs-item" data-orden="' + i + '">' +
        '<span class="obs-tag tag-azul">📋 ' + esc(n.autor || 'Guardia') + '</span>' +
        '<div class="obs-texto">' + esc(n.texto) + '</div>' +
        '<div class="obs-fecha">' + esc(n.fecha) + '</div></div>';
    }).join('') || '<div class="empty">Sin novedades</div>';
  }

  // ── REMISES (sección nueva) ─────────────────────────────────────
  function renderRemises(r) {
    var cont = document.getElementById('listaRemises');
    if (!cont || !r) return;
    function filas(lista) {
      return (lista || []).map(function (p, i) {
        return '<div class="prac-row"><div class="prac-num">' + esc(p.ord || (i + 1)) + '</div>' +
          '<div class="prac-info"><div class="prac-nombre">' + esc(p.nombre) + '</div>' +
          (p.observacion ? '<div class="prac-detalle">' + esc(p.observacion) + '</div>' : '') + '</div></div>';
      }).join('') || '<div class="empty">Sin remiseros cargados</div>';
    }
    function bloque(t, icono) {
      if (!t) return '';
      return '<div class="card" style="margin-top:10px"><div class="card-head"><span>' + icono + '</span><span class="card-head-title">' + esc(t.titulo || '') + (t.fecha ? ' · ' + esc(dmy(t.fecha)) : '') + '</span></div>' + filas(t.remiseros) +
        (t.otros && t.otros.length ? '<div class="sec-label">Franco / otros</div>' + filas(t.otros) : '') +
        (t.novedades && t.novedades.length ? '<div class="sec-label">Novedades</div>' + t.novedades.map(function (n) { return '<div class="varios-row"><div class="varios-icon">📋</div><div><div class="varios-detalle">' + esc(n) + '</div></div></div>'; }).join('') : '') +
        '</div>';
    }
    cont.innerHTML = bloque(r.turno1, '🚕') + bloque(r.turno2, '🌙');
  }

  // ── BUQUES (en tránsito ahora, hoja GUARDIA) ────────────────────
  function tarjetaBuque(b, tipo, meta) {
    // La planilla no trae fecha propia por cada horario de tránsito (sólo
    // "I: 22:30", "S.N: 14:30", etc.) — se agrega la fecha del turno como
    // referencia, ya que es el único dato de fecha real disponible.
    var metaFull = meta + (b.fecha ? ' · ' + esc(b.fecha) : '') + (b.canal ? ' · Canal ' + esc(b.canal) : '');
    return '<div class="desp-card" data-tipo="' + tipo + '" data-nombre="' + esc(b.buque) + '">' +
      '<div class="desp-head ' + tipo + '"><div><div class="desp-buque">' + esc(b.buque) + '</div>' +
      '<div class="desp-meta">' + metaFull + '</div></div>' +
      (b.calado ? '<span class="badge b-azul">' + esc(b.calado) + '</span>' : '') + '</div>' +
      '<div class="desp-body">' +
      (b.practicos && b.practicos.length ? '<div class="desp-practicos">' + b.practicos.map(function (p, i) {
        return '<div class="prac-chip"><div class="chip-pos">' + (i + 1) + '</div><div class="chip-nombre">' + esc(p) + '</div><div class="chip-icon">⚓</div></div>';
      }).join('') + '</div>' : '') +
      // Cada dato ya trae su propia etiqueta tal cual está en la planilla
      // (ej. "S.N:23:30", "ESC:", "I:07:20") — no se le pone un rótulo inventado.
      (b.datos && b.datos.length ? '<div class="desp-datos">' + b.datos.map(function (d) {
        return '<div class="dato-row"><span class="dato-val">' + esc(d) + '</span></div>';
      }).join('') + '</div>' : '') +
      (b.observaciones ? '<div class="buque-obs">' + esc(b.observaciones) + '</div>' : '') +
      '</div></div>';
  }

  function renderBuques(t) {
    var cont = document.getElementById('listaBuquesDesp');
    if (!cont || !t) return;
    var html = '';

    (t.aRosario || []).forEach(function (b) {
      html += tarjetaBuque(b, 'subida', 'En viaje a Rosario');
    });

    (t.deRosario || []).forEach(function (b) {
      html += tarjetaBuque(b, 'bajada', 'En viaje de Rosario');
    });

    (t.aCampana || []).forEach(function (b) { html += tarjetaBuque(b, 'movimiento', 'En viaje a Campana'); });
    (t.deCampana || []).forEach(function (b) { html += tarjetaBuque(b, 'movimiento', 'En viaje de Campana'); });
    (t.movimientos || []).forEach(function (b) { html += tarjetaBuque(b, 'movimiento', 'Movimiento'); });

    if (t.guardiaRosario && t.guardiaRosario.length) {
      html += '<div class="card"><div class="card-head"><span>👥</span><span class="card-head-title">Guardia Rosario (bajada)</span></div>' +
        t.guardiaRosario.map(function (g) {
          // g.tipo trae el estado tal como está en la planilla (p.ej. nombre de
          // buque una vez asignado, o vacío/"en espera" mientras no se confirma).
          var estadoBadge = g.tipo
            ? '<span class="badge b-verde">🚢 ' + esc(g.tipo) + '</span>'
            : '<span class="badge b-naranja">⏳ En espera</span>';
          return '<div class="prac-row"><div class="prac-num">•</div><div class="prac-info"><div class="prac-nombre">' + esc(g.practico) + '</div>' +
            '<div class="prac-detalle">' + esc(g.hora) + '</div></div>' + estadoBadge + '</div>';
        }).join('') + '</div>';
    }

    document.querySelectorAll('#listaBuquesDesp .desp-card, #listaBuquesDesp .card').forEach(function (el) { el.remove(); });
    var vacio = document.getElementById('emptyBuques');
    if (vacio) vacio.insertAdjacentHTML('beforebegin', html);
  }

  document.addEventListener('mp:auth-ok', cargarDatos);
})();
