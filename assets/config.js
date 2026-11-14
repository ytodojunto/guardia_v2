// URL de la "Implementación" del Apps Script (termina en /exec). Vacío = la página funciona solo con "Elegir archivo".
// Implementaciones del Apps Script (terminan en /exec), en orden de preferencia: si la primera falla se usa la siguiente.
window.GUARDIA_API_URLS = [
  "https://script.google.com/macros/s/AKfycbyjP-_agxvmx432_H5LoFNvlOfj5j5UoN6oO6VOjHE7djGOfGlYwgju6gK1hj2cN-tL/exec",
  "https://script.google.com/macros/s/AKfycbxIHyDrV-DTbZzqk6-FBrzPHkBJmZDNZ6dobA2X9_rNYEs-ZNTIFIGMHlCf0-s4A0Qz/exec"
];
window.GUARDIA_API_URL = window.GUARDIA_API_URLS[0];
// Pide la guardia probando cada implementación; devuelve el primer JSON que trae el Excel (o el último error).
window.guardiaPedir = async function (clave) {
  var ultimo = { error: 'No hay URL configurada' };
  for (var i = 0; i < window.GUARDIA_API_URLS.length; i++) {
    var u = window.GUARDIA_API_URLS[i];
    try {
      var r = await fetch(u + (u.indexOf('?') === -1 ? '?' : '&') + 'clave=' + encodeURIComponent(clave), { cache: 'no-store' });
      var d = await r.json();
      if (d && d.xlsx) return d;
      ultimo = d || ultimo;
      if (d && /autorizado/i.test(d.error || '')) return d;
    } catch (e) { ultimo = { error: 'No pude conectar con el script (' + (e && e.message ? e.message : e) + ')' }; }
  }
  return ultimo;
};

// Versión visible en la página (se muestra en el login, el encabezado y la Proyección). Subirla en cada cambio que se publique.
window.GUARDIA_VERSION = "guardia_v2.10";
