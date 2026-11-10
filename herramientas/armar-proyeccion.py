"""Arma proyeccion.html: motor de cálculo de proyeccion-completa.html + diseño de proyeccion-movil.html.
Uso (desde la raíz del repo):  python3 herramientas/armar-proyeccion.py
"""
import re, io
R = '.'
comp = open(R + '/proyeccion-completa.html', encoding='utf8').read()
movil = open(R + '/proyeccion-movil.html', encoding='utf8').read()

# ── 1. motor (scripts) y su DOM oculto, de la versión completa ──
i_wrap = comp.index('<div class="wrap">')
i_scr = comp.index('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx')
motor_dom = comp[i_wrap:i_scr]
motor_dom = re.sub(r'<a href="index.html"[^>]*>.*?</a>', '', motor_dom, count=1)
scripts = comp[i_scr:comp.rindex('</body>')]

# 1a. render() publica su resultado
a = 'var sim = simulate(); window.__sim = sim; window.__T0 = T0;'
assert scripts.count(a) == 1
scripts = scripts.replace(a, a + '\n    window.__ctx = { REF: REF, W0: W0, W1: W1, T0: T0, nFuera: NFUERA, avisos: AVISOS.slice(), modo: MODE, estado: document.getElementById("srcst").textContent };\n    setTimeout(function () { if (window.__publicar) window.__publicar(); }, 0);')

# 1b. "¿Voy a estar?" como función pura (sin DOM): window.__voyFn(nombre, T) -> [lineas con hora estimada, lineas con demora]
b = 'SHIFTMODE = true; SHIFTED = 0; var sim2 = simulate(); SHIFTMODE = false; if (!SHIFTED) sim2 = null;'
assert scripts.count(b) == 1
ini = scripts.index('function describe(S, title) {')
fin = scripts.index('out = outP; outP.appendChild(box);', ini)
cuerpo = scripts[ini:fin]
q0 = cuerpo.index('var q = S.ps.filter')
cuerpo = cuerpo[q0:]
cuerpo = cuerpo.replace('var q = S.ps.filter(function (x) { return x.n === sel.value; })[0];\n        if (!q) return;', 'var q = S.ps.filter(function (x) { return x.n === nm; })[0];\n        if (!q) return lines;')
cuerpo = re.sub(r'function line\(txt, strong\) \{.*?\n', 'function line(txt) { lines.push(txt); }\n', cuerpo, count=1)
voyfn = ('\n      function voyLines(S, nm, T) {\n        var lines = [];\n        ' + cuerpo.rstrip() +
         '\n        return lines;\n      }\n      window.__voyFn = function (nm, T) { var A = voyLines(sim, nm, T); return [A, sim2 ? voyLines(sim2, nm, T) : A]; };\n')
scripts = scripts.replace(b, b + voyfn, 1)

# 1c. arranque: ya no carga datos de ejemplo por su cuenta si hay Drive (ya lo resuelve el bloque loadDrive)
# ── 2. diseño del móvil ──
css = movil[movil.index('<style>') + 7:movil.index('</style>')]
css += """
.pill.uno{background:var(--warn-bg);border-color:var(--warn)}.pill.uno i{background:var(--warn)}.pill.uno b{color:var(--warn)}
.pill.falta{background:var(--bad-bg);border-color:var(--bad)}.pill.falta i{background:var(--bad)}.pill.falta b{color:var(--bad)}
.chip i.uno{background:var(--warn)}.chip i.falta{background:var(--bad)}.chip i.none{background:var(--line)}
.banner{margin:0 0 14px;padding:10px 14px;border-radius:10px;background:var(--warn-bg);color:var(--warn);font-size:15px}
.banner.ok{background:var(--ok-bg);color:var(--ok)}
.avis{margin:6px 0 0;padding-left:18px;font-size:14px;color:var(--mute)}
#motor{display:none}
"""
header = movil[movil.index('<header>'):movil.index('<main>')]
header = re.sub(r'<p>Datos de la guardia[^<]*</p>', '<p id="hdrinfo">Leyendo la guardia…</p>', header)
header = header.replace('<h1>Guardia</h1>', '<h1><a href="index.html" style="color:#fff;text-decoration:none">←</a> Guardia · Proyección</h1>')
main_ini = movil.index('<main>')
hoy_ini = movil.index('<section id="hoy">')
hoy_fin = movil.index('</section>', hoy_ini) + len('</section>')
hoy = '''<section id="hoy">
  <div id="banner"></div>
  <div class="pill" id="pill"><i></i><div><b>Leyendo…</b><br></div></div>
  <div class="strip" id="strip"></div>
  <div class="card nums" id="nums"></div>
  <h2>Qué viene</h2>
  <div class="tl" id="tl"></div>
  <h2>Entran a la lista pronto</h2>
  <div class="card" id="entranHoy" style="padding:6px 14px"></div>
  <p class="foot" style="color:var(--mute);font-size:14px;margin-top:16px"><a href="proyeccion-completa.html" style="color:var(--navy)">Versión completa</a> (parámetros, barcos sin hora, estadísticas de terminales)</p>
</section>'''
resto = movil[hoy_fin:movil.index('<script>')]
# resumen de Análisis dinámico
r_ini = resto.index('<div class="card src" style="margin-bottom:0">')
r_fin = resto.index('</div>', resto.index('<p class="k">Cada práctico vuelve', r_ini)) + 6
resto = resto[:r_ini] + '<div class="card src" style="margin-bottom:0" id="resumen"></div>\n  <ul class="avis" id="avisosL"></ul>' + resto[r_fin:]
# tarjeta nueva en Datos
resto = resto.replace('<h2>De dónde salen los datos</h2>', '''<h2>De dónde salen los datos</h2>
  <div class="card src"><h3>Ahora: la planilla de la guardia en Drive</h3>
    <p>Esta pantalla lee la planilla de la guardia (Google Sheets) cada vez que se abre y se actualiza sola cada 5 minutos. De ahí salen el turno, los viajes en curso, los francos y la programación de buques. Si no logra leerla muestra datos de ejemplo y lo avisa arriba.</p>
    <p class="k">Las tarjetas 1 y 2 de abajo describen la etapa anterior, cuando se cargaba el PDF del sistema y el Excel a mano.</p></div>''', 1)

# ── 3. código de la vista: del script de movil, con CALC/VOY como parámetros ──
sc = movil[movil.index('<script>') + 8:movil.rindex('</script>')]
p1_ini = sc.index('(function(){\n  var DIAS')
p1_fin = sc.index('})();', p1_ini)
vista1 = sc[p1_ini + len('(function(){\n'):p1_fin]
vista1 = vista1.replace("var REF=CALC.REF", "var REF=CALC.REF", 1)
g1="card('Al inicio',fmt(REF),r0.snap,REF,null);"
assert g1 in vista1
vista1=vista1.replace(g1,"if(r0)card('Al inicio',fmt(REF),r0.snap,REF,null);else rot.appendChild(el('div','fila','No hay barcos con hora para calcular la lista.'));")
limpiar = "['resolver','barcos','rotacion','entran','semaforo','gantt','leyenda'].forEach(function(i){var e=document.getElementById(i);e.textContent=''});\n"
hoy_js = r'''
  // ── Hoy ──
  var DIASL=['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  function pilotosTxt(r){return r.taken&&r.taken.length?(r.taken.length===1?r.taken[0]+' (sale solo)':r.taken.join(' y ')):'Sin práctico asignado'}
  function worst(rs){var w='ok';rs.forEach(function(r){var s=stOf(r)[0];if(s==='falta')w='falta';else if(s==='uno'&&w!=='falta')w='uno'});return w}
  var HD=document.getElementById('hdrinfo');
  var win=res.filter(function(r){return r.s.t<24*(Math.floor(REF/24)+4)});
  var prox=res.filter(function(r){return r.s.call<REF+24});
  var pw=prox.length?worst(prox):'ok',pl=document.getElementById('pill');
  pl.className='pill'+(pw==='ok'?'':' '+pw);
  pl.querySelector('div').innerHTML='';
  var pb=el('b','',pw==='ok'?'Cubierto':(pw==='uno'?'Justo':'Falta gente'));pl.querySelector('div').appendChild(pb);pl.querySelector('div').appendChild(document.createElement('br'));
  pl.querySelector('div').appendChild(document.createTextNode(!prox.length?'No hay barcos con hora en las próximas 24 h.':(pw==='ok'?'Todos los barcos de las próximas 24 h tienen sus prácticos.':(pw==='uno'?'Algún barco de las próximas 24 h sale justo de prácticos.':'Hay barcos de las próximas 24 h a los que les faltan prácticos.'))));
  var d0h=Math.floor(REF/24),st=document.getElementById('strip');st.textContent='';
  function despachos(rs){var n=0;rs.forEach(function(r){n+=(MOVS[r.s.route]?1:(r.s.np||(r.taken?r.taken.length:2)))});return n}
  for(var dd=d0h;dd<d0h+4;dd++){
    var rd=res.filter(function(r){return Math.floor(r.s.t/24)===dd});
    var ch=el('div','chip');var ic=el('i',rd.length?(worst(rd)==='ok'?'':worst(rd)):'none');ch.appendChild(ic);
    ch.appendChild(document.createTextNode((dd===d0h?'Hoy':dayLbl(24*dd).split(' ')[0])+' '+dOf(24*dd).getDate()));
    ch.appendChild(el('small','',rd.length?rd.length+(rd.length===1?' servicio':' servicios')+' · '+despachos(rd)+' despachos':'sin servicios'));
    st.appendChild(ch)}
  function onTrip2(q){return q.base.concat(q.segs).some(function(s){return (s[2]==='nav'||s[2]==='asig')&&s[0]<=REF+1e-9&&REF<s[1]-1e-9})}
  var nNav=0,nLis=0;CALC.ps.forEach(function(q){if(onTrip2(q))nNav++;else if(q.hist.some(function(h){return h[0]<=REF+1e-9&&REF<h[1]-1e-9})||q.base.concat(q.segs).some(function(s){return s[2]==='pres'&&s[0]<=REF+1e-9&&REF<s[1]-1e-9}))nLis++});
  var FU=(window.__meta&&window.__meta.fuera)||[];
  var nFr=Math.max(0,(CALC.nFuera||0)-FU.length);
  var nm=document.getElementById('nums');nm.textContent='';
  [[nLis,'en lista'],[nNav,'navegando'],[nFr,'de franco']].concat(FU.length?[[FU.length,'fuera de turno']]:[]).forEach(function(x){var d=el('div');d.appendChild(el('b','',String(x[0])));d.appendChild(el('span','',x[1]));nm.appendChild(d)});
  var tl=document.getElementById('tl');tl.textContent='';
  for(var dj=d0h;dj<d0h+4;dj++){
    var d1=dOf(24*dj),lbl=(dj===d0h?'Hoy · ':(dj===d0h+1?'Mañana · ':''))+DIASL[d1.getDay()]+' '+d1.getDate()+'/'+(d1.getMonth()+1);
    if(dj>d0h+1)lbl=lbl.charAt(0).toUpperCase()+lbl.slice(1);
    tl.appendChild(el('div','day',lbl));
    var rj=res.filter(function(r){return Math.floor(r.s.t/24)===dj}).sort(function(a,b){return a.s.t-b.s.t});
    if(!rj.length)tl.appendChild(el('p','empty','Sin barcos con hora.'));
    rj.forEach(function(r){var ev=el('div','ev'),tp=el('div','top');tp.appendChild(el('span','h',fmtH(r.s.t)));tp.appendChild(el('span','n',r.s.n));tp.appendChild(el('span','t',tipo(r)));ev.appendChild(tp);
      ev.appendChild(el('div','p',pilotosTxt(r)));ev.appendChild(el('div','l','Llamado '+fmtH(r.s.call)));
      if(r.s.vencida)ev.appendChild(el('span','tag','Hora vencida: se supone que sale de un momento a otro'));
      else if(r.s.t>REF+24)ev.appendChild(el('span','tag','Hora a confirmar: puede correrse'));
      if(r.status==='falta')ev.appendChild(el('span','tag','Faltan prácticos'));
      else if(just(r))ev.appendChild(el('span','tag','Justo de prácticos'));
      tl.appendChild(ev)})}
  var eh=document.getElementById('entranHoy');eh.textContent='';
  if(!en.length)eh.appendChild(el('div','fila','Nadie entra en las próximas 12 h'));
  en.forEach(function(x){var f=el('div','fila',x.n);f.appendChild(el('span','',fmtH(x.key)+' · lugar '+x.pos));eh.appendChild(f)});
  // resumen de Análisis
  var su=win.filter(function(r){return !MOVS[r.s.route]&&r.s.dir==='S'}).length,ba=win.filter(function(r){return !MOVS[r.s.route]&&r.s.dir!=='S'}).length;
  var okN=win.filter(function(r){return stOf(r)[0]==='ok'}).length,juN=win.filter(function(r){return stOf(r)[0]==='uno'}).length,faN=win.filter(function(r){return stOf(r)[0]==='falta'}).length;
  var dl=dOf(24*(d0h+3));
  var rsm=document.getElementById('resumen');rsm.textContent='';
  var p1=el('p');p1.appendChild(el('b','','Prácticos: '));p1.appendChild(document.createTextNode(nLis+' en lista, '+nNav+' navegando, '+nFr+' de franco'+(FU.length?', '+FU.length+' fuera de turno (enfermo o licencia): '+FU.map(function(f){return f.practico}).join(', '):'')+'.'));rsm.appendChild(p1);
  var p2_=el('p');p2_.appendChild(el('b','','Hasta el '+DIASL[dl.getDay()]+' '+dl.getDate()+'/'+(dl.getMonth()+1)+': '));p2_.appendChild(document.createTextNode(win.length+' servicios ('+su+' suben, '+ba+(ba===1?' baja':' bajan')+(win.length-su-ba?', '+(win.length-su-ba)+' movimientos':'')+') = '+despachos(win)+' despachos. Cubiertos '+okN+', justos '+juN+', sin cobertura '+faN+'.'));rsm.appendChild(p2_);
  rsm.appendChild(el('p','k','Cada práctico vuelve a estar libre a su hora de presentación más 6 h de descanso. Cada barco se llama 5 h antes de su salida, por orden de lista.'));
'''
vista1 = limpiar + vista1 + hoy_js
voy_js = r'''
(function(){
  var tabs=document.querySelectorAll('nav button');
  tabs.forEach(function(b){b.addEventListener('click',function(){
    tabs.forEach(function(x){x.setAttribute('aria-selected',x===b?'true':'false')});
    ['hoy','voy','analisis','datos'].forEach(function(id){document.getElementById(id).hidden=(id!==b.dataset.t)});
    window.scrollTo(0,0);
  })});
  var pr=document.getElementById('pr'),fh=document.getElementById('fh'),ans=document.getElementById('ans'),quick=document.getElementById('quick');
  function p(n){return String(n).padStart(2,'0')}
  function iso(d){return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate())+'T'+p(d.getHours())+':'+p(d.getMinutes())}
  function head(t){
    if(/fuera de turno/i.test(t))return['Fuera de turno','neu'];
    if(/franco/i.test(t))return['De franco','neu'];
    if(/embarcado/i.test(t))return['Embarcado','warn'];
    if(/llegando/i.test(t))return['Llegando del viaje','warn'];
    if(/descansando/i.test(t))return['Descansando','warn'];
    var m=/lugar (\d+) de (\d+)/.exec(t);return['En la lista'+(m?', lugar '+m[1]+' de '+m[2]:''),'ok'];
  }
  function clean(s){return s.replace(/ ?Si ese día necesitás estar, pedí un cambio\./,'')}
  var C=null;
  function calc(){
    ans.textContent='';
    if(!C||!window.__voyFn||!pr.value){ans.textContent='Esperando los datos de la guardia…';return}
    var d=new Date(fh.value);if(isNaN(d)){ans.textContent='Elegí una fecha y hora.';return}
    var FU=(window.__meta&&window.__meta.fuera)||[],fx=FU.filter(function(f){return String(f.practico).toLowerCase()===pr.value.toLowerCase()})[0];
    var T=(d.getTime()-C.T0)/3600000,L=window.__voyFn(pr.value,T),A=L[0],B=L[1];
    if(fx&&!A.length){
      var m=/^\d{4}-(\d\d)-(\d\d)T(\d\d:\d\d)/.exec(fx.salida||'');
      var hd0=document.createElement('span');hd0.className='hd neu';hd0.textContent='Fuera de turno';ans.appendChild(hd0);
      var q0=document.createElement('p');q0.textContent='Figura fuera de turno (enfermo o licencia)'+(m?' desde el '+(+m[2])+'/'+(+m[1])+' '+m[3]:'')+'. No entra a la lista hasta que se informe su reingreso.';ans.appendChild(q0);
      if(fx.observacion){var q1=document.createElement('p');q1.style.color='var(--mute)';q1.textContent=fx.observacion;ans.appendChild(q1)}
      return;
    }
    if(!A.length){ans.textContent='Sin datos para ese práctico.';return}
    var h=head(A[0]);
    var hd=document.createElement('span');hd.className='hd '+h[1];hd.textContent=h[0];ans.appendChild(hd);
    var need=false;
    A.forEach(function(l){var q=document.createElement('p');q.textContent=clean(l);ans.appendChild(q);if(/embarcado|llamar/i.test(l))need=true});
    if(JSON.stringify(A)!==JSON.stringify(B)){
      var alt=document.createElement('div');alt.className='alt';var b=document.createElement('b');b.textContent='Si se demoran los barcos con hora sin confirmar';alt.appendChild(b);
      B.forEach(function(l){var q=document.createElement('p');q.style.margin='4px 0 0';q.textContent=clean(l);alt.appendChild(q)});ans.appendChild(alt);
    }
    if(need){var f=document.createElement('div');f.className='foot';f.textContent='Si ese día necesitás estar, pedí un cambio.';ans.appendChild(f)}
  }
  window.__armarVoy=function(CALC){
    C=CALC;var sel=pr.value;pr.textContent='';
    var FU=(window.__meta&&window.__meta.fuera)||[];
    function tc(n){n=String(n).trim();return n.charAt(0).toUpperCase()+n.slice(1).toLowerCase()}
    var nombres=CALC.ps.map(function(q){return q.n});
    FU.forEach(function(f){var n=tc(f.practico);if(!nombres.some(function(x){return x.toLowerCase()===n.toLowerCase()}))nombres.push(n)});
    nombres.sort().forEach(function(n){var o=document.createElement('option');o.value=n;o.textContent=n;pr.appendChild(o)});
    if(sel)pr.value=sel;
    var ref=new Date(CALC.T0+Math.round(CALC.REF*60)*60000);
    fh.min=iso(ref);var mx=new Date(ref.getTime()+4*864e5);fh.max=iso(mx);
    if(!fh.value||fh.value<fh.min)fh.value=iso(new Date(new Date(ref.getFullYear(),ref.getMonth(),ref.getDate()+1,8,0)));
    quick.textContent='';
    var DN=['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
    [[1,8,'Mañana 08:00'],[1,20,'Mañana 20:00'],[2,8,DN[(ref.getDay()+2)%7]+' 08:00'],[3,8,DN[(ref.getDay()+3)%7]+' 08:00']].forEach(function(q){
      var b=document.createElement('button');b.type='button';b.textContent=q[2];
      b.addEventListener('click',function(){fh.value=iso(new Date(ref.getFullYear(),ref.getMonth(),ref.getDate()+q[0],q[1],0));calc()});quick.appendChild(b)});
    calc();
  };
  pr.addEventListener('change',calc);fh.addEventListener('change',calc);

  var vb=document.querySelectorAll('#vista button');
  vb.forEach(function(b){b.addEventListener('click',function(){
    vb.forEach(function(x){x.setAttribute('aria-pressed',x===b?'true':'false')});
    var v=b.dataset.v;
    document.querySelectorAll('#tcorr .cs').forEach(function(c){c.style.display=(v==='b')?'none':''});
    document.querySelectorAll('#tcorr .cb').forEach(function(c){c.style.display=(v==='s')?'none':''});
  })});
  var te=document.getElementById('tej'),dir={};
  te.querySelectorAll('th').forEach(function(th){th.addEventListener('click',function(){
    var k=+th.dataset.k,num=th.dataset.num,rows=[].slice.call(te.rows,1);dir[k]=!dir[k];
    rows.sort(function(a,b){var x=a.cells[k].textContent,y=b.cells[k].textContent;var r=num?(parseFloat(x)-parseFloat(y)):x.localeCompare(y,'es',{numeric:true});return dir[k]?r:-r});
    rows.forEach(function(r){te.tBodies[0].appendChild(r)});
  })});
})();

window.__publicar=function(){
  var X=window.__ctx,S=window.__sim;if(!X||!S)return;
  var CALC={W:[X.W0,X.W1],REF:X.REF,T0:X.T0,res:S.res,ps:S.ps,nFuera:X.nFuera};
  try{pintar(CALC)}catch(e){console.error(e);var b=document.getElementById('banner');b.className='banner';b.textContent='No pude dibujar la proyección: '+e.message}
  var b=document.getElementById('banner');
  if(X.modo==='demo'){b.className='banner';b.textContent='Son datos de ejemplo, no la guardia de hoy. Para ver la real entrá desde la pantalla principal con usuario y contraseña.'}
  else{b.className='';b.textContent=''}
  var hi=document.getElementById('hdrinfo');
  var M=window.__meta;
  function dm(iso){var m=/^\d{4}-(\d\d)-(\d\d)/.exec(iso||'');return m?m[2]+'/'+m[1]:''}
  function hm2(d){return d?p2(d.getHours())+':'+p2(d.getMinutes()):''}
  function p2(n){return String(n).padStart(2,'0')}
  if(X.modo!=='demo'&&M&&M.guardia){
    var G=M.guardia,mo=M.modificado?new Date(M.modificado):null,co=new Date();
    hi.textContent='Guardia del '+dm(G.fechaInicio)+' '+(G.horaInicio||'')+' al '+dm(G.fechaFin)+' '+(G.horaFin||'')+(G.operador?' · operador '+G.operador:'')+(mo?' · planilla guardada '+mo.getDate()+'/'+(mo.getMonth()+1)+' '+hm2(mo):'')+' · leída '+hm2(co);
  } else hi.textContent=(X.modo==='demo'?'Datos de ejemplo · ':'')+(X.estado||'');
  if(window.GUARDIA_VERSION)hi.textContent+=' · '+window.GUARDIA_VERSION;
  var av=document.getElementById('avisosL');av.textContent='';
  (X.avisos||[]).forEach(function(t){var li=document.createElement('li');li.textContent=t;av.appendChild(li)});
  if(window.__armarVoy)window.__armarVoy(CALC);
};
'''
vista_fn = 'function pintar(CALC){\n' + vista1 + '\n}\n' + voy_js

html = ('<!doctype html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        '<title>Guardia · Proyección y pronóstico</title>\n<style>' + css + '</style>\n</head>\n<body>\n' + header + '<main>\n' + hoy + resto +
        '</main>\n<div id="motor" hidden>\n' + motor_dom + '\n</div>\n' + scripts +
        '\n<script>\n' + vista_fn + '</script>\n</body>\n</html>\n')
# el motor termina con su arranque; la vista se define después, pero __publicar se invoca recién tras render() (setTimeout)
open(R + '/proyeccion.html', 'w', encoding='utf8').write(html)
print('proyeccion.html', len(html))
