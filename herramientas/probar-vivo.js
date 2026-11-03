// Simula la lectura en vivo: Apps Script falso que devuelve el xlsx en base64.  Uso: node probar-vivo.js archivo.xlsx
const {JSDOM}=require('jsdom'),fs=require('fs');
const buf=fs.readFileSync(process.argv[2]);
const html=fs.readFileSync('../index.html','utf8').replace(/<script src="[^"]+"><\/script>/g,'');
const dom=new JSDOM(html,{runScripts:'outside-only',pretendToBeVisual:true,url:'https://x.github.io/'});
const w=dom.window; w.ExcelJS=require('exceljs'); w.JSZip=require('jszip'); w.GUARDIA_API_URL='https://script.example/exec';
let llamadas=[];
w.fetch=async(u)=>{llamadas.push(u);const ok=/clave=bien/.test(u);return{json:async()=>ok?{nombre:'GUARDIA.xlsx',modificado:new Date().toISOString(),consultado:new Date().toISOString(),xlsx:buf.toString('base64')}:{error:'No autorizado'}};};
const scripts=[...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m=>m[1]);
scripts.forEach(c=>w.eval(c));
const d=w.document, $=id=>d.getElementById(id);
setTimeout(()=>{
  console.log('1) sin clave -> pide clave:',!$('pedirClave').hidden);
  $('clave').value='mal'; $('entrar').click();
  setTimeout(()=>{
    console.log('2) clave mala -> mensaje:',$('estado').textContent);
    $('clave').value='bien'; $('entrar').click();
    setTimeout(()=>{
      console.log('3) clave buena -> secciones:',[...$('salida').querySelectorAll('h2')].length,'| estado:',$('estado').textContent.slice(0,40));
      w.document.defaultView.eval("1"); $('actualizar').click();
      setTimeout(()=>{console.log('4) actualizar -> llamadas:',llamadas.length);process.exit(0)},800);
    },1500);
  },500);
},300);
