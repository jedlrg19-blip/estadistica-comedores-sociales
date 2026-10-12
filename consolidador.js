/* ============================================================================
   Consolidador DPS — genera el libro «dps.xlsx» correcto a partir de los
   archivos crudos que envía cada programa (ZIP «Ejecuciones»).

   Funciona 100 % en el navegador (SheetJS + JSZip) y también en Node (pruebas).
   Principios:
   · Cada archivo se identifica por su CONTENIDO, no por su nombre.
   · Se parte del dps anterior (el publicado) y solo se reemplaza lo que trae
     un crudo; lo que no viene en el ZIP se conserva y se avisa.
   · Todo número derivado se calcula del crudo; las constantes humanas
     (presupuestos, ejecución contable) se piden/conservan y se marcan.
   · Los meses ya cerrados se comparan contra el dps anterior y cualquier
     cambio se reporta.
   ========================================================================== */
(function(root){
'use strict';
const XLSX = root.XLSX || (typeof require!=='undefined'? require('xlsx') : null);
const JSZipLib = root.JSZip || (typeof require!=='undefined'? require('jszip') : null);
const XU = XLSX.utils;
const MESES=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const MESES_U=MESES.map(m=>m.toUpperCase());
const DEPS=['ALTA VERAPAZ','BAJA VERAPAZ','CHIMALTENANGO','CHIQUIMULA','EL PROGRESO','ESCUINTLA','GUATEMALA','HUEHUETENANGO','IZABAL','JALAPA','JUTIAPA','PETÉN','QUETZALTENANGO','QUICHÉ','RETALHULEU','SACATEPÉQUEZ','SAN MARCOS','SANTA ROSA','SOLOLÁ','SUCHITEPÉQUEZ','TOTONICAPÁN','ZACAPA'];

/* ---------- utilidades ---------- */
const norm = s => String(s==null?'':s).normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\s+/g,' ').trim().toUpperCase();
const clean = s => String(s==null?'':s).replace(/\s+/g,' ').trim();
const num = v => (typeof v==='number' && isFinite(v)) ? v : 0;
const r2 = v => Math.round((v+Number.EPSILON)*100)/100;
const colL = c => XU.encode_col(c);
const A = (r,c) => XU.encode_cell({r,c});
const CI = l => XU.decode_col(l);
const DEPCAN = (()=>{ const m={}; DEPS.forEach(d=>m[norm(d)]=d); m['PETEN']='PETÉN'; return m; })();
const canonDep = s => DEPCAN[norm(s)] || clean(s).toUpperCase();

function cget(ws,r,c){ return ws[A(r,c)]; }
function vget(ws,r,c){ const x=ws[A(r,c)]; return x?x.v:undefined; }
function cset(ws,r,c,v,o){
  const a=A(r,c); o=o||{};
  if(v===null||v===undefined||v===''&&!o.keepEmpty){ delete ws[a]; return; }
  const x={v, t: typeof v==='number'?'n':typeof v==='boolean'?'b':'s'};
  if(o.f) x.f=o.f; if(o.z) x.z=o.z; ws[a]=x;
}
function copyCell(x){ if(!x) return x; const y={t:x.t,v:x.v}; if(x.f) y.f=x.f; if(x.z) y.z=x.z; if(x.w) y.w=x.w; return y; }
function fixRef(ws){
  let R=0,C=0,r0=1e9,c0=1e9;
  for(const k of Object.keys(ws)){ if(k[0]==='!') continue; const p=XU.decode_cell(k); if(p.r>R)R=p.r; if(p.c>C)C=p.c; if(p.r<r0)r0=p.r; if(p.c<c0)c0=p.c; }
  (ws['!merges']||[]).forEach(m=>{ if(m.e.r>R)R=m.e.r; if(m.e.c>C)C=m.e.c; });
  ws['!ref']=XU.encode_range({s:{r:0,c:0},e:{r:R,c:C}});
}
function aoa(ws){ const e=XU.decode_range(ws['!ref']||'A1').e; return XU.sheet_to_json(ws,{header:1,defval:null,raw:true,blankrows:true,range:{s:{r:0,c:0},e}}); }
function findSheet(wb,pred){ for(const n of wb.SheetNames){ if(pred(n,wb.Sheets[n])) return n; } return null; }
function topText(ws,rows,cols){ let t=''; for(let r=0;r<(rows||6);r++) for(let c=0;c<(cols||8);c++){ const v=vget(ws,r,c); if(typeof v==='string') t+=' '+v; } return norm(t); }
function hasHeader(R,max,...must){ for(let i=0;i<Math.min(max,R.length);i++){ const nn=(R[i]||[]).map(norm); if(must.every(m=>nn.some(x=>x.length<=40&&x.includes(m)))) return i; } return -1; }

/* ---------- identificación por contenido ---------- */
function identificarLibro(wb){
  const out=[]; const vistos=new Set();
  for(const sn of wb.SheetNames){
    const ws=wb.Sheets[sn]; if(!ws||!ws['!ref']) continue;
    const R=aoa(ws); const t=topText(ws,8,8);
    if(norm(sn).includes('APORTES TMCA') || (t.includes('BOLSA SOCIAL')&&t.includes('TRANSFERENCIA')&&hasHeader(R,12,'DEPARTAMENTO','MUNICIPIO','PLANILLA')>=0)){ out.push({tipo:'tmca',hoja:sn}); continue; }
    if(t.includes('BECA EDUCACION SUPERIOR') && hasHeader(R,6,'DEPARTAMENTO','MUNICIPIO')>=0){ out.push({tipo:'bes',hoja:sn}); continue; }
    if(t.includes('BECA EDUCACION MEDIA') && hasHeader(R,6,'DEPARTAMENTO','MUNICIPIO')>=0){ out.push({tipo:'bm',hoja:sn}); continue; }
    if(hasHeader(R,15,'DEPARTAMENTO','MUNICIPIO','BENEFICIARIOS','PRESUPUESTO EJECUTADO')>=0){ out.push({tipo:'jp',hoja:sn}); continue; }
    if(hasHeader(R,10,'DEPARTAMENTO','MUNICIPIO','UNIDAD')>=0 && (t.includes('COMEDOR')||norm(sn).includes('COMEDOR'))){ out.push({tipo:'comedores',hoja:sn}); continue; }
  }
  return out.filter(x=>{ if(vistos.has(x.tipo)) return false; vistos.add(x.tipo); return true; });
}
async function leerZip(buf){
  const zip=await JSZipLib.loadAsync(buf);
  const files=[], ignorados=[];
  const names=Object.keys(zip.files).filter(n=>!zip.files[n].dir);
  for(const n of names){
    const base=n.split('/').pop();
    if(/^~\$|^\./.test(base) || /__MACOSX/.test(n)) continue;
    if(!/\.xlsx?$/i.test(base)){ ignorados.push(base); continue; }
    try{
      const data=await zip.files[n].async('arraybuffer');
      const wb=XLSX.read(data,{type:'array',cellFormula:true,cellNF:true});
      files.push({nombre:base.normalize('NFC'),ruta:n,wb,bytes:data.byteLength,ids:identificarLibro(wb)});
    }catch(e){ ignorados.push(base+' (no se pudo leer)'); }
  }
  return {files,ignorados};
}

/* ---------- COMEDORES ---------- */
function parseComedores(ws){
  const R=aoa(ws); const h=hasHeader(R,12,'DEPARTAMENTO','MUNICIPIO','UNIDAD'); if(h<0) return null;
  const hd=R[h].map(norm), cDep=hd.indexOf('DEPARTAMENTO'), cMun=hd.findIndex(x=>x.includes('MUNICIPIO')), cUM=hd.findIndex(x=>x.includes('UNIDAD'));
  const cAC=hd.findIndex(x=>x.includes('TOTAL')&&(x.includes('2026')||x.includes('EJECUCION')));
  const first=h+2, rows=[]; let last=first-1, cur=null;
  for(let i=first;i<R.length;i++){
    const r=R[i]; const mun=r[cMun], um=r[cUM];
    if(r[cDep]!=null&&String(r[cDep]).trim()) cur=r[cDep];
    if(!(mun!=null&&String(mun).trim()&&um!=null&&String(um).trim())) { if(rows.length) break; else continue; }
    last=i;
    rows.push({r:i,dep:cur,mun:String(mun),um:String(um).trim(),fis:[],fin:[]});
    for(let m=0;m<12;m++){ rows[rows.length-1].fis.push(r[4+2*m]); rows[rows.length-1].fin.push(r[5+2*m]); }
  }
  let mesUlt=-1; for(let m=0;m<12;m++){ if(rows.some(x=>num(x.fis[m])>0||num(x.fin[m])>0)) mesUlt=m; }
  return {h,first,last,totRow:last+1,cAC:cAC<0?28:cAC,rows,mesUlt};
}
const ukey = x => canonDep(x.dep)+'|'+norm(x.mun)+'|'+norm(x.um).split(' ')[0];

function construirComedores(raw, base, ctx){
  const P=ctx.prog('comedores','Comedor Social');
  const wsR=raw.ws, pr=parseComedores(wsR); if(!pr){ P.error('estructura','No se encontró la tabla DEPARTAMENTO / MUNICIPIO / UNIDAD.'); return null; }
  const rows=pr.rows;
  if(rows.length%2) P.error('pares','Número impar de filas de datos ('+rows.length+'): cada comedor debe tener Desayuno y Almuerzo.');
  // pares
  for(let i=0;i<rows.length;i+=2){ const d=rows[i], a=rows[i+1]; if(!a) break;
    if(!norm(d.um).startsWith('DESAYUNO')||!norm(a.um).startsWith('ALMUERZO')) P.error('pares','Fila '+(d.r+1)+': se esperaba Desayuno/Almuerzo y hay «'+d.um+'» / «'+a.um+'».');
    if(canonDep(d.dep)!==canonDep(a.dep)||norm(d.mun)!==norm(a.mun)) { if(norm(d.mun).replace(/ /g,'')!==norm(a.mun).replace(/ /g,'')) P.aviso('par_nombre','Fila '+(d.r+1)+': desayuno «'+clean(d.mun)+'» y almuerzo «'+clean(a.mun)+'» no coinciden en nombre.'); }
  }
  // celdas
  let nForm=0;
  for(const x of rows){ for(let m=0;m<12;m++){ for(const t of ['fis','fin']){ const v=x[t][m]; if(v==null||v==='') continue; if(typeof v!=='number') P.error('celda_no_numerica','Fila '+(x.r+1)+' '+MESES[m]+' ('+t+'): «'+v+'» no es numérico.'); else if(v<0) P.error('negativo','Fila '+(x.r+1)+' '+MESES[m]+': valor negativo '+v+'.'); } } }
  const keys=rows.map(ukey); const seen={}; keys.forEach((k,i)=>{ if(seen[k]!=null && i%2===0) P.error('duplicado','Comedor repetido: '+clean(rows[i].dep)+' / '+clean(rows[i].mun)); seen[k]=i; });
  if(pr.mesUlt<0){ P.error('sin_datos','El archivo no trae ninguna ración ni monto.'); return null; }
  // AC/AD vs suma
  let nAC=0; for(const x of rows){ const cac=vget(wsR,x.r,pr.cAC); const s=x.fis.reduce((a,v)=>a+num(v),0); if(cac!=null&&Math.abs(num(cac)-s)>0.005) nAC++; }
  if(nAC) P.aviso('total_fila','En '+nAC+' fila(s) el total del archivo no coincide con la suma de sus meses; se recalcula con la suma.');
  // totales por mes
  const tot=[]; for(let m=0;m<=pr.mesUlt;m++){ tot.push({mes:m,fis:rows.reduce((a,x)=>a+num(x.fis[m]),0),fin:r2(rows.reduce((a,x)=>a+num(x.fin[m]),0))}); }
  // fila de totales del crudo
  const diffTot=[]; tot.forEach(t=>{ const c=vget(wsR,pr.totRow,4+2*t.mes); if(c!=null&&typeof c==='number'&&Math.abs(c-t.fis)>0.5) diffTot.push(MESES[t.mes]); });
  if(diffTot.length) P.aviso('fila_total_crudo','La fila de totales del archivo no coincide con la suma de filas en: '+diffTot.join(', ')+'.');
  // comparación con el dps anterior
  let prev=null; if(base.Sheets['COMEDORES']) prev=parseComedores(base.Sheets['COMEDORES']);
  const cmp={nuevos:[],bajas:[],cambios:[],prevMes:-1};
  if(prev){
    cmp.prevMes=prev.mesUlt;
    const pk=new Map(prev.rows.map(x=>[ukey(x),x])), nk=new Map(rows.map(x=>[ukey(x),x]));
    const nuevos=new Set(), bajas=new Set();
    nk.forEach((x,k)=>{ if(!pk.has(k)) nuevos.add(canonDep(x.dep)+' / '+clean(x.mun)); });
    pk.forEach((x,k)=>{ if(!nk.has(k)) bajas.add(canonDep(x.dep)+' / '+clean(x.mun)); });
    cmp.nuevos=[...nuevos]; cmp.bajas=[...bajas];
    // cambios en meses cerrados, comparando por departamento (robusto a fusiones/renombres)
    const agg=(R,mx)=>{ const o={}; R.forEach(x=>{ const d=canonDep(x.dep); const a=(o[d]=o[d]||{fis:Array(12).fill(0),fin:Array(12).fill(0)}); for(let m=0;m<=mx;m++){ a.fis[m]+=num(x.fis[m]); a.fin[m]+=num(x.fin[m]); } }); return o; };
    const mx=Math.min(prev.mesUlt,pr.mesUlt), A1=agg(prev.rows,mx), A2=agg(rows,mx);
    new Set([...Object.keys(A1),...Object.keys(A2)]).forEach(d=>{ for(let m=0;m<=mx;m++){ const a=(A1[d]||{fis:[],fin:[]}), b=(A2[d]||{fis:[],fin:[]}); const df=num(b.fis[m])-num(a.fis[m]), dn=num(b.fin[m])-num(a.fin[m]); if(Math.abs(df)>0.5||Math.abs(dn)>0.005) cmp.cambios.push({dep:d,mes:MESES[m],dFis:df,dFin:r2(dn)}); } });
    if(pr.mesUlt<=prev.mesUlt) P.aviso('mes_no_avanza','El crudo llega hasta '+MESES[pr.mesUlt]+' y el dps anterior ya llegaba hasta '+MESES[prev.mesUlt]+'.');
    if(nuevos.size) P.aviso('nuevos',nuevos.size+' comedor(es) nuevo(s) frente al dps anterior: '+cmp.nuevos.slice(0,14).join('; ')+(nuevos.size>14?'…':'')+'.');
    if(bajas.size) P.aviso('bajas',bajas.size+' unidad(es) del dps anterior ya no existen con ese nombre (fusión o cierre): '+cmp.bajas.join('; ')+'. Se verificó que no cambia ningún total de meses cerrados: '+(cmp.cambios.length?'¡SÍ HAY DIFERENCIAS!':'sin diferencias')+'.');
    if(cmp.cambios.length) P.aviso('cambios_cerrados',cmp.cambios.length+' diferencia(s) por departamento en meses ya cerrados (ej.: '+cmp.cambios.slice(0,3).map(c=>c.dep+' '+c.mes+' '+c.dFis+' rac.').join('; ')+').');
    else P.info('cerrados','Los meses ya cerrados ('+MESES[0]+'–'+MESES[mx]+') no cambiaron respecto al dps anterior.');
  }
  rows.forEach(x=>{ if(/ACATENANCO/.test(norm(x.mun))) { if(!P._acat){ P._acat=1; P.aviso('errata','El nombre «'+clean(x.mun)+'» parece tener una errata («ACATENANCO» por «ACATENANGO»). Afecta su ubicación en el mapa; se respeta tal como viene.'); } } });
  // anomalías
  let sinMonto=0, montoSinRac=0; const lstMS=[];
  for(let i=0;i<rows.length;i+=2){ const d=rows[i], a=rows[i+1]||{fis:[],fin:[]}; for(let m=0;m<=pr.mesUlt;m++){ const rac=num(d.fis[m])+num(a.fis[m]), mon=num(d.fin[m]); if(rac>0&&mon===0) sinMonto++; if(mon>0&&rac===0){ montoSinRac++; if(m===pr.mesUlt) lstMS.push(clean(d.mun)+' (Q'+mon.toLocaleString('en-US',{maximumFractionDigits:0})+')'); } } }
  if(sinMonto) P.info('rac_sin_monto',sinMonto+' comedor-mes tienen raciones sin pago registrado (el pago suele reportarse con 1–2 meses de rezago).');
  if(montoSinRac) P.aviso('monto_sin_rac',montoSinRac+' comedor-mes tienen pago sin raciones'+(lstMS.length?'; en '+MESES[pr.mesUlt]+': '+lstMS.slice(0,8).join(', ')+(lstMS.length>8?'…':''):'')+'.');
  const zeroNow=[]; for(let i=0;i<rows.length;i+=2){ const d=rows[i],a=rows[i+1]||{fis:[]}; const k=pr.mesUlt; if(num(d.fis[k])+num(a.fis[k])===0 && k>0 && num(d.fis[k-1])+num(a.fis[k-1])>0) zeroNow.push(clean(d.mun)); }
  if(zeroNow.length) P.aviso('dejaron_operar',zeroNow.length+' comedor(es) sin raciones en '+MESES[pr.mesUlt]+' tras operar en '+MESES[pr.mesUlt-1]+': '+zeroNow.slice(0,10).join(', ')+(zeroNow.length>10?'…':'')+'.');
  if(tot.length>1){ const a=tot[tot.length-2], b=tot[tot.length-1]; const ch=a.fis?((b.fis-a.fis)/a.fis*100):0; if(Math.abs(ch)>=10) P.aviso('variacion_mes','Las raciones de '+MESES[b.mes]+' ('+b.fis.toLocaleString('en-US')+') '+(ch<0?'bajan ':'suben ')+Math.abs(ch).toFixed(1)+' % frente a '+MESES[a.mes]+' ('+a.fis.toLocaleString('en-US')+').'); }
  if(P.fail) return null;

  /* construir la hoja nueva */
  const ws={}; const lastRow=pr.last, nC=30;
  for(let r=0;r<=lastRow;r++) for(let c=0;c<nC;c++){
    const x=cget(wsR,r,c); if(!x) continue;
    if(r>=pr.first&&(c===28||c===29)) continue;
    const y=copyCell(x);
    if(r>=pr.first&&c>=4&&c<28&&y.f){ nForm++; delete y.f; }
    ws[A(r,c)]=y;
  }
  for(const x of rows){
    const r=x.r; const sumF=[...Array(12).keys()].map(i=>colL(4+2*i)+(r+1)).reverse().join('+');
    cset(ws,r,28,x.fis.reduce((a,v)=>a+num(v),0),{f:sumF,z:(cget(wsR,r,28)||{}).z});
    if(norm(x.um).startsWith('DESAYUNO')){ const sf=[...Array(12).keys()].map(i=>colL(5+2*i)+(r+1)).reverse().join('+'); cset(ws,r,29,r2(x.fin.reduce((a,v)=>a+num(v),0)),{f:sf,z:(cget(wsR,r,29)||{}).z}); }
  }
  const tr=pr.totRow;
  for(let c=4;c<=29;c++){
    let s=0; for(const x of rows){ if(c<28){ const v=(c%2===0)?x.fis[(c-4)/2]:x.fin[(c-5)/2]; s+=num(v); } else if(c===28) s+=x.fis.reduce((a,v)=>a+num(v),0); else if(norm(x.um).startsWith('DESAYUNO')) s+=x.fin.reduce((a,v)=>a+num(v),0); }
    cset(ws,tr,c,r2(s),{f:'SUM('+colL(c)+(pr.first+1)+':'+colL(c)+(pr.last+1)+')',z:(cget(wsR,tr,c)||{}).z||'#,##0.00'});
  }
  cset(ws,tr,0,'TOTAL');
  ws['!merges']=(wsR['!merges']||[]).filter(m=>m.e.r<=lastRow&&m.e.c<nC).map(m=>({s:{...m.s},e:{...m.e}}));
  if(wsR['!cols']) ws['!cols']=wsR['!cols'].slice(0,nC);
  if(wsR['!rows']) ws['!rows']=wsR['!rows'].slice(0,tr+1);
  fixRef(ws);
  if(nForm) P.info('formulas_a_valor',nForm+' celda(s) mensuales eran fórmulas de constantes (p. ej. =80000+74000); se guardó su valor.');
  const totFis=rows.reduce((a,x)=>a+x.fis.reduce((b,v)=>b+num(v),0),0), totFin=r2(rows.reduce((a,x)=>a+x.fin.reduce((b,v)=>b+num(v),0),0));
  P.kv('Comedores (unidades)',rows.length/2); P.kv('Último mes con datos',MESES[pr.mesUlt]);
  P.kv('Raciones acumuladas',totFis); P.kv('Ejecución acumulada (Q)',totFin);
  P.kv('Raciones del mes',tot[pr.mesUlt].fis); P.kv('Ejecución del mes (Q)',tot[pr.mesUlt].fin);
  return {ws, mesUlt:pr.mesUlt, totFis, totFin, tot, rows, pr, cmp};
}

/* ---------- BECAS (Superior / Media) ---------- */
function construirBeca(raw, base, ctx, cfg){
  const P=ctx.prog(cfg.id,cfg.nombre);
  const wsR=raw.ws, wsB=base.Sheets[cfg.hoja];
  if(!wsB){ P.error('sin_base','El dps anterior no tiene la hoja '+cfg.hoja+'.'); return null; }
  const R=aoa(wsR), B=aoa(wsB);
  const hR=hasHeader(R,6,'DEPARTAMENTO','MUNICIPIO'), hB=hasHeader(B,6,'DEPARTAMENTO','MUNICIPIO');
  const keyRows=(AO,h)=>{ const out=[]; let dep=null; for(let i=h+2;i<AO.length;i++){ const r=AO[i]; if(r[1]!=null&&String(r[1]).trim()) dep=r[1]; const mun=r[2]; if(!(mun!=null&&String(mun).trim())){ if(out.length) break; else continue; } if(norm(r[0]).includes('TOTAL')||norm(mun).includes('TOTAL')) break; out.push({i,k:canonDep(dep)+'|'+norm(mun)}); } return out; };
  // la tabla empieza en la fila siguiente a los encabezados de 2 filas
  const kr=keyRows(R,hR), kb=keyRows(B,hB);
  if(kr.length!==kb.length || kr.some((x,j)=>x.k!==kb[j].k || x.i!==kb[j].i)){
    const nuevos=kr.filter(x=>!kb.some(y=>y.k===x.k)).length, perd=kb.filter(x=>!kr.some(y=>y.k===x.k)).length;
    P.error('claves','La lista de municipios no coincide con el dps anterior ('+kr.length+' vs '+kb.length+' filas; '+nuevos+' nuevos, '+perd+' ausentes). No se inserta automáticamente para no romper BASE CONSOLIDADA; se conserva la hoja anterior.');
    return null;
  }
  const first=kr[0].i, last=kr[kr.length-1].i, nC=XU.decode_col('AS')+1;
  // valores
  let neg=0, text=0; const dif=[];
  const mesCol=m=>6+3*m; // física de mes m (0=ene)
  // último mes con datos
  let mesUlt=-1; for(let m=0;m<12;m++){ for(let i=first;i<=last;i++){ if(num(vget(wsR,i,mesCol(m)))>0){ mesUlt=m; break; } } }
  for(let i=first;i<=last;i++){ for(let c=4;c<nC;c++){ const v=vget(wsR,i,c); if(v==null) continue; if(typeof v!=='number'){ text++; } else if(v<0) neg++; } }
  if(text) P.error('texto',text+' celda(s) de datos con texto en lugar de número.');
  if(neg) P.error('negativo',neg+' valor(es) negativo(s).');
  if(P.fail) return null;
  // cambios en meses cerrados (mes < último)
  for(let i=first;i<=last;i++){ for(let m=0;m<mesUlt;m++){ for(let k=0;k<3;k++){ const c=mesCol(m)+k; const a=num(vget(wsB,i,c)), b=num(vget(wsR,i,c)); if(Math.abs(a-b)>0.005) dif.push({i,c,a,b}); } } }
  // sumas
  const sum=c=>{ let s=0; for(let i=first;i<=last;i++) s+=num(vget(wsR,i,c)); return s; };
  const S={fisU:sum(mesCol(mesUlt)), finU:sum(mesCol(mesUlt)+1), benU:sum(mesCol(mesUlt)+2), fisT:sum(CI('AQ')), finT:sum(CI('AR')), benT:sum(CI('AS'))};
  // validaciones específicas
  let rateBad=0, benLess=0; const tarifa=cfg.tarifa;
  for(let i=first;i<=last;i++){ const fi=num(vget(wsR,i,mesCol(mesUlt))), fn=num(vget(wsR,i,mesCol(mesUlt)+1)), bn=num(vget(wsR,i,mesCol(mesUlt)+2)); if(Math.abs(fn-fi*tarifa)>0.5) rateBad++; if(bn<fi-1e-9) benLess++; }
  if(rateBad) P.aviso('tarifa',rateBad+' fila(s) donde el financiero de '+MESES[mesUlt]+' no es física × Q'+tarifa+'.');
  if(benLess) P.aviso('benef_menor',benLess+' fila(s) con beneficiarios menores que la física.');
  // total declarado vs suma de meses
  const sumMes=[...Array(mesUlt+1).keys()].reduce((a,m)=>a+sum(mesCol(m)),0);
  if(Math.abs(sumMes-S.fisT)>0.5) P.aviso('total_no_suma','El «total 2026» que trae el archivo ('+S.fisT.toLocaleString('en-US',{maximumFractionDigits:2})+') NO es la suma de los meses ('+sumMes.toLocaleString('en-US',{maximumFractionDigits:2})+'): el archivo declara solo el corte de '+MESES[mesUlt]+'. Se respeta lo declarado; confirme con el programa que '+MESES[mesUlt]+' es acumulado a la fecha.');
  if(dif.length){ const tot=dif.reduce((a,x)=>a+(x.b-x.a),0); P.aviso('correcciones','El archivo corrige '+dif.length+' celda(s) de meses ya cerrados (ej.: '+dif.slice(0,3).map(d=>A(d.i,d.c)+': '+d.a+' → '+d.b).join('; ')+').'); }
  // resumen del archivo (hoja pequeña de resumen), si existe
  const rs=findSheet(raw.wb,(n,w)=>n!==raw.hoja && /SUPERIOR|RESUMEN/i.test(n) && topText(w,8,8).includes('SEPTIEMBRE'));
  if(rs){ const RR=aoa(raw.wb.Sheets[rs]); const tr=RR.find(r=>r&&norm(r[0]).includes('TOTAL')); if(tr){ const nums=tr.filter(v=>typeof v==='number'); if(nums.length>=3 && Math.abs(nums[nums.length-2]-S.finU)>1) P.aviso('resumen_difiere','La hoja resumen del archivo ('+rs.trim()+') indica Q'+nums[nums.length-2].toLocaleString('en-US')+' y la hoja principal suma Q'+S.finU.toLocaleString('en-US')+'.'); else P.info('resumen_ok','Cuadra con la hoja resumen del archivo ('+rs.trim()+').'); } }
  // construir: copia de la base con datos del crudo
  const ws={}; for(const k of Object.keys(wsB)){ if(k[0]==='!') continue; ws[k]=copyCell(wsB[k]); }
  ['!merges','!cols','!rows'].forEach(k=>{ if(wsB[k]) ws[k]=JSON.parse(JSON.stringify(wsB[k])); });
  for(let i=first;i<=last;i++) for(let c=4;c<nC;c++){ const x=cget(wsR,i,c); if(x){ const y=copyCell(x); ws[A(i,c)]=y; } else delete ws[A(i,c)]; }
  const tr=last+1;
  // fila de totales: SUM de cada columna numérica
  for(let c=4;c<nC;c++){ const s=sum(c); const old=cget(wsB,tr,c); if(old||s!==0||c>=6){ cset(ws,tr,c,r2(s*1e4)/1e4,{f:'SUM('+colL(c)+(first+1)+':'+colL(c)+(last+1)+')',z:old&&old.z}); } }
  fixRef(ws);
  P.kv('Municipios',kr.length); P.kv('Último mes con datos',MESES[mesUlt]);
  P.kv('Física del mes',S.fisU); P.kv('Financiero del mes (Q)',S.finU); P.kv('Beneficiarios del mes',S.benU);
  P.kv('Total 2026 declarado · física',S.fisT); P.kv('Total 2026 declarado · Q',S.finT);
  return {ws, mesUlt, S, first, last};
}

/* ---------- JÓVENES PROTAGONISTAS ---------- */
const JP_ALIAS={'PALESTIAN DE LOS ALTOS':'PALESTINA DE LOS ALTOS','SAN MIGUEL PANAM':'SAN MIGUEL PANAN'};
function construirJP(raw, base, ctx){
  const P=ctx.prog('jp','Jóvenes Protagonistas');
  const wsR=raw.ws, wsB=base.Sheets['JOVENES PROTAGONISTAS'];
  if(!wsB){ P.error('sin_base','El dps anterior no tiene la hoja JOVENES PROTAGONISTAS.'); return null; }
  const R=aoa(wsR), h=hasHeader(R,15,'DEPARTAMENTO','MUNICIPIO','BENEFICIARIOS');
  const hd=R[h].map(norm), cB=hd.findIndex(x=>x.includes('BENEFICIARIOS')), cE=hd.findIndex(x=>x.includes('PRESUPUESTO'));
  const dat=[]; let totGen=null, dep=null;
  for(let i=h+1;i<R.length;i++){ const r=R[i]; const a=norm(r[0]);
    if(a==='TOTAL'){ const nxt=R.slice(i+1).find(x=>x&&norm(x[0])!==''); if(!r[1]&&!r[2]&&(norm((R[i+1]||[])[0]).startsWith('FUENTE')||!R[i+1]||!R[i+1][0])) { totGen={d:r[cB],e:r[cE],i}; break; } continue; }
    if(a.startsWith('FUENTE')) break;
    if(r[1]!=null&&String(r[1]).trim()) dep=r[1];
    if(r[2]!=null&&String(r[2]).trim()&&typeof r[cB]==='number') dat.push({i,dep:canonDep(dep),mun:clean(r[2]),ben:r[cB],ppto:num(r[cE])});
  }
  if(!dat.length){ P.error('sin_datos','No se encontraron municipios con beneficiarios.'); return null; }
  const sb=dat.reduce((a,x)=>a+x.ben,0), se=dat.reduce((a,x)=>a+x.ppto,0);
  if(totGen){ if(Math.abs(totGen.d-sb)>0.5) P.error('total','La suma de municipios ('+sb+') no coincide con el TOTAL del archivo ('+totGen.d+').'); if(Math.abs(totGen.e-se)>1) P.aviso('total_q','El total en Q del archivo ('+totGen.e.toLocaleString('en-US')+') difiere de la suma de municipios por '+Math.abs(totGen.e-se).toFixed(2)+' (redondeo).'); }
  else P.aviso('sin_total','No se encontró la fila TOTAL general del archivo; el total se calculó sumando municipios.');
  if(P.fail) return null;
  const tarifas=dat.filter(x=>x.ben>0).map(x=>x.ppto/x.ben); const t0=tarifas[0]; const uniforme=tarifas.every(t=>Math.abs(t-t0)<1e-6);
  if(uniforme) P.aviso('prorrateo','El presupuesto del archivo es beneficiarios × Q'+t0.toFixed(4)+' en TODOS los municipios: es un prorrateo uniforme, no el gasto real por municipio. Úselo solo a nivel nacional/departamental.');
  const ac=wsB; const Bd=aoa(wsB);
  // filas de base
  const baseRows=[]; let d2=null; for(let i=4;i<=131&&i<Bd.length;i++){ const r=Bd[i]; if(r[1]!=null&&String(r[1]).trim()) d2=r[1]; if(r[2]!=null&&String(r[2]).trim()&&!norm(r[2]).includes('TOTAL')) baseRows.push({i,dep:canonDep(d2),mun:clean(r[2]),acum:num(r[18]),fin:num(r[19])}); }
  const kn=m=>{ const n=norm(m); return JP_ALIAS[n]||n; };
  const idxR=new Map(); dat.forEach(x=>idxR.set(x.dep+'|'+kn(x.mun),x));
  const byMun=new Map(); dat.forEach(x=>{ const k=kn(x.mun); byMun.set(k,(byMun.get(k)||[]).concat(x)); });
  const baseByMun=new Map(); baseRows.forEach(x=>{ const k=kn(x.mun); baseByMun.set(k,(baseByMun.get(k)||[]).concat(x)); });
  const usados=new Set(), pares=[]; const sinPar=[];
  baseRows.forEach(b=>{ let x=idxR.get(b.dep+'|'+kn(b.mun)); let nota='';
    if(!x){ const c=byMun.get(kn(b.mun)); if(c&&c.length===1&&(baseByMun.get(kn(b.mun))||[]).length===1){ x=c[0]; nota='departamento distinto en el crudo ('+x.dep+')'; } }
    if(x&&!usados.has(x)){ usados.add(x); pares.push({b,x,nota}); } else sinPar.push(b); });
  const nuevos=dat.filter(x=>!usados.has(x));
  // menor que acumulado de julio
  const menores=pares.filter(p=>p.x.ben<p.b.acum-1e-9);
  if(menores.length) P.aviso('acum_menor',menores.length+' municipio(s) con acumulado de septiembre MENOR al de julio: '+menores.slice(0,6).map(p=>p.b.mun+' ('+p.x.ben+' < '+p.b.acum+')').join(', ')+'.');
  const dups=[...byMun.entries()].filter(([k,v])=>v.length>1).map(([k,v])=>v[0].mun+' ('+v.map(z=>z.dep).join(' / ')+')');
  if(nuevos.length) P.aviso('municipios_nuevos',nuevos.length+' municipio(s) en el archivo sin histórico en el dps ('+nuevos.reduce((a,x)=>a+x.ben,0)+' beneficiarios): '+nuevos.map(x=>x.mun+' ('+x.dep+')').join('; ')+'. Se agregan al bloque «Nuevos en el corte» con enero–julio sin dato (no cero).');
  if(sinPar.length) P.aviso('municipios_ausentes',sinPar.length+' municipio(s) del dps anterior no aparecen en el archivo: '+sinPar.map(x=>x.mun).join(', ')+'. Se conservan con su detalle enero–julio.');
  if(dups.length) P.aviso('nombre_duplicado','Nombres de municipio repetidos en departamentos distintos: '+dups.join('; ')+'. Verifique que no sea un error de departamento.');
  P.aviso('sin_detalle_mensual','El archivo solo trae el ACUMULADO enero–septiembre, sin detalle por mes. El detalle enero–julio del dps se conserva intacto; el acumulado nuevo se agrega en columnas aparte (AC:AF) y no se inventa agosto ni septiembre por separado.');
  const ctr=aoa(wsR).slice(-6).map(r=>r&&r.join(' ')).join(' ');
  // construir
  const ws={}; for(const k of Object.keys(wsB)){ if(k[0]==='!') continue; ws[k]=copyCell(wsB[k]); }
  ['!merges','!cols','!rows'].forEach(k=>{ if(wsB[k]) ws[k]=JSON.parse(JSON.stringify(wsB[k])); });
  const cs=CI('AC');
  cset(ws,2,cs,'ACUMULADO AL CORTE (archivo de '+(raw.corte||'septiembre')+')'); ws['!merges']=(ws['!merges']||[]).concat([{s:{r:2,c:cs},e:{r:2,c:cs+3}}]);
  ['Beneficiarios acumulados','Presupuesto ejecutado acumulado (Q)','Aumento vs. julio (beneficiarios)','Observación'].forEach((t,j)=>cset(ws,3,cs+j,t));
  pares.forEach(p=>{ const r=p.b.i; cset(ws,r,cs,p.x.ben); cset(ws,r,cs+1,r2(p.x.ppto)); cset(ws,r,cs+2,p.x.ben-p.b.acum,{f:colL(cs)+(r+1)+'-S'+(r+1)}); if(p.nota) cset(ws,r,cs+3,p.nota); });
  sinPar.forEach(b=>cset(ws,b.i,cs+3,'No aparece en el archivo del corte'));
  // bloque de nuevos
  const r0=148; cset(ws,r0,1,'MUNICIPIOS NUEVOS EN EL CORTE (sin detalle enero–julio)');
  nuevos.forEach((x,j)=>{ const r=r0+1+j; cset(ws,r,1,x.dep); cset(ws,r,2,x.mun); cset(ws,r,3,'Persona'); cset(ws,r,cs,x.ben); cset(ws,r,cs+1,r2(x.ppto)); cset(ws,r,cs+3,'Nuevo en el corte'); });
  const rt=r0+1+nuevos.length;
  cset(ws,rt,2,'TOTAL ACUMULADO (archivo)'); cset(ws,rt,cs,sb,{f:'SUM('+colL(cs)+'5:'+colL(cs)+'133)+SUM('+colL(cs)+(r0+2)+':'+colL(cs)+rt+')'}); cset(ws,rt,cs+1,r2(se),{f:'SUM('+colL(cs+1)+'5:'+colL(cs+1)+'133)+SUM('+colL(cs+1)+(r0+2)+':'+colL(cs+1)+rt+')'});
  fixRef(ws);
  const julio=baseRows.reduce((a,b)=>a+b.acum,0);
  P.kv('Municipios en el archivo',dat.length); P.kv('Beneficiarios acumulados',sb); P.kv('Presupuesto ejecutado acumulado (Q)',r2(se));
  P.kv('Acumulado en el dps anterior (jul)',julio); P.kv('Aumento (ago+sep combinados)',sb-julio);
  return {ws, sb, se:r2(se), julio, nuevos:nuevos.length, sinPar:sinPar.length};
}

/* ---------- BOLSA SOCIAL (TMCA) ---------- */
function construirBolsa(raw, base, ctx){
  const P=ctx.prog('bolsa','Bolsa Social');
  const wsR=raw.ws, R=aoa(wsR);
  const h=hasHeader(R,12,'DEPARTAMENTO','MUNICIPIO'); if(h<0){ P.error('estructura','No se encontró el encabezado DEPARTAMENTO / MUNICIPIO.'); return null; }
  const hd=R[h].map(norm);
  const cMun=hd.findIndex(x=>x.includes('MUNICIPIO')), cDep=hd.findIndex(x=>x.includes('DEPARTAMENTO'));
  const cTotA=hd.findIndex(x=>x.includes('TOTAL APORTES')), cTotM=hd.findIndex(x=>x.includes('MONTO TOTAL'));
  if(cTotA<0||cTotM<0){ P.error('columnas','No se encontraron las columnas «TOTAL APORTES EFECTUADOS» / «MONTO TOTAL EJECUTADO».'); return null; }
  // mes por columna (fila de meses = h-1, con celdas combinadas)
  const mrow=h-1; const merges=wsR['!merges']||[]; const mesCol={};
  const labelAt=(r,c)=>{ const v=vget(wsR,r,c); if(v!=null) return norm(v); const m=merges.find(m=>m.s.r<=r&&m.e.r>=r&&m.s.c<=c&&m.e.c>=c); return m?norm(vget(wsR,m.s.r,m.s.c)):''; };
  let prev=-1; const cIni=Math.min(...[6]), cFin=cTotA-1;
  for(let c=cIni;c<=cFin;c++){ let m=MESES_U.findIndex(x=>labelAt(mrow,c)===x); if(m<0) m=MESES_U.findIndex(x=>norm(vget(wsR,h,c)).includes(x)); if(m<0) m=prev+1; mesCol[c]=m; prev=m; }
  // filas de datos
  const dat=[]; let dep=null, totRow=-1;
  for(let i=h+1;i<R.length;i++){ const r=R[i]; if(norm(r[2])==='TOTAL'||norm(r[0])==='TOTAL'){ totRow=i; break; } if(r[cDep]!=null&&String(r[cDep]).trim()) dep=r[cDep]; if(r[cMun]!=null&&String(r[cMun]).trim()) dat.push({i,dep:canonDep(dep),mun:clean(r[cMun])}); }
  if(!dat.length||totRow<0){ P.error('estructura','No se encontraron municipios o la fila TOTAL.'); return null; }
  const price=(()=>{ const x=num(vget(wsR,dat[0].i,cTotA)), y=num(vget(wsR,dat[0].i,cTotM)); return x>0?y/x:325; })();
  // recomputar
  let nBad=0; const monthTot=Array(12).fill(0); let totAp=0;
  dat.forEach(d=>{ let s=0; for(let c=cIni;c<=cFin;c++){ const v=vget(wsR,d.i,c); if(v==null||c===6) { if(typeof v==='number'&&c===6){ monthTot[0]+=v; } continue; } if(typeof v!=='number'){ P.error('texto','Celda '+A(d.i,c)+' no es numérica.'); continue; } if(v<0) P.error('negativo','Celda '+A(d.i,c)+' negativa.'); s+=v; monthTot[mesCol[c]]+=v; }
    d.tot=s; totAp+=s; const cache=num(vget(wsR,d.i,cTotA)); if(Math.abs(cache-s)>0.5) nBad++; });
  if(P.fail) return null;
  if(nBad) P.aviso('total_fila',nBad+' fila(s) donde «TOTAL APORTES» del archivo no coincide con la suma de sus planillas (se recalcula; enero se excluye como en el original).');
  let mesUlt=-1; monthTot.forEach((v,m)=>{ if(v>0) mesUlt=m; });
  const monto=r2(totAp*price);
  // validaciones con hojas auxiliares
  const wsRA=raw.wb.Sheets[findSheet(raw.wb,n=>norm(n).includes('RESUMEN ANUAL'))||''], wsPY=raw.wb.Sheets[findSheet(raw.wb,n=>norm(n).includes('PROYECCION'))||''];
  let metaFis=null, vigente=num(vget(wsR,totRow+4,cTotM));
  // presupuesto asignado: fila con texto «Presupuesto Asignado» en la columna cTotA
  for(let i=totRow;i<Math.min(R.length,totRow+10);i++){ if(norm(vget(wsR,i,cTotA)).includes('PRESUPUESTO ASIGNADO')){ vigente=num(vget(wsR,i,cTotM)); break; } }
  if(wsRA){ const RA=aoa(wsRA); const tr=RA.find(r=>r&&norm(r[2])==='TOTAL'||norm(r[0])==='TOTAL'); const nums=tr?tr.filter(v=>typeof v==='number'):[]; if(nums.length>=2){ if(Math.abs(nums[1]-totAp)>0.5) P.aviso('resumen_anual','La hoja RESUMEN ANUAL indica '+nums[1].toLocaleString('en-US')+' aportes y la hoja principal suma '+totAp.toLocaleString('en-US')+'.'); else P.info('resumen_anual','Cuadra con la hoja RESUMEN ANUAL del archivo ('+totAp.toLocaleString('en-US')+' aportes).'); } }
  if(wsPY){ const PY=aoa(wsPY); const tr=PY.find(r=>r&&r.some(v=>norm(v)==='')===false&&0);
    const hy=hasHeader(PY,12,'EJECUCION FISICA'); if(hy>=0){ let s=0; for(let i=hy+1;i<PY.length;i++){ const r=PY[i]; const mm=MESES_U.findIndex(x=>x===norm(r[5])); if(mm>=0&&typeof r[6]==='number'){ s+=r[6]; if(mm<=mesUlt&&Math.abs(r[6]-monthTot[mm])>0.5) P.aviso('proyeccion','PROYECCIÓN indica '+r[6].toLocaleString('en-US')+' para '+MESES[mm]+' y las planillas suman '+monthTot[mm].toLocaleString('en-US')+'.'); } } if(s>0) metaFis=s; } }
  if(metaFis==null) { metaFis=Math.floor(vigente/price); P.aviso('meta_estimada','No hay hoja PROYECCIÓN: la meta física se estimó como presupuesto ÷ Q'+price+'.'); }
  if(monto>vigente&&vigente>0) P.aviso('excede_presupuesto','El monto ejecutado (Q'+monto.toLocaleString('en-US')+') supera el presupuesto asignado (Q'+vigente.toLocaleString('en-US')+').');
  // cambios en meses cerrados vs dps anterior
  const wsB=base.Sheets['BOLSA']; let difs=0, comp=0;
  if(wsB){ const RB=aoa(wsB), hb=hasHeader(RB,12,'DEPARTAMENTO','MUNICIPIO'); const mb={};
    // columnas del base con mes ≤ mesUlt-1
    for(let c=6;c<=cFin;c++){ const bh=norm(vget(wsB,hb,c)); if(c<=16){ mb[c]=true; } }
    const bidx=new Map(); for(let i=hb+1;i<RB.length;i++){ if(norm(RB[i][2])==='TOTAL') break; if(RB[i][4]!=null) bidx.set(norm(RB[i][4]),i); }
    dat.forEach(d=>{ const bi=bidx.get(norm(d.mun)); if(bi==null) return; for(const c of Object.keys(mb).map(Number)){ if(mesCol[c]!==undefined&&mesCol[c]<mesUlt-1){ comp++; if(Math.abs(num(vget(wsR,d.i,c))-num(vget(wsB,bi,c)))>0.5) difs++; } } });
    const nuevos=dat.filter(d=>!bidx.has(norm(d.mun))).length; if(nuevos) P.aviso('municipios_nuevos',nuevos+' municipio(s) nuevo(s) frente al dps anterior.');
  }
  if(difs) P.aviso('cambios_cerrados',difs+' celda(s) de meses cerrados cambian respecto al dps anterior.'); else if(comp) P.info('cerrados','Los aportes de enero a '+MESES[Math.max(0,mesUlt-2)]+' no cambiaron respecto al dps anterior ('+comp+' celdas comparadas).');
  // planillas con datos repetidos
  const colsMes={}; Object.entries(mesCol).forEach(([c,m])=>{ (colsMes[m]=colsMes[m]||[]).push(+c); });
  const multi=Object.entries(colsMes).filter(([m,cs])=>cs.length>1&&+m<=mesUlt&&+m>=mesUlt-1).map(([m,cs])=>MESES[m]+' ('+cs.length+' planillas)');
  if(multi.length) P.info('planillas','Meses con varias planillas, sumadas por mes: '+multi.join(', ')+'.');
  const hidden=(wsR['!cols']||[]).map((c,i)=>c&&c.hidden?colL(i):null).filter(Boolean).filter(l=>{ const c=CI(l); return c>=cIni&&c<=cFin&&dat.some(d=>num(vget(wsR,d.i,c))!==0); });
  if(hidden.length) P.info('ocultas','Columnas ocultas con datos en el archivo: '+hidden.join(', ')+' (se conservan y se cuentan).');
  // construir (copia del layout del crudo, con totales recalculados)
  const ws={}; for(const k of Object.keys(wsR)){ if(k[0]==='!') continue; ws[k]=copyCell(wsR[k]); }
  ['!merges','!cols','!rows'].forEach(k=>{ if(wsR[k]) ws[k]=JSON.parse(JSON.stringify(wsR[k])); });
  const ltA=colL(cTotA), ltM=colL(cTotM);
  dat.forEach(d=>{ cset(ws,d.i,cTotA,d.tot,{f:'SUM(H'+(d.i+1)+':'+colL(cFin)+(d.i+1)+')',z:'#,##0'}); cset(ws,d.i,cTotM,r2(d.tot*price),{f:'+'+ltA+(d.i+1)+'*'+price,z:(cget(wsR,d.i,cTotM)||{}).z}); });
  for(let c=7;c<=cFin;c++){ let s=0; dat.forEach(d=>s+=num(vget(wsR,d.i,c))); if(s!==0||cget(wsR,totRow,c)) cset(ws,totRow,c,s,{f:'SUM('+colL(c)+(dat[0].i+1)+':'+colL(c)+(dat[dat.length-1].i+1)+')',z:'#,##0'}); }
  cset(ws,totRow,cTotA,totAp,{f:'SUM('+ltA+(dat[0].i+1)+':'+ltA+(dat[dat.length-1].i+1)+')'}); cset(ws,totRow,cTotM,monto,{f:'SUM('+ltM+(dat[0].i+1)+':'+ltM+(dat[dat.length-1].i+1)+')'});
  for(let c=7;c<=cFin;c++){ const x=cget(ws,totRow+1,c); if(x&&x.f){ const s=num(cget(ws,totRow,c)&&cget(ws,totRow,c).v); x.v=r2(s*price); x.f='+'+colL(c)+(totRow+1)+'*'+price; } }
  fixRef(ws);
  P.kv('Municipios',dat.length); P.kv('Último mes con aportes',MESES[mesUlt]); P.kv('Aportes del mes',monthTot[mesUlt]); P.kv('Aportes acumulados',totAp); P.kv('Monto ejecutado (Q)',monto); P.kv('Presupuesto asignado (Q)',vigente); P.kv('Precio por aporte (Q)',price);
  return {ws, mesUlt, monthTot, totAp, monto, vigente, metaFis, price, dat, mesCol, cIni, cFin, cTotA, cTotM, totRow};
}

/* ---------- INTEGRADO ---------- */
const INT_FILAS={'COMEDOR SOCIAL':'comedores','BOLSA SOCIAL':'bolsa','JOVENES PROTAGONISTAS':'jp','BECA MEDIA':'bm','BECA SUPERIOR':'bes'};
function actualizarIntegrado(wb, res, ctx, manual){
  const ws=wb.Sheets['INTEGRADO']; const P=ctx.prog('integrado','Hoja INTEGRADO');
  if(!ws){ P.error('sin_hoja','El dps anterior no tiene la hoja INTEGRADO.'); return; }
  const filas={}; const R=aoa(ws);
  R.forEach((r,i)=>{ const k=INT_FILAS[norm(r&&r[3])]; if(k) filas[k]=i; });
  const cambios=ctx.cambios;
  const set=(k,col,val,origen,estado)=>{ const r=filas[k]; if(r==null) return; const c=CI(col); const old=vget(ws,r,c); if(old===val || (typeof old==='number'&&typeof val==='number'&&Math.abs(old-val)<1e-9)) return; const x=cget(ws,r,c); cset(ws,r,c,val,{z:x&&x.z}); cambios.push({prog:k,celda:col+(r+1),antes:old,despues:val,origen,estado:estado||'derivado'}); };
  const pend=(k,txt)=>ctx.pendientes.push({prog:k,txt});
  // Mes de reporte = el más avanzado entre los programas con detalle mensual
  const m=[res.comedores&&res.comedores.mesUlt, res.bolsa&&res.bolsa.mesUlt, res.bes&&res.bes.mesUlt].filter(v=>v!=null&&v>=0);
  const mes=Math.max(...m); ctx.mes=mes;
  const i2=A(1,CI('I')); { const x=ws[i2]; const txt='EJECUCION DE '+MESES_U[mes]+' GS 2026 DPS'; if(!x||x.v!==txt){ ctx.cambios.push({prog:'general',celda:'I2',antes:x&&x.v,despues:txt,origen:'mes del último dato',estado:'derivado'}); ws[i2]={t:'s',v:txt}; } }
  if(res.comedores){ const c=res.comedores, k=c.mesUlt; set('comedores','I',c.tot[k].fis,'COMEDORES · raciones de '+MESES[k]); set('comedores','J',c.tot[k].fin,'COMEDORES · Q de '+MESES[k]); set('comedores','L',c.totFis,'COMEDORES · raciones acumuladas'); set('comedores','O',c.totFin,'COMEDORES · Q acumulado'); }
  if(res.bolsa){ const b=res.bolsa, k=b.mesUlt; set('bolsa','I',b.monthTot[k],'BOLSA · aportes de '+MESES[k]); set('bolsa','J',r2(b.monthTot[k]*b.price),'BOLSA · aportes × Q'+b.price); set('bolsa','K',b.metaFis,'PROYECCIÓN · meta física del año'); set('bolsa','L',b.totAp,'BOLSA · aportes acumulados'); set('bolsa','N',b.vigente,'BOLSA · presupuesto asignado'); set('bolsa','O',b.monto,'BOLSA · monto ejecutado'); }
  if(res.bes){ const b=res.bes, S=b.S; set('bes','I',S.fisU,'BECA SUPERIOR · física de '+MESES[b.mesUlt]); set('bes','J',S.finU,'BECA SUPERIOR · Q de '+MESES[b.mesUlt]); set('bes','L',S.fisT,'BECA SUPERIOR · total declarado'); set('bes','O',S.finT,'BECA SUPERIOR · total declarado (Q)'); }
  if(res.bm){ const b=res.bm, S=b.S; set('bm','I',S.fisU,'BECA MEDIA · física de '+MESES[b.mesUlt]); set('bm','J',S.finU,'BECA MEDIA · Q de '+MESES[b.mesUlt]); set('bm','L',S.fisT,'BECA MEDIA · total declarado'); set('bm','O',S.finT,'BECA MEDIA · total declarado (Q)'); }
  if(res.jp){ const j=res.jp; set('jp','L',j.sb,'JÓVENES · beneficiarios acumulados del archivo'); set('jp','G',j.se,'JÓVENES · presupuesto ejecutado acumulado'); set('jp','O',j.se,'JÓVENES · presupuesto ejecutado acumulado'); pend('jp','I6 y J6 (ejecución del último mes) no se pueden obtener: el archivo solo trae el acumulado. Quedan con el dato de julio.'); }
  if(!res.bm) pend('bm','Beca Media no vino en el ZIP: su hoja y su fila de INTEGRADO se conservan con los datos del dps anterior.');
  // entradas manuales
  const nom={comedores:'COMEDOR SOCIAL',bolsa:'BOLSA SOCIAL',jp:'JOVENES PROTAGONISTAS',bm:'BECA MEDIA',bes:'BECA SUPERIOR'};
  Object.keys(filas).forEach(k=>{ const mm=(manual||{})[k]||{}; Object.keys(mm).forEach(col=>{ const v=mm[col]; if(typeof v==='number'&&isFinite(v)) set(k,col,v,'Dato ingresado por el administrador','manual'); }); });
  // fórmulas H, M, P con valores
  Object.values(filas).forEach(r=>{ [['H','G','F'],['M','L','K'],['P','O','N']].forEach(([d,n,q])=>{ const a=num(vget(ws,r,CI(n))), b=num(vget(ws,r,CI(q))); const x=cget(ws,r,CI(d)); cset(ws,r,CI(d),b?a/b:0,{f:n+(r+1)+'/'+q+(r+1),z:x&&x.z}); }); });
  // coherencia
  Object.entries(filas).forEach(([k,r])=>{ const g=num(vget(ws,r,CI('G'))), o=num(vget(ws,r,CI('O'))), f=num(vget(ws,r,CI('F'))), n=num(vget(ws,r,CI('N'))), kk=num(vget(ws,r,CI('K'))), l=num(vget(ws,r,CI('L')));
    const manualG = ((manual||{})[k]||{}).G!=null;
    if(g<o-0.5 && !manualG && k!=='jp') pend(k,'La «Ejecución financiera» contable (G'+(r+1)+' = Q'+g.toLocaleString('en-US',{maximumFractionDigits:0})+') es menor que lo ejecutado según el programa (Q'+o.toLocaleString('en-US',{maximumFractionDigits:0})+'): G es un dato contable que no viene en los archivos y está sin actualizar. Ingréselo para que el % de ejecución general sea correcto.');
    if(l>kk&&kk>0) ctx.avisos.push({nivel:'aviso',prog:'integrado',cod:'sobre_meta',msg:k+': lo ejecutado físico ('+l.toLocaleString('en-US')+') supera la meta vigente ('+kk.toLocaleString('en-US')+').'}); });
  ctx.filasInt=filas;
}

/* ---------- BASE CONSOLIDADA y RESUMEN EJECUTIVO ---------- */
function construirBase(wb, res, ctx){
  const rows=[]; const lab=n=>(n<10?'0':'')+n+' '+MESES[n-1];
  const push=(prog,dep,mun,um,m,fF,fN,vF,vN)=>rows.push([prog,dep,mun,um,lab(m),m,{f:fF,v:vF},{f:fN,v:vN}]);
  const clean2=clean;
  const programa=(prog,hoja,first,last,colF0,step,um,nm,depCol,munCol)=>{
    const ws=wb.Sheets[hoja]; if(!ws) return; let dep=null;
    for(let r=first;r<=last;r++){ const d=vget(ws,r,depCol); if(d!=null&&String(d).trim()) dep=d; const mu=vget(ws,r,munCol); if(mu==null||!String(mu).trim()) continue;
      for(let m=1;m<=nm;m++){ const c=colF0+step*(m-1); push(prog,canonDep(dep),clean2(mu),um,m,"'"+hoja+"'!"+A(r,c),"'"+hoja+"'!"+A(r,c+1),num(vget(ws,r,c)),num(vget(ws,r,c+1))); } }
  };
  const ult={};
  const lastData=(ws,from,col)=>{ let r=from; while(vget(ws,r+1,col)!=null&&String(vget(ws,r+1,col)).trim()&&!norm(vget(ws,r+1,0)).includes('TOTAL')) r++; return r; };
  const nm=ctx.mes+1;
  ['BECA MEDIA','BECA SUPERIOR'].forEach(h=>{ const ws=wb.Sheets[h]; if(!ws) return; const last=lastData(ws,3,2); const mUlt=(h==='BECA SUPERIOR'&&res.bes)?res.bes.mesUlt+1:(h==='BECA MEDIA'&&res.bm)?res.bm.mesUlt+1:Math.min(nm,7);
    // meses con dato real del hoja, no más allá del mes global
    let mx=0; for(let m=1;m<=12;m++){ for(let r=3;r<=last;r++){ if(num(vget(ws,r,6+3*(m-1)))>0||num(vget(ws,r,7+3*(m-1)))>0){ mx=m; break; } } }
    const nmes=Math.min(Math.max(mx,0)||0,12); ult[h]=nmes;
    programa(h==='BECA MEDIA'?'Beca Educación Media':'Beca Educación Superior',h,3,last,6,3,'Beca',nmes,1,2); });
  { const ws=wb.Sheets['JOVENES PROTAGONISTAS']; ult.jp=7; programa('Jóvenes Protagonistas','JOVENES PROTAGONISTAS',4,131,4,2,'Persona',7,1,2); }
  { const ws=wb.Sheets['COMEDORES']; let dep=null; const c=res.comedores; const nmes=c?c.mesUlt+1:nm; ult.com=nmes; const first=4, last=c?c.pr.last:358;
    for(let r=first;r<=last;r+=2){ const d=vget(ws,r,1); if(d!=null&&String(d).trim()) dep=d; for(const rr of [r,r+1]){ const mu=clean(vget(ws,rr,2)); const um=norm(vget(ws,rr,3)).startsWith('DES')?'Desayuno':'Almuerzo'; for(let m=1;m<=nmes;m++){ const cc=4+2*(m-1); push('Comedor Social',canonDep(dep),mu,um,m,'COMEDORES!'+A(rr,cc),'COMEDORES!'+A(rr,cc+1),num(vget(ws,rr,cc)),num(vget(ws,rr,cc+1))); } } } }
  if(res.bolsa){ const b=res.bolsa, ws=wb.Sheets['BOLSA']; ult.bolsa=b.mesUlt+1;
    b.dat.forEach(d=>{ for(let m=0;m<=b.mesUlt;m++){ const cols=Object.entries(b.mesCol).filter(([c,mm])=>mm===m&&+c>=6).map(([c])=>+c).sort((x,y)=>x-y); if(!cols.length) continue; const rng=cols.length===1?A(d.i,cols[0]):A(d.i,cols[0])+':'+A(d.i,cols[cols.length-1]); const s=cols.reduce((a,c)=>a+num(vget(ws,d.i,c)),0); const fF='SUM(BOLSA!'+rng+')', fN='SUM(BOLSA!'+rng+')*'+b.price; rows.push(['Bolsa Social',d.dep,d.mun,'Aporte',lab(m+1),m+1,{f:fF,v:s},{f:fN,v:r2(s*b.price)}]); } }); }
  const ws={}; const hdr=['Programa','Departamento','Municipio','Unidad de Medida','Mes','No. Mes','Física','Financiera'];
  hdr.forEach((t,c)=>cset(ws,0,c,t));
  rows.forEach((r,i)=>{ for(let c=0;c<6;c++) cset(ws,i+1,c,r[c]); cset(ws,i+1,6,r[6].v,{f:r[6].f,z:'#,##0.00'}); cset(ws,i+1,7,r[7].v,{f:r[7].f,z:'#,##0.00'}); });
  ws['!cols']=[21.2,17,40.1,18.9,14.6,10.2,12,14].map(w=>({wch:w}));
  ws['!autofilter']={ref:'A1:H'+(rows.length+1)};
  fixRef(ws); ctx.baseRows=rows; ctx.ultMes=ult;
  wb.Sheets['BASE CONSOLIDADA']=ws; if(!wb.SheetNames.includes('BASE CONSOLIDADA')) wb.SheetNames.push('BASE CONSOLIDADA');
  // duplicados
  const seen=new Set(); let dup=0; rows.forEach(r=>{ const k=[r[0],r[1],r[2],r[3],r[5]].join('|'); if(seen.has(k)) dup++; seen.add(k); });
  if(dup) ctx.avisos.push({nivel:'aviso',prog:'base',cod:'duplicados',msg:'BASE CONSOLIDADA tiene '+dup+' fila(s) repetida(s) (mismo programa, departamento, municipio, unidad y mes). Suele ser un nombre de municipio repetido.'});
}
function construirResumen(wb, ctx){
  const wi=wb.Sheets['INTEGRADO']; const ws={}; const f=ctx.filasInt; const ord=['comedores','bolsa','jp','bm','bes'];
  cset(ws,0,0,'MINISTERIO DE DESARROLLO SOCIAL  |  DIRECCIÓN DE PREVENCIÓN SOCIAL');
  cset(ws,1,0,'RESUMEN EJECUTIVO DE EJECUCIÓN FÍSICA Y FINANCIERA — '+MESES_U[ctx.mes]+' 2026');
  ['Programa Social','Presupuesto Vigente (Q)','Ejecución Financiera (Q)','% Ejec. Financiera','Meta Física Vigente','Ejecución Física','% Ejec. Física'].forEach((t,c)=>cset(ws,3,c,t));
  let sB=0,sC=0,sE=0,sF=0;
  ord.forEach((k,j)=>{ const r=f[k]; if(r==null) return; const rr=4+j, ir=r+1; const B=num(vget(wi,r,CI('N'))), C=num(vget(wi,r,CI('O'))), E=num(vget(wi,r,CI('K'))), F=num(vget(wi,r,CI('L')));
    cset(ws,rr,0,vget(wi,r,CI('D')),{f:'INTEGRADO!D'+ir}); cset(ws,rr,1,B,{f:'INTEGRADO!N'+ir,z:'#,##0.00'}); cset(ws,rr,2,C,{f:'INTEGRADO!O'+ir,z:'#,##0.00'}); cset(ws,rr,3,B?C/B:0,{f:'INTEGRADO!O'+ir+'/INTEGRADO!N'+ir,z:'0.0%'}); cset(ws,rr,4,E,{f:'INTEGRADO!K'+ir,z:'#,##0'}); cset(ws,rr,5,F,{f:'INTEGRADO!L'+ir,z:'#,##0'}); cset(ws,rr,6,E?F/E:0,{f:'INTEGRADO!L'+ir+'/INTEGRADO!K'+ir,z:'0.0%'});
    sB+=B; sC+=C; sE+=E; sF+=F; });
  const n=ord.filter(k=>f[k]!=null).length, tr=4+n;
  cset(ws,tr,0,'TOTAL DPS'); cset(ws,tr,1,sB,{f:'SUM(B5:B'+(4+n)+')',z:'#,##0.00'}); cset(ws,tr,2,sC,{f:'SUM(C5:C'+(4+n)+')',z:'#,##0.00'}); cset(ws,tr,3,sB?sC/sB:0,{f:'C'+(tr+1)+'/B'+(tr+1),z:'0.0%'}); cset(ws,tr,4,sE,{f:'SUM(E5:E'+(4+n)+')',z:'#,##0'}); cset(ws,tr,5,sF,{f:'SUM(F5:F'+(4+n)+')',z:'#,##0'}); cset(ws,tr,6,sE?sF/sE:0,{f:'F'+(tr+1)+'/E'+(tr+1),z:'0.0%'});
  cset(ws,tr+1,0,'Nota: la meta y la ejecución física suman unidades distintas (raciones, aportes, becas, personas); solo los porcentajes por programa son comparables.');
  // bloque mensual con SUMIFS sobre BASE CONSOLIDADA
  const r0=tr+4; cset(ws,r0-1,0,'EJECUCIÓN MENSUAL POR PROGRAMA (sumas sobre BASE CONSOLIDADA)');
  cset(ws,r0,0,'Programa'); cset(ws,r0,1,'Concepto'); for(let m=1;m<=ctx.mes+1;m++){ cset(ws,r0,1+m,MESES[m-1]); cset(ws,r0+1,1+m,m); }
  cset(ws,r0+1,0,'(No. de mes)');
  const progs=['Comedor Social','Bolsa Social','Jóvenes Protagonistas','Beca Educación Media','Beca Educación Superior'];
  const agg={}; ctx.baseRows.forEach(r=>{ const k=r[0]+'|'+r[5]; const a=agg[k]=agg[k]||{f:0,n:0}; a.f+=r[6].v; a.n+=r[7].v; });
  const ultP={'Comedor Social':ctx.ultMes.com,'Bolsa Social':ctx.ultMes.bolsa,'Jóvenes Protagonistas':ctx.ultMes.jp,'Beca Educación Media':ctx.ultMes['BECA MEDIA'],'Beca Educación Superior':ctx.ultMes['BECA SUPERIOR']};
  let rr=r0+2;
  progs.forEach(p=>{ [['Física',6,'G'],['Financiera (Q)',7,'H']].forEach(([lbl,idx,col])=>{ cset(ws,rr,0,p); cset(ws,rr,1,lbl);
    for(let m=1;m<=ctx.mes+1;m++){ if(m>(ultP[p]||0)) continue; const a=agg[p+'|'+m]||{f:0,n:0}; const cl=colL(1+m); cset(ws,rr,1+m,r2(idx===6?a.f:a.n),{f:"SUMIFS('BASE CONSOLIDADA'!$"+col+":$"+col+",'BASE CONSOLIDADA'!$A:$A,$A"+(rr+1)+",'BASE CONSOLIDADA'!$F:$F,"+cl+"$"+(r0+2)+")",z:'#,##0.00'}); }
    rr++; }); });
  cset(ws,rr+1,0,'Las celdas vacías significan que el programa no reporta ese mes en este archivo.');
  ws['!cols']=[26,22,16,16,16,16,16,16,16,16,16,16].map(w=>({wch:w}));
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:6}},{s:{r:1,c:0},e:{r:1,c:6}}];
  fixRef(ws); wb.Sheets['RESUMEN EJECUTIVO']=ws; if(!wb.SheetNames.includes('RESUMEN EJECUTIVO')) wb.SheetNames.unshift('RESUMEN EJECUTIVO');
}
function construirNotas(wb, ctx){
  const ws={}; let r=0;
  const L=(a,b,c,d)=>{ cset(ws,r,0,a); if(b!==undefined) cset(ws,r,1,b); if(c!==undefined) cset(ws,r,2,c); if(d!==undefined) cset(ws,r,3,d); r++; };
  L('NOTAS DE CONSOLIDACIÓN — generado automáticamente'); L('Corte',MESES[ctx.mes]+' 2026'); L('Generado',new Date().toISOString().slice(0,16).replace('T',' ')+' UTC'); r++;
  L('Cifras de INTEGRADO que cambiaron'); L('Programa','Celda','Antes','Después','Origen');
  ctx.cambios.forEach(c=>{ cset(ws,r,0,c.prog); cset(ws,r,1,c.celda); cset(ws,r,2,c.antes==null?'':c.antes); cset(ws,r,3,c.despues); cset(ws,r,4,c.origen+(c.estado==='manual'?' (manual)':'')); r++; }); r++;
  L('Pendientes de la persona administradora'); ctx.pendientes.forEach(p=>L(p.prog,p.txt)); r++;
  L('Avisos de validación'); ctx.avisos.forEach(a=>L(a.nivel.toUpperCase(),a.prog,a.msg));
  ws['!cols']=[{wch:22},{wch:14},{wch:70},{wch:22},{wch:50}]; fixRef(ws); wb.Sheets['NOTAS DE CONSOLIDACIÓN']=ws; if(!wb.SheetNames.includes('NOTAS DE CONSOLIDACIÓN')) wb.SheetNames.push('NOTAS DE CONSOLIDACIÓN');
}

/* ---------- orquestación ---------- */
function newCtx(){
  const ctx={avisos:[],cambios:[],pendientes:[],programas:[],mes:0};
  ctx.prog=(id,nombre)=>{ const P={id,nombre,fail:false,kvs:[],estado:'ok',
    error(cod,msg){ P.fail=true; P.estado='error'; ctx.avisos.push({nivel:'error',prog:id,cod,msg}); },
    aviso(cod,msg){ ctx.avisos.push({nivel:'aviso',prog:id,cod,msg}); },
    info(cod,msg){ ctx.avisos.push({nivel:'info',prog:id,cod,msg}); },
    kv(k,v){ P.kvs.push([k,v]); } };
    ctx.programas.push(P); return P; };
  return ctx;
}
/* archivos: [{nombre, wb, ids}] ; baseWb: libro dps anterior (SheetJS) ; opciones.manual: {comedores:{E,F,G,K,N}, ...} */
function consolidar(archivos, baseWb, opciones){
  opciones=opciones||{}; const ctx=newCtx();
  // clonar la base (copia profunda de celdas)
  const wb={SheetNames:baseWb.SheetNames.slice(),Sheets:{}};
  baseWb.SheetNames.forEach(n=>{ const s=baseWb.Sheets[n], o={}; Object.keys(s).forEach(k=>{ o[k]=k[0]==='!'?s[k]:copyCell(s[k]); }); if(s['!merges']) o['!merges']=JSON.parse(JSON.stringify(s['!merges'])); wb.Sheets[n]=o; });
  const pick=t=>{ for(const f of archivos){ const id=(f.ids||[]).find(x=>x.tipo===t); if(id) return {wb:f.wb,ws:f.wb.Sheets[id.hoja],hoja:id.hoja,nombre:f.nombre}; } return null; };
  const raws={comedores:pick('comedores'),bes:pick('bes'),bm:pick('bm'),jp:pick('jp'),tmca:pick('tmca')};
  const res={};
  if(raws.comedores){ const r=construirComedores(raws.comedores,baseWb,ctx); if(r){ wb.Sheets['COMEDORES']=r.ws; res.comedores=r; } }
  if(raws.bes){ const r=construirBeca(raws.bes,baseWb,ctx,{id:'bes',nombre:'Beca Educación Superior',hoja:'BECA SUPERIOR',tarifa:2500}); if(r){ wb.Sheets['BECA SUPERIOR']=r.ws; res.bes=r; } }
  if(raws.bm){ const r=construirBeca(raws.bm,baseWb,ctx,{id:'bm',nombre:'Beca Educación Media',hoja:'BECA MEDIA',tarifa:2500}); if(r){ wb.Sheets['BECA MEDIA']=r.ws; res.bm=r; } }
  if(raws.jp){ const r=construirJP(raws.jp,baseWb,ctx); if(r){ wb.Sheets['JOVENES PROTAGONISTAS']=r.ws; res.jp=r; } }
  if(raws.tmca){ const r=construirBolsa(raws.tmca,baseWb,ctx); if(r){ wb.Sheets['BOLSA']=r.ws; res.bolsa=r; } }
  const faltan=['comedores','bes','jp','tmca'].filter(k=>!raws[k]);
  const nombres={comedores:'Comedor Social',bes:'Beca Educación Superior',jp:'Jóvenes Protagonistas',tmca:'Bolsa Social (TMCA)'};
  faltan.forEach(k=>{ const P=ctx.prog(k,nombres[k]); P.estado='conservado'; ctx.avisos.push({nivel:'aviso',prog:k,cod:'no_viene',msg:'No vino en el ZIP: se conserva '+nombres[k]+' del dps anterior.'}); });
  if(!raws.bm){ const P=ctx.prog('bm','Beca Educación Media'); P.estado='conservado'; }
  if(!Object.keys(res).length){ return {ok:false,ctx,wb:null,res,error:'Ningún archivo del ZIP se pudo procesar.'}; }
  // programas con error → conservados
  ctx.programas.forEach(P=>{ if(P.fail) P.estado='error'; });
  actualizarIntegrado(wb,res,ctx,opciones.manual);
  construirBase(wb,res,ctx);
  construirResumen(wb,ctx);
  construirNotas(wb,ctx);
  // orden de hojas
  const orden=['RESUMEN EJECUTIVO','INTEGRADO','BECA MEDIA','BECA SUPERIOR','JOVENES PROTAGONISTAS','BOLSA','COMEDORES','BASE CONSOLIDADA','NOTAS DE CONSOLIDACIÓN'];
  wb.SheetNames=orden.filter(n=>wb.Sheets[n]).concat(wb.SheetNames.filter(n=>!orden.includes(n)));
  const errores=ctx.avisos.filter(a=>a.nivel==='error');
  return {ok:true,ctx,wb,res,errores,mes:ctx.mes,mesNombre:MESES[ctx.mes]};
}
function escribir(wb){ return XLSX.write(wb,{bookType:'xlsx',type:'array',cellStyles:true,compression:true}); }

const API={consolidar,leerZip,identificarLibro,escribir,MESES,norm,parseComedores};
if(typeof module!=='undefined'&&module.exports) module.exports=API; else root.Consolidador=API;
})(typeof window!=='undefined'?window:globalThis);
