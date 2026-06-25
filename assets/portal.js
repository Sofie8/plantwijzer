'use strict';
const $=(s)=>document.querySelector(s), $$=(s)=>[...document.querySelectorAll(s)];
const IMG='https://pub-bb204453b9b642598d8514f7ac4f68be.r2.dev/images';
const S={config:null,plants:[],nativeLoaded:false,fytoRows:[],saved:new Map(),expanded:{flowers:false,woody:false,fyto:false},current:null,detailRows:[]};

const norm=v=>String(v??'').replace(/\u00a0/g,' ').trim().replace(/\s+/g,' ');
const low=v=>norm(v).toLowerCase();
const latin=r=>norm(r['Latijnse naam']||r['Latijnse naam ']||r['latijnse naam']||r['Scientific name']||r['Plant Scientific Name']||r['Soort']);
const dutch=r=>norm(r['Nederlandse naam']||r['Nederlandse naam ']||r['nederlandse naam']||r['Plant Common Name']||r['Common name']);
const xlsxPath=r=>r.endsWith('.xlsx')?r:r+'.xlsx';
async function workbook(path){const res=await fetch(path);if(!res.ok)throw new Error(`${path} (${res.status})`);return XLSX.read(await res.arrayBuffer(),{type:'array'})}
const rows=(wb,sheet=wb.SheetNames[0])=>XLSX.utils.sheet_to_json(wb.Sheets[sheet],{defval:''});
function bestSheet(wb){for(const n of wb.SheetNames){const h=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:''})[0]||[];if(h.map(low).some(x=>['latijnse naam','nederlandse naam','plant scientific name'].includes(x)))return n}return wb.SheetNames[0]}
function fytoSheet(wb,medium,pollutant){const target=`${medium}_${pollutant}`.toLowerCase();return wb.SheetNames.find(x=>x.toLowerCase()===target)||bestSheet(wb)}
function fotoIds(r){
  const raw=norm(r.foto_ids||r.foto_id||r.foto||r.image_paths||'');
  if(!raw)return[];
  return raw.split('|').map(x=>norm(x).replace(/^images[\\/]/i,'').replace(/\.(jpg|jpeg|png|webp)$/i,'')).filter(Boolean);
}
function parseNumber(v){const m=norm(v).replace(',','.').match(/-?\d+(\.\d+)?/);return m?Number(m[0]):0}
function cleanPostcode(v){
  if(v===undefined||v===null||v==='')return'';
  const t=norm(v);
  if(!t)return'';
  const n=Number(t.replace(',','.'));
  if(Number.isFinite(n))return String(Math.trunc(n));
  return t.replace(/\.0$/,'');
}
function postcodesFromValue(v){
  const t=norm(v);
  if(!t)return[];
  return t.split(/[|,;]/).map(cleanPostcode).filter(Boolean);
}
function postcodeMeta(r){
  const direct=postcodesFromValue(r.postcode||r.Postcode||r.postcodes||r.Postcodes||r.kenmerk_postcode);
  const from=cleanPostcode(r.postcode_van||r['postcode van']||r.PostcodeVan||r.Postcode_van);
  const to=cleanPostcode(r.postcode_tot||r['postcode tot']||r.PostcodeTot||r.Postcode_tot);
  return {direct,from:from?parseInt(from,10):null,to:to?parseInt(to,10):null};
}
function postcodeMatch(r,pcRaw){
  const pc=cleanPostcode(pcRaw), n=parseInt(pc,10), m=postcodeMeta(r);
  if(m.direct.length){
    return m.direct.some(x=>x.endsWith('*')?pc.startsWith(x.slice(0,-1)):x===pc);
  }
  if(m.from!==null&&m.to!==null&&Number.isFinite(n))return n>=m.from&&n<=m.to;
  return null;
}
function rawBiodiversity(r){let sum=0;for(const [k,v] of Object.entries(r)){const key=low(k);if(/vogel|vlinder|mot|bij|gastheer|biodivers|keystone|insect/.test(key))sum+=parseNumber(v)}return sum}
function category(typ){return /Bomen|Haagplanten/.test(typ)?'woody':/bloemenweide|inheemse_planten|grasland_weide|wadi/.test(typ)?'flowers':'other'}
function baseScore(p){return Math.round((p.amber?30:0)+(p.regional?25:0)+Math.min(p.bio,45))}
function uniquePlants(arr){const m=new Map();for(const p of arr){const k=low(p.latin)||`nl:${low(p.dutch)}`;if(!m.has(k))m.set(k,p);else{const e=m.get(k);e.amber||=p.amber;e.regional||=p.regional;e.bio=Math.max(e.bio,p.bio);e.score=baseScore(e);Object.assign(e.raw,Object.fromEntries(Object.entries(p.raw).filter(([k,v])=>!e.raw[k]&&v)))}}return[...m.values()]}

async function loadIndex(refs){
  const map=new Set();
  for(const ref of refs||[]){
    const finalRef=ref.includes('/')?ref:`regionale_soortenlijst/${ref}`;
    try{const wb=await workbook(`data/layers/${xlsxPath(finalRef)}`);for(const r of rows(wb,bestSheet(wb))){if(latin(r))map.add(low(latin(r)));if(dutch(r))map.add(`nl:${low(dutch(r))}`)}}catch(e){console.warn(e)}
  }
  return map;
}
async function loadNative(){
  if(S.nativeLoaded)return;
  $('#nativeStatus').textContent='Planten laden…';
  const amber=await loadIndex(S.config.layers.amberlijst||[]);
  const regional=await loadIndex(S.config.layers.regionale_soortenlijst||[]);
  const out=[];
  for(const [typ,node] of Object.entries(S.config.typologies)){
    for(const [sub,refs] of Object.entries(node.subtypes||{})){
      for(const ref of refs){
        try{
          const wb=await workbook(`data/typologies/${xlsxPath(ref)}`);
          for(const r of rows(wb,bestSheet(wb))){
            const la=latin(r);if(!la)continue;
            const nl=dutch(r), key=low(la);
            const p={latin:la,dutch:nl,typ,sub,raw:r,photos:fotoIds(r),amber:amber.has(key)||amber.has(`nl:${low(nl)}`),regional:regional.has(key)||regional.has(`nl:${low(nl)}`),bio:rawBiodiversity(r)};
            p.score=baseScore(p);out.push(p);
          }
        }catch(e){console.warn('Typologie niet geladen',ref,e)}
      }
    }
  }
  S.plants=uniquePlants(out).sort((a,b)=>b.score-a.score||a.dutch.localeCompare(b.dutch));
  S.nativeLoaded=true;
}
function firstImage(p,host){
  if(!host||!p.photos.length)return;
  const tryExt=(id,ext)=>{const im=new Image();im.onload=()=>{host.innerHTML='';host.append(im)};im.onerror=()=>{if(ext==='jpg')tryExt(id,'png')};im.src=`${IMG}/${id}.${ext}`};
  tryExt(p.photos[0],'jpg');
}
function isSaved(p){return S.saved.has(low(p.latin))}
function card(p,mode='native'){
  const el=document.createElement('article');el.className='plantCard';
  el.innerHTML=`<button class="heart ${isSaved(p)?'saved':''}" aria-label="bewaar">${isSaved(p)?'♥':'♡'}</button><div class="cardImage">${mode==='fyto'?'🧪':category(p.typ)==='woody'?'🌳':'🌼'}</div><div class="cardBody"><h4>${p.dutch||p.latin}</h4><div class="latin">${p.latin}</div>${mode==='native'?`<div class="scoreLine"><span class="rankBadge">${p.score}</span><span>lokale & ecologische score</span></div><div class="tagRow">${p.amber?'<span class="tag">AMBER</span>':''}${p.regional?'<span class="tag">streekeigen</span>':''}${p.bio?`<span class="tag">biodiversiteit ${p.bio}</span>`:''}</div>`:`<div class="tagRow"><span class="tag">${p.pollutantLabel}</span><span class="tag">${p.mediumLabel}</span><span class="tag">${p.typ.replace(/^\d+\./,'')}</span></div>`}</div>`;
  firstImage(p,el.querySelector('.cardImage'));
  el.addEventListener('click',e=>{if(!e.target.closest('.heart'))openDrawer(p,mode)});
  el.querySelector('.heart').onclick=e=>{e.stopPropagation();toggleSave(p)};
  return el;
}
function renderNative(){
  const pc=cleanPostcode($('#postcode').value);
  const hasMeta=S.plants.some(p=>{const m=postcodeMeta(p.raw);return m.direct.length||m.from!==null});
  let selection=S.plants;
  if(pc&&hasMeta)selection=selection.filter(p=>postcodeMatch(p.raw,pc)===true);
  $('#nativeStatus').textContent = pc ? (hasMeta?`${selection.length} planten gevonden voor postcode ${pc}.`:`${selection.length} planten geladen. Er zijn nog geen postcodevelden gevonden in de geladen Excelbestanden.`) : 'Vul een postcode in.';
  for(const group of ['flowers','woody']){
    const list=selection.filter(p=>category(p.typ)===group);
    const limit=S.expanded[group]?list.length:8, grid=$(`#${group}Grid`);grid.innerHTML='';
    list.slice(0,limit).forEach(p=>grid.append(card(p)));
    const b=$(`[data-more="${group}"]`);b.hidden=list.length<=8;b.textContent=S.expanded[group]?'Toon minder':'Meer planten';
  }
}
function toggleSave(p){const k=low(p.latin);if(S.saved.has(k))S.saved.delete(k);else S.saved.set(k,p);localStorage.setItem('plantwijzer_saved',JSON.stringify([...S.saved.values()]));updateSaved();if(S.nativeLoaded)renderNative();if(S.fytoRows.length)renderFyto()}
function updateSaved(){$('#savedCount').textContent=S.saved.size;const g=$('#savedGrid');if(!g)return;g.innerHTML='';[...S.saved.values()].forEach(p=>g.append(card(p,p.fytoRow?'fyto':'native')))}
function setView(v){$$('.view').forEach(x=>x.classList.toggle('active',v==='home'?x.id==='home':x.id===`view-${v}`));window.scrollTo({top:0,behavior:'smooth'});if(v==='saved')updateSaved()}
async function loadFyto(){
  $('#fytoStatus').textContent='Fytoremediatiegegevens laden…';
  const typ=$('#fytoType').value, medium=$('#fytoMedium').value, pollutant=$('#fytoPollutant').value;
  const nameQ=low($('#fytoNameSearch').value), allQ=low($('#fytoSearchAll').value);
  const out=[];
  for(const [t,refs] of Object.entries(S.config.layers.fytoremediatie||{})){
    if(typ!=='ALL'&&typ!==t)continue;
    for(const ref of refs){
      try{
        const wb=await workbook(`data/layers/${xlsxPath(ref)}`), sh=fytoSheet(wb,medium,pollutant);
        for(const r of rows(wb,sh)){
          const la=latin(r);if(!la)continue;const nl=dutch(r), blob=low(Object.values(r).join(' '));
          if(nameQ && !(low(la).includes(nameQ)||low(nl).includes(nameQ)))continue;
          if(allQ && !blob.includes(allQ))continue;
          out.push({latin:la,dutch:nl,typ:t,sub:'',raw:r,photos:fotoIds(r),fytoRow:r,pollutant,medium,pollutantLabel:$('#fytoPollutant').selectedOptions[0].text,mediumLabel:$('#fytoMedium').selectedOptions[0].text,score:0,bio:0});
        }
      }catch(e){console.warn(e)}
    }
  }
  S.fytoRows=uniquePlants(out);S.expanded.fyto=false;renderFyto();
}
function renderFyto(){const g=$('#fytoGrid');g.innerHTML='';const list=S.fytoRows, limit=S.expanded.fyto?list.length:8;list.slice(0,limit).forEach(p=>g.append(card(p,'fyto')));$('#fytoStatus').textContent=`${list.length} planten gevonden.`;$('#fytoMore').hidden=list.length<=8;$('#fytoMore').textContent=S.expanded.fyto?'Toon minder':'Meer resultaten'}
function rawTable(r){return Object.entries(r||{}).filter(([k,v])=>norm(k)&&norm(v)).sort((a,b)=>a[0].localeCompare(b[0])).map(([k,v])=>`<b>${k}</b><span>${norm(v)}</span>`).join('')}
async function detailRowsFor(p){if(!p.fytoRow)return[];const base=(S.config.layers.fytoremediatie[p.typ]?.[0]||'').split('/').pop();const path=`data/layers/fytoremediatie/detail/${base}_${p.medium}_${p.pollutant}_detail.xlsx`;try{const wb=await workbook(path);return rows(wb,bestSheet(wb)).filter(r=>low(latin(r))===low(p.latin)||low(dutch(r))===low(p.dutch))}catch(e){console.warn(e);return[]}}
function renderDetailTable(list){S.detailRows=list;const q=low($('#detailSearch').value), filtered=q?list.filter(r=>low(Object.values(r).join(' ')).includes(q)):list;$('#detailMeta').textContent=`${filtered.length} / ${list.length} studies`;const h=filtered.length?[...new Set(filtered.flatMap(Object.keys))]:[];$('#detailTable thead').innerHTML=h.length?`<tr>${h.map(x=>`<th>${x}</th>`).join('')}</tr>`:'';$('#detailTable tbody').innerHTML=filtered.map(r=>`<tr>${h.map(x=>`<td>${norm(r[x])||'—'}</td>`).join('')}</tr>`).join('')}
function csv(rows){if(!rows.length)return'';const h=[...new Set(rows.flatMap(Object.keys))],q=v=>`"${String(v??'').replaceAll('"','""')}"`;return[h.map(q).join(','),...rows.map(r=>h.map(x=>q(r[x])).join(','))].join('\n')}
function download(name,text){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type:'text/csv'}));a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),200)}
async function openDrawer(p,mode){S.current=p;$('#drawerTitle').textContent=p.dutch||p.latin;$('#drawerSubtitle').textContent=p.latin;$('#drawerKicker').textContent=p.typ.replace(/^\d+\./,'');$('#drawerScore').innerHTML=mode==='native'?`<strong>Score ${p.score}</strong><br>${p.amber?'AMBER · ':''}${p.regional?'streekeigen · ':''}${p.bio?`biodiversiteit ${p.bio}`:''}`:'Onderzoeksselectie voor zuiverende werking';$('#rawDetails').innerHTML=rawTable(p.raw);$('#drawerImages').innerHTML=category(p.typ)==='woody'?'🌳':'🌼';firstImage(p,$('#drawerImages'));const fy=$('#drawerFyto');fy.hidden=!p.fytoRow;if(p.fytoRow){$('#fytoSummary').innerHTML=rawTable(p.fytoRow);$('#detailMeta').textContent='Detailstudies laden…';renderDetailTable(await detailRowsFor(p))}$('#savePlant').textContent=isSaved(p)?'♥ Verwijder uit mijn lijst':'♡ Bewaar deze plant';$('#plantDrawer').classList.add('open');$('#plantDrawer').setAttribute('aria-hidden','false')}
async function init(){
  S.config=await fetch('data/config.json').then(r=>{if(!r.ok)throw new Error('data/config.json niet gevonden');return r.json()});
  try{JSON.parse(localStorage.getItem('plantwijzer_saved')||'[]').forEach(p=>S.saved.set(low(p.latin),p))}catch{}
  updateSaved();
  const typeSel=$('#fytoType');
  for(const typ of Object.keys(S.config.layers.fytoremediatie||{})){const o=document.createElement('option');o.value=typ;o.textContent=typ.replace(/^\d+\./,'');typeSel.append(o)}
  if(typeSel.options.length<=1){$('#fytoStatus').textContent='Geen fytoremediatiebestanden gevonden in data/config.json.'}
  $$('[data-link]').forEach(a=>{const u=S.config.portal?.links?.[a.dataset.link]||'#';a.href=u;a.target=u==='#'?'_self':'_blank';a.rel='noopener'});
  $$('[data-view]').forEach(b=>b.onclick=()=>setView(b.dataset.view));
  $('#postcodeForm').onsubmit=async e=>{e.preventDefault();await loadNative();renderNative();setView('native')};
  $$('.moreBtn').forEach(b=>b.onclick=()=>{S.expanded[b.dataset.more]=!S.expanded[b.dataset.more];renderNative()});
  $('#runFyto').onclick=loadFyto;$('#fytoMore').onclick=()=>{S.expanded.fyto=!S.expanded.fyto;renderFyto()};
  $$('[data-goto]').forEach(b=>b.onclick=()=>setView(b.dataset.goto));
  $('#drawerClose').onclick=()=>$('#plantDrawer').classList.remove('open');
  $('#savePlant').onclick=()=>{if(S.current)toggleSave(S.current);$('#savePlant').textContent=isSaved(S.current)?'♥ Verwijder uit mijn lijst':'♡ Bewaar deze plant'};
  $('#detailSearch').oninput=()=>renderDetailTable(S.detailRows);
  $('#detailCsv').onclick=()=>download(`fyto_detail_${S.current?.latin||'plant'}.csv`,csv(S.detailRows));
  $('#extendedCsv').onclick=()=>download(`extended_${S.current?.latin||'plant'}.csv`,csv([Object.assign({},S.current?.raw||{},S.current?.fytoRow||{})])+'\n\nFYTO_DETAIL\n'+csv(S.detailRows));
  $('#clearSaved').onclick=()=>{S.saved.clear();localStorage.removeItem('plantwijzer_saved');updateSaved()};
  $('#tourBtn').onclick=()=>$('#tourModal').hidden=false;$('#tourClose').onclick=()=>$('#tourModal').hidden=true;
}
window.addEventListener('DOMContentLoaded',()=>init().catch(e=>{console.error(e);document.body.insertAdjacentHTML('afterbegin',`<div style="padding:12px;background:#fee;color:#900">${e.message}</div>`)}));
