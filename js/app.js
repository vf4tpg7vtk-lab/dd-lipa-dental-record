// ── Constants ────────────────────────────────────────────────────────────────
const UPPER_TEETH = [18,17,16,15,14,13,12,11,21,22,23,24,25,26,27,28];
const LOWER_TEETH = [48,47,46,45,44,43,42,41,31,32,33,34,35,36,37,38];
const LOCAL_KEY   = 'dentarecord_patients';
const TOKEN_KEY   = 'dentarecord_od_token';
const OD_FOLDER   = 'DentaRecords';

const CONDITIONS = {
  healthy:   { color: '#ffffff', label: 'Healthy'   },
  cavity:    { color: '#e74c3c', label: 'Cavity'    },
  filled:    { color: '#3498db', label: 'Filled'    },
  crown:     { color: '#f39c12', label: 'Crown'     },
  missing:   { color: '#95a5a6', label: 'Missing'   },
  extracted: { color: '#2c3e50', label: 'Extracted' },
  implant:   { color: '#27ae60', label: 'Implant'   },
};

// ── LocalStorage ─────────────────────────────────────────────────────────────
const LS = {
  all:    () => { try { return JSON.parse(localStorage.getItem(LOCAL_KEY)||'{}'); } catch { return {}; } },
  save:   p  => { const a=LS.all(); a[p.id]={...p,updatedAt:new Date().toISOString()}; localStorage.setItem(LOCAL_KEY,JSON.stringify(a)); },
  del:    id => { const a=LS.all(); delete a[id]; localStorage.setItem(LOCAL_KEY,JSON.stringify(a)); },
  list:   () => Object.values(LS.all()).sort((a,b)=>new Date(b.updatedAt)-new Date(a.updatedAt)),
  count:  () => Object.keys(LS.all()).length,
};

// ── OneDrive API ──────────────────────────────────────────────────────────────
async function odEnsureFolder(tok) {
  const r = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${OD_FOLDER}`,
    {headers:{Authorization:`Bearer ${tok}`}});
  if (!r.ok) await fetch('https://graph.microsoft.com/v1.0/me/drive/root/children',{
    method:'POST',
    headers:{Authorization:`Bearer ${tok}`,'Content-Type':'application/json'},
    body:JSON.stringify({name:OD_FOLDER,folder:{},'@microsoft.graph.conflictBehavior':'rename'})
  });
}
async function odSave(patient, tok) {
  await odEnsureFolder(tok);
  const fn = `dental_${patient.id}_${(patient.name||'unnamed').replace(/\s+/g,'_')}.json`;
  const r  = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${OD_FOLDER}/${fn}:/content`,
    {method:'PUT',headers:{Authorization:`Bearer ${tok}`,'Content-Type':'application/json'},body:JSON.stringify(patient,null,2)});
  if (!r.ok) throw new Error(`OneDrive save failed (${r.status})`);
  return r.json();
}
async function odList(tok) {
  await odEnsureFolder(tok);
  const r = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${OD_FOLDER}:/children?$select=id,name,lastModifiedDateTime`,
    {headers:{Authorization:`Bearer ${tok}`}});
  if (!r.ok) throw new Error('Failed to list OneDrive files');
  const d = await r.json();
  return (d.value||[]).filter(f=>f.name.startsWith('dental_'))
    .map(f=>({id:f.id,name:f.name,modifiedTime:f.lastModifiedDateTime}));
}
async function odLoad(fileId, tok) {
  const m  = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`,{headers:{Authorization:`Bearer ${tok}`}});
  if (!m.ok) throw new Error('Metadata error');
  const md = await m.json();
  const r  = await fetch(md['@microsoft.graph.downloadUrl']);
  if (!r.ok) throw new Error('Download failed');
  return r.json();
}
async function odDel(fileId, tok) {
  const r = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`,
    {method:'DELETE',headers:{Authorization:`Bearer ${tok}`}});
  if (!r.ok && r.status!==204) throw new Error('Delete failed');
}

// ── App State ─────────────────────────────────────────────────────────────────
const state = {
  view:         'home',
  listSource:   'local',
  patient:      newPatient(),
  selectedCond: 'cavity',
  token:        localStorage.getItem(TOKEN_KEY)||null,
  tokenInput:   '',
  odFiles:      [],
  localList:    LS.list(),
  search:       '',
  status:       null,
  autoSaved:    false,
  confirm:      null,
  autoTimer:    null,
  dirHandle:    null,
  dirName:      localStorage.getItem('dentarecord_dirname')||null,
};

function newPatient() {
  return {
    id: Date.now().toString(), name:'', address:'', telephone:'', age:'',
    occupation:'', status:'', complaint:'', teeth:{}, treatments:[],
    createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
  };
}
function newTreatment() {
  return {
    id: Date.now().toString(), date: new Date().toISOString().split('T')[0],
    toothNo:'', description:'', time:'', debit:'', creditDate:'', creditAmount:'', balance:''
  };
}

// ── Render helpers ────────────────────────────────────────────────────────────
function h(tag, attrs={}, ...children) {
  const el = document.createElement(tag);
  for (const [k,v] of Object.entries(attrs)) {
    if (k==='class') el.className=v;
    else if (k.startsWith('on')) el[k]=v;
    else if (k==='style'&&typeof v==='object') Object.assign(el.style,v);
    else el.setAttribute(k,v);
  }
  for (const c of children.flat()) {
    if (c==null||c===false) continue;
    el.appendChild(typeof c==='string'?document.createTextNode(c):c);
  }
  return el;
}
const svg = (content, w=22, h_=22, vb='0 0 24 24') => {
  const el = document.createElementNS('http://www.w3.org/2000/svg','svg');
  el.setAttribute('width',w); el.setAttribute('height',h_);
  el.setAttribute('viewBox',vb); el.setAttribute('fill','none');
  el.setAttribute('stroke','currentColor');
  el.setAttribute('stroke-width','1.8');
  el.setAttribute('stroke-linecap','round');
  el.setAttribute('stroke-linejoin','round');
  el.innerHTML=content;
  return el;
};

// ── Toast ─────────────────────────────────────────────────────────────────────
let toastTimer=null;
function toast(type, msg) {
  state.status={type,msg}; render();
  if (toastTimer) clearTimeout(toastTimer);
  if (type!=='loading') toastTimer=setTimeout(()=>{ state.status=null; render(); },3500);
}

// ── Auto-save: silent background save, never calls render() ──────────────────
function scheduleAutoSave() {
  if (state.autoTimer) clearTimeout(state.autoTimer);
  state.autoTimer=setTimeout(()=>{
    const p=state.patient;
    const isEmpty=!p.name&&!p.complaint&&!(p.treatments||[]).length&&!Object.keys(p.teeth||{}).length;
    if (isEmpty) return;
    LS.save(p); state.localList=LS.list();
    // Show badge by direct DOM — no render(), no focus disruption
    const old=document.getElementById('autosave-badge'); if(old) old.remove();
    const badge=document.createElement('div');
    badge.id='autosave-badge'; badge.className='autosave-badge'; badge.textContent='✓ Auto-saved';
    document.body.appendChild(badge);
    setTimeout(()=>{ const el=document.getElementById('autosave-badge'); if(el) el.remove(); },2000);
    // Silently update the status pill text
    const pill=document.querySelector('#record-status-bar .status-pill');
    if (pill) { pill.className='status-pill pill-saved'; pill.textContent='🖥️ Saved locally'; }
  },2000);
}

// ── Handlers ──────────────────────────────────────────────────────────────────
function setView(v, opts={}) {
  state.view=v; state.search='';
  if (v==='record'&&opts.fresh) state.patient=newPatient();
  render();
}

// Silent patch — for text inputs. No render, no lost focus.
function patchPatient(patch) {
  Object.assign(state.patient, patch);
  scheduleAutoSave();
}

// Full update + render — for structural changes only (teeth clicks, row add/remove).
function updatePatient(patch) {
  state.patient={...state.patient,...patch};
  scheduleAutoSave(); render();
}

// Update a single treatment cell silently — no render
function updateTreatment(i, key, value) {
  if (state.patient.treatments[i]) {
    state.patient.treatments[i][key]=value;
    scheduleAutoSave();
  }
}

function addTreatment() {
  state.patient.treatments=[...(state.patient.treatments||[]),newTreatment()];
  scheduleAutoSave();
  const wrap=document.getElementById('treatment-table-wrap');
  if (wrap) { wrap.innerHTML=''; wrap.appendChild(buildTreatmentTable()); }
}

function removeTreatment(i) {
  state.patient.treatments=state.patient.treatments.filter((_,idx)=>idx!==i);
  scheduleAutoSave();
  const wrap=document.getElementById('treatment-table-wrap');
  if (wrap) { wrap.innerHTML=''; wrap.appendChild(buildTreatmentTable()); }
}

function clickTooth(num) {
  const teeth={...state.patient.teeth};
  if (state.selectedCond==='healthy') delete teeth[num];
  else teeth[num]=state.selectedCond;
  updatePatient({teeth});
}

function saveLocal() {
  if (!state.patient.name&&!(state.patient.treatments||[]).length) return toast('error','Add a patient name first.');
  LS.save(state.patient); state.localList=LS.list();
  const pill=document.querySelector('#record-status-bar .status-pill');
  if (pill) { pill.className='status-pill pill-saved'; pill.textContent='🖥️ Saved locally'; }
  toast('success','Saved to browser storage ✓');
}

// File System Access API folder picker
const hasFSA = 'showDirectoryPicker' in window;

async function pickFolder() {
  if (!hasFSA) return toast('error','Folder picker needs Chrome or Edge on desktop. Files will go to Downloads.');
  try {
    state.dirHandle = await window.showDirectoryPicker({mode:'readwrite'});
    state.dirName   = state.dirHandle.name;
    localStorage.setItem('dentarecord_dirname', state.dirName);
    toast('success',`Save folder: ${state.dirName} ✓`);
    const el=document.getElementById('folder-setting');
    if (el) el.replaceWith(buildFolderSetting());
  } catch(e) { if (e.name!=='AbortError') toast('error','Could not access folder.'); }
}

async function saveToFolder(p=state.patient) {
  if (!state.dirHandle) { downloadFile(p); return; }
  try {
    const fname=`dental_${(p.name||'patient').replace(/\s+/g,'_')}_${p.id}.json`;
    const fh=await state.dirHandle.getFileHandle(fname,{create:true});
    const w=await fh.createWritable();
    await w.write(JSON.stringify(p,null,2)); await w.close();
    toast('success',`Saved to ${state.dirName}/${fname} ✓`);
  } catch { state.dirHandle=null; state.dirName=null; downloadFile(p); }
}

async function exportAllToFolder() {
  if (!state.dirHandle) { exportAll(); return; }
  try {
    const records=LS.list();
    for (const p of records) {
      const fname=`dental_${(p.name||'patient').replace(/\s+/g,'_')}_${p.id}.json`;
      const fh=await state.dirHandle.getFileHandle(fname,{create:true});
      const w=await fh.createWritable();
      await w.write(JSON.stringify(p,null,2)); await w.close();
    }
    toast('success',`Exported ${records.length} records to ${state.dirName} ✓`);
  } catch { exportAll(); }
}

function buildFolderSetting() {
  const wrap=h('div',{id:'folder-setting'});
  if (!hasFSA) {
    const note=h('p',{style:{fontSize:'12px',color:'var(--muted)',fontFamily:'var(--font-mono)',marginTop:'8px'}});
    note.textContent='📁 Folder picker requires Chrome or Edge on desktop. Files download to your Downloads folder.';
    wrap.appendChild(note); return wrap;
  }
  const row=h('div',{class:'row',style:{marginTop:'10px',gap:'10px',flexWrap:'wrap'}});
  if (state.dirName) {
    const info=h('div',{class:'connected-row',style:{flex:'1',background:'rgba(26,122,62,0.08)',padding:'8px 12px',borderRadius:'8px',border:'1px solid rgba(26,122,62,0.2)'}});
    const txt=h('span',{style:{fontFamily:'var(--font-mono)',fontSize:'12px',color:'var(--green)',fontWeight:'600',marginLeft:'4px'}});
    txt.textContent=`📁 Saving to: ${state.dirName}/`;
    const change=h('button',{class:'btn btn-outline btn-sm',onclick:pickFolder,style:{marginLeft:'auto'}});
    change.textContent='Change'; info.appendChild(txt); info.appendChild(change);
    row.appendChild(info);
  } else {
    const btn=h('button',{class:'btn btn-outline',onclick:pickFolder});
    btn.textContent='📁 Choose Save Folder';
    const note=h('span',{style:{fontSize:'11px',color:'var(--muted)',fontFamily:'var(--font-mono)'}});
    note.textContent='Save files directly to a folder on your computer';
    row.appendChild(btn); row.appendChild(note);
  }
  wrap.appendChild(row); return wrap;
}
  const blob=new Blob([JSON.stringify(p,null,2)],{type:'application/json'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url;
  a.download=`dental_${(p.name||'patient').replace(/\s+/g,'_')}_${p.id}.json`;
  a.click(); URL.revokeObjectURL(url); toast('success','File downloaded ✓');
}

function exportAll() {
  const blob=new Blob([JSON.stringify(LS.all(),null,2)],{type:'application/json'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url;
  a.download=`dentarecord_backup_${new Date().toISOString().split('T')[0]}.json`;
  a.click(); URL.revokeObjectURL(url); toast('success','All records exported ✓');
}

function uploadFile(file) {
  const reader=new FileReader();
  reader.onload=ev=>{
    try {
      const d=JSON.parse(ev.target.result);
      if (Array.isArray(d)) { // backup file with multiple records
        Object.values(d).forEach(p=>LS.save(p));
        state.localList=LS.list(); setView('list');
        toast('success',`Imported ${Object.values(d).length} records ✓`);
      } else {
        state.patient=d; LS.save(d); state.localList=LS.list();
        setView('record'); toast('success','Patient record loaded ✓');
      }
    } catch { toast('error','Invalid file.'); }
  };
  reader.readAsText(file);
}

async function saveOneDrive() {
  if (!state.token) return toast('error','Connect OneDrive first.');
  toast('loading','Saving to OneDrive…');
  try { await odSave({...state.patient,updatedAt:new Date().toISOString()},state.token); toast('success',`Saved "${state.patient.name||'Patient'}" to OneDrive ✓`); }
  catch(e) { toast('error',e.message); }
}

async function loadODList() {
  if (!state.token) return toast('error','Connect OneDrive first.');
  toast('loading','Loading from OneDrive…');
  try { state.odFiles=await odList(state.token); state.listSource='onedrive'; state.view='list'; state.status=null; render(); }
  catch(e) { toast('error',e.message); }
}

async function loadODFile(fileId) {
  toast('loading','Opening…');
  try {
    const d=await odLoad(fileId,state.token);
    LS.save(d); state.localList=LS.list(); state.patient=d;
    state.view='record'; state.status=null; render();
  } catch(e) { toast('error',e.message); }
}

async function deleteOD(fileId) {
  toast('loading','Deleting…');
  try {
    await odDel(fileId,state.token);
    state.odFiles=state.odFiles.filter(f=>f.id!==fileId);
    state.confirm=null; toast('success','Deleted from OneDrive ✓');
  } catch(e) { toast('error',e.message); }
}

function connectToken() {
  const t=state.tokenInput.trim();
  if (!t) return;
  state.token=t; state.tokenInput='';
  localStorage.setItem(TOKEN_KEY,t);
  toast('success','OneDrive connected ✓');
}

function disconnectToken() {
  state.token=null; localStorage.removeItem(TOKEN_KEY); render();
}

// ── SVG Icons ─────────────────────────────────────────────────────────────────
const ICONS = {
  home:    `<path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>`,
  plus:    `<line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>`,
  list:    `<line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/>`,
  cloud:   `<polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0018 9h-1.26A8 8 0 103 16.3"/>`,
  chevron: `<polyline points="9 18 15 12 9 6"/>`,
  trash:   `<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 011-1h4a1 1 0 011 1v2"/>`,
  download:`<path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>`,
  search:  `<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>`,
};

// ── Tooth SVG ─────────────────────────────────────────────────────────────────
function renderTooth(num, condition='healthy') {
  const cond = CONDITIONS[condition]||CONDITIONS.healthy;
  const isMolar = [6,7,8,16,17,18].some(n=>String(num).endsWith(String(n)));
  const wrapper = h('div',{class:'tooth',onclick:()=>clickTooth(num),title:`Tooth ${num}: ${cond.label}`});
  wrapper.appendChild(h('span',{class:'tooth-num'},String(num)));

  const svgEl = document.createElementNS('http://www.w3.org/2000/svg','svg');
  svgEl.setAttribute('width','32'); svgEl.setAttribute('height','32'); svgEl.setAttribute('viewBox','0 0 40 40');

  if (condition==='missing'||condition==='extracted') {
    svgEl.innerHTML=`<line x1="8" y1="8" x2="32" y2="32" stroke="#aaa" stroke-width="2.5" stroke-linecap="round"/>
      <line x1="32" y1="8" x2="8" y2="32" stroke="#aaa" stroke-width="2.5" stroke-linecap="round"/>`;
  } else {
    let inner = `<rect x="4" y="4" width="32" height="32" rx="${isMolar?6:10}"
      fill="${cond.color}" stroke="#c0392b" stroke-width="2"
      ${condition!=='healthy'?'filter="url(#shadow)"':''}/>`;
    if (isMolar) inner+=`
      <line x1="20" y1="8" x2="20" y2="32" stroke="#c0392b" stroke-width="1" opacity="0.4"/>
      <line x1="8" y1="20" x2="32" y2="20" stroke="#c0392b" stroke-width="1" opacity="0.4"/>
      <circle cx="14" cy="14" r="2" fill="#c0392b" opacity="0.3"/>
      <circle cx="26" cy="14" r="2" fill="#c0392b" opacity="0.3"/>
      <circle cx="14" cy="26" r="2" fill="#c0392b" opacity="0.3"/>
      <circle cx="26" cy="26" r="2" fill="#c0392b" opacity="0.3"/>`;
    if (condition==='crown') inner+=`<path d="M12 10 L20 6 L28 10 L28 30 L12 30 Z" fill="#f39c12" opacity="0.4"/>`;
    svgEl.innerHTML=`<defs><filter id="shadow"><feDropShadow dx="0" dy="0" stdDeviation="2" flood-opacity="0.2"/></filter></defs>${inner}`;
  }
  wrapper.appendChild(svgEl);
  return wrapper;
}

// ── Dental Chart Component ────────────────────────────────────────────────────
function renderChart() {
  const wrap = h('div',{class:'chart-wrap'});
  const labelRow = h('div',{class:'chart-label-row'});
  ['RIGHT','UPPER','LEFT'].forEach(t=>labelRow.appendChild(h('span',{class:'chart-label'},t)));
  wrap.appendChild(labelRow);

  const upperRow = h('div',{class:'chart-row'});
  UPPER_TEETH.forEach(t=>upperRow.appendChild(renderTooth(t,state.patient.teeth[t]||'healthy')));
  wrap.appendChild(upperRow);

  const divRow = h('div',{class:'chart-divider'});
  const lowerRow = h('div',{class:'chart-row'});
  LOWER_TEETH.forEach(t=>lowerRow.appendChild(renderTooth(t,state.patient.teeth[t]||'healthy')));
  divRow.appendChild(lowerRow); wrap.appendChild(divRow);

  const centerLabel = h('div',{class:'chart-center-label'});
  centerLabel.appendChild(h('span',{class:'chart-label'},'LOWER'));
  wrap.appendChild(centerLabel);
  return wrap;
}

// ── Legend ────────────────────────────────────────────────────────────────────
function renderLegend() {
  const wrap = h('div',{class:'legend'});
  for (const [key,val] of Object.entries(CONDITIONS)) {
    const item = h('div',{
      class:`legend-item${state.selectedCond===key?' active':''}`,
      onclick:()=>{ state.selectedCond=key; render(); }
    });
    const dot = h('span',{class:'legend-dot',style:{background:val.color}});
    item.appendChild(dot);
    item.appendChild(document.createTextNode(val.label));
    wrap.appendChild(item);
  }
  return wrap;
}

// ── Treatment Table ───────────────────────────────────────────────────────────
function buildTreatmentTable() {
  const frag=document.createDocumentFragment();
  const table=h('table',{class:'treatment-table'});
  const thead=h('thead');
  const tr1=h('tr');
  [['DATE','100px'],['NO.','50px'],['DESCRIPTION OF TREATMENT',''],['TIME','70px'],['DEBIT','90px']].forEach(([label,w])=>{
    const th=h('th'); th.textContent=label; if(w)th.style.width=w; tr1.appendChild(th);
  });
  const cth=h('th'); cth.colSpan=3; cth.textContent='CREDIT'; tr1.appendChild(cth);
  tr1.appendChild(h('th',{style:{width:'36px'}}));
  thead.appendChild(tr1);
  const tr2=h('tr');
  const blank=h('td'); blank.colSpan=5; blank.style.background='rgba(255,248,240,0.9)'; tr2.appendChild(blank);
  ['DATE','AMOUNT','BALANCE'].forEach(label=>{ const th=h('th'); th.textContent=label; th.style.width='90px'; tr2.appendChild(th); });
  tr2.appendChild(h('td')); thead.appendChild(tr2); table.appendChild(thead);
  const tbody=h('tbody');
  const treatments=state.patient.treatments||[];
  if (!treatments.length) {
    const empty=h('tr'); const td=h('td'); td.colSpan=9; td.className='empty-state'; td.style.fontSize='13px'; td.textContent='No treatments yet.'; empty.appendChild(td); tbody.appendChild(empty);
  }
  treatments.forEach((row,i)=>{
    const tr=h('tr');
    [['date','100px','date'],['toothNo','50px','text'],['description','','text'],['time','70px','text'],['debit','90px','text'],['creditDate','90px','date'],['creditAmount','90px','text'],['balance','90px','text']].forEach(([key,w,type])=>{
      const td=h('td');
      const inp=h('input',{type,class:'t-input',value:row[key]||''});
      if(w) td.style.width=w;
      // 'input' event — silent patch, no render, no focus loss
      inp.addEventListener('input',e=>updateTreatment(i,key,e.target.value));
      td.appendChild(inp); tr.appendChild(td);
    });
    const delTd=h('td',{style:{textAlign:'center',padding:'4px'}});
    const delBtn=h('button',{class:'btn',onclick:()=>removeTreatment(i),style:{color:'#e74c3c',background:'none',border:'none',cursor:'pointer',fontSize:'18px',padding:'2px 6px'}});
    delBtn.textContent='×'; delTd.appendChild(delBtn); tr.appendChild(delTd);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody); frag.appendChild(table);
  const addBtn=h('button',{class:'btn btn-ghost',style:{marginTop:'10px'},onclick:addTreatment});
  addBtn.textContent='+ Add Treatment Row'; frag.appendChild(addBtn);
  return frag;
}
function renderTreatmentTable() { return buildTreatmentTable(); }

// ── Patient Form ──────────────────────────────────────────────────────────────
function renderPatientForm() {
  const grid=h('div',{class:'form-grid'});
  const fields=[
    {label:'Name',key:'name',span:2,type:'text',placeholder:'Full name'},
    {label:'Age',key:'age',span:1,type:'number',placeholder:'—'},
    {label:'Address',key:'address',span:2,type:'text',placeholder:'Full address'},
    {label:'Telephone',key:'telephone',span:1,type:'tel',placeholder:'+1 000 000 0000'},
    {label:'Occupation',key:'occupation',span:1,type:'text',placeholder:'—'},
    {label:'Status',key:'status',span:1,type:'text',placeholder:'Single / Married…'},
    {label:'Chief Complaint',key:'complaint',span:2,type:'text',placeholder:'Describe the main concern…'},
  ];
  fields.forEach(f=>{
    const field=h('div',{class:`form-field span-${f.span}`});
    const lbl=h('label'); lbl.textContent=f.label; field.appendChild(lbl);
    const inp=h('input',{type:f.type,class:'form-input',value:state.patient[f.key]||'',placeholder:f.placeholder||''});
    inp.addEventListener('input',e=>patchPatient({[f.key]:e.target.value}));
    field.appendChild(inp);
    grid.appendChild(field);
  });
  return grid;
}

// ── Home View ─────────────────────────────────────────────────────────────────
function renderHome() {
  const frag=document.createDocumentFragment();

  // Hero
  const hero=h('div',{class:'home-hero fade-in'});
  hero.appendChild(h('h1',{},'Digital Dental Records'));
  hero.appendChild(h('p',{},'Modern dental chart system with automatic local saving, file export, and optional OneDrive cloud sync.'));
  frag.appendChild(hero);

  // Storage overview
  const storageCard=h('div',{class:'card fade-in'});
  const sb=h('div',{class:'card-body'});
  const slbl=h('div',{class:'section-label',style:{marginBottom:'14px'}},'Storage Options');
  sb.appendChild(slbl);
  const sg=h('div',{class:'storage-grid'});
  [
    {icon:'🖥️',title:'Browser Storage',count:`${LS.count()} record${LS.count()!==1?'s':''} saved`,action:()=>{ state.listSource='local'; setView('list'); }},
    {icon:'📁',title:'Local Folder',count:state.dirName?`→ ${state.dirName}`:'Choose a folder',action:pickFolder},
    {icon:'☁️',title:'OneDrive',count:state.token?'Connected':'Tap to connect',action:()=>state.token?loadODList():document.getElementById('token-input-home')?.focus()},
  ].forEach(item=>{
    const sc=h('div',{class:'storage-card',onclick:item.action});
    sc.appendChild(h('div',{class:'icon'},item.icon));
    sc.appendChild(h('div',{class:'s-title'},item.title));
    sc.appendChild(h('div',{class:'s-count'},item.count));
    sg.appendChild(sc);
  });
  // Folder picker row
  sb.appendChild(sg);
  sb.appendChild(h('div',{class:'divider',style:{margin:'14px 0'}}));
  sb.appendChild(h('div',{class:'section-label',style:{marginBottom:'8px'}},'Save Folder'));
  sb.appendChild(buildFolderSetting());
  storageCard.appendChild(sb); frag.appendChild(storageCard);

  // OneDrive connect
  const odCard=h('div',{class:'card connect-card fade-in'});
  const ob=h('div',{class:'card-body'});
  ob.appendChild(h('div',{class:'section-label',style:{marginBottom:'14px'}},'Microsoft OneDrive'));
  if (state.token) {
    const row=h('div',{class:'connected-row'});
    row.appendChild(h('div',{class:'dot-green'}));
    const txt=h('span',{class:'mono',style:{fontSize:'13px',color:'var(--green)',fontWeight:'600'}},'Connected — syncing to OneDrive/'+OD_FOLDER+'/');
    row.appendChild(txt);
    const disc=h('button',{class:'btn btn-outline btn-sm',onclick:disconnectToken,style:{marginLeft:'auto'}});
    disc.textContent='Disconnect'; row.appendChild(disc);
    ob.appendChild(row);
  } else {
    const info=h('p',{style:{color:'var(--muted)',fontSize:'13px',marginBottom:'12px',lineHeight:'1.7'}});
    info.innerHTML='Sign in at <a href="https://developer.microsoft.com/en-us/graph/graph-explorer" target="_blank" style="color:var(--blue)">Microsoft Graph Explorer</a> with your Outlook account, then copy the Access Token and paste it below.';
    ob.appendChild(info);
    const row=h('div',{class:'token-row'});
    const inp=h('input',{id:'token-input-home',class:'token-input',type:'password',placeholder:'Paste Microsoft OAuth2 access token…',value:state.tokenInput});
    inp.addEventListener('input',e=>{ state.tokenInput=e.target.value; });
    const btn=h('button',{class:'btn btn-blue',onclick:connectToken});
    btn.textContent='Connect'; row.appendChild(inp); row.appendChild(btn); ob.appendChild(row);
  }
  odCard.appendChild(ob); frag.appendChild(odCard);

  // Quick actions
  const qg=h('div',{class:'quick-grid fade-in'});
  [
    {icon:'📋',title:'New Patient',desc:'Create a new dental record with full chart',action:()=>setView('record',{fresh:true})},
    {icon:'🗂️',title:'Patient List',desc:`Browse ${LS.count()} saved records`,action:()=>{ state.listSource='local'; setView('list'); }},
    {icon:'📥',title:'Import File',desc:'Load a saved .json record file',action:()=>document.getElementById('upload-input').click()},
  ].forEach(item=>{
    const c=h('div',{class:'quick-card',onclick:item.action});
    c.appendChild(h('div',{class:'q-icon'},item.icon));
    c.appendChild(h('div',{class:'q-title'},item.title));
    c.appendChild(h('div',{class:'q-desc'},item.desc));
    qg.appendChild(c);
  });
  frag.appendChild(qg);
  frag.appendChild(h('input',{id:'upload-input',type:'file',accept:'.json',style:{display:'none'},onchange:e=>{ if(e.target.files[0]) uploadFile(e.target.files[0]); e.target.value=''; }}));

  return frag;
}

// ── Record View ───────────────────────────────────────────────────────────────
function renderRecord() {
  const frag=document.createDocumentFragment();
  const p=state.patient;
  const isSaved=!!LS.all()[p.id];

  // Toolbar
  const toolbar=h('div',{class:'row-between'});
  const titleEl=h('h2',{style:{fontSize:'20px',color:'var(--red-dim)',flex:'1',fontFamily:'var(--font-display)'}}); titleEl.textContent=p.name||'New Patient';
  toolbar.appendChild(titleEl);
  const actions=h('div',{class:'row',style:{gap:'6px'}});

  const clearBtn=h('button',{class:'btn btn-outline btn-sm',onclick:()=>{ state.patient=newPatient(); render(); }});
  clearBtn.textContent='Clear';
  const saveBtn=h('button',{class:'btn btn-outline btn-sm',onclick:saveLocal});
  saveBtn.textContent='💾 Save';
  const dlBtn=h('button',{class:'btn btn-ghost btn-sm',onclick:()=>saveToFolder()});
  dlBtn.textContent='📁 Export';
  const odBtn=h('button',{class:'btn btn-blue btn-sm',onclick:saveOneDrive});
  odBtn.textContent='☁ OneDrive';
  [clearBtn,saveBtn,dlBtn,odBtn].forEach(b=>actions.appendChild(b));
  toolbar.appendChild(actions); frag.appendChild(toolbar);

  // Status bar
  const status=h('div',{id:'record-status-bar',class:'record-status-bar'});
  const pill1=h('span',{class:`status-pill ${isSaved?'pill-saved':'pill-unsaved'}`});
  pill1.textContent=isSaved?'🖥️ Saved locally':'⚠ Not yet saved';
  status.appendChild(pill1);
  if (state.dirName) { const p2=h('span',{class:'status-pill',style:{background:'rgba(26,122,62,0.1)',color:'var(--green)'}}); p2.textContent=`📁 ${state.dirName}`; status.appendChild(p2); }
  if (state.token) { const pill2=h('span',{class:'status-pill pill-cloud'}); pill2.textContent='☁ OneDrive ready'; status.appendChild(pill2); }
  frag.appendChild(status);

  // Patient info
  const infoCard=h('div',{class:'card fade-in'});
  const ih=h('div',{class:'card-header'}); ih.appendChild(h('div',{class:'section-label'},'Patient Information'));
  const ib=h('div',{class:'card-body',style:{paddingTop:'16px'}});
  ib.appendChild(renderPatientForm());
  infoCard.appendChild(ih); infoCard.appendChild(ib); frag.appendChild(infoCard);

  // Dental chart
  const chartCard=h('div',{class:'card fade-in'});
  const ch=h('div',{class:'card-header'}); ch.appendChild(h('div',{class:'section-label'},'Dental Chart — Tap a tooth to mark'));
  const cb=h('div',{class:'card-body'});
  cb.appendChild(renderLegend());
  cb.appendChild(h('div',{style:{height:'12px'}}));
  cb.appendChild(renderChart());

  const chartFooter=h('div',{class:'row',style:{marginTop:'10px',gap:'8px'}});
  const resetBtn=h('button',{class:'btn btn-outline btn-sm',onclick:()=>{ updatePatient({teeth:{}}); }});
  resetBtn.textContent='Reset All';
  const countSpan=h('span',{class:'mono text-muted',style:{fontSize:'10px'}}); countSpan.textContent=`${Object.keys(p.teeth||{}).length} teeth marked`;
  chartFooter.appendChild(resetBtn); chartFooter.appendChild(countSpan);
  cb.appendChild(chartFooter);
  chartCard.appendChild(ch); chartCard.appendChild(cb); frag.appendChild(chartCard);

  // Treatment table — in a named wrapper so add/remove rows can update without full render
  const txCard=h('div',{class:'card fade-in'});
  const txh=h('div',{class:'card-header'}); txh.appendChild(h('div',{class:'section-label'},'Treatment History & Billing Ledger'));
  const txb=h('div',{class:'card-body'});
  const txWrap=h('div',{id:'treatment-table-wrap',class:'treatment-wrap'});
  txWrap.appendChild(buildTreatmentTable());
  txb.appendChild(txWrap); txCard.appendChild(txh); txCard.appendChild(txb); frag.appendChild(txCard);

  // Bottom save bar
  const bar=h('div',{class:'row',style:{justifyContent:'flex-end',paddingBottom:'16px',gap:'8px'}});
  const b1=h('button',{class:'btn btn-outline',onclick:saveLocal}); b1.textContent='💾 Save Local';
  const b2=h('button',{class:'btn btn-ghost',onclick:()=>saveToFolder()}); b2.textContent='📁 Export File';
  const b3=h('button',{class:'btn btn-blue',onclick:saveOneDrive}); b3.textContent='☁ Save to OneDrive';
  [b1,b2,b3].forEach(b=>bar.appendChild(b));
  frag.appendChild(bar);

  return frag;
}

// ── List View ─────────────────────────────────────────────────────────────────
function renderList() {
  const frag=document.createDocumentFragment();
  const isLocal=state.listSource==='local';

  // Tab bar
  const tabs=h('div',{class:'tab-bar'});
  [['local','🖥️ Local'],['onedrive','☁ OneDrive']].forEach(([src,label])=>{
    const btn=h('button',{class:`tab-btn${state.listSource===src?' active':''}`,onclick:()=>{
      state.listSource=src; state.search='';
      if(src==='onedrive') loadODList(); else { state.localList=LS.list(); render(); }
    }}); btn.textContent=label; tabs.appendChild(btn);
  });
  frag.appendChild(tabs);

  // Header
  const hrow=h('div',{class:'row-between'});
  const title=h('h2',{style:{fontSize:'18px',color:'var(--red-dim)'}}); title.textContent=isLocal?`Local (${state.localList.length})`:`OneDrive (${state.odFiles.length})`;
  hrow.appendChild(title);
  const btnRow=h('div',{class:'row',style:{gap:'6px'}});
  if (isLocal) { const exp=h('button',{class:'btn btn-ghost btn-sm',onclick:exportAll}); exp.textContent='⬇ Backup All'; btnRow.appendChild(exp); }
  const newBtn=h('button',{class:'btn btn-primary btn-sm',onclick:()=>setView('record',{fresh:true})}); newBtn.textContent='+ New Patient';
  btnRow.appendChild(newBtn); hrow.appendChild(btnRow); frag.appendChild(hrow);

  // Search
  const sw=h('div',{class:'search-wrap'});
  sw.appendChild(svg(ICONS.search,16,16,'0 0 24 24'));
  sw.querySelector('svg').style.cssText='position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--muted);pointer-events:none';
  const si=h('input',{class:'search-input',type:'search',placeholder:'Search patients…',value:state.search});
  si.addEventListener('input',e=>{ state.search=e.target.value; render(); });
  sw.appendChild(si); frag.appendChild(sw);

  // List card
  const card=h('div',{class:'card mb-0 fade-in'});
  const q=state.search.toLowerCase();

  if (isLocal) {
    const filtered=state.localList.filter(p=>(p.name||'').toLowerCase().includes(q)||(p.complaint||'').toLowerCase().includes(q));
    if (!filtered.length) {
      const empty=h('div',{class:'empty-state'});
      empty.textContent=state.localList.length?'No matching records.':'No records yet — create a new patient to get started.';
      card.appendChild(empty);
    } else {
      filtered.forEach(p=>{
        const row=h('div',{class:'patient-row',onclick:()=>{ state.patient=p; setView('record'); }});
        const av=h('div',{class:'patient-avatar'}); av.textContent=(p.name||'?')[0].toUpperCase();
        const info=h('div',{class:'patient-info'});
        const name=h('div',{class:'patient-name'}); name.textContent=p.name||'Unnamed';
        const meta=h('div',{class:'patient-meta'});
        const parts=[];
        if(p.age) parts.push(`Age ${p.age}`);
        if(p.complaint) parts.push(p.complaint.substring(0,40)+(p.complaint.length>40?'…':''));
        meta.textContent=parts.join(' · ')||new Date(p.updatedAt).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
        info.appendChild(name); info.appendChild(meta);

        const btnWrap=h('div',{class:'row',style:{gap:'4px'}});
        const dl=h('button',{class:'btn btn-ghost btn-icon',onclick:e=>{ e.stopPropagation(); downloadFile(p); },title:'Download'});
        dl.appendChild(svg(ICONS.download,14,14));
        const del=h('button',{class:'btn btn-icon',onclick:e=>{ e.stopPropagation(); state.confirm={id:p.id,label:p.name||'Unnamed',source:'local'}; render(); },title:'Delete',style:{color:'#e74c3c',background:'none',border:'none',cursor:'pointer'}});
        del.appendChild(svg(ICONS.trash,14,14));
        btnWrap.appendChild(dl); btnWrap.appendChild(del);

        row.appendChild(av); row.appendChild(info); row.appendChild(btnWrap);
        row.appendChild(svg(ICONS.chevron,16,16));
        card.appendChild(row);
      });
    }
  } else {
    if (!state.token) {
      const empty=h('div',{class:'empty-state'}); empty.textContent='Connect OneDrive on the Home screen.'; card.appendChild(empty);
    } else {
      const filtered=state.odFiles.filter(f=>f.name.toLowerCase().includes(q));
      if (!filtered.length) {
        const empty=h('div',{class:'empty-state'}); empty.textContent=`No records in OneDrive/${OD_FOLDER}/`; card.appendChild(empty);
      } else {
        filtered.forEach(f=>{
          const cleanName=f.name.replace('dental_','').replace('.json','').replace(/_/g,' ');
          const row=h('div',{class:'patient-row',onclick:()=>loadODFile(f.id)});
          const av=h('div',{class:'patient-avatar',style:{background:'linear-gradient(135deg,#005a9e,#0078d4)'}}); av.textContent=cleanName[0]?.toUpperCase()||'?';
          const info=h('div',{class:'patient-info'});
          const name=h('div',{class:'patient-name'}); name.textContent=cleanName;
          const meta=h('div',{class:'patient-meta'}); meta.textContent='OneDrive · '+new Date(f.modifiedTime).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
          info.appendChild(name); info.appendChild(meta);
          const del=h('button',{class:'btn btn-icon',onclick:e=>{ e.stopPropagation(); state.confirm={id:f.id,label:cleanName,source:'onedrive'}; render(); },style:{color:'#e74c3c',background:'none',border:'none',cursor:'pointer'}});
          del.appendChild(svg(ICONS.trash,14,14));
          row.appendChild(av); row.appendChild(info); row.appendChild(del);
          row.appendChild(svg(ICONS.chevron,16,16));
          card.appendChild(row);
        });
      }
    }
  }
  frag.appendChild(card);
  return frag;
}

// ── Bottom Nav ────────────────────────────────────────────────────────────────
function renderBottomNav() {
  const nav=document.getElementById('bottom-nav');
  if (!nav) return;
  nav.innerHTML='';
  const inner=h('div',{class:'bottom-nav-inner'});
  const items=[
    {id:'home',  label:'Home',    icon:ICONS.home,  action:()=>setView('home')},
    {id:'record',label:'New',     icon:ICONS.plus,  action:()=>setView('record',{fresh:true})},
    {id:'list',  label:'Records', icon:ICONS.list,  action:()=>{ state.listSource='local'; setView('list'); }},
    {id:'cloud', label:'OneDrive',icon:ICONS.cloud, action:()=>state.token?loadODList():setView('home')},
  ];
  items.forEach(item=>{
    const isActive = (item.id==='home'&&state.view==='home')
      ||(item.id==='record'&&state.view==='record')
      ||(item.id==='list'&&state.view==='list')
      ||(item.id==='cloud'&&state.view==='list'&&state.listSource==='onedrive');
    const btn=h('button',{class:`nav-item${isActive?' active':''}`,onclick:item.action});
    btn.appendChild(svg(item.icon,22,22));
    btn.appendChild(document.createTextNode(item.label));
    if (item.id==='list'&&LS.count()>0) {
      const badge=h('span',{class:'nav-badge'}); badge.textContent=LS.count(); btn.appendChild(badge);
    }
    inner.appendChild(btn);
  });
  nav.appendChild(inner);
}

// ── Main Render ────────────────────────────────────────────────────────────────
function render() {
  const main=document.getElementById('main-content');
  if (!main) return;
  main.innerHTML='';

  const isMobile=window.innerWidth<=768;
  if (isMobile) main.className='main with-bottom-nav'; else main.className='main';

  let content;
  if (state.view==='home')   content=renderHome();
  else if (state.view==='record') content=renderRecord();
  else content=renderList();
  main.appendChild(content);

  // Header active states
  document.querySelectorAll('.header-btn').forEach(btn=>{
    btn.classList.remove('active');
    const v=btn.dataset.view;
    if (v===state.view||(v==='list'&&state.view==='list')) btn.classList.add('active');
  });

  // Toast
  const existing=document.getElementById('toast');
  if (existing) existing.remove();
  if (state.status) {
    const t=h('div',{id:'toast',class:`toast toast-${state.status.type}`});
    t.textContent=(state.status.type==='loading'?'⏳ ':'')+state.status.msg;
    document.body.appendChild(t);
  }

  // Auto-save badge
  const existingBadge=document.getElementById('autosave-badge');
  if (existingBadge) existingBadge.remove();
  if (state.autoSaved&&state.view==='record') {
    const badge=h('div',{id:'autosave-badge',class:'autosave-badge'},'✓ Auto-saved locally');
    document.body.appendChild(badge);
  }

  // Confirm modal
  const existingModal=document.getElementById('modal-overlay');
  if (existingModal) existingModal.remove();
  if (state.confirm) {
    const overlay=h('div',{id:'modal-overlay',class:'modal-overlay',onclick:e=>{ if(e.target===overlay){ state.confirm=null; render(); } }});
    const modal=h('div',{class:'modal'});
    modal.appendChild(h('div',{class:'modal-handle'}));
    modal.appendChild(h('div',{class:'modal-title'},'Delete Record?'));
    const sub=h('div',{class:'modal-sub'}); sub.textContent=state.confirm.label||'This record will be permanently deleted.'; modal.appendChild(sub);
    const acts=h('div',{class:'modal-actions'});
    const cancelBtn=h('button',{class:'btn btn-outline btn-full',onclick:()=>{ state.confirm=null; render(); }}); cancelBtn.textContent='Cancel';
    const delBtn=h('button',{class:'btn btn-full',style:{background:'#e74c3c',color:'white'},onclick:()=>{
      if(state.confirm.source==='local'){ LS.del(state.confirm.id); state.localList=LS.list(); state.confirm=null; toast('success','Deleted.'); }
      else deleteOD(state.confirm.id);
    }}); delBtn.textContent='Delete';
    acts.appendChild(cancelBtn); acts.appendChild(delBtn); modal.appendChild(acts);
    overlay.appendChild(modal); document.body.appendChild(overlay);
  }

  renderBottomNav();
}

// ── Init ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', ()=>{
  // Build static shell
  document.getElementById('app').innerHTML=`
    <header class="app-header">
      <div class="header-brand">
        <svg class="header-logo" viewBox="0 0 40 40" fill="none">
          <circle cx="20" cy="20" r="20" fill="rgba(255,255,255,0.15)"/>
          <path d="M20 8C15 8 10 12 10 17C10 22 13 24 13 28C13 30 14 32 15.5 32C17 32 17.5 28 20 28C22.5 28 23 32 24.5 32C26 32 27 30 27 28C27 24 30 22 30 17C30 12 25 8 20 8Z" fill="white" opacity="0.9"/>
        </svg>
        <div>
          <div class="header-title">DentaRecord</div>
          <div class="header-sub">DIGITAL DENTAL CHART</div>
        </div>
      </div>
      <div class="header-actions">
        <button class="header-btn" data-view="home" onclick="state.view='home';render()">Home</button>
        <button class="header-btn" data-view="record" onclick="state.patient=newPatient();state.view='record';render()">New Patient</button>
        <button class="header-btn" data-view="list" onclick="state.listSource='local';state.view='list';render()">Records (${LS.count()})</button>
        <button class="header-btn" style="background:rgba(0,120,212,0.3)" onclick="state.token?loadODList():state.view='home'&&render()">☁ OneDrive</button>
      </div>
    </header>
    <main id="main-content" class="main"></main>
    <nav class="bottom-nav" id="bottom-nav"></nav>
  `;

  // Register service worker
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(()=>{});
  }

  // Handle resize
  window.addEventListener('resize', render);

  render();
});
