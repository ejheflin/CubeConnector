/* CubeConnector wizard UI. Talks to the C# WizardBridge host object ("cc"). */
const cc = window.chrome.webview.hostObjects.cc;
let MODEL = { measures: [], columns: [] };
let CURRENT = null;            // function being built/edited
let modelCombo, measureCombo;  // searchable pickers

async function call(p){ const s = await p; const o = JSON.parse(s); if(o.error) throw new Error(o.error); return o; }
function esc(s){ return (s==null?'':String(s)).replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function $(id){ return document.getElementById(id); }

/* ---------- searchable combobox ---------- */
function Combo(mount, opts){
  let items = [], value = null, activeIdx = -1, collapsed = new Set(), loading = false;
  const root = document.createElement('div'); root.className = 'combo';
  root.innerHTML =
    `<div class="combo-trigger field" tabindex="0"><span class="val placeholder"></span><span class="chev">▾</span></div>
     <div class="combo-pop"><input class="combo-search" type="text"><div class="combo-options"></div></div>`;
  mount.innerHTML = ''; mount.appendChild(root);
  const trigger = root.querySelector('.combo-trigger'),
        valEl   = root.querySelector('.val'),
        search  = root.querySelector('.combo-search'),
        optsEl  = root.querySelector('.combo-options');
  valEl.textContent = opts.placeholder || 'Select…';
  search.placeholder = opts.searchPlaceholder || 'Type to filter…';

  function open(){ root.classList.add('open'); search.value=''; render(''); setTimeout(()=>search.focus(),0); }
  function close(){ root.classList.remove('open'); }
  trigger.addEventListener('click', ()=> root.classList.contains('open') ? close() : open());
  trigger.addEventListener('keydown', e=>{ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); open(); }});
  search.addEventListener('input', ()=> render(search.value));
  search.addEventListener('keydown', e=>{
    const vis = [...optsEl.querySelectorAll('.combo-opt')];
    if(e.key==='ArrowDown'){ e.preventDefault(); activeIdx=Math.min(activeIdx+1,vis.length-1); paint(vis); }
    else if(e.key==='ArrowUp'){ e.preventDefault(); activeIdx=Math.max(activeIdx-1,0); paint(vis); }
    else if(e.key==='Enter'){ e.preventDefault(); (vis[activeIdx]||vis[0])?.click(); }
    else if(e.key==='Escape'){ close(); }
  });
  document.addEventListener('click', e=>{ if(!root.contains(e.target)) close(); });

  function paint(vis){ vis.forEach((el,i)=>el.classList.toggle('active', i===activeIdx)); vis[activeIdx]?.scrollIntoView({block:'nearest'}); }
  function match(it,q){ q=q.toLowerCase(); return !q || (it.label||'').toLowerCase().includes(q) || (it.sub||'').toLowerCase().includes(q) || (it.group||'').toLowerCase().includes(q); }
  function optEl(it){
    const d=document.createElement('div'); d.className='combo-opt';
    d.setAttribute('aria-selected', String(it.value===value));
    d.innerHTML = `<span>${esc(it.label)}</span>` + (it.sub?`<span class="sub">${esc(it.sub)}</span>`:'');
    d.addEventListener('click', ()=>{ choose(it, true); close(); });
    return d;
  }
  function render(q){
    activeIdx = -1;
    optsEl.innerHTML = '';
    if(loading && !items.length){ optsEl.innerHTML = `<div class="combo-none">Loading…</div>`; return; }
    const list = items.filter(it=>match(it,q));
    if(!list.length){ optsEl.innerHTML = `<div class="combo-none">No matches</div>`; return; }
    if(opts.grouped){
      const searching = !!q.trim();
      const groups = {};
      list.forEach(it=>{ const g=it.group||''; (groups[g]=groups[g]||[]).push(it); });
      Object.keys(groups).forEach(g=>{
        const isCol = !searching && collapsed.has(g);
        if(g){
          const h=document.createElement('div'); h.className='combo-group';
          h.innerHTML = `<span class="gchev">${isCol?'▸':'▾'}</span><span class="gname">${esc(g)}</span><span class="gcount">${groups[g].length}</span>`;
          h.addEventListener('click', e=>{ e.stopPropagation(); if(collapsed.has(g)) collapsed.delete(g); else collapsed.add(g); render(search.value); });
          optsEl.appendChild(h);
        }
        if(!isCol) groups[g].forEach(it=>optsEl.appendChild(optEl(it)));
      });
    } else list.forEach(it=>optsEl.appendChild(optEl(it)));
  }
  function setLabel(label){ if(label){ valEl.textContent=label; valEl.classList.remove('placeholder'); } else { valEl.textContent=opts.placeholder||'Select…'; valEl.classList.add('placeholder'); } }
  function choose(it, fromUser){ value = it.value; setLabel(it.label); if(fromUser && opts.onSelect) opts.onSelect(it.value, it); }

  return {
    setItems(list){ items = list || []; loading = false; if(opts.defaultCollapsed) collapsed = new Set(items.map(i=>i.group||'')); if(root.classList.contains('open')) render(search.value); },
    setLoading(v){ loading = !!v; if(root.classList.contains('open')) render(search.value); },
    getValue(){ return value; },
    setValueLabel(v, label){ value = v; setLabel(label); },
    selectByValue(v){ const it = items.find(x=>x.value===v); if(it){ choose(it, false); return it; } return null; },
    clear(){ value=null; setLabel(null); }
  };
}

/* ---------- boot ---------- */
async function boot(){
  modelCombo = Combo($('modelCombo'), { placeholder:'Choose your data…', searchPlaceholder:'Search models or workspaces…', grouped:true, onSelect:onModelPicked });
  measureCombo = Combo($('measureCombo'), { placeholder:'Choose a number…', searchPlaceholder:'Search measures…', grouped:true,
    onSelect:(v)=>{ CURRENT.MeasureName = '['+v+']'; renderPreview(); } });
  try { const a = await call(cc.GetAccount()); $('account').textContent = 'Signed in: ' + (a.upn||'(unknown)'); }
  catch(e){ $('account').textContent = 'Not signed in'; }
  await refreshLibrary();
}

async function switchAccount(){
  try {
    const r = await call(cc.SignInDifferent());
    $('account').textContent = 'Signed in: ' + (r.upn||'(unknown)');
    showStatus('Switched account. Pick a model to continue.');
  } catch(e){ showStatus('Sign-in failed: ' + e.message); }
}

/* ---------- library ---------- */
async function refreshLibrary(){
  const o = await call(cc.GetFunctions());
  const list = $('functionList'); list.innerHTML = '';
  const fns = o.functions || [];
  if(!fns.length){ list.innerHTML = `<div class="empty">No formulas yet. Click “+ New formula”, or import a set someone shared with you.</div>`; return; }
  fns.forEach(f => {
    const div = document.createElement('div'); div.className = 'func-card';
    const filters = (f.Parameters||[]).filter(p=>p.FilterType!=='RangeEnd').length;
    div.innerHTML =
      `<div class="name">${esc(f.FunctionName)}</div>
       <div class="meta">${esc(f.ModelName||'')}${f.ModelName?' · ':''}${esc(f.MeasureName||'')} · ${filters} filter${filters===1?'':'s'}
         <span style="flex:1"></span>
         <a onclick="editFunction('${esc(f.FunctionName)}')">Edit</a>
         <a class="danger" onclick="delFunction('${esc(f.FunctionName)}')">Delete</a>
       </div>`;
    list.appendChild(div);
  });
}

function showLibrary(){ $('editorView').style.display='none'; $('libraryView').style.display='block'; }
function showEditor(){ $('libraryView').style.display='none'; $('editorView').style.display='block'; }

/* ---------- builder ---------- */
async function newFunction(){
  CURRENT = { FunctionName:'', MeasureName:'', DatasetId:'', _group:'', ModelName:'', Parameters:[] };
  showEditor();
  $('friendlyName').value = '';
  modelCombo.clear(); measureCombo.clear();
  MODEL = { measures:[], columns:[] };
  renderFilters(); renderPreview();
  modelCombo.setLoading(true);
  await loadModels();
}

async function loadModels(){
  const o = await call(cc.ListDatasets());
  modelCombo.setItems((o.datasets||[]).map(d => ({
    value: d.Id, label: d.Name, group: d.WorkspaceName || 'My workspace', wsId: d.WorkspaceId || ''
  })));
}

function onModelPicked(id, it){
  CURRENT.DatasetId = id; CURRENT._group = it.wsId || ''; CURRENT.ModelName = it.label;
  measureCombo.clear(); CURRENT.MeasureName = '';
  loadMeasures(id, it.wsId || '');
}

async function loadMeasures(id, wsId){
  measureCombo.setLoading(true);
  try { MODEL = await call(cc.GetModel(id, wsId)); }
  catch(e){ MODEL = { measures:[], columns:[] }; measureCombo.setItems([]); showStatus("Couldn't read this model — you may not have access."); renderFilters(); return; }
  measureCombo.setItems((MODEL.measures||[]).map(m => ({ value:m.Name, label:m.Name, group:m.Table || 'Measures', sub:m.Description||'' })));
  renderFilters(); renderPreview();
}

function addFilter(){
  CURRENT.Parameters.push({ Name:'', TableName:'', FieldName:'', DataType:'text', FilterType:'List', IsOptional:true, _kind:'match' });
  renderFilters(); renderPreview();
}
function renderFilters(){
  const wrap = $('filterList'); wrap.innerHTML='';
  const cols = MODEL.columns || [];
  // field options grouped by table, for the searchable field picker
  const fieldItems = cols.map(c => ({ value:`${c.Table}||${c.Name}||${c.DataType}`, label:c.Name, group:c.Table, sub:c.Description||'' }));
  CURRENT.Parameters.filter(p=>!p._isEnd).forEach(p=>{
    const idx = CURRENT.Parameters.indexOf(p);
    const card = document.createElement('div'); card.className='filter-card'; card.dataset.idx = idx;
    card.innerHTML =
      `<div class="row"><span class="drag" title="Drag to reorder">⠿</span><div class="fieldcombo" style="flex:1"></div>
         <button class="icon-btn" title="Remove" onclick="removeFilter(${idx})">✕</button></div>
       <div class="row">
         <span class="seg">
           <label class="${p._kind!=='range'?'on':''}"><input type="radio" name="k${idx}" hidden ${p._kind!=='range'?'checked':''} onchange="setKind(${idx},'match')">Match value(s)</label>
           <label class="${p._kind==='range'?'on':''}"><input type="radio" name="k${idx}" hidden ${p._kind==='range'?'checked':''} onchange="setKind(${idx},'range')">Date range</label>
         </span>
         <input class="field" style="flex:1" placeholder="filter name" value="${esc(p.Name||'')}" oninput="setName(${idx}, this.value)">
       </div>`;
    wrap.appendChild(card);
    const combo = Combo(card.querySelector('.fieldcombo'), {
      placeholder:'choose a field…', searchPlaceholder:'Search fields…', grouped:true,
      onSelect:(v)=>setField(idx, v)
    });
    combo.setItems(fieldItems);
    if(p.TableName && p.FieldName) combo.setValueLabel(`${p.TableName}||${p.FieldName}||${p.DataType}`, p.FieldName);
    wireDrag(card);
  });
}

let dragSrc = null;
function wireDrag(card){
  const handle = card.querySelector('.drag');
  handle.draggable = true;
  handle.addEventListener('dragstart', e=>{ dragSrc = +card.dataset.idx; card.classList.add('dragging');
    e.dataTransfer.effectAllowed='move'; try{ e.dataTransfer.setData('text/plain', card.dataset.idx); }catch(_){} });
  handle.addEventListener('dragend', ()=>{ card.classList.remove('dragging');
    document.querySelectorAll('.filter-card.dragover').forEach(c=>c.classList.remove('dragover')); });
  card.addEventListener('dragover', e=>{ e.preventDefault(); e.dataTransfer.dropEffect='move'; card.classList.add('dragover'); });
  card.addEventListener('dragleave', ()=> card.classList.remove('dragover'));
  card.addEventListener('drop', e=>{ e.preventDefault(); card.classList.remove('dragover');
    const to = +card.dataset.idx;
    if(dragSrc!=null && dragSrc!==to){ const a=CURRENT.Parameters; const [m]=a.splice(dragSrc,1); a.splice(to,0,m); renderFilters(); renderPreview(); }
    dragSrc=null; });
}
function setField(i,v){ const [t,f,dt]=v.split('||'); const p=CURRENT.Parameters[i]; p.TableName=t; p.FieldName=f;
  p.DataType=mapType(dt); if(!p.Name) p.Name=suggest(f); renderFilters(); renderPreview(); }
function setName(i,v){ CURRENT.Parameters[i].Name=v; renderPreview(); }
function setKind(i,k){ CURRENT.Parameters[i]._kind=k; renderFilters(); renderPreview(); }
function removeFilter(i){ CURRENT.Parameters.splice(i,1); renderFilters(); renderPreview(); }
function mapType(dt){ dt=(dt||'').toLowerCase(); if(dt.includes('date')||dt.includes('time'))return 'date';
  if(['integer','int64','number','double','decimal','currency'].includes(dt))return 'number'; return 'text'; }
function suggest(f){ return (f||'param').replace(/[^A-Za-z0-9]/g,'').toLowerCase(); }

function paramNames(){
  const out=[];
  CURRENT.Parameters.filter(p=>!p._isEnd).forEach(p=>{
    if(p._kind==='range') out.push((p.Name||'from')+'_start',(p.Name||'to')+'_end');
    else out.push(p.Name||'value');
  });
  return out;
}
function renderPreview(){
  const friendly = ($('friendlyName').value||'Formula');
  const fnName = 'CC.' + friendly.replace(/[^A-Za-z0-9_]/g,'');
  $('nameHint').innerHTML = "In Excel you'll type <b>="+esc(fnName)+"(…)</b>";
  const measure = (measureCombo.getValue()) || 'the value';
  const names = paramNames();
  // tinted formula
  $('formula').innerHTML = `<span class="fn">=${esc(fnName)}</span>(` +
    names.map(n=>`<span class="arg">${esc(n)}</span>`).join(', ') + ')';
  $('explain').innerHTML = `Returns <b>${esc(measure)}</b>` + (names.length? `, filtered by ${esc(names.join(', '))}.` : '.');
  const ex = names.map(n => /date|start|end/.test(n) ? '"1/1/2025"' : '"4000"');
  $('example').innerHTML = names.length
    ? 'e.g. <span class="lit">=' + esc(fnName) + '(' + ex.join(', ') + ')</span>'
    : '';
}

async function saveFunction(){
  const friendly = $('friendlyName').value.trim();
  const measure = measureCombo.getValue();
  if(!measure || !friendly){ showStatus('Pick the number you want and give the formula a name.'); return; }
  const params=[]; let pos=0;
  CURRENT.Parameters.filter(p=>!p._isEnd).forEach(p=>{
    if(p._kind==='range'){
      params.push({Name:(p.Name||'from')+'_start',Position:pos++,TableName:p.TableName,FieldName:p.FieldName,DataType:'date',FilterType:'RangeStart',IsOptional:true});
      params.push({Name:(p.Name||'to')+'_end',Position:pos++,TableName:p.TableName,FieldName:p.FieldName,DataType:'date',FilterType:'RangeEnd',IsOptional:true});
    } else {
      params.push({Name:p.Name||'value',Position:pos++,TableName:p.TableName,FieldName:p.FieldName,DataType:p.DataType||'text',FilterType:'List',IsOptional:true});
    }
  });
  const dto = { FunctionName:'CC.'+friendly.replace(/[^A-Za-z0-9_]/g,''), MeasureName:'['+measure+']',
    DatasetId:CURRENT.DatasetId, TenantId:'', ModelName:CURRENT.ModelName||'', Parameters:params };
  try {
    await call(cc.SaveFunction(JSON.stringify(dto)));
    await refreshLibrary(); showLibrary();
    try { await call(cc.ReloadFunctions()); showStatus('✓ Saved — your formula is ready to use in Excel.'); }
    catch(e){ showStatus('Saved. Reload into Excel failed: ' + e.message); }
  } catch(e){ showStatus('Save failed: ' + e.message); }
}

async function editFunction(name){
  const o = await call(cc.GetFunctions());
  const f = (o.functions||[]).find(x=>x.FunctionName===name); if(!f) return;
  CURRENT = JSON.parse(JSON.stringify(f)); CURRENT._group='';
  // collapse RangeStart/RangeEnd pairs back to one "range" filter for editing
  CURRENT.Parameters = (f.Parameters||[]).filter(p=>p.FilterType!=='RangeEnd')
    .map(p=>({ ...p, _kind: p.FilterType==='RangeStart' ? 'range' : 'match' }));
  showEditor();
  $('friendlyName').value = name.replace(/^CC\./,'');
  const measName = (f.MeasureName||'').replace(/^\[|\]$/g,'');
  measureCombo.setValueLabel(measName, measName);   // instant — before measures load
  CURRENT.MeasureName = f.MeasureName;
  renderFilters(); renderPreview();
  modelCombo.setLoading(true);
  await loadModels();
  const it = modelCombo.selectByValue(f.DatasetId);     // programmatic: shows real model name, doesn't reset measure
  if(it){ CURRENT.DatasetId=f.DatasetId; CURRENT._group=it.wsId||''; CURRENT.ModelName=it.label; await loadMeasures(f.DatasetId, it.wsId||''); }
  else { modelCombo.setValueLabel(f.DatasetId, f.ModelName||'(model not in your list)'); await loadMeasures(f.DatasetId, ''); }
}

async function delFunction(name){
  if(!confirm('Delete '+name+'?')) return;
  await call(cc.DeleteFunction(name));
  await refreshLibrary();
  try { const r = await call(cc.ReloadFunctions()); if(r.removedNeedRestart) $('restart').style.display='block'; }
  catch(e){ $('restart').style.display='block'; }
}

/* ---------- import / export / reload ---------- */
async function doImport(){
  let path;
  try { const pick = await call(cc.PickImportPath()); if(pick.canceled) return; path = pick.path; }
  catch(e){ showStatus('Import failed: ' + e.message); return; }
  const policy = confirm('Overwrite formulas that already exist?\n\nOK = overwrite,  Cancel = keep both') ? 'Overwrite' : 'KeepBoth';
  try {
    const r = await call(cc.ImportFunctions(path, policy));
    await refreshLibrary();
    try { const rel = await call(cc.ReloadFunctions());
      if(rel.removedNeedRestart) $('restart').style.display='block';
      else showStatus(`✓ Imported ${r.added} new, ${r.overwritten} replaced, ${r.skipped} skipped — ready to use.`);
    } catch(e){ $('restart').style.display='block'; }
  } catch(e){ showStatus('Import failed: ' + e.message); }
}
async function doExport(){
  let path;
  try { const pick = await call(cc.PickExportPath('CubeConnector-formulas.json')); if(pick.canceled) return; path = pick.path; }
  catch(e){ showStatus('Export failed: ' + e.message); return; }
  try { await call(cc.ExportFunctions(JSON.stringify([]), path)); showStatus('✓ Exported to ' + path); }
  catch(e){ showStatus('Export failed: ' + e.message); }
}
async function reloadIntoExcel(){
  try { const r = await call(cc.ReloadFunctions());
    if(r.removedNeedRestart) $('restart').style.display='block';
    else showStatus('✓ Loaded — your formulas are ready to use in Excel.');
  } catch(e){ showStatus('Reload failed: ' + e.message); }
}

function showStatus(msg){
  const el = $('status'); el.textContent = msg; el.style.display='block';
  clearTimeout(el._t); el._t = setTimeout(()=>{ el.style.display='none'; }, 6000);
}

boot();
