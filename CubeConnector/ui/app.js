const cc = window.chrome.webview.hostObjects.cc;
let MODEL = { measures: [], columns: [] };
let CURRENT = null; // function being edited

async function call(p){ const s = await p; const o = JSON.parse(s); if(o.error) throw new Error(o.error); return o; }

async function boot(){
  try { const a = await call(cc.GetAccount()); document.getElementById('account').textContent = 'Signed in: ' + (a.upn||'(unknown)'); }
  catch(e){ document.getElementById('account').textContent = 'Not signed in'; }
  await refreshLibrary();
}

async function refreshLibrary(){
  const o = await call(cc.GetFunctions());
  const list = document.getElementById('functionList'); list.innerHTML = '';
  (o.functions||[]).forEach(f => {
    const div = document.createElement('div'); div.className='function-item';
    div.innerHTML = `<div class="function-name">${f.FunctionName}</div>
      <div class="function-meta">${f.MeasureName||''} · ${(f.Parameters||[]).length} filters
      <a href="#" onclick="editFunction('${f.FunctionName}');return false;">Edit</a>
      <a href="#" onclick="delFunction('${f.FunctionName}');return false;">Delete</a></div>`;
    list.appendChild(div);
  });
}

function showLibrary(){ document.getElementById('editorView').style.display='none';
  document.getElementById('libraryView').style.display='block'; }
function showEditor(){ document.getElementById('libraryView').style.display='none';
  document.getElementById('editorView').style.display='block'; }

async function newFunction(){
  CURRENT = { FunctionName:'', MeasureName:'', DatasetId:'', TenantId:'', Parameters:[] };
  showEditor(); await loadModels(); document.getElementById('friendlyName').value=''; renderFilters(); renderPreview();
}

async function loadModels(preselectId){
  const sel = document.getElementById('modelSelect'); sel.innerHTML = '<option>Loading…</option>';
  const o = await call(cc.ListDatasets());
  sel.innerHTML='';
  (o.datasets||[]).forEach(d => { const opt=document.createElement('option');
    opt.value = JSON.stringify({id:d.Id, group:d.WorkspaceId, name:d.Name});
    opt.textContent = (d.WorkspaceName||'') + ' ▸ ' + d.Name; sel.appendChild(opt); });
  if (sel.options.length) {
    // When editing, pre-select the formula's saved model so we don't silently re-point it.
    if (preselectId) {
      for (const opt of sel.options) {
        try { if (JSON.parse(opt.value).id === preselectId) { sel.value = opt.value; break; } } catch(_){}
      }
    }
    await onModelChange();
  }
}

async function onModelChange(){
  const sel = document.getElementById('modelSelect'); if(!sel.value) return;
  const {id, group, name} = JSON.parse(sel.value);
  CURRENT.DatasetId = id; CURRENT._group = group; CURRENT.ModelName = name;
  const ms = document.getElementById('measureSelect'); ms.innerHTML='<option>Loading…</option>';
  try { MODEL = await call(cc.GetModel(id, group||'')); }
  catch(e){ ms.innerHTML='<option>Couldn\'t read this data</option>'; return; }
  ms.innerHTML='';
  MODEL.measures.forEach(m => { const o=document.createElement('option'); o.value=m.Name; o.textContent=m.Name; ms.appendChild(o); });
  renderFilters(); renderPreview();
}

function addFilter(){
  CURRENT.Parameters.push({ Name:'', TableName:'', FieldName:'', DataType:'text', FilterType:'List', IsOptional:true, _kind:'match' });
  renderFilters(); renderPreview();
}
function renderFilters(){
  const wrap = document.getElementById('filterList'); wrap.innerHTML='';
  CURRENT.Parameters.filter(p=>!p._isEnd).forEach((p) => {
    const idx = CURRENT.Parameters.indexOf(p);
    const fields = MODEL.columns.map(c => `<option value='${c.Table}||${c.Name}||${c.DataType}'
      ${p.TableName===c.Table&&p.FieldName===c.Name?'selected':''}>${c.Table} · ${c.Name}</option>`).join('');
    const div = document.createElement('div'); div.className='parameter-card';
    div.innerHTML = `<select class="cc-input" onchange="setField(${idx}, this.value)"><option value="">choose a field…</option>${fields}</select>
      <label><input type="radio" name="kind${idx}" ${p._kind!=='range'?'checked':''} onchange="setKind(${idx},'match')"> Match value(s)</label>
      <label><input type="radio" name="kind${idx}" ${p._kind==='range'?'checked':''} onchange="setKind(${idx},'range')"> Date range</label>
      <input class="cc-input" placeholder="filter name" value="${p.Name||''}" oninput="setName(${idx}, this.value)">
      <a href="#" onclick="removeFilter(${idx});return false;">remove</a>`;
    wrap.appendChild(div);
  });
}
function setField(i,v){ const [t,f,dt]=v.split('||'); CURRENT.Parameters[i].TableName=t; CURRENT.Parameters[i].FieldName=f;
  CURRENT.Parameters[i].DataType = mapType(dt); if(!CURRENT.Parameters[i].Name) CURRENT.Parameters[i].Name=suggest(f); renderPreview(); }
function setName(i,v){ CURRENT.Parameters[i].Name=v; renderPreview(); }
function setKind(i,k){ CURRENT.Parameters[i]._kind=k; renderPreview(); }
function removeFilter(i){ CURRENT.Parameters.splice(i,1); renderFilters(); renderPreview(); }
function mapType(dt){ dt=(dt||'').toLowerCase(); if(dt.includes('date')||dt.includes('time'))return 'date';
  if(['integer','int64','number','double','decimal','currency'].includes(dt))return 'number'; return 'text'; }
function suggest(f){ return (f||'param').replace(/[^A-Za-z0-9]/g,'').toLowerCase(); }

function paramNames(){
  const out=[];
  CURRENT.Parameters.filter(p=>!p._isEnd).forEach(p=>{
    if(p._kind==='range'){ out.push((p.Name||'from')+'_start',(p.Name||'to')+'_end'); }
    else out.push(p.Name||'value');
  });
  return out;
}
function renderPreview(){
  const friendly = document.getElementById('friendlyName').value || 'Formula';
  const name = 'CC.' + friendly.replace(/[^A-Za-z0-9_]/g,'') ;
  document.getElementById('nameHint').innerHTML = "In Excel you'll type: <b>="+name+"(…)</b>";
  const measure = document.getElementById('measureSelect').value || 'the value';
  const names = paramNames();
  const tmpl = '='+name+'('+names.join(', ')+')';
  const ex = '='+name+'('+names.map(n=>n.includes('date')||n.includes('start')||n.includes('end')?'"1/1/2025"':'"4000"').join(',')+')';
  document.getElementById('preview').innerHTML =
    `<div><b>How you'll use it</b></div><div>Returns <b>${measure}</b>${names.length?`, filtered by ${names.join(', ')}`:''}.</div>
     <code>${tmpl}</code><div class="subtitle">Example:</div><code>${ex}</code>`;
}

async function saveFunction(){
  const friendly = document.getElementById('friendlyName').value.trim();
  if(!document.getElementById('measureSelect').value || !friendly){ alert('Pick a number and give it a name.'); return; }
  // Expand range filters into RangeStart/RangeEnd pairs; assign positions.
  const params=[]; let pos=0;
  CURRENT.Parameters.filter(p=>!p._isEnd).forEach(p=>{
    if(p._kind==='range'){
      params.push({Name:(p.Name||'from')+'_start',Position:pos++,TableName:p.TableName,FieldName:p.FieldName,DataType:'date',FilterType:'RangeStart',IsOptional:true});
      params.push({Name:(p.Name||'to')+'_end',Position:pos++,TableName:p.TableName,FieldName:p.FieldName,DataType:'date',FilterType:'RangeEnd',IsOptional:true});
    } else {
      params.push({Name:p.Name||'value',Position:pos++,TableName:p.TableName,FieldName:p.FieldName,DataType:p.DataType||'text',FilterType:'List',IsOptional:true});
    }
  });
  const dto = { FunctionName:'CC.'+friendly.replace(/[^A-Za-z0-9_]/g,''), MeasureName:'['+document.getElementById('measureSelect').value+']',
    DatasetId:CURRENT.DatasetId, TenantId:CURRENT.TenantId||'', ModelName:CURRENT.ModelName||'', Parameters:params };
  await call(cc.SaveFunction(JSON.stringify(dto)));
  await refreshLibrary(); showLibrary();
  // Reload functions into Excel without restarting (save never removes, so no restart needed)
  try {
    const r = await call(cc.ReloadFunctions());
    showStatus('✓ Loaded — your formula is ready to use in Excel.');
  } catch(e) {
    showStatus('Saved. Reload into Excel failed: ' + e.message);
  }
}

async function editFunction(name){
  const o = await call(cc.GetFunctions());
  const f = (o.functions||[]).find(x=>x.FunctionName===name); if(!f) return;
  CURRENT = JSON.parse(JSON.stringify(f)); CURRENT._group='';
  showEditor();
  // Pre-select the saved model (loads its measures/fields and restores ModelName/DatasetId).
  await loadModels(f.DatasetId);
  // Restore the saved measure selection (measureSelect option values are bare names).
  const measName = (f.MeasureName||'').replace(/^\[|\]$/g,'');
  const ms = document.getElementById('measureSelect');
  for (const opt of ms.options) { if (opt.value === measName) { ms.value = measName; break; } }
  document.getElementById('friendlyName').value = name.replace(/^CC\./,'');
  // collapse RangeStart/RangeEnd pairs back to one 'range' filter for editing
  const collapsed=[]; (f.Parameters||[]).forEach(p=>{ if(p.FilterType==='RangeEnd')return;
    collapsed.push({...p,_kind:p.FilterType==='RangeStart'?'range':'match'}); });
  CURRENT.Parameters = collapsed; renderFilters(); renderPreview();
}
async function delFunction(name){ if(!confirm('Delete '+name+'?'))return; await call(cc.DeleteFunction(name));
  await refreshLibrary();
  // Reload so the bridge detects the removal and tells us a restart is needed
  try {
    const r = await call(cc.ReloadFunctions());
    if (r.removedNeedRestart) document.getElementById('restart').style.display='block';
  } catch(e) {
    document.getElementById('restart').style.display='block';
  }
}

async function doImport(){
  const path = prompt('Path to the shared formulas file (.json):'); if(!path) return;
  const policy = confirm('Overwrite formulas that already exist? (Cancel = keep both)') ? 'Overwrite' : 'KeepBoth';
  const r = await call(cc.ImportFunctions(path, policy));
  alert(`Imported: ${r.added} new, ${r.overwritten} replaced, ${r.skipped} skipped.`);
  await refreshLibrary();
  try {
    const rel = await call(cc.ReloadFunctions());
    if (rel.removedNeedRestart) document.getElementById('restart').style.display='block';
    else showStatus('✓ Imported formulas are ready to use in Excel.');
  } catch(e) {
    document.getElementById('restart').style.display='block';
  }
}
async function doExport(){
  const path = prompt('Save shared file to (.json):'); if(!path) return;
  await call(cc.ExportFunctions(JSON.stringify([]), path)); alert('Exported to '+path);
}

function showStatus(msg){
  const el = document.getElementById('status-msg');
  el.textContent = msg; el.style.display='block';
  clearTimeout(el._t); el._t = setTimeout(()=>{ el.style.display='none'; }, 5000);
}

async function reloadIntoExcel(){
  try {
    const r = await call(cc.ReloadFunctions());
    if (r.removedNeedRestart) {
      document.getElementById('restart').style.display='block';
    } else {
      showStatus('✓ Loaded — your formulas are ready to use in Excel.');
    }
  } catch(e) {
    showStatus('Reload failed: ' + e.message);
  }
}

boot();
