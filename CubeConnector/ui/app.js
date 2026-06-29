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
  let items = [], value = null, activeIdx = -1, collapsed = new Set(), loading = false, curLabel = null;
  const root = document.createElement('div'); root.className = 'combo';
  root.innerHTML =
    `<div class="combo-trigger field" tabindex="0"><span class="val placeholder"></span><span class="combo-spin" aria-hidden="true"></span><span class="chev">▾</span></div>
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
  function setLabel(label){ curLabel = label; if(label){ valEl.textContent=label; valEl.classList.remove('placeholder'); } else { valEl.textContent=opts.placeholder||'Select…'; valEl.classList.add('placeholder'); } }
  function choose(it, fromUser){ value = it.value; setLabel(it.label); if(fromUser && opts.onSelect) opts.onSelect(it.value, it); }

  return {
    setItems(list){ items = list || []; loading = false; root.classList.remove('loading'); setLabel(curLabel); if(opts.defaultCollapsed) collapsed = new Set(items.map(i=>i.group||'')); if(root.classList.contains('open')) render(search.value); },
    setLoading(v){ loading = !!v; root.classList.toggle('loading', loading);
      if(loading){ valEl.textContent='Loading…'; valEl.classList.remove('placeholder'); } else { setLabel(curLabel); }
      if(root.classList.contains('open')) render(search.value); },
    getValue(){ return value; },
    setValueLabel(v, label){ value = v; setLabel(label); },
    selectByValue(v){ const it = items.find(x=>x.value===v); if(it){ choose(it, false); return it; } return null; },
    clear(){ value=null; setLabel(null); }
  };
}

/* ---------- boot ---------- */
let _authPoll = null;
let _pendingAuthRetry = null;   // re-run once sign-in becomes ready

function renderAccount(a){
  const el = $('account'); const st = a && a.status;
  if(st==='ready') el.innerHTML = 'Signed in: ' + esc(a.upn||'(unknown)');
  else if(st==='signing-in') el.innerHTML = 'Signing in… <a onclick="cancelSignIn()">Cancel</a>';
  else if(st==='error') el.innerHTML = 'Sign-in failed <a onclick="retrySignIn()">Retry</a>';
  else el.innerHTML = 'Not signed in <a onclick="retrySignIn()">Sign in</a>';
}
function pollAuth(){
  if(_authPoll) return;
  _authPoll = setInterval(async ()=>{
    let a; try { a = await call(cc.GetAuthState()); } catch(e){ return; }
    renderAccount(a);
    if(a.status!=='signing-in'){
      clearInterval(_authPoll); _authPoll = null;
      if(a.status==='ready' && _pendingAuthRetry){ const fn=_pendingAuthRetry; _pendingAuthRetry=null; fn(); }
    }
  }, 600);
}
function onReadyRetry(fn){ _pendingAuthRetry = fn; pollAuth(); }
function cancelSignIn(){ call(cc.CancelSignIn()).then(renderAccount).catch(()=>{}); }
function retrySignIn(){ call(cc.GetAccount()).then(a=>{ renderAccount(a); if(a.status==='signing-in') pollAuth(); }).catch(()=>{}); }

async function boot(){
  modelCombo = Combo($('modelCombo'), { placeholder:'Choose your data…', searchPlaceholder:'Search models or workspaces…', grouped:true, onSelect:onModelPicked });
  measureCombo = Combo($('measureCombo'), { placeholder:'Choose a number…', searchPlaceholder:'Search measures…', grouped:true,
    onSelect:(v)=>{ CURRENT.MeasureName = '['+v+']'; renderPreview(); } });
  try { const a = await call(cc.GetAccount()); renderAccount(a); if(a.status==='signing-in') pollAuth(); }
  catch(e){ $('account').textContent = 'Not signed in'; }
  await refreshLibrary();
}

async function switchAccount(){
  try { const a = await call(cc.SignInDifferent()); renderAccount(a); pollAuth(); showStatus('Signing in… pick a model once connected.'); }
  catch(e){ showStatus('Sign-in failed: ' + e.message); }
}

/* ---------- library ---------- */
async function refreshLibrary(){
  const o = await call(cc.GetFunctions());
  const list = $('functionList'); list.innerHTML = '';
  const fns = o.functions || [];
  if(!fns.length){ list.innerHTML = `<div class="empty">No formulas yet. Click “+ New formula”, or import a set someone shared with you.</div>`; return; }
  fns.forEach(f => {
    const div = document.createElement('div'); div.className = 'func-card';
    const filters = (f.Parameters||[]).length;
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
  if(o.needAuth){ modelCombo.setLoading(true); onReadyRetry(loadModels); return; }
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
  let o;
  try { o = await call(cc.GetModel(id, wsId)); }
  catch(e){ MODEL = { measures:[], columns:[] }; measureCombo.setItems([]); showStatus("Couldn't read this model — you may not have access."); renderFilters(); return; }
  if(o.needAuth){ onReadyRetry(()=>loadMeasures(id, wsId)); return; }
  MODEL = o;
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
  CURRENT.Parameters.forEach((p, idx)=>{
    const card = document.createElement('div'); card.className='filter-card'; card.dataset.idx = idx;
    card.innerHTML =
      `<div class="row"><span class="drag" title="Drag to reorder">⠿</span><div class="fieldcombo" style="flex:1"></div>
         <button class="icon-btn" title="Remove" onclick="removeFilter(${idx})">✕</button></div>
       <div class="row sample-row">${sampleHint(p)}</div>
       <div class="row">
         <span class="seg">
           <label class="${p._kind==='match'?'on':''}"><input type="radio" name="k${idx}" hidden ${p._kind==='match'?'checked':''} onchange="setKind(${idx},'match')">Match</label>
           <label class="${p._kind==='start'?'on':''}"><input type="radio" name="k${idx}" hidden ${p._kind==='start'?'checked':''} onchange="setKind(${idx},'start')">Start ≥</label>
           <label class="${p._kind==='end'?'on':''}"><input type="radio" name="k${idx}" hidden ${p._kind==='end'?'checked':''} onchange="setKind(${idx},'end')">End ≤</label>
         </span>
       </div>
       <div class="row">
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
function suggestName(field, kind){ const base=suggest(field); return base + (kind==='start'?'_start':kind==='end'?'_end':''); }
function setField(i,v){ const [t,f,dt]=v.split('||'); const p=CURRENT.Parameters[i]; p.TableName=t; p.FieldName=f;
  p.DataType=mapType(dt); if(!p.Name) p.Name=suggestName(f, p._kind); fetchSample(i); renderFilters(); renderPreview(); }

function truncate(s,n){ s=String(s); return s.length>n ? s.slice(0,n)+'…' : s; }
// Inline "e.g. <value>" hint for a filter's chosen field. '' = nothing, '…' = loading.
function sampleHint(p){
  if(!p.FieldName) return '';
  if(p._sample===undefined) return '<span class="sample">…</span>';
  if(typeof p._sample==='string' && p._sample!=='') return '<span class="sample">e.g. '+esc(truncate(p._sample,40))+'</span>';
  return '';
}
// Fetch one example value for the field at index i (async, cached server-side, stale-guarded).
async function fetchSample(i){
  const p = CURRENT.Parameters[i];
  if(!p || !p.TableName || !p.FieldName || !CURRENT.DatasetId){ if(p) p._sample=null; return; }
  const key = p.TableName+'||'+p.FieldName;
  p._sampleKey = key; p._sample = undefined;   // loading
  try {
    const r = await call(cc.GetSampleValue(CURRENT.DatasetId, CURRENT._group||'', p.TableName, p.FieldName));
    if(p._sampleKey !== key) return;            // field changed since we asked — ignore stale result
    if(r.needAuth){ onReadyRetry(()=>fetchSample(i)); return; }   // retry once signed in
    p._sample = (r.value===undefined || r.value===null) ? null : String(r.value);
  } catch(e){ if(p._sampleKey===key) p._sample = null; }
  renderFilters(); renderPreview();
}
function setName(i,v){ CURRENT.Parameters[i].Name=v; renderPreview(); }
function setKind(i,k){
  const p = CURRENT.Parameters[i];
  const prevAuto = p.FieldName ? suggestName(p.FieldName, p._kind) : '';
  p._kind = k;
  if(p.FieldName && (!p.Name || p.Name === prevAuto)) p.Name = suggestName(p.FieldName, k);
  renderFilters(); renderPreview();
}
function removeFilter(i){ CURRENT.Parameters.splice(i,1); renderFilters(); renderPreview(); }
function mapType(dt){ dt=(dt||'').toLowerCase(); if(dt.includes('date')||dt.includes('time'))return 'date';
  if(['integer','int64','number','double','decimal','currency'].includes(dt))return 'number'; return 'text'; }
function suggest(f){ return (f||'param').replace(/[^A-Za-z0-9]/g,'').toLowerCase(); }

function paramNames(){
  return CURRENT.Parameters.map(p => p.Name || suggestName(p.FieldName, p._kind) || 'value');
}
// Excel built-in worksheet functions (English). A UDF sharing one of these names is shadowed
// by the built-in and won't work, so we warn live and block on save. Case-insensitive. Covers
// the standard set (incl. legacy/compatibility names); extend as Excel adds functions.
const RESERVED = new Set([
  // math & trig
  "ABS","ACOS","ACOSH","ACOT","ACOTH","AGGREGATE","ARABIC","ASIN","ASINH","ATAN","ATAN2","ATANH","BASE","CEILING","COMBIN","COMBINA","COS","COSH","COT","COTH","CSC","CSCH","DECIMAL","DEGREES","EVEN","EXP","FACT","FACTDOUBLE","FLOOR","GCD","INT","LCM","LN","LOG","LOG10","MDETERM","MINVERSE","MMULT","MOD","MROUND","MULTINOMIAL","MUNIT","ODD","PI","POWER","PRODUCT","QUOTIENT","RADIANS","RAND","RANDARRAY","RANDBETWEEN","ROMAN","ROUND","ROUNDDOWN","ROUNDUP","SEC","SECH","SERIESSUM","SIGN","SIN","SINH","SQRT","SQRTPI","SUBTOTAL","SUM","SUMIF","SUMIFS","SUMPRODUCT","SUMSQ","SUMX2MY2","SUMX2PY2","SUMXMY2","TAN","TANH","TRUNC",
  // statistical (incl. legacy names)
  "AVEDEV","AVERAGE","AVERAGEA","AVERAGEIF","AVERAGEIFS","CORREL","COUNT","COUNTA","COUNTBLANK","COUNTIF","COUNTIFS","COVAR","CONFIDENCE","CRITBINOM","DEVSQ","FISHER","FISHERINV","FORECAST","FREQUENCY","GAMMA","GAMMALN","GAUSS","GEOMEAN","GROWTH","HARMEAN","INTERCEPT","KURT","LARGE","LINEST","LOGEST","MAX","MAXA","MAXIFS","MEDIAN","MIN","MINA","MINIFS","MODE","PEARSON","PERCENTILE","PERCENTRANK","PERMUT","PERMUTATIONA","PHI","PROB","QUARTILE","RANK","RSQ","SKEW","SLOPE","SMALL","STANDARDIZE","STDEV","STDEVA","STDEVP","STDEVPA","STEYX","TREND","TRIMMEAN","VAR","VARA","VARP","VARPA","ZTEST","NORMDIST","NORMINV","NORMSDIST","NORMSINV","LOGNORMDIST","LOGINV","BINOMDIST","NEGBINOMDIST","POISSON","EXPONDIST","WEIBULL","HYPGEOMDIST","BETADIST","BETAINV","CHIDIST","CHIINV","CHITEST","FDIST","FINV","FTEST","TDIST","TINV","TTEST","GAMMADIST","GAMMAINV",
  // text
  "ARRAYTOTEXT","ASC","BAHTTEXT","CHAR","CLEAN","CODE","CONCAT","CONCATENATE","DBCS","DOLLAR","EXACT","FIND","FINDB","FIXED","LEFT","LEFTB","LEN","LENB","LOWER","MID","MIDB","NUMBERVALUE","PROPER","REPLACE","REPLACEB","REPT","RIGHT","RIGHTB","SEARCH","SEARCHB","SUBSTITUTE","T","TEXT","TEXTAFTER","TEXTBEFORE","TEXTJOIN","TEXTSPLIT","TRIM","UNICHAR","UNICODE","UPPER","VALUE","VALUETOTEXT",
  // logical
  "AND","BYCOL","BYROW","FALSE","IF","IFERROR","IFNA","IFS","LAMBDA","LET","MAKEARRAY","MAP","NOT","OR","REDUCE","SCAN","SWITCH","TRUE","XOR",
  // lookup & reference
  "ADDRESS","AREAS","CHOOSE","CHOOSECOLS","CHOOSEROWS","COLUMN","COLUMNS","DROP","EXPAND","FILTER","FORMULATEXT","GETPIVOTDATA","HLOOKUP","HSTACK","HYPERLINK","INDEX","INDIRECT","LOOKUP","MATCH","OFFSET","ROW","ROWS","RTD","SORT","SORTBY","TAKE","TOCOL","TOROW","TRANSPOSE","UNIQUE","VLOOKUP","VSTACK","WRAPCOLS","WRAPROWS","XLOOKUP","XMATCH",
  // date & time
  "DATE","DATEDIF","DATEVALUE","DAY","DAYS","DAYS360","EDATE","EOMONTH","HOUR","ISOWEEKNUM","MINUTE","MONTH","NETWORKDAYS","NOW","SECOND","TIME","TIMEVALUE","TODAY","WEEKDAY","WEEKNUM","WORKDAY","YEAR","YEARFRAC",
  // financial
  "ACCRINT","ACCRINTM","AMORDEGRC","AMORLINC","COUPDAYBS","COUPDAYS","COUPDAYSNC","COUPNCD","COUPNUM","COUPPCD","CUMIPMT","CUMPRINC","DB","DDB","DISC","DOLLARDE","DOLLARFR","DURATION","EFFECT","FV","FVSCHEDULE","INTRATE","IPMT","IRR","ISPMT","MDURATION","MIRR","NOMINAL","NPER","NPV","ODDFPRICE","ODDFYIELD","ODDLPRICE","ODDLYIELD","PDURATION","PMT","PPMT","PRICE","PRICEDISC","PRICEMAT","PV","RATE","RECEIVED","RRI","SLN","SYD","TBILLEQ","TBILLPRICE","TBILLYIELD","VDB","XIRR","XNPV","YIELD","YIELDDISC","YIELDMAT",
  // information
  "CELL","INFO","ISBLANK","ISERR","ISERROR","ISEVEN","ISFORMULA","ISLOGICAL","ISNA","ISNONTEXT","ISNUMBER","ISODD","ISREF","ISTEXT","N","NA","SHEET","SHEETS","TYPE",
  // database
  "DAVERAGE","DCOUNT","DCOUNTA","DGET","DMAX","DMIN","DPRODUCT","DSTDEV","DSTDEVP","DSUM","DVAR","DVARP",
  // engineering
  "BIN2DEC","BIN2HEX","BIN2OCT","BITAND","BITLSHIFT","BITOR","BITRSHIFT","BITXOR","COMPLEX","CONVERT","DEC2BIN","DEC2HEX","DEC2OCT","DELTA","ERF","ERFC","GESTEP","HEX2BIN","HEX2DEC","HEX2OCT","IMABS","IMAGINARY","IMREAL","OCT2BIN","OCT2DEC","OCT2HEX",
  // web & cube
  "ENCODEURL","FILTERXML","WEBSERVICE","CUBEKPIMEMBER","CUBEMEMBER","CUBEMEMBERPROPERTY","CUBERANKEDMEMBER","CUBESET","CUBESETCOUNT","CUBEVALUE"
]);
function isReserved(name){ return RESERVED.has((name || '').toUpperCase()); }

// Excel function name: letters/digits/dot/underscore, not starting with a digit or dot.
// No forced prefix — the user controls the entire name.
function cleanName(s){ let n=(s||'').replace(/[^A-Za-z0-9_.]/g,''); if(/^[0-9.]/.test(n)) n='_'+n; return n; }
function renderPreview(){
  const friendly = ($('friendlyName').value||'Formula');
  const fnName = cleanName(friendly);
  $('nameHint').innerHTML = isReserved(fnName)
    ? '<span style="color:var(--danger)">⚠ <b>'+esc(fnName)+'</b> is a built-in Excel function — choose a different name.</span>'
    : "In Excel you'll type <b>="+esc(fnName)+"(…)</b>";
  const measure = (measureCombo.getValue()) || 'the value';
  const names = paramNames();
  // tinted formula
  $('formula').innerHTML = `<span class="fn">=${esc(fnName)}</span>(` +
    names.map(n=>`<span class="arg">${esc(n)}</span>`).join(', ') + ')';
  $('explain').innerHTML = `Returns <b>${esc(measure)}</b>` + (names.length? `, filtered by ${esc(names.join(', '))}.` : '.');
  const ex = CURRENT.Parameters.map(p => {
    const s = (typeof p._sample==='string' && p._sample!=='') ? p._sample : null;
    if(s !== null) return p.DataType==='number' ? esc(s) : '"'+esc(s)+'"';
    return p.DataType==='date' ? '"1/1/2025"' : p.DataType==='number' ? '"4000"' : '"East"';
  });
  $('example').innerHTML = names.length
    ? 'e.g. <span class="lit">=' + esc(fnName) + '(' + ex.join(', ') + ')</span>'
    : '';
}

async function saveFunction(){
  const friendly = $('friendlyName').value.trim();
  const measure = measureCombo.getValue();
  if(!measure || !friendly){ showStatus('Pick the number you want and give the formula a name.'); return; }
  const params=[]; let pos=0;
  const kindToFilter = { match:'List', start:'RangeStart', end:'RangeEnd' };
  CURRENT.Parameters.forEach(p=>{
    params.push({
      Name: p.Name || suggestName(p.FieldName, p._kind),
      Position: pos++,
      TableName: p.TableName,
      FieldName: p.FieldName,
      DataType: p.DataType || 'text',
      FilterType: kindToFilter[p._kind] || 'List',
      IsOptional: true
    });
  });
  const fnName = cleanName(friendly);
  if(!fnName){ showStatus('Use letters, numbers, dots or underscores for the formula name.'); return; }
  if(isReserved(fnName)){ showStatus('“'+fnName+'” is a built-in Excel function — please choose a different name.'); return; }
  const dto = { FunctionName:fnName, MeasureName:'['+measure+']',
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
  // each stored parameter becomes one card; map its FilterType to a UI kind
  const filterToKind = { List:'match', RangeStart:'start', RangeEnd:'end' };
  CURRENT.Parameters = (f.Parameters||[]).map(p=>({ ...p, _kind: filterToKind[p.FilterType] || 'match' }));
  showEditor();
  $('friendlyName').value = name;
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
