const tabs=document.querySelectorAll('.tab');
const panels=document.querySelectorAll('.panel');

function activateTab(tabName){
  const targetTab=document.querySelector(`.tab[data-tab="${tabName}"]`);
  const targetPanel=document.getElementById(tabName);
  if(!targetTab||!targetPanel)return;
  tabs.forEach(t=>t.classList.remove('active'));
  panels.forEach(p=>p.classList.remove('active'));
  targetTab.classList.add('active');
  targetPanel.classList.add('active');
}

tabs.forEach(tab=>tab.addEventListener('click',()=>{
  activateTab(tab.dataset.tab);
  window.scrollTo({top:0,behavior:'smooth'});
}));

window.addEventListener('load',()=>{
  activateTab('home');
  setTimeout(()=>window.scrollTo({top:0,left:0,behavior:'auto'}),0);
  loadExcelData();
});

document.querySelectorAll('[data-target-tab][data-scroll-target]').forEach(el=>{
  el.addEventListener('click',()=>{
    activateTab(el.dataset.targetTab);
    setTimeout(()=>{
      const target=document.getElementById(el.dataset.scrollTarget);
      if(target){
        target.scrollIntoView({behavior:'smooth',block:'start'});
        target.classList.add('highlight');
        setTimeout(()=>target.classList.remove('highlight'),1600);
      }
    },80);
  });
});

const toast=document.getElementById('toast');
document.querySelectorAll('.copy-btn').forEach(btn=>btn.addEventListener('click',async()=>{
  try{await navigator.clipboard.writeText(btn.dataset.copy);}catch(e){}
  toast.classList.add('show');
  setTimeout(()=>toast.classList.remove('show'),1000);
}));

function esc(value){
  return String(value ?? '').replace(/[&<>"']/g, s=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}

function sheetToRows(workbook, name){
  const sheet=workbook.Sheets[name];
  if(!sheet)return [];
  return XLSX.utils.sheet_to_json(sheet,{defval:''});
}

function paymentBadge(status){
  const raw=String(status||'Pending').trim();
  const cls=raw.toLowerCase();
  return `<span class="payment ${esc(cls)}">${esc(raw)}</span>`;
}

function statusBadge(status){
  const raw=String(status||'Pending').trim();
  const low=raw.toLowerCase();
  let cls='neutral';
  if(!raw||low.includes('pending')) cls='pending';
  else if(low.includes('complete')||low==='paid') cls='complete';
  else if(low.includes('individual')) cls='individual';
  else if(low.includes('see')) cls='linkout';
  return `<span class="status-badge ${cls}">${esc(raw||'Pending')}</span>`;
}

function maybeLink(value, type){
  const v=String(value||'');
  if(!v)return '';
  if(type==='email') return `<a href="mailto:${esc(v)}">${esc(v)}</a>`;
  if(type==='phone') return `<a href="tel:${esc(v)}">${esc(v)}</a>`;
  return esc(v);
}

function money(v){
  if(v===''||v==null)return '';
  const num=Number(v);
  if(Number.isFinite(num))return `$${num.toFixed(2)}`;
  return esc(v);
}

function renderRows(id, rows, columns){
  const el=document.getElementById(id);
  if(!el)return;
  el.innerHTML=rows.map(row=>{
    const cells=columns.map(col=>{
      const value=row[col.key] ?? '';
      let content='';
      if(col.type==='payment') content=paymentBadge(value);
      else if(col.type==='status') content=statusBadge(value);
      else if(col.type==='email') content=maybeLink(value,'email');
      else if(col.type==='phone') content=maybeLink(value,'phone');
      else if(col.type==='money') content=money(value);
      else content=esc(value);
      return `<td>${content}</td>`;
    }).join('');
    return `<tr>${cells}</tr>`;
  }).join('');
}

function paceToSeconds(pace){
  const parts=String(pace||'').split(':').map(Number);
  if(parts.length!==2||parts.some(Number.isNaN))return null;
  return parts[0]*60+parts[1];
}

function secondsToPace(sec){
  if(!Number.isFinite(sec))return '—';
  const m=Math.floor(sec/60);
  const s=Math.round(sec%60);
  return `${m}:${String(s).padStart(2,'0')}`;
}

function renderTeam(id, avgId, rows){
  const el=document.getElementById(id);
  const avgEl=document.getElementById(avgId);
  if(!el)return;
  const active=rows.filter(r=>String(r.Name||'').toUpperCase()!=='DRAFT' && r['Runner #']!=='');
  el.innerHTML=active.map(r=>`<tr><td class="runner-num">${esc(r['Runner #'])}</td><td>${esc(r.Name)}</td><td>${esc(r.Ave)}</td></tr>`).join('')
    + `<tr class="volunteer-row"><td></td><td>Volunteer</td><td>Pending</td></tr>`;
  const secs=active.map(r=>paceToSeconds(r.Ave)).filter(v=>v!=null);
  if(avgEl) avgEl.textContent=secs.length ? secondsToPace(secs.reduce((a,b)=>a+b,0)/secs.length) : '—';
}

function renderChecklist(rows){
  const host=document.getElementById('checklist-sections');
  if(!host)return;
  const areas=[...new Set(rows.map(r=>r.Area).filter(Boolean))];
  host.innerHTML=areas.map(area=>{
    const areaRows=rows.filter(r=>r.Area===area);
    const body=areaRows.map(r=>`<tr><td>${esc(r.Item)}</td><td>${esc(r.Req)}</td><td>${statusBadge(r.Status)}</td><td>${esc(r.Owner)}</td><td>${esc(r.Qty)}</td></tr>`).join('');
    return `<details class="check-section" open><summary>${esc(area)} <span>${areaRows.length} items</span></summary><div class="table-scroll compact-list"><table><thead><tr><th>Item</th><th>Req</th><th>Status</th><th>Owner</th><th>Qty</th></tr></thead><tbody>${body}</tbody></table></div></details>`;
  }).join('');
}

function renderLinks(rows){
  const host=document.getElementById('links-cards');
  if(!host || !rows.length)return;
  const icons={"Trail Maps":"🗺️","Runner's Guide":"📖"};
  host.innerHTML=rows.map(r=>`<a class="link-card" href="${esc(r.URL)}" target="_blank" rel="noopener"><span>${icons[r.Name]||'🔗'}</span><strong>${esc(r.Name)}</strong><small>Link</small></a>`).join('');
}

async function loadExcelData(){
  const status=document.getElementById('data-status');
  try{
    if(typeof XLSX==='undefined') throw new Error('XLSX library did not load');
    const response=await fetch('Relay2026_Data.xlsx?cache=' + Date.now());
    if(!response.ok) throw new Error('Could not load Relay2026_Data.xlsx');
    const buffer=await response.arrayBuffer();
    const workbook=XLSX.read(buffer,{type:'array'});

    const runners=sheetToRows(workbook,'Runners');
    const sherpas=sheetToRows(workbook,'Sherpas');
    const teams=sheetToRows(workbook,'Teams');
    const emergency=sheetToRows(workbook,'Emergency');
    const tent=sheetToRows(workbook,'TentList');
    const carpool=sheetToRows(workbook,'Carpool');
    const checklist=sheetToRows(workbook,'Checklist');
    const links=sheetToRows(workbook,'Links');

    renderRows('runners-body',runners,[{key:'#'},{key:'Name'},{key:'Payment',type:'payment'},{key:'Amount',type:'money'},{key:'Initials'},{key:'Average'},{key:'Email',type:'email'},{key:'Phone',type:'phone'}]);
    renderRows('sherpas-body',sherpas,[{key:'Name'},{key:'Email',type:'email'},{key:'Phone',type:'phone'},{key:'Role'}]);
    renderRows('emergency-body',emergency,[{key:'#'},{key:'Runner'},{key:'Emergency Contact'},{key:'Relation'},{key:'Phone',type:'phone'}]);
    renderRows('tent-body',tent,[{key:'Item'},{key:'Required'},{key:'Status',type:'status'},{key:'Who'},{key:'Total'}]);
    renderRows('carpool-body',carpool,[{key:'Item'},{key:'Status',type:'status'},{key:'Who'},{key:'Total'}]);
    renderChecklist(checklist);
    renderLinks(links);

    renderTeam('team-no-egrets','avg-no-egrets',teams.filter(r=>r.Team==='No Egrets'));
    renderTeam('team-still-full','avg-still-full',teams.filter(r=>r.Team==='Still Full of Egrets'));

    const paid=runners.filter(r=>String(r.Payment).toLowerCase()==='paid').length;
    const pending=runners.filter(r=>String(r.Payment).toLowerCase()==='pending').length;
    document.getElementById('metric-runners').textContent=`${runners.length} / ${runners.length}`;
    document.getElementById('metric-paid').textContent=`${paid} Paid`;
    document.getElementById('metric-pending').textContent=`${pending} Pending`;

    const complete=checklist.filter(r=>String(r.Status).toLowerCase().includes('complete')).length;
    document.getElementById('metric-checklist').textContent=`${complete} / ${checklist.length}`;

    if(status) status.textContent='Data loaded from Relay2026_Data.xlsx';
  }catch(error){
    console.error(error);
    if(status) status.textContent='Could not load Excel data. Check that Relay2026_Data.xlsx is uploaded to GitHub.';
  }
}



/* OFFICIAL v8: Excel data source */
const RELAY_EXCEL_FILE = 'Relay2026_Data.xlsx';

function relayEsc(value){
  return String(value ?? '').replace(/[&<>"']/g, s=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}
function relayMoney(v){
  if(v===''||v==null)return '';
  const num=Number(v);
  return Number.isFinite(num) ? `$${num.toFixed(2)}` : relayEsc(v);
}
function relayPaymentBadge(status){
  const raw=String(status||'Pending').trim();
  return `<span class="payment ${relayEsc(raw.toLowerCase())}">${relayEsc(raw)}</span>`;
}
function relayStatusBadge(status){
  const raw=String(status||'Pending').trim();
  const low=raw.toLowerCase();
  let cls='neutral';
  if(!raw || low.includes('pending')) cls='pending';
  else if(low.includes('complete') || low==='paid') cls='complete';
  else if(low.includes('individual')) cls='individual';
  else if(low.includes('see')) cls='linkout';
  return `<span class="status-badge ${cls}">${relayEsc(raw||'Pending')}</span>`;
}
function relayLink(value,type){
  const v=String(value||'');
  if(!v)return '';
  if(type==='email') return `<a href="mailto:${relayEsc(v)}">${relayEsc(v)}</a>`;
  if(type==='phone') return `<a href="tel:${relayEsc(v)}">${relayEsc(v)}</a>`;
  return relayEsc(v);
}
function relayRows(workbook, sheetName){
  const sheet=workbook.Sheets[sheetName];
  return sheet ? XLSX.utils.sheet_to_json(sheet,{defval:''}) : [];
}
function relayRenderRows(id, rows, columns){
  const el=document.getElementById(id);
  if(!el)return;
  el.innerHTML=rows.map(row=>`<tr>${columns.map(col=>{
    const v=row[col.key] ?? '';
    let c='';
    if(col.type==='payment') c=relayPaymentBadge(v);
    else if(col.type==='status') c=relayStatusBadge(v);
    else if(col.type==='money') c=relayMoney(v);
    else if(col.type==='email') c=relayLink(v,'email');
    else if(col.type==='phone') c=relayLink(v,'phone');
    else c=relayEsc(v);
    return `<td>${c}</td>`;
  }).join('')}</tr>`).join('');
}
function relayPaceToSeconds(pace){
  const parts=String(pace||'').split(':').map(Number);
  if(parts.length!==2 || parts.some(Number.isNaN))return null;
  return parts[0]*60+parts[1];
}
function relaySecondsToPace(sec){
  if(!Number.isFinite(sec))return '—';
  const m=Math.floor(sec/60), s=Math.round(sec%60);
  return `${m}:${String(s).padStart(2,'0')}`;
}
function relayRenderTeam(id, avgId, rows){
  const el=document.getElementById(id);
  const avgEl=document.getElementById(avgId);
  if(!el)return;
  const active=rows.filter(r=>String(r.Name||'').toUpperCase()!=='DRAFT' && String(r['#']||r['Runner #']||'')!=='');
  el.innerHTML=active.map(r=>`<tr><td class="runner-num">${relayEsc(r['#']||r['Runner #'])}</td><td>${relayEsc(r.Name)}</td><td>${relayEsc(r.Ave)}</td></tr>`).join('')
    + `<tr class="volunteer-row"><td></td><td>Volunteer</td><td>Pending</td></tr>`;
  const secs=active.map(r=>relayPaceToSeconds(r.Ave)).filter(v=>v!=null);
  if(avgEl) avgEl.textContent=secs.length ? relaySecondsToPace(secs.reduce((a,b)=>a+b,0)/secs.length) : '—';
}
function relayRenderChecklist(rows){
  const host=document.getElementById('checklist-sections');
  if(!host)return;
  const areas=[...new Set(rows.map(r=>r.Area).filter(Boolean))];
  host.innerHTML=areas.map(area=>{
    const areaRows=rows.filter(r=>r.Area===area);
    const body=areaRows.map(r=>`<tr><td>${relayEsc(r.Item)}</td><td>${relayEsc(r.Req)}</td><td>${relayStatusBadge(r.Status)}</td><td>${relayEsc(r.Owner)}</td><td>${relayEsc(r.Qty)}</td></tr>`).join('');
    return `<details class="check-section" open><summary>${relayEsc(area)} <span>${areaRows.length} items</span></summary><div class="table-scroll compact-list"><table><thead><tr><th>Item</th><th>Req</th><th>Status</th><th>Owner</th><th>Qty</th></tr></thead><tbody>${body}</tbody></table></div></details>`;
  }).join('');
}
function relayRenderLinks(rows){
  const host=document.getElementById('links-cards');
  if(!host || !rows.length)return;
  const icons={"Trail Maps":"🗺️","Runner's Guide":"📖"};
  host.innerHTML=rows.map(r=>`<a class="link-card" href="${relayEsc(r.URL)}" target="_blank" rel="noopener"><span>${icons[r.Name]||'🔗'}</span><strong>${relayEsc(r.Name)}</strong><small>Link</small></a>`).join('');
}
async function loadRelayExcelData(){
  const status=document.getElementById('data-status');
  try{
    if(typeof XLSX==='undefined') throw new Error('XLSX library did not load');
    const response=await fetch(RELAY_EXCEL_FILE + '?v=' + Date.now());
    if(!response.ok) throw new Error('Could not load ' + RELAY_EXCEL_FILE);
    const buffer=await response.arrayBuffer();
    const workbook=XLSX.read(buffer,{type:'array'});

    const runners=relayRows(workbook,'Runners');
    const sherpas=relayRows(workbook,'Sherpas');
    const teams=relayRows(workbook,'Teams');
    const emergency=relayRows(workbook,'Emergency');
    const tent=relayRows(workbook,'TentList');
    const carpool=relayRows(workbook,'Carpool');
    const checklist=relayRows(workbook,'Checklist');
    const links=relayRows(workbook,'Links');

    relayRenderRows('runners-body',runners,[{key:'#'},{key:'Name'},{key:'Payment',type:'payment'},{key:'Amount',type:'money'},{key:'Initials'},{key:'Average'},{key:'Email',type:'email'},{key:'Phone',type:'phone'}]);
    relayRenderRows('sherpas-body',sherpas,[{key:'Name'},{key:'Email',type:'email'},{key:'Phone',type:'phone'},{key:'Role'}]);
    relayRenderRows('emergency-body',emergency,[{key:'#'},{key:'Runner'},{key:'Emergency Contact'},{key:'Relation'},{key:'Phone',type:'phone'}]);
    relayRenderRows('tent-body',tent,[{key:'Item'},{key:'Required'},{key:'Status',type:'status'},{key:'Who'},{key:'Total'}]);
    relayRenderRows('carpool-body',carpool,[{key:'Item'},{key:'Status',type:'status'},{key:'Who'},{key:'Total'}]);
    relayRenderChecklist(checklist);
    relayRenderLinks(links);

    relayRenderTeam('team-no-egrets','avg-no-egrets',teams.filter(r=>r.Team==='No Egrets'));
    relayRenderTeam('team-still-full','avg-still-full',teams.filter(r=>r.Team==='Still Full of Egrets'));

    const paid=runners.filter(r=>String(r.Payment).toLowerCase()==='paid').length;
    const pending=runners.filter(r=>String(r.Payment).toLowerCase()==='pending').length;
    const metricRunners=document.getElementById('metric-runners');
    const metricPaid=document.getElementById('metric-paid');
    const metricPending=document.getElementById('metric-pending');
    const metricChecklist=document.getElementById('metric-checklist');
    if(metricRunners) metricRunners.textContent=`${runners.length} / ${runners.length}`;
    if(metricPaid) metricPaid.textContent=`${paid} Paid`;
    if(metricPending) metricPending.textContent=`${pending} Pending`;
    if(metricChecklist) {
      const complete=checklist.filter(r=>String(r.Status).toLowerCase().includes('complete')).length;
      metricChecklist.textContent=`${complete} / ${checklist.length}`;
    }

    if(status) status.textContent='Data loaded from Relay2026_Data.xlsx';
  }catch(error){
    console.error(error);
    if(status) status.textContent='Could not load Excel data. Check that Relay2026_Data.xlsx is uploaded to GitHub root.';
  }
}
window.addEventListener('load', loadRelayExcelData);
