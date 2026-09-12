(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const state = {initialized:false,demo:false,report:null,properties:[],request:0,activeTab:'event',pathStep:1};
  const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const numeric = value => typeof value === 'number' && Number.isFinite(value);
  const fmtInt = value => numeric(value) ? Math.round(value).toLocaleString('en-US') : '—';
  const fmtPct = value => numeric(value) ? `${(value*100).toFixed(1)}%` : '—';
  const fmtSeconds = value => numeric(value) ? value >= 60 ? `${Math.floor(value/60)}m ${Math.round(value%60)}s` : `${Math.round(value)}s` : '—';
  const fmtMoney = value => numeric(value) ? new Intl.NumberFormat('en-US',{style:'currency',currency:state.report?.property?.currency||'USD',maximumFractionDigits:0}).format(value) : '—';
  const presetLabels = {ecommerce:'Ecommerce purchase',checkout:'Checkout completion',lead:'Lead generation'};
  const stepName = index => state.report?.steps?.find(step=>step.step===Number(index))?.name || `Step ${index}`;

  async function api(url,options={}) {
    const response=await fetch(url,{headers:{Accept:'application/json','Content-Type':'application/json'},...options});
    const data=await response.json().catch(()=>({}));
    if(!response.ok){const error=new Error(data.detail?.message||data.message||'Google Analytics could not return this request.');error.code=data.detail?.code||data.code||'ga4_error';throw error;}
    return data;
  }
  function status(message,type=''){$('faStatus').textContent=message;$('faStatus').className='fa-status'+(type?` ${type}`:'');}
  function clearReport(){state.report=null;$('faReport').hidden=true;$('faWelcome').hidden=false;}
  function dateParams(){
    if($('faPeriod').value!=='custom')return {days:$('faPeriod').value};
    if(!$('faStart').value||!$('faEnd').value)throw new Error('Choose both a start and end date.');
    if($('faStart').value>$('faEnd').value)throw new Error('The start date must be before the end date.');
    return {start_date:$('faStart').value,end_date:$('faEnd').value};
  }
  async function connect(){
    const request=++state.request;clearReport();$('faProperty').disabled=true;$('faRefresh').disabled=true;status('Checking your Google Analytics connection…');
    try{
      const data=await api('/api/ga4/properties');if(request!==state.request)return;
      state.demo=false;state.properties=data.properties||[];$('faDemo').textContent='Explore sample data';
      $('faSource').textContent=data.connected?'CONNECTED':'NOT CONNECTED';$('faSource').className='fa-badge '+(data.connected?'live':'');
      $('faConnect').textContent=data.connected?'Change Google account ↗':'Connect Google Analytics ↗';
      $('faConnectionText').textContent=data.connected?(data.email||'Google Analytics connected'):'Connect a GA4 property to explore ordered event funnels, next actions and aggregate user quality.';
      $('faProperty').innerHTML='<option value="">'+(data.connected?'Select your GA4 property':'Connect an account first')+'</option>'+state.properties.map(item=>`<option value="${esc(item.id)}">${esc(item.name)} · ${esc(item.id)}</option>`).join('');
      $('faProperty').disabled=!state.properties.length;$('faRefresh').disabled=false;
      if(!data.connected){status('Sign in with Google to load your funnel, or explore the clearly labelled sample.');return;}
      if(!state.properties.length){status('No GA4 properties are available to this Google account. Ask an Analytics administrator for Viewer access, then reconnect.','error');return;}
      const selected=state.properties.some(item=>item.id===data.selected_property)?data.selected_property:state.properties.length===1?state.properties[0].id:'';
      $('faProperty').value=selected;
      if(selected)await load();else status('Choose a GA4 property. Its first funnel report will load automatically.');
    }catch(error){if(request!==state.request)return;$('faProperty').innerHTML='<option value="">GA4 connection required</option>';$('faSource').textContent='NOT CONNECTED';$('faSource').className='fa-badge';$('faRefresh').disabled=false;status(error.message,'error');}
  }
  async function load(){
    const request=++state.request,property=$('faProperty').value;
    if(!state.demo&&!property){clearReport();status('Choose a GA4 property to load its funnel.');return;}
    let range;try{range=dateParams();}catch(error){status(error.message,'error');return;}
    clearReport();$('faWelcome').hidden=true;$('faRefresh').disabled=true;$('funnelWorkspaceContainer').setAttribute('aria-busy','true');
    status('Building ordered funnel, paths and user segments…');
    const params=new URLSearchParams({...range,preset:$('faPreset').value,breakdown:$('faBreakdown').value,open_funnel:$('faType').value==='open'?'true':'false',...(state.demo?{sample:'true'}:{property_id:property})});
    try{
      const report=await api('/api/ga4/funnel-report?'+params);if(request!==state.request)return;
      state.report=report;state.pathStep=report.steps?.[0]?.step||1;$('faReport').hidden=false;$('faWelcome').hidden=true;
      $('faSource').textContent=state.demo?'SAMPLE DATA':'GA4 DATA';$('faSource').className='fa-badge '+(state.demo?'sample':'live');
      status(state.demo?'Sample data only — these figures are illustrative and do not belong to your GA4 property.':report.steps?.some(step=>step.users)?'Funnel report updated from Google Analytics 4.':'No users completed this sequence in the selected range. Try an open funnel or another journey.',state.demo?'sample':'');
      render();
      if(!state.demo&&report.property?.id===property){try{await api('/api/ga4/selection',{method:'POST',body:JSON.stringify({property_id:property})});}catch{}}
    }catch(error){if(request!==state.request)return;$('faSource').textContent='REPORT UNAVAILABLE';$('faSource').className='fa-badge';status(error.message,'error');$('faWelcome').hidden=false;}
    finally{if(request===state.request){$('funnelWorkspaceContainer').setAttribute('aria-busy','false');$('faRefresh').disabled=false;}}
  }
  function render(){
    const report=state.report,summary=report.summary||{};
    $('faPeriodLabel').textContent=`Performance: ${report.start_date} – ${report.end_date} · Prior: ${report.previous_start} – ${report.previous_end}`;
    $('faPropertyLabel').textContent=`${report.property?.name||'GA4 property'} · ${report.source==='sample'?'Illustrative sample':`Property ${report.property?.id}`}`;
    const change=summary.conversion_change_pp;
    const delta=numeric(change)?`<span class="fa-delta ${change<0?'down':''}">${change>=0?'+':''}${change.toFixed(2)} pp vs prior</span>`:'No prior comparison';
    $('faKpis').innerHTML=[
      ['ENTRY USERS',fmtInt(summary.entry_users),`${presetLabels[report.preset]||'Journey'} baseline`,''],
      ['COMPLETED',fmtInt(summary.completion_users),`Reached ${esc(report.steps?.at(-1)?.name||'final step')}`,'green'],
      ['FUNNEL CONVERSION',fmtPct(summary.conversion_rate),delta,'blue'],
      ['ABANDONMENTS',fmtInt(summary.total_abandonments),`${stepName(summary.largest_drop_step)} is the largest loss`,'red'],
    ].map(([label,value,note,cls])=>`<article class="fa-kpi ${cls}"><div class="fa-kpi-label">${label}</div><div class="fa-kpi-value">${value}</div><div class="fa-kpi-note">${note}</div></article>`).join('');
    $('faFunnelNote').textContent=`${report.is_open?'Open':'Closed'} funnel · ${report.start_date} to ${report.end_date} · active users completing events in order.`;
    renderFunnel();renderBreakdown();renderPathSteps();renderPath();renderLanding();renderUsers();renderWarnings();switchTab(state.activeTab);
  }
  function renderFunnel(){
    const steps=state.report.steps||[],max=Math.max(1,...steps.map(step=>step.users||0));
    $('faFunnelChart').innerHTML=steps.map(step=>{
      const width=Math.max(step.users?7:0,(step.users||0)/max*100),loss=numeric(step.abandonment_rate)?`${fmtPct(step.abandonment_rate)} abandon · ${fmtInt(step.abandonments)} users`:'Final journey step';
      return `<div class="fa-funnel-row"><div class="fa-step-name"><strong>${step.step}. ${esc(step.name)}</strong><span>${esc(step.event)}</span></div><div class="fa-bar-track"><div class="fa-bar-fill" style="--width:${width.toFixed(1)}%">${width>19?fmtInt(step.users):''}</div></div><div class="fa-step-metric"><strong>${fmtInt(step.users)}</strong><span>${loss}</span></div></div>`;
    }).join('')||'<div class="fa-empty">No funnel steps were returned.</div>';
    const largest=state.report.summary?.largest_drop_step,step=steps.find(item=>item.step===largest);
    $('faDropInsight').innerHTML=step?`<span class="fa-eyebrow">LARGEST LOSS</span><h3>After ${esc(step.name)}</h3><p>${fmtInt(step.abandonments)} active users did not reach the next configured event. Inspect their next actions and compare the selected breakdown before prioritising a UX change.</p><div class="fa-insight-stat"><strong>${fmtPct(step.abandonment_rate)}</strong><span>step abandonment rate reported by GA4</span></div>`:'<span class="fa-eyebrow">FUNNEL STATUS</span><h3>No measured loss</h3><p>GA4 did not return enough ordered users for a bottleneck calculation.</p>';
  }
  function renderBreakdown(){
    const label=state.report.breakdown?.label||'Segment';$('faBreakdownEyebrow').textContent=(state.report.breakdown?.dimension||'BREAKDOWN').replace(/([A-Z])/g,' $1').toUpperCase();$('faBreakdownTitle').textContent=`Conversion by ${label.toLowerCase()}`;
    $('faBreakdownBody').innerHTML=(state.report.breakdown_rows||[]).map(row=>`<tr><td>${esc(row.name)}</td><td>${fmtInt(row.entry_users)}</td><td>${fmtInt(row.completion_users)}</td><td><span class="fa-ratebar"><i style="--rate:${numeric(row.conversion_rate)?Math.min(100,row.conversion_rate*100):0}%"></i><strong>${fmtPct(row.conversion_rate)}</strong></span></td><td class="fa-loss">${row.largest_drop_step?`After ${esc(stepName(row.largest_drop_step))}`:'—'}</td></tr>`).join('')||'<tr><td colspan="5">No breakdown rows were returned for this journey.</td></tr>';
  }
  function renderPathSteps(){
    $('faPathSteps').innerHTML=(state.report.steps||[]).slice(0,-1).map(step=>`<button type="button" class="${step.step===state.pathStep?'active':''}" data-path-step="${step.step}">${step.step}. ${esc(step.name)}</button>`).join('');
    $('faPathSteps').querySelectorAll('button').forEach(button=>button.addEventListener('click',()=>{state.pathStep=Number(button.dataset.pathStep);renderPathSteps();renderPath();}));
  }
  function renderPath(){
    const step=state.report.steps?.find(item=>item.step===state.pathStep),rows=state.report.next_actions?.[String(state.pathStep)]||[];
    if(!step){$('faPathFlow').innerHTML='<div class="fa-empty">Choose a funnel step.</div>';return;}
    const nodes=rows.map(row=>`<div class="fa-path-node"><div><strong>${esc(row.name)}</strong><i style="--share:${numeric(row.share)?row.share*100:0}%"></i><span>${fmtPct(row.share)} of returned next actions</span></div><b>${fmtInt(row.users)}</b></div>`).join('')||'<div class="fa-empty">GA4 returned no next actions between these two funnel steps.</div>';
    $('faPathFlow').innerHTML=`<div class="fa-origin-node"><span>${esc(step.event)}</span><strong>${esc(step.name)}</strong><em>${fmtInt(step.users)} active users reached this step</em></div><div class="fa-flow-line" aria-hidden="true"></div><div class="fa-path-nodes">${nodes}</div>`;
  }
  function renderLanding(){
    $('faLandingBody').innerHTML=(state.report.landing_pages||[]).map(row=>`<tr><td title="${esc(row.path)}">${esc(row.path)}</td><td>${fmtInt(row.sessions)}</td><td>${fmtInt(row.activeUsers)}</td><td>${fmtPct(row.engagementRate)}</td><td>${fmtPct(row.bounceRate)}</td><td>${fmtSeconds(row.averageSessionDuration)}</td><td>${fmtInt(row.keyEvents)}</td></tr>`).join('')||'<tr><td colspan="7">Landing-page metrics are unavailable for this property.</td></tr>';
  }
  function renderUsers(){
    const rows=state.report.user_segments||[];
    $('faUserCards').innerHTML=rows.map(row=>`<article class="fa-user-card"><div class="fa-ring" style="--ring:${numeric(row.engagementRate)?row.engagementRate*100:0}%"><strong>${fmtPct(row.engagementRate)}</strong></div><div><h3>${esc(row.name)} users</h3><div class="fa-user-metrics"><div><span>Active users</span><strong>${fmtInt(row.activeUsers)}</strong></div><div><span>Avg. session</span><strong>${fmtSeconds(row.averageSessionDuration)}</strong></div><div><span>Purchases</span><strong>${fmtInt(row.ecommercePurchases)}</strong></div></div></div></article>`).join('')||'<div class="fa-empty">User segment metrics are unavailable.</div>';
    $('faUserBody').innerHTML=rows.map(row=>`<tr><td>${esc(row.name)} users</td><td>${fmtInt(row.activeUsers)}</td><td>${fmtInt(row.sessions)}</td><td>${fmtInt(row.engagedSessions)}</td><td>${fmtPct(row.engagementRate)}</td><td>${fmtPct(row.bounceRate)}</td><td>${fmtSeconds(row.averageSessionDuration)}</td><td>${numeric(row.eventsPerSession)?row.eventsPerSession.toFixed(1):'—'}</td><td>${fmtInt(row.ecommercePurchases)}</td><td>${fmtMoney(row.totalRevenue)}</td></tr>`).join('')||'<tr><td colspan="10">User segment metrics are unavailable for this property.</td></tr>';
  }
  function renderWarnings(){
    const fixed=['GA4 Funnel Reporting uses the Google Analytics Data API v1alpha Early Preview endpoint; Google may change its schema or availability.','Funnel counts are active users who completed configured events in order; they are not raw event totals.','Path exploration shows the supported next action between configured steps. Exact full-session paths and pseudonymous user-level histories require a separately configured BigQuery export.'];
    $('faWarnings').innerHTML=[...(state.report.warnings||[]),...fixed].map(item=>`<div class="fa-warning">${esc(item)}</div>`).join('');
  }
  function switchTab(tab){
    state.activeTab=tab;
    const map={event:'faEventPanel',path:'faPathPanel',user:'faUserPanel',method:'faMethodPanel'};
    document.querySelectorAll('.fa-tabbar button').forEach(button=>{const active=button.dataset.tab===tab;button.classList.toggle('active',active);button.setAttribute('aria-selected',String(active));});
    Object.entries(map).forEach(([name,id])=>{$(id).hidden=name!==tab;});
  }
  function exportCsv(){
    if(!state.report)return;
    const rows=[['Step','Event','Active users','Completion rate','Abandonments','Abandonment rate'],...(state.report.steps||[]).map(step=>[step.name,step.event,step.users,step.completion_rate??'',step.abandonments,step.abandonment_rate??''])];
    const csv=rows.map(row=>row.map(value=>`"${String(value).replaceAll('"','""')}"`).join(',')).join('\n');
    const link=document.createElement('a');link.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));link.download=`${state.report.source==='sample'?'sample-':''}ga4-${state.report.preset}-funnel.csv`;link.click();URL.revokeObjectURL(link.href);
  }
  function demo(){
    state.demo=!state.demo;$('faDemo').textContent=state.demo?'Exit sample':'Explore sample data';
    if(state.demo){$('faProperty').innerHTML='<option value="sample">Illustrative GA4 property · Sample data</option>';$('faProperty').disabled=true;load();}else connect();
  }
  function syncGlobalDate(start,end){$('faPeriod').value='custom';$('faCustomDates').hidden=false;$('faStart').value=start;$('faEnd').value=end;if(document.body.classList.contains('funnel-analysis-active'))load();}
  function init(){
    if(state.initialized)return;state.initialized=true;
    $('faDemo').addEventListener('click',demo);$('faRefresh').addEventListener('click',load);$('faExport').addEventListener('click',exportCsv);
    $('faProperty').addEventListener('change',load);['faPreset','faBreakdown','faType'].forEach(id=>$(id).addEventListener('change',load));
    $('faPeriod').addEventListener('change',()=>{$('faCustomDates').hidden=$('faPeriod').value!=='custom';if($('faPeriod').value!=='custom')load();});
    $('faStart').addEventListener('change',()=>{if($('faEnd').value)load();});$('faEnd').addEventListener('change',()=>{if($('faStart').value)load();});
    document.querySelectorAll('.fa-tabbar button').forEach(button=>button.addEventListener('click',()=>switchTab(button.dataset.tab)));
    connect();
  }
  window.FunnelAnalysis={init,syncGlobalDate};
  window.initFunnelWorkspace=init;
  document.addEventListener('DOMContentLoaded',()=>{if(document.body.classList.contains('funnel-analysis-active'))init();});
})();
