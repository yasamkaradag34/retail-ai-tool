(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const pageSize = 15;
  const state = {initialized:false,demo:false,report:null,page:0,sort:'clicks',desc:true,tab:'products',request:0,controller:null};
  const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const numeric = value => typeof value === 'number' && Number.isFinite(value);
  const fmtNumber = value => numeric(value) ? value.toLocaleString('en-US',{maximumFractionDigits:value % 1 ? 2 : 0}) : '—';
  const fmtPercent = value => numeric(value) ? `${value.toFixed(1)}%` : '—';
  const fmtRate = value => numeric(value) ? `${(value*100).toFixed(2)}%` : '—';
  const fmtMoney = money => numeric(money?.amount) ? new Intl.NumberFormat('en-US',{style:'currency',currency:money.currency || 'USD',maximumFractionDigits:2}).format(money.amount) : '—';
  const statusLabel = value => ({above:'Above benchmark',below:'Below benchmark',at_market:'Competitively priced',unavailable:'No benchmark',in_stock:'In stock',out_of_stock:'Out of stock',preorder:'Preorder',backorder:'Backorder',unknown:'Unknown'})[value] || String(value || 'Unknown').replaceAll('_',' ');
  async function api(url, options={}) {
    const response = await fetch(url,{headers:{'Accept':'application/json','Content-Type':'application/json'},...options});
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.detail?.message || data.message || 'Merchant Center could not return this request.');
    return data;
  }
  function status(message,type='') { $('spStatus').textContent=message; $('spStatus').className='sp-status'+(type?` ${type}`:''); }
  function clearReport() { state.report=null; $('spReport').hidden=true; $('spWelcome').hidden=false; }
  function dateParams() {
    if ($('spPeriod').value !== 'custom') return {days:$('spPeriod').value};
    if (!$('spStart').value || !$('spEnd').value) throw new Error('Choose both a start and end date.');
    if ($('spStart').value > $('spEnd').value) throw new Error('The start date must be before the end date.');
    return {start_date:$('spStart').value,end_date:$('spEnd').value};
  }
  function delta(current,previous,{rate=false,inverse=false}={}) {
    if (!numeric(current)||!numeric(previous)) return '<span class="sp-delta">No prior comparison</span>';
    let value,label;
    if (rate) { value=(current-previous)*100; label=`${value>=0?'+':''}${value.toFixed(2)} pp vs prior period`; }
    else if (previous===0) return '<span class="sp-delta">No comparable prior value</span>';
    else { value=(current-previous)/Math.abs(previous)*100; label=`${value>=0?'+':''}${value.toFixed(1)}% vs prior period`; }
    const good=inverse?value<0:value>0;
    return `<span class="sp-delta ${value===0?'':good?'up':'down'}">${esc(label)}</span>`;
  }
  async function connect() {
    state.controller?.abort(); const request=++state.request;
    clearReport(); $('spAccount').disabled=true; $('spRefresh').disabled=true;
    status('Checking your Merchant Center connection…');
    try {
      const data=await api('/api/merchant/accounts'); if(request!==state.request)return;
      state.demo=false; $('spDemo').textContent='Explore sample data';
      $('spSource').textContent=data.connected?'CONNECTED':'NOT CONNECTED'; $('spSource').className='sp-badge '+(data.connected?'live':'');
      $('spConnect').textContent=data.connected?'Change Google account ↗':'Connect Merchant Center ↗';
      $('spConnectionText').textContent=data.connected?(data.email||'Google account connected'):'Connect your Merchant Center account to analyse product availability, Shopping performance and market pricing.';
      $('spAccount').innerHTML='<option value="">'+(data.connected?'Select a Merchant Center account':'Connect an account first')+'</option>'+(data.accounts||[]).map(item=>`<option value="${esc(item.id)}">${esc(item.name)} · ${esc(item.id)}</option>`).join('');
      $('spAccount').disabled=!(data.accounts||[]).length; $('spRefresh').disabled=false;
      if(!data.connected){status('Sign in with Google to load Merchant Center reports, or explore the clearly labelled sample.');return;}
      if(!data.accounts?.length){status('No Merchant Center accounts were returned. Reconnect with Merchant Center permission and confirm your account access role.','error');return;}
      const selected=data.accounts.some(item=>item.id===data.selected_account)?data.selected_account:data.accounts.length===1?data.accounts[0].id:'';
      $('spAccount').value=selected;
      if(selected) await load(); else status('Choose a Merchant Center account above. Its report will load automatically.');
    } catch(error) {
      if(request!==state.request)return;
      $('spAccount').innerHTML='<option value="">Merchant permission required</option>'; $('spRefresh').disabled=false;
      status(error.message,'error');
    }
  }
  async function load() {
    state.controller?.abort(); const request=++state.request; const account=$('spAccount').value;
    if(!state.demo&&!account){clearReport();status('Choose a Merchant Center account to load its report.');return;}
    let range; try{range=dateParams();}catch(error){status(error.message,'error');return;}
    const controller=new AbortController(); state.controller=controller; clearReport(); $('spWelcome').hidden=true; $('spRefresh').disabled=true;
    $('stockWorkspaceContainer').setAttribute('aria-busy','true'); $('spSource').textContent=state.demo?'SAMPLE DATA':'CONNECTED';
    status('Loading catalog, market pricing and Shopping performance…');
    const params=new URLSearchParams({...range,...(state.demo?{sample:'true'}:{account_id:account})});
    const timer=setTimeout(()=>controller.abort(),100000);
    try {
      const report=await api('/api/merchant/insights?'+params,{signal:controller.signal}); if(request!==state.request)return;
      state.report=report; state.page=0; $('spReport').hidden=false; $('spWelcome').hidden=true; $('spExport').disabled=!report.rows.length;
      $('spSource').textContent=state.demo?'SAMPLE DATA':'MERCHANT DATA'; $('spSource').className='sp-badge '+(state.demo?'sample':'live');
      status(state.demo?'Sample data only — these illustrative figures do not belong to your store. Connect Merchant Center to see your own reports.':report.rows.length?'Merchant report updated. Catalog and benchmark values are current snapshots.':'No products were returned. Check the account selection and Merchant Center product feed.',state.demo?'sample':'');
      render();
      if(!state.demo&&state.report?.account.id===account){try{await api('/api/merchant/selection',{method:'POST',body:JSON.stringify({account_id:account})});}catch{} }
    } catch(error) {
      if(request!==state.request)return;
      $('spSource').textContent='REPORT UNAVAILABLE'; $('spSource').className='sp-badge';
      status(error.name==='AbortError'?'Merchant Center took too long to respond. Please refresh the report.':error.message,'error'); $('spWelcome').hidden=false;
    } finally {clearTimeout(timer);if(request===state.request){$('stockWorkspaceContainer').setAttribute('aria-busy','false');$('spRefresh').disabled=false;}}
  }
  function sortValue(row,key){
    if(key==='title')return row.title;
    if(key==='availability'||key==='price_status')return row[key];
    if(key==='price')return row.price?.amount;
    if(key==='benchmark_price')return row.benchmark_price?.amount;
    if(key==='price_gap_percent')return row.price_gap_percent;
    return row.performance?.[key];
  }
  function filteredRows(){
    const query=$('spSearch').value.trim().toLocaleLowerCase(),price=$('spPriceFilter').value,stock=$('spStockFilter').value;
    return (state.report?.rows||[]).filter(row=>(!query||`${row.title} ${row.brand} ${row.offer_id} ${row.category}`.toLocaleLowerCase().includes(query))&&(!price||row.price_status===price)&&(!stock||row.availability===stock)).sort((a,b)=>{
      const left=sortValue(a,state.sort),right=sortValue(b,state.sort);
      if(typeof left==='string')return left.localeCompare(right||'')*(state.desc?-1:1);
      if(!numeric(left))return numeric(right)?1:0;if(!numeric(right))return-1;
      return(left-right)*(state.desc?-1:1);
    });
  }
  function productHead(){
    const columns=[['title','Product'],['availability','Availability'],['price_status','Price position'],['price','Your price'],['benchmark_price','Benchmark'],['price_gap_percent','Price gap'],['clicks','Shopping clicks'],['conversions','Free-listing conv.'],['conversion_value','Free-listing value']];
    return '<tr>'+columns.map(([key,label])=>`<th scope="col" aria-sort="${state.sort===key?(state.desc?'descending':'ascending'):'none'}"><button type="button" data-sp-sort="${key}">${label} <span aria-hidden="true">${state.sort===key?(state.desc?'↓':'↑'):'↕'}</span></button></th>`).join('')+'</tr>';
  }
  function renderProducts(){
    const rows=filteredRows(),pages=Math.max(1,Math.ceil(rows.length/pageSize));state.page=Math.min(state.page,pages-1);const subset=rows.slice(state.page*pageSize,(state.page+1)*pageSize);
    $('spProductHead').innerHTML=productHead();
    $('spProductBody').innerHTML=subset.length?subset.map(row=>`<tr><td class="sp-entity"><strong>${esc(row.title)}</strong><small>${esc(row.brand)} · ${esc(row.offer_id)} · ${esc(row.category)}</small></td><td><span class="sp-pill ${esc(row.availability)}">${esc(statusLabel(row.availability))}</span></td><td><span class="sp-pill ${esc(row.price_status)}">${esc(statusLabel(row.price_status))}</span></td><td>${fmtMoney(row.price)}</td><td>${fmtMoney(row.benchmark_price)}${row.benchmark_country?`<small class="sp-country"> ${esc(row.benchmark_country)}</small>`:''}</td><td>${fmtPercent(row.price_gap_percent)}</td><td>${fmtNumber(row.performance.clicks)}${delta(row.performance.clicks,row.previous.clicks)}</td><td>${fmtNumber(row.performance.conversions)}${delta(row.performance.conversions,row.previous.conversions)}</td><td>${fmtMoney({amount:row.performance.conversion_value,currency:row.performance.conversion_currency})}${delta(row.performance.conversion_value,row.previous.conversion_value)}</td></tr>`).join(''):'<tr><td class="sp-empty" colspan="9">No products match these filters.</td></tr>';
    $('spRowCount').textContent=`${rows.length.toLocaleString('en-US')} ${rows.length===1?'product':'products'}`;$('spPage').textContent=`${state.page+1} / ${pages}`;$('spPrevious').disabled=state.page===0;$('spNext').disabled=state.page+1>=pages;
    $('spTableFootnote').textContent=`${state.report.list_complete?'All returned rows':'Partial list — export contains loaded rows only'} · Market benchmark threshold ±${state.report.definitions.at_market_threshold_percent}%`;
  }
  function renderBrands(){
    $('spBrandList').innerHTML=(state.report.brands||[]).map(row=>{
      const total=row.benchmarked||0,below=total?row.below/total*100:0,at=total?row.at_market/total*100:0,above=total?row.above/total*100:0;
      return `<article class="sp-brand"><div class="sp-brand-name"><strong>${esc(row.brand)}</strong><small>${fmtNumber(row.benchmarked)} of ${fmtNumber(row.products)} products benchmarked</small></div><div><div class="sp-distribution" aria-label="${below.toFixed(0)}% below, ${at.toFixed(0)}% competitive, ${above.toFixed(0)}% above"><span class="below" style="width:${below}%"></span><span class="at" style="width:${at}%"></span><span class="above" style="width:${above}%"></span></div><div class="sp-legend"><span>${below.toFixed(0)}% below</span><span>${at.toFixed(0)}% competitive</span><span>${above.toFixed(0)}% above</span></div></div><div class="sp-brand-demand"><strong>${fmtNumber(row.clicks)} clicks</strong><br>${fmtNumber(row.impressions)} impressions</div></article>`;
    }).join('')||'<div class="sp-empty">Brand distribution will appear when product rows are available.</div>';
  }
  function actionCard(row,type){
    const configs={oos:['OUT-OF-STOCK DEMAND','Restore availability',`${fmtNumber(row.performance.clicks)} Shopping clicks were recorded in the selected period while the current catalog status is out of stock.`],above:['PRICING REVIEW','Review price premium',`Your price is ${fmtPercent(row.price_gap_percent)} above Google’s current market benchmark.`],below:['MARGIN OPPORTUNITY','Validate the discount',`Your price is ${fmtPercent(Math.abs(row.price_gap_percent))} below benchmark. Check whether demand supports recovering margin.`],suggestion:['GOOGLE PRICE INSIGHT',`${row.effectiveness.toLowerCase()} effectiveness suggestion`,`Google suggests ${fmtMoney(row.suggested_price)} for this eligible product.`]};
    const [badge,title,copy]=configs[type];
    return `<article class="sp-action"><span>${esc(badge)}</span><h4>${esc(row.title)}</h4><p>${esc(copy)}</p><dl><div><dt>Offer ID</dt><dd>${esc(row.offer_id)}</dd></div><div><dt>Selected-period clicks</dt><dd>${fmtNumber(row.performance.clicks)}</dd></div><div><dt>Your price</dt><dd>${fmtMoney(row.price)}</dd></div><div><dt>Benchmark</dt><dd>${fmtMoney(row.benchmark_price)}</dd></div></dl></article>`;
  }
  function renderActions(){
    const o=state.report.opportunities,groups=[...(o.out_of_stock_demand||[]).slice(0,2).map(row=>[row,'oos']),...(o.above_benchmark||[]).slice(0,2).map(row=>[row,'above']),...(o.below_benchmark||[]).slice(0,2).map(row=>[row,'below']),...(o.price_suggestions||[]).slice(0,2).map(row=>[row,'suggestion'])];
    $('spActions').innerHTML=groups.length?groups.map(([row,type])=>actionCard(row,type)).join(''):'<div class="sp-empty">No actionable price or availability signals were returned for this account.</div>';$('spActionCount').textContent=groups.length||'';
  }
  function renderCoverage(){
    const labels={catalog:'Catalog & availability',price_competitiveness:'Market benchmark',performance:'Shopping performance',previous_performance:'Prior-period comparison',price_suggestions:'Google price suggestions'};
    $('spCapabilities').innerHTML=Object.entries(labels).map(([key,label])=>`<article class="sp-capability ${state.report.capabilities[key]?'available':''}"><strong>${label}</strong><span>${state.report.capabilities[key]?'AVAILABLE':'NOT RETURNED'}</span></article>`).join('');
    $('spWarnings').innerHTML=(state.report.warnings||[]).map(warning=>`<div class="sp-warning">${esc(warning)}</div>`).join('');
  }
  function render(){
    const r=state.report,s=r.summary,p=s.performance,prior=s.previous_performance,positions=s.price_positions;
    $('spReportScope').textContent=`Performance: ${r.start_date} – ${r.end_date} · Prior: ${r.previous_start} – ${r.previous_end}`;
    $('spFreshness').textContent=`${esc(r.account.name)} · ${state.demo?'Illustrative sample':'Updated '+new Date(r.fetched_at).toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'})}`;
    const kpis=[['Catalog products',fmtNumber(s.catalog_products),'Current Merchant product snapshot',''],['Benchmark coverage',numeric(s.benchmark_coverage_rate)?fmtRate(s.benchmark_coverage_rate):'—',`${fmtNumber(s.benchmark_products)} products with a market benchmark`,''],['Shopping clicks',fmtNumber(p.clicks),'Selected date range',delta(p.clicks,prior.clicks)],['Out of stock',fmtNumber(s.out_of_stock),'Current catalog status','']];
    $('spKpis').innerHTML=kpis.map(([label,value,note,change])=>`<article class="sp-kpi"><span class="sp-kpi-label">${label}</span><strong class="sp-kpi-value">${value}</strong>${change}<small>${note}</small></article>`).join('');
    $('spBelowCount').textContent=fmtNumber(positions.below.count);$('spAtCount').textContent=fmtNumber(positions.at_market.count);$('spAboveCount').textContent=fmtNumber(positions.above.count);
    $('spBelowGap').textContent=numeric(positions.below.average_gap_percent)?`Average ${Math.abs(positions.below.average_gap_percent).toFixed(1)}% below benchmark`:'Your price is lower';
    $('spAboveGap').textContent=numeric(positions.above.average_gap_percent)?`Average ${positions.above.average_gap_percent.toFixed(1)}% above benchmark`:'Your price is higher';
    renderProducts();renderBrands();renderActions();renderCoverage();setTab(state.tab);
  }
  function setTab(tab){
    state.tab=tab;
    document.querySelectorAll('[data-sp-tab]').forEach(button=>{const active=button.dataset.spTab===tab;button.classList.toggle('active',active);button.setAttribute('aria-selected',active);button.tabIndex=active?0:-1;});
    for(const key of ['products','brands','actions','coverage']) $('sp'+key[0].toUpperCase()+key.slice(1)+'Panel').hidden=key!==tab;
    $('spExport').hidden=tab!=='products';
  }
  function exportCSV(){
    if(!state.report)return;const cell=value=>{let s=String(value??'');if(/^[=+@\-\t\r]/.test(s))s="'"+s;return'"'+s.replace(/"/g,'""')+'"';};
    const header=['source','merchant_account_id','start_date','end_date','offer_id','title','brand','category','availability','listing_status','price','currency','benchmark_price','benchmark_currency','benchmark_country','price_status','price_gap_percent','clicks','impressions','ctr','conversions','conversion_rate','conversion_value','conversion_currency','suggested_price','suggestion_effectiveness'];
    const rows=filteredRows().map(row=>[state.report.source,state.report.account.id,state.report.start_date,state.report.end_date,row.offer_id,row.title,row.brand,row.category,row.availability,row.listing_status,row.price.amount,row.price.currency,row.benchmark_price.amount,row.benchmark_price.currency,row.benchmark_country,row.price_status,row.price_gap_percent,row.performance.clicks,row.performance.impressions,row.performance.click_through_rate,row.performance.conversions,row.performance.conversion_rate,row.performance.conversion_value,row.performance.conversion_currency,row.suggested_price.amount,row.effectiveness]);
    const blob=new Blob(['\ufeff'+[header,...rows].map(row=>row.map(cell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8;'}),url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=`${state.demo?'SAMPLE-':''}merchant-stock-price-${state.report.start_date}.csv`;link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  }
  function init(){
    if(state.initialized)return;state.initialized=true;
    $('spPeriod').addEventListener('change',()=>{$('spCustomDates').hidden=$('spPeriod').value!=='custom';if($('spPeriod').value!=='custom'&&(state.report||$('spAccount').value))load();});
    $('spRefresh').addEventListener('click',()=>!state.demo&&!$('spAccount').value?connect():load());$('spAccount').addEventListener('change',load);
    $('spDemo').addEventListener('click',()=>{if(state.demo){connect();return;}state.controller?.abort();state.demo=true;$('spDemo').textContent='Exit sample';$('spAccount').innerHTML='<option value="sample">Sample Merchant Center store · Illustrative data</option>';$('spAccount').disabled=true;$('spConnectionText').textContent='Exploring an illustrative Merchant Center store';load();});
    for(const id of ['spSearch','spPriceFilter','spStockFilter']) $(id).addEventListener(id==='spSearch'?'input':'change',()=>{state.page=0;if(state.report)renderProducts();});
    $('spPrevious').addEventListener('click',()=>{state.page--;renderProducts();});$('spNext').addEventListener('click',()=>{state.page++;renderProducts();});$('spExport').addEventListener('click',exportCSV);
    $('stockWorkspaceContainer').addEventListener('click',event=>{const sort=event.target.closest('[data-sp-sort]'),tab=event.target.closest('[data-sp-tab]');if(sort){state.desc=state.sort===sort.dataset.spSort?!state.desc:true;state.sort=sort.dataset.spSort;renderProducts();}if(tab)setTab(tab.dataset.spTab);});
    document.querySelector('.sp-tabs').addEventListener('keydown',event=>{const tabs=[...document.querySelectorAll('[data-sp-tab]')],index=tabs.indexOf(document.activeElement);if(index<0||!['ArrowLeft','ArrowRight','Home','End'].includes(event.key))return;event.preventDefault();const next=event.key==='Home'?0:event.key==='End'?tabs.length-1:(index+(event.key==='ArrowRight'?1:-1)+tabs.length)%tabs.length;setTab(tabs[next].dataset.spTab);tabs[next].focus();});
    const end=new Date(Date.now()-86400000),start=new Date(end.getTime()-29*86400000),iso=date=>date.toISOString().slice(0,10),max=new Date().toISOString().slice(0,10);$('spStart').value=iso(start);$('spEnd').value=iso(end);$('spStart').max=max;$('spEnd').max=max;connect();
  }
  window.StockPriceAnalysis={init,refresh:load,reconnect:connect};
  document.addEventListener('DOMContentLoaded',()=>{if(document.body.classList.contains('stock-price-active'))init();});
})();
