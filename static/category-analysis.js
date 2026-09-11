(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const state = {initialized: false, demo: false, view: 'categories', tab: 'performance', category: '', properties: [], categories: [], report: null, sort: 'itemRevenue', desc: true, page: 0, request: 0, controller: null};
  const pageSize = 15;
  const metrics = {
    itemsViewed: ['Items viewed', 'number'], itemsAddedToCart: ['Added to cart', 'number'], itemsPurchased: ['Items purchased', 'number'], itemRevenue: ['Item revenue', 'money'], cartToViewRate: ['Cart-to-view', 'percent'], purchaseToViewRate: ['Purchase-to-view', 'percent'],
    sessions: ['Sessions', 'number'], activeUsers: ['Active users', 'number'], engagedSessions: ['Engaged sessions', 'number'], bounceRate: ['Bounce rate', 'percent'], engagementRate: ['Engagement rate', 'percent'], averageSessionDuration: ['Avg. session duration', 'duration']
  };
  const performanceKeys = ['itemsViewed', 'itemsAddedToCart', 'itemsPurchased', 'itemRevenue', 'cartToViewRate', 'purchaseToViewRate'];
  const qualityKeys = ['sessions', 'activeUsers', 'engagedSessions', 'bounceRate', 'engagementRate', 'averageSessionDuration'];
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[c]));
  const numeric = value => typeof value === 'number' && Number.isFinite(value);
  function format(value, type) {
    if (!numeric(value)) return '—';
    if (type === 'percent') return (value * 100).toFixed(1) + '%';
    if (type === 'duration') { const seconds = Math.round(value); return `${Math.floor(seconds / 60)}m ${String(seconds % 60).padStart(2, '0')}s`; }
    if (type === 'money') { const currency = state.report?.property.currency; if (!currency) return Math.round(value).toLocaleString('en'); try { return new Intl.NumberFormat('en', {style:'currency', currency, maximumFractionDigits:0}).format(value); } catch { return `${currency} ${Math.round(value).toLocaleString('en')}`; } }
    return Math.round(value).toLocaleString('en');
  }
  function change(current, previous, key) {
    if (!numeric(current)) return '<span class="ca-delta">Unavailable</span>';
    if (!numeric(previous)) return '<span class="ca-delta">No prior data</span>';
    const rate = metrics[key][1] === 'percent';
    if (!rate && previous === 0) return `<span class="ca-delta">${current ? 'No prior baseline' : 'No change'}</span>`;
    const value = rate ? (current - previous) * 100 : (current / previous - 1) * 100;
    const favorable = key === 'bounceRate' ? value < 0 : value > 0;
    const color = Math.abs(value) < .05 || key === 'averageSessionDuration' ? '' : favorable ? 'up' : 'down';
    return `<span class="ca-delta ${color}">${value > 0 ? '+' : ''}${value.toFixed(1)}${rate ? ' pp' : '%'} vs prior</span>`;
  }
  function status(text, style = '') { $('caStatus').textContent = text; $('caStatus').className = 'ca-status ' + style; }
  async function api(url, options = {}) {
    const response = await fetch(url, {...options, credentials:'same-origin', headers:{'Content-Type':'application/json', ...options.headers}});
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      const error = new Error(data.detail?.message || (typeof data.detail === 'string' ? data.detail : '') || 'The report could not be loaded. Please try again.');
      error.code = data.detail?.code; error.status = response.status; throw error;
    }
    return data;
  }
  function dateRange() {
    if ($('caPeriod').value === 'custom') {
      const start = $('caStart').value, end = $('caEnd').value;
      if (!start || !end || start > end) throw new Error('Choose a valid start and end date.');
      if ((new Date(end) - new Date(start)) / 86400000 > 365) throw new Error('Choose a date range of 366 days or fewer.');
      return {start_date:start, end_date:end};
    }
    return {days:$('caPeriod').value};
  }
  function clearReport() { state.report = null; $('caReport').hidden = true; $('caExport').disabled = true; }
  function clearCategory() { state.category = ''; $('caCategoryFilter').value = ''; $('caClearCategory').hidden = true; }
  async function connect() {
    const request = ++state.request;
    state.controller?.abort(); clearReport(); $('categoryWorkspaceContainer').setAttribute('aria-busy', 'false');
    $('caWelcome').hidden = false; $('caRefresh').disabled = true; $('caProperty').disabled = true;
    $('caProperty').innerHTML = '<option value="">Loading your properties…</option>';
    status('Checking your Google Analytics connection…');
    try {
      const data = await api('/api/ga4/properties');
      if (request !== state.request) return;
      state.properties = data.properties || [];
      state.demo = false; clearCategory(); state.categories = [];
      $('caDemo').textContent = 'Explore sample data';
      $('caSource').textContent = data.connected ? 'CONNECTED' : 'NOT CONNECTED'; $('caSource').className = 'ca-badge ' + (data.connected ? 'live' : '');
      $('caConnect').textContent = data.connected ? 'Change Google account ↗' : 'Connect Google Analytics ↗';
      $('caConnectionText').textContent = data.connected ? data.email : 'Connect your store to load its categories and products.';
      $('caProperty').innerHTML = '<option value="">' + (data.connected ? 'Select your GA4 property' : 'Connect an account first') + '</option>' + state.properties.map(p => `<option value="${esc(p.id)}">${esc(p.name)} · ${esc(p.id)}</option>`).join('');
      $('caProperty').disabled = !state.properties.length;
      if (!data.connected) { status('Sign in with Google to load your own analytics, or explore the clearly labelled sample.'); return; }
      if (!state.properties.length) { status('No GA4 properties are available to this Google account. Ask your Analytics administrator for Viewer access, then reconnect.', 'error'); return; }
      const selected = state.properties.some(p => p.id === data.selected_property) ? data.selected_property : state.properties.length === 1 ? state.properties[0].id : '';
      $('caProperty').value = selected;
      if (selected) await load(); else status('Choose a GA4 property above. Its first report will load automatically.');
    } catch (error) {
      if (request !== state.request) return;
      $('caProperty').innerHTML = '<option value="">Connection unavailable</option>';
      status(error.message, 'error'); $('caRefresh').disabled = false;
    }
  }
  async function load() {
    state.controller?.abort();
    const request = ++state.request;
    clearReport(); $('categoryWorkspaceContainer').setAttribute('aria-busy', 'false'); $('caRefresh').disabled = false;
    const property = $('caProperty').value;
    if (!state.demo && !property) { clearReport(); status('Choose a GA4 property to load its report.'); return; }
    let range;
    try { range = dateRange(); } catch(error) { clearReport(); status(error.message, 'error'); return; }
    const controller = new AbortController(); state.controller = controller;
    clearReport(); $('caWelcome').hidden = true; $('caRefresh').disabled = true;
    $('categoryWorkspaceContainer').setAttribute('aria-busy', 'true');
    $('caSource').textContent = state.demo ? 'SAMPLE DATA' : 'CONNECTED';
    status('Loading performance and engagement from Google Analytics…');
    const timer = setTimeout(() => controller.abort(), 90000);
    try {
      const report = state.demo ? demoReport(range) : await api('/api/ga4/commerce-report?' + new URLSearchParams({property_id:property, view:state.view, category:state.category, ...range}), {signal:controller.signal});
      if (request !== state.request) return;
      state.report = report; state.page = 0;
      if (state.view === 'categories') {
        state.categories = report.rows.map(r => r.name);
        $('caCategoryFilter').innerHTML = '<option value="">All categories</option>' + state.categories.map(c => `<option>${esc(c)}</option>`).join('');
      }
      $('caCategoryFilter').value = state.category;
      $('caSource').textContent = state.demo ? 'SAMPLE DATA' : 'GA4 DATA'; $('caSource').className = 'ca-badge ' + (state.demo ? 'sample' : 'live');
      $('caReport').hidden = false; $('caExport').disabled = !report.rows.length;
      status(state.demo ? 'Sample data only — these illustrative figures do not belong to your store. Connect Google Analytics to see your own results.' : report.rows.length ? 'Report updated. Changes compare the immediately preceding period.' : 'No ecommerce items were returned for this period. Check Measurement for event coverage and tracking guidance.', state.demo ? 'sample' : '');
      render();
    } catch (error) {
      if (request !== state.request) return;
      $('caSource').textContent = 'REPORT UNAVAILABLE'; $('caSource').className = 'ca-badge';
      status(error.name === 'AbortError' ? 'Google Analytics took too long to respond. Please refresh the report.' : error.message, 'error');
      $('caWelcome').hidden = false;
    } finally { clearTimeout(timer); if (request === state.request) { $('categoryWorkspaceContainer').setAttribute('aria-busy', 'false'); $('caRefresh').disabled = false; } }
  }
  function filteredRows() {
    const q = $('caSearch').value.trim().toLocaleLowerCase();
    return (state.report?.rows || []).filter(r => `${r.name} ${r.item_id || ''} ${r.category || ''}`.toLocaleLowerCase().includes(q)).sort((a,b) => {
      if (state.sort === 'name') return a.name.localeCompare(b.name) * (state.desc ? -1 : 1);
      const x = a.current[state.sort], y = b.current[state.sort];
      if (!numeric(x)) return numeric(y) ? 1 : 0;
      if (!numeric(y)) return -1;
      return (x - y) * (state.desc ? -1 : 1);
    });
  }
  function heading(keys) { return '<tr>' + ['name', ...keys].map(k => `<th scope="col" aria-sort="${state.sort === k ? state.desc ? 'descending' : 'ascending' : 'none'}"><button type="button" data-ca-sort="${k}">${k === 'name' ? state.view === 'categories' ? 'Category' : 'Product' : metrics[k][0]} <span aria-hidden="true">${state.sort === k ? state.desc ? '↓' : '↑' : '↕'}</span></button></th>`).join('') + '</tr>'; }
  function rowHTML(row, keys) {
    const title = state.view === 'categories' ? `<button type="button" data-ca-category="${esc(row.name)}" title="View products in ${esc(row.name)}">${esc(row.name)} <span aria-hidden="true">↗</span></button>` : esc(row.name);
    return `<tr><td class="ca-entity">${title}${state.view === 'products' ? `<small>${esc(row.item_id || 'No item ID')} · ${esc(row.category)}</small>` : ''}</td>` + keys.map(k => `<td class="${k === 'itemRevenue' ? 'ca-money' : ''}">${format(row.current[k], metrics[k][1])}${change(row.current[k], row.previous?.[k], k)}</td>`).join('') + '</tr>';
  }
  function renderTable() {
    const rows = filteredRows(); const pages = Math.max(1, Math.ceil(rows.length / pageSize)); state.page = Math.min(state.page, pages - 1);
    const subset = rows.slice(state.page * pageSize, (state.page + 1) * pageSize);
    $('caTableHead').innerHTML = heading(performanceKeys); $('caQualityHead').innerHTML = heading(qualityKeys);
    const empty = `<tr><td class="ca-empty" colspan="7">${state.report.rows.length ? 'No matches. Try another search.' : 'No items were recorded in this period. Open Measurement to check your ecommerce events.'}</td></tr>`;
    $('caTableBody').innerHTML = subset.length ? subset.map(r => rowHTML(r, performanceKeys)).join('') : empty;
    $('caQualityBody').innerHTML = subset.length ? subset.map(r => rowHTML(r, qualityKeys)).join('') : empty;
    $('caRowCount').textContent = `${rows.length.toLocaleString('en')} ${rows.length === 1 ? state.view === 'categories' ? 'category' : 'product' : state.view}`;
    $('caPage').textContent = `${state.page + 1} / ${pages}`;
    $('caPrevious').disabled = state.page === 0; $('caNext').disabled = state.page + 1 >= pages;
    $('caTableFootnote').textContent = `${state.report.list_complete ? 'All returned rows' : 'Partial list — export contains the loaded rows only'} · Item quantities, not orders`;
  }
  function render() {
    const r = state.report;
    $('caReportScope').textContent = `${r.start_date} – ${r.end_date} · Prior: ${r.previous_start} – ${r.previous_end}`;
    $('caFreshness').textContent = `${r.property.time_zone} · ${r.property.currency || 'Currency unavailable'} · ${state.demo ? 'Illustrative sample' : 'Updated ' + new Date(r.fetched_at).toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'})}`;
    const kpis = [['itemRevenue',r.summary,'Item revenue',state.category || 'All items'], ['purchaseToViewRate',r.summary,'Purchase-to-view rate',state.category || 'All items'], ['bounceRate',r.property_quality,'Bounce rate','Whole property'], ['averageSessionDuration',r.property_quality,'Avg. session duration','Whole property']];
    $('caKpis').innerHTML = kpis.map(([key, source, label, scope]) => `<article class="ca-kpi"><span class="ca-kpi-label">${label}</span><strong class="ca-kpi-value">${format(source?.current[key],metrics[key][1])}</strong>${change(source?.current[key],source?.previous?.[key],key)}<small>${esc(scope)}</small></article>`).join('');
    $('caTableTitle').textContent = state.view === 'categories' ? 'Category performance' : state.category ? `${state.category} · Products` : 'Product performance';
    $('caCategoryFilterWrap').hidden = state.view !== 'products' || !state.categories.length;
    $('caClearCategory').hidden = !state.category;
    $('caClearCategory').textContent = state.category ? `${state.category} ×` : '';
    $('caCategories').setAttribute('aria-pressed', state.view === 'categories'); $('caProducts').setAttribute('aria-pressed',state.view === 'products');
    $('caSearch').placeholder = state.view === 'categories' ? 'Search categories…' : 'Search products or item IDs…';
    const missing = qualityKeys.filter(k => !r.quality_metrics.includes(k)).map(k => metrics[k][0]);
    $('caQualityNote').textContent = missing.length ? `Unavailable for this breakdown: ${missing.join(', ')}. Whole-property bounce rate and session duration remain visible in the cards above when available.` : 'Session and user metrics associated with each category or product, as reported by GA4. These rows overlap and are not additive.';
    renderTable(); renderCharts(); renderMeasurement(); setTab(state.tab);
  }
  function renderCharts() {
    const r = state.report, rows = [...r.rows].sort((a,b) => (b.current.itemRevenue || 0) - (a.current.itemRevenue || 0));
    const revenue = r.summary.current.itemRevenue;
    $('caRevenueBars').innerHTML = rows.slice(0,6).map(row => {
      const share = numeric(revenue) && revenue > 0 && numeric(row.current.itemRevenue) ? row.current.itemRevenue / revenue * 100 : null;
      return `<div class="ca-bar"><div class="ca-bar-label"><span>${esc(row.name)}</span><strong>${format(row.current.itemRevenue,'money')} · ${share === null ? '—' : share.toFixed(1) + '%'}</strong></div><div class="ca-bar-track"><div class="ca-bar-fill" style="width:${Math.max(0, Math.min(100,share || 0))}%"></div></div></div>`;
    }).join('') || '<p class="ca-note">Revenue will appear once GA4 returns ecommerce items.</p>';
    const insights = [];
    const rate = r.summary.current.purchaseToViewRate;
    const opportunities = r.rows.filter(row => row.current.itemsViewed >= 100 && numeric(row.current.purchaseToViewRate) && numeric(rate) && row.current.purchaseToViewRate < rate).sort((a,b) => b.current.itemsViewed - a.current.itemsViewed);
    if (opportunities.length) {
      const row = opportunities[0];
      insights.push(`<div class="ca-insight"><span class="ca-badge">DEMAND WITHOUT CONVERSION</span><h4>${esc(row.name)}</h4><p>${format(row.current.itemsViewed,'number')} items viewed, with a ${format(row.current.purchaseToViewRate,'percent')} purchase-to-view rate versus ${format(rate,'percent')} in this report’s scope. Review product availability, pricing and the detail-page experience.</p></div>`);
    }
    if (rows.length && numeric(rows[0].current.itemRevenue) && numeric(revenue) && revenue > 0) insights.push(`<div class="ca-insight"><span class="ca-badge">${r.list_complete ? 'REVENUE LEADER' : 'TOP LOADED ROW'}</span><h4>${esc(rows[0].name)}</h4><p>Contributes ${((rows[0].current.itemRevenue / revenue) * 100).toFixed(1)}% of item revenue in this report. Check its product mix before drawing conclusions from category averages.</p></div>`);
    const unmapped = r.rows.find(row => ['(not set)','(unset)',''].includes(row.category || row.name));
    if (unmapped) insights.push('<div class="ca-insight"><h4>Some items have no category</h4><p>Review item_category in your ecommerce tracking so those items can be analysed in the correct group.</p></div>');
    $('caInsights').innerHTML = insights.join('') || '<p class="ca-note">There is not enough recorded activity for a useful comparison yet. Try a longer date range.</p>';
  }
  function renderMeasurement() {
    const r = state.report;
    $('caEventCoverage').innerHTML = [['view_item','Product detail'],['add_to_cart','Add to cart'],['begin_checkout','Begin checkout'],['purchase','Purchase']].map(([key,label]) => `<article class="ca-event"><span class="ca-eyebrow">${label}</span><strong>${format(r.events?.current?.[key],'number')}</strong><p>Recorded events</p><code>${key}</code></article>`).join('');
    const warnings = [...r.warnings];
    if (state.demo) warnings.unshift('This is sample data for exploring the interface. No customer account or Google Analytics report was used.');
    warnings.push('GA4 data may be delayed or revised during processing. These reports are not real-time.', 'Event counts can repeat for one user and do not prove a completed, ordered user journey.');
    $('caHealthCount').textContent = r.warnings.length || '';
    $('caWarnings').innerHTML = warnings.map(w => `<div class="ca-warning">${esc(w)}</div>`).join('');
  }
  function setTab(tab) {
    state.tab = tab;
    for (const key of ['performance','engagement','measurement']) {
      const title = key[0].toUpperCase() + key.slice(1);
      const button = $('ca' + title + 'Tab');
      button.classList.toggle('active',key === tab); button.setAttribute('aria-selected',key === tab); button.tabIndex = key === tab ? 0 : -1;
      $('ca' + title + 'Panel').hidden = key !== tab;
    }
    $('caExplorerControls').hidden = tab === 'measurement';
    // Keep the same pagination controls accessible under either data table.
    const footer = document.querySelector('.ca-table-footer');
    $(tab === 'engagement' ? 'caEngagementPanel' : 'caPerformancePanel').querySelector('.ca-table-card').appendChild(footer);
    if (state.report) $('caTableFootnote').textContent = `${state.report.list_complete ? 'All returned rows' : 'Partial list — loaded rows only'} · ${tab === 'engagement' ? 'Session and user rows overlap' : 'Item quantities, not orders'}`;
  }
  async function setView(view, category = '') {
    state.view = view; state.category = category; state.page = 0; $('caSearch').value = '';
    if (state.report || state.demo || $('caProperty').value) await load();
  }
  function exportCSV() {
    if (!state.report) return;
    const r = state.report, keys = state.tab === 'engagement' ? qualityKeys : performanceKeys;
    const cell = value => { let s = String(value ?? ''); if (/^[=+@\-\t\r]/.test(s)) s = "'" + s; return '"' + s.replace(/"/g,'""') + '"'; };
    const header = ['source','property_id','currency','time_zone','start_date','end_date','previous_start','previous_end','complete_list','name','item_id','category', ...keys.flatMap(k => [k,k + '_previous'])];
    const rows = filteredRows().map(row => [r.source,r.property.id,r.property.currency,r.property.time_zone,r.start_date,r.end_date,r.previous_start,r.previous_end,r.list_complete,row.name,row.item_id,row.category,...keys.flatMap(k => [row.current[k],row.previous?.[k]])]);
    const blob = new Blob(['\ufeff' + [header,...rows].map(row => row.map(cell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob); const link = document.createElement('a'); link.href = url; link.download = `${state.demo ? 'SAMPLE-' : ''}${state.view}-${state.tab}-${r.start_date}.csv`; link.click(); setTimeout(() => URL.revokeObjectURL(url),1000);
  }
  function demoReport(range) {
    const end = range.end_date ? new Date(range.end_date + 'T12:00:00Z') : new Date(Date.now() - 86400000);
    const start = range.start_date ? new Date(range.start_date + 'T12:00:00Z') : new Date(end.getTime() - (Number(range.days) - 1) * 86400000);
    const length = Math.round((end-start)/86400000) + 1, iso = d => d.toISOString().slice(0,10);
    const categories = [['Electronics',24830,6840,1975,286420,.239,.064, .39,176],['Home & Living',18640,4820,1420,148620,.217,.059,.34,203],['Fashion',32480,5120,1150,94250,.143,.028,.52,118],['Beauty & Care',14760,4680,1630,82740,.287,.092,.28,224],['Sports & Outdoors',9840,1820,420,63150,.174,.036,.45,146],['Accessories',12240,2940,910,32680,.211,.061,.41,158]];
    const items = { 'Electronics':['Wireless Headphones','Portable Speaker','Smart Watch','USB-C Hub'], 'Home & Living':['Cotton Bed Set','Desk Lamp','Ceramic Dinner Set','Storage Basket'], 'Fashion':['Everyday Sneakers','Cotton Shirt','Denim Jacket','Linen Trousers'], 'Beauty & Care':['Daily Moisturiser','Facial Cleanser','Sun Cream SPF 50','Night Serum'], 'Sports & Outdoors':['Running Shoes','Training Mat','Stainless Bottle','Daypack'], 'Accessories':['Leather Wallet','Canvas Tote','Phone Case','Sunglasses'] };
    const make = (c,i,factor=1,name=c[0],id='') => {
      const current = {itemsViewed:Math.round(c[1]*factor),itemsAddedToCart:Math.round(c[2]*factor),itemsPurchased:Math.round(c[3]*factor),itemRevenue:Math.round(c[4]*factor),cartToViewRate:c[5],purchaseToViewRate:c[6],sessions:Math.round(c[1]*.72*factor),activeUsers:Math.round(c[1]*.61*factor),bounceRate:c[7],engagementRate:1-c[7],averageSessionDuration:c[8]};
      current.engagedSessions = Math.round(current.sessions * current.engagementRate);
      const previous = Object.fromEntries(Object.entries(current).map(([key,value]) => [key,metrics[key][1] === 'percent' ? Math.max(0,value - (i%2 ? -.013 : .012)) : value / (i%2 ? .94 : 1.128)]));
      previous.bounceRate = 1 - previous.engagementRate;
      previous.engagedSessions = Math.round(previous.sessions * previous.engagementRate);
      return {name,key:id || c[0],category:c[0],item_id:id,current,previous};
    };
    let rows = state.view === 'categories' ? categories.map((c,i) => make(c,i)) : categories.filter(c => !state.category || c[0] === state.category).flatMap((c,i) => items[c[0]].map((name,j) => make(c,i,[.4,.3,.2,.1][j],name,`DP-${categories.indexOf(c)+1}0${j+1}`)));
    const summary = {current:{},previous:{}};
    for (const period of ['current','previous']) {
      for (const key of performanceKeys.slice(0,4)) summary[period][key] = rows.reduce((sum,r) => sum+r[period][key],0);
      summary[period].cartToViewRate = state.category && rows.length ? rows[0][period].cartToViewRate : period === 'current' ? .212 : .201;
      summary[period].purchaseToViewRate = state.category && rows.length ? rows[0][period].purchaseToViewRate : period === 'current' ? .054 : .049;
    }
    return {source:'sample',property:{id:'sample',name:'Sample retail store',currency:'USD',time_zone:'UTC'},start_date:iso(start),end_date:iso(end),previous_start:iso(new Date(start.getTime()-length*86400000)),previous_end:iso(new Date(start.getTime()-86400000)),fetched_at:new Date().toISOString(),rows,summary,property_quality:{current:{bounceRate:.418,averageSessionDuration:167},previous:{bounceRate:.449,averageSessionDuration:151}},quality_metrics:qualityKeys,events:{current:{view_item:112790,add_to_cart:26220,begin_checkout:12280,purchase:6450}},warnings:[],list_complete:true};
  }
  function init() {
    if (state.initialized) return;
    state.initialized = true;
    $('caPeriod').addEventListener('change',() => { $('caCustomDates').hidden = $('caPeriod').value !== 'custom'; if ($('caPeriod').value !== 'custom' && (state.report || $('caProperty').value)) load(); });
    $('caRefresh').addEventListener('click',() => !state.demo && !$('caProperty').value ? connect() : load());
    $('caProperty').addEventListener('change',async () => {
      const selectedProperty = $('caProperty').value;
      clearCategory(); state.categories = []; $('caCategoryFilter').innerHTML = '<option value="">All categories</option>'; $('caSearch').value = '';
      await load();
      if (state.report?.property.id === selectedProperty && !state.demo) { try { await api('/api/ga4/selection',{method:'POST',body:JSON.stringify({property_id:selectedProperty})}); } catch { if ($('caProperty').value === selectedProperty) status('Report loaded, but your property preference could not be saved. Select it again next time.', 'error'); } }
    });
    $('caDemo').addEventListener('click',() => {
      if (state.demo) { connect(); return; }
      state.controller?.abort(); state.demo = true; clearCategory(); state.view = 'categories'; state.categories = []; $('caSearch').value = '';
      $('caProperty').innerHTML = '<option value="sample">Sample retail store · Illustrative data</option>'; $('caProperty').disabled = true;
      $('caConnectionText').textContent = 'Exploring an illustrative retail store'; $('caDemo').textContent = 'Exit sample'; load();
    });
    $('caCategories').addEventListener('click',() => setView('categories'));
    $('caProducts').addEventListener('click',() => setView('products'));
    $('caCategoryFilter').addEventListener('change',() => setView('products',$('caCategoryFilter').value));
    $('caClearCategory').addEventListener('click',() => setView('products'));
    $('caSearch').addEventListener('input',() => { state.page=0; if (state.report) renderTable(); });
    $('caPrevious').addEventListener('click',() => { state.page--; renderTable(); }); $('caNext').addEventListener('click',() => { state.page++; renderTable(); });
    $('caExport').addEventListener('click',exportCSV);
    $('categoryWorkspaceContainer').addEventListener('click',event => {
      const sort = event.target.closest('[data-ca-sort]'), category = event.target.closest('[data-ca-category]'), tab = event.target.closest('[data-ca-tab]');
      if (sort) { state.desc = state.sort === sort.dataset.caSort ? !state.desc : true; state.sort = sort.dataset.caSort; renderTable(); }
      if (category) setView('products',category.dataset.caCategory);
      if (tab) setTab(tab.dataset.caTab);
    });
    document.querySelector('.ca-tabs').addEventListener('keydown',event => {
      const tabs = [...document.querySelectorAll('[data-ca-tab]')], i = tabs.indexOf(document.activeElement);
      if (i < 0 || !['ArrowLeft','ArrowRight','Home','End'].includes(event.key)) return;
      event.preventDefault(); const next = event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length-1 : (i+(event.key==='ArrowRight'?1:-1)+tabs.length)%tabs.length;
      setTab(tabs[next].dataset.caTab); tabs[next].focus();
    });
    const end = new Date(Date.now()-86400000), start = new Date(end.getTime()-29*86400000);
    $('caStart').value = start.toISOString().slice(0,10); $('caEnd').value = end.toISOString().slice(0,10);
    const max = new Date().toISOString().slice(0,10); $('caStart').max=max; $('caEnd').max=max;
    connect();
  }
  window.CategoryAnalysis = {init,refresh:load,reconnect:connect};
  document.addEventListener('DOMContentLoaded',() => { if (document.body.classList.contains('category-analysis-active')) init(); });
})();
