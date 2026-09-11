(function () {
  'use strict';
  let initialized = false;
  const NA = 'Not available';
  const state = { baseline: null, rivals: [], country: 'tr', snapshot: null, reviews: null, busy: false, reviewsBusy: false, compareController: null, reviewsController: null, compareSequence: 0, reviewsSequence: 0 };
  const searches = { own: { sequence: 0, timer: null, controller: null }, rival: { sequence: 0, timer: null, controller: null } };
  const ui = {};

  function el(tag, className, text) {
    const item = document.createElement(tag);
    if (className) item.className = className;
    if (text !== undefined && text !== null) item.textContent = String(text);
    return item;
  }
  function numeric(value) { return value !== null && value !== undefined && value !== '' && Number.isFinite(Number(value)) ? Number(value) : null; }
  function number(value) { const n = numeric(value); return n === null ? NA : new Intl.NumberFormat('en-US').format(n); }
  function textValue(value) { return value === null || value === undefined || value === '' ? NA : String(value); }
  function dateValue(value) { const parsed = value ? new Date(value) : null; return parsed && Number.isFinite(parsed.getTime()) ? parsed : null; }
  function dateText(value) { const d = dateValue(value); return d ? new Intl.DateTimeFormat('en-GB', { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' }).format(d) : NA; }
  function rating(value, count) { const n = numeric(value); return n === null ? NA : numeric(count) === 0 ? 'No ratings yet' : n.toFixed(2) + ' / 5'; }
  function price(app) {
    const n = numeric(app.price);
    if (n === null) return NA;
    if (n === 0) return 'Free';
    try { return new Intl.NumberFormat('en-US', { style: 'currency', currency: app.currency || 'XXX' }).format(n); }
    catch (_) { return n.toFixed(2) + (app.currency ? ' ' + app.currency : ''); }
  }
  function updateAge(value) {
    const d = dateValue(value), now = dateValue(state.snapshot && state.snapshot.fetched_at) || new Date();
    if (!d) return null;
    const days = Math.floor((now.getTime() - d.getTime()) / 86400000);
    return days < 0 ? null : days === 0 ? 'Updated today' : days + (days === 1 ? ' day ago' : ' days ago');
  }
  function safeURL(value, kind) {
    try {
      const url = new URL(value);
      if (url.protocol !== 'https:' || url.username || url.password) return null;
      const host = url.hostname.toLowerCase();
      if (kind === 'icon' ? (host === 'mzstatic.com' || host.endsWith('.mzstatic.com')) : (host === 'apps.apple.com' || host === 'itunes.apple.com')) return url.href;
    } catch (_) { /* External values are never used as HTML or unchecked URLs. */ }
    return null;
  }
  function icon(app) {
    const frame = el('span', 'ab-icon-frame', String(app.name || '?').slice(0, 1).toUpperCase());
    frame.setAttribute('aria-hidden', 'true');
    const url = safeURL(app.icon_url, 'icon');
    if (url) { const img = el('img'); img.src = url; img.alt = ''; img.loading = 'lazy'; img.referrerPolicy = 'no-referrer'; img.addEventListener('error', () => img.remove(), { once: true }); frame.append(img); }
    return frame;
  }
  function chosen() { return (state.baseline ? [state.baseline] : []).concat(state.rivals); }
  function isSelected(app) { return chosen().some(item => String(item.id) === String(app.id)); }
  function marketName() { return ui.country.options[ui.country.selectedIndex].textContent; }
  function mainStatus(message, error) {
    ui.mainStatus.textContent = message || '';
    ui.mainStatus.hidden = !message;
    ui.mainStatus.classList.toggle('ab-error', !!error);
  }
  function updateButton() {
    ui.compare.disabled = !state.baseline || state.rivals.length < 1 || state.busy;
    ui.compare.textContent = state.busy ? 'Fetching store data…' : 'Compare apps  →';
    ui.compare.classList.toggle('ab-loading-button', state.busy);
    ui.compare.setAttribute('aria-busy', String(state.busy));
  }
  function clearSearch(role, clearInput) {
    const search = searches[role], view = ui[role];
    search.sequence += 1;
    clearTimeout(search.timer);
    if (search.controller) search.controller.abort();
    search.controller = null;
    view.results.replaceChildren();
    view.results.hidden = true;
    view.status.textContent = '';
    if (clearInput) view.input.value = '';
  }
  function invalidateSnapshot() {
    state.compareSequence += 1;
    state.reviewsSequence += 1;
    if (state.compareController) state.compareController.abort();
    if (state.reviewsController) state.reviewsController.abort();
    state.compareController = null;
    state.reviewsController = null;
    state.busy = false;
    state.reviewsBusy = false;
    state.snapshot = null;
    state.reviews = null;
    ui.results.hidden = true;
    ui.empty.hidden = false;
    ui.export.hidden = true;
    if (ui.signalsResults) renderSignals([]);
    if (ui.reviewsResults) ui.reviewsResults.replaceChildren();
    if (ui.reviewsStatus) ui.reviewsStatus.textContent = 'Compare at least two apps, then open this tab to load public reviews.';
    mainStatus('');
    updateButton();
  }
  function removeSelection(id, isOwn) {
    if (isOwn) state.baseline = null;
    else state.rivals = state.rivals.filter(app => String(app.id) !== String(id));
    clearSearch('own', false);
    clearSearch('rival', false);
    invalidateSnapshot();
    renderSelections();
    (isOwn ? ui.own.input : ui.rival.input).focus();
  }
  function removeButton(app, isOwn) {
    const button = el('button', 'ab-remove', '×');
    button.type = 'button';
    button.setAttribute('aria-label', 'Remove ' + textValue(app.name) + (isOwn ? ' as your app' : ' from competitors'));
    button.addEventListener('click', () => removeSelection(app.id, isOwn));
    return button;
  }
  function renderSelections() {
    ui.own.selection.replaceChildren();
    if (state.baseline) {
      const selected = el('div', 'ab-selection');
      const info = el('div', 'ab-selection-info');
      info.append(el('div', 'ab-selection-name', state.baseline.name), el('div', 'ab-selection-subtitle', 'Your app · ' + textValue(state.baseline.developer)));
      selected.append(icon(state.baseline), info, removeButton(state.baseline, true));
      ui.own.selection.append(selected);
    }
    ui.rival.selection.replaceChildren();
    state.rivals.forEach(app => { const chip = el('div', 'ab-rival-chip'); chip.append(el('span', '', app.name), removeButton(app, false)); ui.rival.selection.append(chip); });
    ui.own.input.placeholder = state.baseline ? 'Search to replace your app' : 'Search an app or publisher';
    ui.rival.input.disabled = !state.baseline || state.rivals.length >= 4;
    ui.rival.hint.textContent = !state.baseline ? 'Choose your app first, then add up to 4 competitors.' : state.rivals.length >= 4 ? 'Your set is full. Remove an app to add a different competitor.' : 'Add ' + (4 - state.rivals.length) + ' more competitor' + (4 - state.rivals.length === 1 ? '' : 's') + ' from the same market.';
    ui.rival.count.textContent = state.rivals.length + ' / 4';
    ui.summary.textContent = !state.baseline ? 'Choose your app and at least one competitor to get started.' : !state.rivals.length ? 'Your app is ready. Add a competitor to compare.' : chosen().length + ' apps selected · ' + marketName() + ' App Store';
    updateButton();
  }
  function selectApp(role, app) {
    if (!app || !app.id || isSelected(app)) return;
    if (role === 'own') state.baseline = app;
    else { if (!state.baseline || state.rivals.length >= 4) return; state.rivals.push(app); }
    clearSearch('own', role === 'own');
    clearSearch('rival', true);
    invalidateSnapshot();
    renderSelections();
    if (!ui.rival.input.disabled) ui.rival.input.focus();
    else ui.compare.focus();
  }
  function errorMessage(error) { return error && error.message ? error.message : 'The store could not be reached. Please try again.'; }
  async function request(url, options, controller, collection) {
    let timedOut = false;
    const timer = window.setTimeout(() => { timedOut = true; controller.abort(); }, 25000);
    try {
      const response = await fetch(url, Object.assign({ credentials: 'same-origin', signal: controller.signal, headers: { 'Accept': 'application/json' } }, options || {}));
      let data;
      try { data = await response.json(); } catch (_) { throw new Error('The server returned an unexpected response. Please try again.'); }
      if (!response.ok) {
        if (response.status === 401 || response.status === 403) throw new Error('Please sign in again to use app comparison.');
        if (response.status === 429) throw new Error('Too many store requests. Please wait a moment and try again.');
        const detail = typeof data.detail === 'string' ? data.detail : (data.detail && typeof data.detail.message === 'string' ? data.detail.message : null);
        throw new Error(detail ? detail.slice(0, 300) : 'App Store data is temporarily unavailable. Please try again.');
      }
      const key = collection || 'apps';
      if (!data || !Array.isArray(data[key])) throw new Error('The server returned an unexpected response. Please try again.');
      return data;
    } catch (error) {
      if (timedOut) throw new Error('The store took too long to respond. Please try again.');
      if (error.name === 'AbortError') throw error;
      if (error instanceof TypeError) throw new Error('Cannot reach the server. Check your connection and try again.');
      throw error;
    } finally { window.clearTimeout(timer); }
  }
  function renderSearch(role, apps) {
    const view = ui[role];
    view.results.replaceChildren();
    const validApps = apps.filter(app => app && app.id && app.name);
    if (!validApps.length) { view.results.hidden = true; view.status.textContent = 'No apps found in ' + marketName() + '. Try another name or publisher.'; return; }
    validApps.forEach(app => {
      const button = el('button', 'ab-search-result'); button.type = 'button';
      const info = el('span', 'ab-result-info');
      info.append(el('span', 'ab-result-name', app.name), el('span', 'ab-result-meta', [app.developer, app.category].filter(Boolean).join(' · ')));
      button.append(icon(app), info, el('span', 'ab-result-add', isSelected(app) ? '✓' : '+'));
      button.disabled = isSelected(app);
      button.setAttribute('aria-label', (isSelected(app) ? 'Already selected: ' : role === 'own' ? 'Select as your app: ' : 'Add competitor: ') + app.name + (app.developer ? ', ' + app.developer : ''));
      button.addEventListener('click', () => selectApp(role, app));
      view.results.append(button);
    });
    view.results.hidden = false;
    view.status.textContent = validApps.length + ' apps found. Select an app below.';
  }
  function queueSearch(role) {
    clearSearch(role, false);
    const view = ui[role], query = view.input.value.trim();
    if (query.length < 2) { if (query.length) view.status.textContent = 'Type at least 2 characters to search.'; return; }
    const sequence = searches[role].sequence, country = state.country;
    view.status.textContent = 'Searching the ' + marketName() + ' App Store…';
    searches[role].timer = window.setTimeout(async () => {
      const controller = new AbortController(); searches[role].controller = controller;
      try {
        const data = await request('/api/app-benchmark/search?' + new URLSearchParams({ q: query, country: country }), null, controller);
        if (sequence !== searches[role].sequence || country !== state.country) return;
        renderSearch(role, data.apps);
      } catch (error) {
        if (sequence !== searches[role].sequence || error.name === 'AbortError') return;
        view.status.textContent = errorMessage(error) + ' Edit your search to retry.';
      } finally { if (sequence === searches[role].sequence) searches[role].controller = null; }
    }, 350);
  }
  const metrics = [
    { group: 'CUSTOMER SENTIMENT' },
    { label: 'Store rating', value: app => rating(app.rating, app.rating_count) },
    { label: 'Rating count', value: app => number(app.rating_count) },
    { label: 'Current version rating', value: app => rating(app.current_rating, app.current_rating_count) },
    { label: 'Current version rating count', value: app => number(app.current_rating_count) },
    { group: 'PRICING & RELEASES' },
    { label: 'Download price', value: price, note: () => 'Excludes in-app purchases' },
    { label: 'Latest update', value: app => dateText(app.updated_at), note: app => updateAge(app.updated_at) },
    { label: 'Current version', value: app => textValue(app.version) },
    { label: 'First release', value: app => dateText(app.released_at) },
    { group: 'APP DETAILS' },
    { label: 'App size', value: app => numeric(app.size_bytes) === null ? NA : (Number(app.size_bytes) / 1000000).toFixed(1) + ' MB' },
    { label: 'Minimum iOS', value: app => textValue(app.min_os) },
    { label: 'Languages', value: app => Array.isArray(app.languages) ? number(app.languages.length) : NA },
    { label: 'Age rating', value: app => textValue(app.content_rating) },
    { label: 'Category', value: app => textValue(app.category) }
  ];
  function appCards(apps) {
    ui.cards.replaceChildren();
    apps.forEach((app, index) => {
      const card = el('article', 'ab-app-card' + (index === 0 ? ' ab-own-card' : ''));
      const top = el('div', 'ab-app-card-top'); top.append(icon(app), el('span', 'ab-card-label', index === 0 ? 'YOUR APP' : 'COMPETITOR'));
      card.append(top, el('h4', '', app.name), el('p', 'ab-card-developer', textValue(app.developer)));
      const url = safeURL(app.store_url, 'store');
      if (url) { const link = el('a', 'ab-card-link', 'View in App Store ↗'); link.href = url; link.target = '_blank'; link.rel = 'noopener noreferrer'; link.setAttribute('aria-label', 'View ' + app.name + ' in App Store'); card.append(link); }
      ui.cards.append(card);
    });
  }
  function barRow(app, amount, label, index) {
    const row = el('div', 'ab-bar-row');
    const heading = el('div', 'ab-bar-label'); heading.append(el('span', '', app.name), el('strong', '', label));
    const track = el('div', 'ab-bar-track'); track.setAttribute('aria-hidden', 'true');
    const fill = el('div', 'ab-bar-fill' + (index === 0 ? ' ab-own-bar' : '')); fill.style.width = Math.max(0, Math.min(100, amount || 0)) + '%';
    track.append(fill); row.append(heading, track); return row;
  }
  function clamp(value, minimum, maximum) { return Math.max(minimum, Math.min(maximum, value)); }
  function daysSince(value) {
    const date = dateValue(value), now = dateValue(state.snapshot && state.snapshot.fetched_at) || new Date();
    if (!date) return null;
    return Math.max(0, (now.getTime() - date.getTime()) / 86400000);
  }
  const signalModels = [
    { key: 'rating', label: 'Customer rating', note: 'Current public storefront rating', direction: 'Higher is better', value: app => numeric(app.rating), display: value => value === null ? NA : value.toFixed(2) + ' / 5', bar: value => value === null ? 0 : value * 20 },
    { key: 'ratings', label: 'Rating footprint', note: 'Published rating count, not installs or users', direction: 'Higher is larger', value: app => numeric(app.rating_count), display: value => value === null ? NA : number(value) },
    { key: 'update', label: 'Update recency', note: 'Days since the latest public version', direction: 'Lower is fresher', inverse: true, value: app => daysSince(app.updated_at), display: value => value === null ? NA : Math.round(value) + (Math.round(value) === 1 ? ' day' : ' days') },
    { key: 'languages', label: 'Localization reach', note: 'Languages listed by the storefront', direction: 'Higher is broader', value: app => Array.isArray(app.languages) ? app.languages.length : null, display: value => value === null ? NA : number(value) + (value === 1 ? ' language' : ' languages') }
  ];
  function signalData(apps) {
    return signalModels.map(model => {
      const raw = apps.map(model.value);
      const available = raw.filter(value => value !== null);
      const maximum = Math.max.apply(null, available.concat([0]));
      const total = model.key === 'ratings' ? available.reduce((sum, value) => sum + value, 0) : null;
      return { model, values: raw.map((value, index) => ({ raw: value, app: apps[index], amount: model.bar ? model.bar(value) : value === null || maximum <= 0 ? 0 : value / maximum * 100, share: total && value !== null ? value / total * 100 : null })) };
    });
  }
  function renderSignals(apps) {
    if (!ui.signalsResults) return;
    ui.signalsResults.replaceChildren();
    if (!apps || apps.length < 2) { ui.signalsResults.append(el('div', 'ab-proxy-empty', 'Compare at least two apps to see public store signals.')); return; }
    signalData(apps).forEach(({ model, values }) => {
      const card = el('article', 'ab-proxy-card');
      const heading = el('div', 'ab-proxy-card-heading');
      const title = el('div'); title.append(el('h4', '', model.label), el('p', '', model.note));
      heading.append(title, el('span', model.inverse ? 'ab-direction ab-direction-risk' : 'ab-direction', model.direction));
      card.append(heading);
      const baselineRaw = values[0].raw;
      values.forEach((item, index) => {
        const row = el('div', 'ab-proxy-row');
        const label = el('div', 'ab-proxy-label');
        const ratio = index === 0 ? 'Your app · Baseline' : item.raw !== null && baselineRaw !== null && baselineRaw !== 0 ? (item.raw / baselineRaw).toFixed(2) + '× baseline' : 'Baseline ratio unavailable';
        const share = model.key === 'ratings' && item.share !== null ? ' · ' + item.share.toFixed(1) + '% of selected ratings' : '';
        label.append(el('span', '', item.app.name), el('small', '', ratio + share));
        const score = el('strong', item.raw === null ? 'ab-unavailable' : '', model.display(item.raw));
        const track = el('div', 'ab-proxy-track');
        const fill = el('div', 'ab-proxy-fill' + (index === 0 ? ' ab-proxy-fill-own' : '') + (model.inverse ? ' ab-proxy-fill-risk' : '')); fill.style.width = clamp(item.amount, 0, 100) + '%';
        track.append(fill); row.append(label, score, track); card.append(row);
      });
      ui.signalsResults.append(card);
    });
  }
  function sentimentLabel(value) { return value === 'positive' ? 'Positive' : value === 'neutral' ? 'Neutral' : 'Negative'; }
  function reviewItem(review) {
    const item = el('article', 'ab-review-item');
    const top = el('div', 'ab-review-item-top');
    const stars = el('span', 'ab-review-stars', '★'.repeat(Math.max(1, Math.min(5, Number(review.rating) || 0))));
    stars.setAttribute('aria-label', number(review.rating) + ' out of 5 stars');
    top.append(stars, el('span', 'ab-sentiment-pill ab-sentiment-' + review.sentiment, sentimentLabel(review.sentiment)));
    const title = el('h5', '', textValue(review.title));
    const body = el('p', '', String(review.body || '').slice(0, 420));
    const date = el('time', '', dateText(review.created_at));
    if (review.created_at) date.dateTime = review.created_at;
    item.append(top, title, body, date);
    return item;
  }
  function renderReviews(data) {
    ui.reviewsResults.replaceChildren();
    const byId = new Map(state.snapshot.apps.map(app => [String(app.id), app]));
    data.reviews.forEach((result, index) => {
      const app = byId.get(String(result.app_id));
      if (!app) return;
      const card = el('section', 'ab-review-card' + (index === 0 ? ' ab-review-card-own' : ''));
      const header = el('div', 'ab-review-card-header');
      const identity = el('div', 'ab-review-identity'); identity.append(icon(app));
      const name = el('div'); name.append(el('h4', '', app.name), el('span', '', index === 0 ? 'YOUR APP · BASELINE' : 'COMPETITOR')); identity.append(name);
      header.append(identity, el('span', 'ab-review-sample', number(result.sample_count) + ' PUBLIC REVIEWS'));
      const stats = el('div', 'ab-review-stats');
      const countStat = el('div'); countStat.append(el('span', '', 'Written reviews fetched'), el('strong', '', number(result.sample_count)));
      const ratingStat = el('div'); ratingStat.append(el('span', '', 'Sample average'), el('strong', '', result.average_rating === null ? NA : Number(result.average_rating).toFixed(2) + ' / 5'));
      stats.append(countStat, ratingStat);
      card.append(header, stats);
      if (!result.sample_count) {
        card.append(el('p', 'ab-review-empty', 'No written reviews are available in the public feed for this storefront.'));
        ui.reviewsResults.append(card); return;
      }
      const bar = el('div', 'ab-sentiment-bar');
      ['positive', 'neutral', 'negative'].forEach(kind => { const segment = el('span', 'ab-sentiment-segment ab-sentiment-segment-' + kind); segment.style.width = Number(result.sentiment_percent[kind] || 0) + '%'; bar.append(segment); });
      bar.setAttribute('aria-label', 'Sentiment distribution');
      const legend = el('div', 'ab-sentiment-legend');
      ['positive', 'neutral', 'negative'].forEach(kind => { const entry = el('span', 'ab-legend-' + kind); entry.append(el('i'), el('span', '', sentimentLabel(kind)), el('strong', '', Number(result.sentiment_percent[kind] || 0).toFixed(1) + '%'), el('small', '', number(result.sentiment[kind]) + ' reviews')); legend.append(entry); });
      const reviewList = el('div', 'ab-review-list'); reviewList.append(el('h5', 'ab-review-list-title', 'Latest written reviews'));
      result.reviews.slice(0, 5).forEach(review => reviewList.append(reviewItem(review)));
      card.append(bar, legend, reviewList);
      ui.reviewsResults.append(card);
    });
    ui.reviewsStatus.textContent = 'Sentiment calculated from ' + number(data.reviews.reduce((sum, item) => sum + Number(item.sample_count || 0), 0)) + ' public written reviews across the selected apps.';
  }
  async function loadReviews() {
    if (!state.snapshot || state.snapshot.apps.length < 2) {
      ui.reviewsStatus.textContent = 'Compare at least two apps, then open this tab to load public reviews.';
      return;
    }
    if (state.reviews || state.reviewsBusy) return;
    const sequence = ++state.reviewsSequence, country = state.country;
    const ids = state.snapshot.apps.map(app => Number(app.id));
    const controller = new AbortController(); state.reviewsController = controller; state.reviewsBusy = true;
    ui.reviewsStatus.textContent = 'Loading the latest public written reviews…';
    try {
      const data = await request('/api/app-benchmark/reviews', { method: 'POST', headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' }, body: JSON.stringify({ app_ids: ids, country: country }) }, controller, 'reviews');
      if (sequence !== state.reviewsSequence || country !== state.country) return;
      state.reviews = data; renderReviews(data);
    } catch (error) {
      if (sequence !== state.reviewsSequence || error.name === 'AbortError') return;
      ui.reviewsStatus.textContent = errorMessage(error) + ' Select this tab again to retry.';
    } finally {
      if (sequence === state.reviewsSequence) { state.reviewsBusy = false; state.reviewsController = null; }
    }
  }
  function charts(apps) {
    ui.ratingChart.replaceChildren(); ui.shareChart.replaceChildren();
    apps.forEach((app, index) => ui.ratingChart.append(barRow(app, numeric(app.rating_count) === 0 ? 0 : (numeric(app.rating) || 0) * 20, rating(app.rating, app.rating_count), index)));
    if (apps.some(app => numeric(app.rating_count) === null)) { ui.shareChart.append(el('p', 'ab-chart-note', 'Rating share is unavailable because a selected app has no rating count.')); return; }
    const total = apps.reduce((sum, app) => sum + Number(app.rating_count), 0);
    if (total <= 0) { ui.shareChart.append(el('p', 'ab-chart-note', 'These apps have no ratings in this market yet.')); return; }
    apps.forEach((app, index) => { const share = Number(app.rating_count) / total * 100; ui.shareChart.append(barRow(app, share, share.toFixed(1) + '% · ' + number(app.rating_count), index)); });
  }
  function table(apps) {
    const tableNode = el('table', 'ab-table');
    const thead = el('thead'); const head = el('tr'); const title = el('th', '', 'Metric'); title.scope = 'col'; head.append(title);
    apps.forEach((app, index) => { const th = el('th', index === 0 ? 'ab-own-column' : '', app.name); th.scope = 'col'; if (!index) th.append(el('span', 'ab-cell-note', 'Your app · Baseline')); head.append(th); });
    thead.append(head); tableNode.append(thead);
    const tbody = el('tbody');
    metrics.forEach(metric => {
      const row = el('tr', metric.group ? 'ab-table-row-category' : '');
      if (metric.group) { const th = el('th', '', metric.group); th.colSpan = apps.length + 1; row.append(th); }
      else {
        const th = el('th', '', metric.label); th.scope = 'row'; row.append(th);
        apps.forEach((app, index) => {
          const value = metric.value(app); const td = el('td', index === 0 ? 'ab-own-column' : '', value);
          if (value === NA) td.classList.add('ab-unavailable');
          const note = metric.note && metric.note(app);
          if (note && value !== NA) td.append(el('span', 'ab-cell-note', note));
          row.append(td);
        });
      }
      tbody.append(row);
    });
    tableNode.append(tbody); ui.table.replaceChildren(tableNode);
  }
  function renderSnapshot() {
    const snapshot = state.snapshot;
    ui.meta.replaceChildren();
    ui.meta.append(el('strong', '', 'Apple App Store · ' + marketName() + ' · ' + snapshot.apps.length + ' apps'));
    const fetched = dateValue(snapshot.fetched_at);
    ui.meta.append(el('span', '', fetched ? 'Retrieved ' + new Intl.DateTimeFormat('en-GB', { dateStyle: 'medium', timeStyle: 'short' }).format(fetched) + ' (your local time)' : 'Retrieval time not available'));
    state.reviews = null; ui.reviewsResults.replaceChildren(); ui.reviewsStatus.textContent = 'Open Review sentiment to load the latest public written reviews.';
    appCards(snapshot.apps); charts(snapshot.apps); table(snapshot.apps); renderSignals(snapshot.apps);
    ui.empty.hidden = true; ui.results.hidden = false; ui.export.hidden = false;
    setTab('public');
  }
  async function compare() {
    if (state.busy || !state.baseline || state.rivals.length < 1) return;
    clearSearch('own', false); clearSearch('rival', false);
    const sequence = ++state.compareSequence, country = state.country, ids = chosen().map(app => String(app.id));
    const controller = new AbortController(); state.compareController = controller; state.busy = true;
    mainStatus('Fetching current public store data for your selected apps…'); updateButton();
    try {
      const data = await request('/api/app-benchmark/compare', { method: 'POST', headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' }, body: JSON.stringify({ app_ids: ids.map(Number), country: country }) }, controller);
      if (sequence !== state.compareSequence || country !== state.country) return;
      const byId = new Map(data.apps.filter(app => app && app.id).map(app => [String(app.id), app]));
      const apps = ids.map(id => byId.get(id)).filter(Boolean);
      if (!byId.has(ids[0])) throw new Error('Your baseline app is unavailable in this market. Choose another app and try again.');
      if (apps.length < 2) throw new Error('The selected competitors are unavailable in this market. Choose another competitor and try again.');
      state.snapshot = Object.assign({}, data, { apps: apps, country: country });
      renderSnapshot();
      const missing = ids.length - apps.length;
      mainStatus(missing ? missing + ' selected app' + (missing === 1 ? ' is' : 's are') + ' unavailable in this market. Showing ' + apps.length + ' available apps; your selection is preserved.' : '', !!missing);
    } catch (error) {
      if (sequence !== state.compareSequence || error.name === 'AbortError') return;
      mainStatus(errorMessage(error) + ' Your app selections have been kept.', true);
    } finally {
      if (sequence === state.compareSequence) { state.busy = false; state.compareController = null; updateButton(); }
    }
  }
  function setTab(name) {
    const publicActive = name === 'public', signalsActive = name === 'signals', reviewsActive = name === 'reviews';
    ui.publicPanel.hidden = !publicActive; ui.signalsPanel.hidden = !signalsActive; ui.reviewsPanel.hidden = !reviewsActive;
    [[ui.publicTab, publicActive], [ui.signalsTab, signalsActive], [ui.reviewsTab, reviewsActive]].forEach(([tab, active]) => { tab.classList.toggle('ab-tab-active', active); tab.setAttribute('aria-selected', String(active)); tab.tabIndex = active ? 0 : -1; });
    ui.export.hidden = !publicActive || !state.snapshot;
    if (reviewsActive) loadReviews();
  }
  function csvCell(value) {
    let text = value === null || value === undefined ? NA : String(value);
    if (/^[\s\u0000-\u001f]*[=+@-]/.test(text) || /^[\t\r\n]/.test(text)) text = "'" + text;
    return '"' + text.replace(/"/g, '""') + '"';
  }
  function exportCSV() {
    if (!state.snapshot) return;
    const apps = state.snapshot.apps;
    const rows = [['Metric', ...apps.map(app => app.name)], ['Role', ...apps.map((_, index) => index ? 'Competitor' : 'Your app / baseline')], ['Source', ...apps.map(() => 'Apple App Store public API')], ['Market', ...apps.map(() => state.snapshot.country.toUpperCase())], ['Retrieved at', ...apps.map(() => textValue(state.snapshot.fetched_at))], ['App ID', ...apps.map(app => app.id)], ['Developer', ...apps.map(app => textValue(app.developer))], ['Store URL', ...apps.map(app => safeURL(app.store_url, 'store') || NA)]];
    metrics.filter(metric => !metric.group).forEach(metric => rows.push([metric.label, ...apps.map(app => metric.value(app))]));
    signalData(apps).forEach(({ model, values }) => {
      rows.push([model.label, ...values.map(item => model.display(item.raw))]);
      rows.push([model.label + ' vs baseline', ...values.map((item, index) => index === 0 ? 'Baseline' : item.raw !== null && values[0].raw !== null && values[0].raw !== 0 ? (item.raw / values[0].raw).toFixed(2) + 'x' : NA)]);
    });
    if (state.reviews) {
      const reviewsById = new Map(state.reviews.reviews.map(result => [String(result.app_id), result]));
      const reviewValue = (app, read) => { const result = reviewsById.get(String(app.id)); return result ? read(result) : NA; };
      rows.push(['Written review sample count', ...apps.map(app => reviewValue(app, result => result.sample_count))]);
      rows.push(['Sample review average', ...apps.map(app => reviewValue(app, result => result.average_rating === null ? NA : result.average_rating))]);
      rows.push(['Positive review share', ...apps.map(app => reviewValue(app, result => result.sentiment_percent.positive + '%'))]);
      rows.push(['Neutral review share', ...apps.map(app => reviewValue(app, result => result.sentiment_percent.neutral + '%'))]);
      rows.push(['Negative review share', ...apps.map(app => reviewValue(app, result => result.sentiment_percent.negative + '%'))]);
    }
    rows.push(['Download price note', ...apps.map(() => 'Excludes in-app purchases and subscriptions')]);
    rows.push(['Data limitations', ...apps.map(() => 'Current public store snapshot; exact installs, uninstalls, revenue, active users and engagement are not public')]);
    const csv = '\ufeff' + rows.map(row => row.map(csvCell).join(',')).join('\r\n');
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
    const link = el('a'); link.href = url; link.download = 'app-benchmark-' + state.country + '-' + new Date().toISOString().slice(0, 10) + '.csv';
    document.body.append(link); link.click(); link.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  function init() {
    if (initialized) return;
    const root = document.getElementById('appBenchmarkWorkspaceContainer');
    if (!root) return;
    initialized = true;
    const get = id => root.querySelector('#' + id);
    Object.assign(ui, { country: get('abCountry'), compare: get('abCompareButton'), summary: get('abSelectionSummary'), mainStatus: get('abMainStatus'), empty: get('abEmptyState'), results: get('abComparisonResults'), cards: get('abAppCards'), meta: get('abSnapshotMeta'), ratingChart: get('abRatingChart'), shareChart: get('abShareChart'), table: get('abComparisonTable'), export: get('abExportButton'), publicTab: get('abPublicTab'), signalsTab: get('abSignalsTab'), reviewsTab: get('abReviewsTab'), publicPanel: get('abPublicPanel'), signalsPanel: get('abSignalsPanel'), reviewsPanel: get('abReviewsPanel'), signalsResults: get('abSignalsResults'), reviewsResults: get('abReviewsResults'), reviewsStatus: get('abReviewsStatus') });
    ui.own = { input: get('abOwnSearch'), status: get('abOwnStatus'), results: get('abOwnResults'), selection: get('abOwnSelection') };
    ui.rival = { input: get('abRivalSearch'), status: get('abRivalStatus'), results: get('abRivalResults'), selection: get('abRivalSelections'), hint: get('abRivalHint'), count: get('abRivalCount') };
    state.country = ui.country.value;
    ['own', 'rival'].forEach(role => {
      ui[role].input.addEventListener('input', () => queueSearch(role));
      ui[role].input.addEventListener('keydown', event => {
        if (event.key === 'Escape') clearSearch(role, false);
        if (event.key === 'ArrowDown' && !ui[role].results.hidden) { const first = ui[role].results.querySelector('button:not(:disabled)'); if (first) { event.preventDefault(); first.focus(); } }
        if (event.key === 'Enter') { event.preventDefault(); queueSearch(role); }
      });
      ui[role].results.addEventListener('keydown', event => {
        if (event.key === 'Escape') { clearSearch(role, false); ui[role].input.focus(); return; }
        if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
        const buttons = Array.from(ui[role].results.querySelectorAll('button:not(:disabled)')); const index = buttons.indexOf(document.activeElement);
        if (index < 0 || !buttons.length) return;
        event.preventDefault(); buttons[(index + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length].focus();
      });
    });
    ui.country.addEventListener('change', () => {
      state.country = ui.country.value; state.baseline = null; state.rivals = [];
      clearSearch('own', true); clearSearch('rival', true); invalidateSnapshot(); renderSelections();
      mainStatus('Market changed to ' + marketName() + '. Choose apps from this storefront to start a new comparison.');
    });
    ui.compare.addEventListener('click', compare);
    ui.export.addEventListener('click', exportCSV);
    ui.publicTab.addEventListener('click', () => setTab('public'));
    ui.signalsTab.addEventListener('click', () => setTab('signals'));
    ui.reviewsTab.addEventListener('click', () => setTab('reviews'));
    const tabs = [ui.publicTab, ui.signalsTab, ui.reviewsTab], tabNames = ['public', 'signals', 'reviews'];
    tabs.forEach(tab => tab.addEventListener('keydown', event => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault();
      const current = tabs.indexOf(tab);
      const next = event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length - 1 : (current + (event.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
      setTab(tabNames[next]); tabs[next].focus();
    }));
    renderSelections();
  }
  window.AppBenchmark = Object.freeze({ init: init });
})();
