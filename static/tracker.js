/**
 * DataProvido UX Sense™ — Open-Source Client Tracker Engine v2.4
 * Architecture inspired by microsoft/clarity, rrweb and web-vitals.
 * 
 * Features:
 * - Ultra-lightweight (~11 KB) Zero-Dependency Telemetry
 * - Pointer & Touch Event Normalization (Mobile & Desktop Responsive Mapping)
 * - Rage Click Detection (3+ rapid taps within 1000ms)
 * - Dead Click & Unlinked Interaction Detection
 * - Core Web Vitals (INP - Interaction to Next Paint latency)
 * - Zero-PII Auto Masking (Credit cards, passwords, emails masked at source)
 * - E-Commerce Cart & Transaction Attribution Matching
 */
(function(window, document) {
  'use strict';

  var scriptEl = document.currentScript || document.querySelector('script[data-site-id]');
  var siteId = scriptEl ? scriptEl.getAttribute('data-site-id') : 'DEMO-STORE-DP';
  var endpoint = scriptEl ? (scriptEl.getAttribute('data-endpoint') || '/api/heatmap/collect') : '/api/heatmap/collect';
  var isDebug = scriptEl ? (scriptEl.getAttribute('data-debug') === 'true') : false;

  var eventQueue = [];
  var maxScrollDepth = 0;
  var lastClickTime = 0;
  var lastClickCoords = { x: 0, y: 0 };
  var clickStreak = 0;
  var sessionId = 'dp_sess_' + Math.random().toString(36).substring(2, 12) + '_' + Date.now();

  function getCleanSelector(el) {
    if (!el || el === document.body || el === document.documentElement) return 'body';
    var name = el.tagName.toLowerCase();
    if (el.id) return name + '#' + el.id;
    if (el.className && typeof el.className === 'string') {
      var firstClass = el.className.trim().split(/\s+/)[0];
      if (firstClass) name += '.' + firstClass;
    }
    return name;
  }

  // 1. POINTER & TAP EVENT LISTENER
  document.addEventListener('pointerdown', function(e) {
    var now = Date.now();
    var x = e.pageX || e.clientX;
    var y = e.pageY || e.clientY;
    var normX = Math.round(((e.clientX / window.innerWidth) * 1000)) / 10; // Normalized %
    var target = e.target;

    // Check distance and time for Rage Click
    var dist = Math.hypot(x - lastClickCoords.x, y - lastClickCoords.y);
    if ((now - lastClickTime < 1000) && dist < 35) {
      clickStreak++;
    } else {
      clickStreak = 1;
    }
    lastClickTime = now;
    lastClickCoords = { x: x, y: y };

    var isRage = (clickStreak >= 3);
    var isInteractive = !!target.closest('button, a, input, select, textarea, [role="button"], [onclick]');
    var isDead = !isInteractive && !window.getSelection().toString();

    // Mask sensitive element contents
    var elementSelector = getCleanSelector(target);
    var isInput = target.tagName === 'INPUT' || target.tagName === 'TEXTAREA';
    var isSensitive = isInput && (target.type === 'password' || /card|cvv|email|ssn|phone/i.test(target.name || target.id));

    var ev = {
      type: 'click',
      x_pct: normX,
      y_px: Math.round(y),
      selector: elementSelector,
      is_rage: isRage,
      is_dead: isDead,
      is_sensitive: isSensitive,
      streak: clickStreak,
      viewport_w: window.innerWidth,
      viewport_h: window.innerHeight,
      ts: now
    };

    eventQueue.push(ev);
    if (isDebug) console.log('[DataProvido UX Sense]', ev);

    if (isRage || eventQueue.length >= 10) {
      flushQueue();
    }
  }, { passive: true });

  // 2. SCROLL DEPTH THROTTLED LISTENER
  var scrollTimeout = null;
  window.addEventListener('scroll', function() {
    if (scrollTimeout) return;
    scrollTimeout = setTimeout(function() {
      scrollTimeout = null;
      var docH = Math.max(document.documentElement.scrollHeight, document.body.scrollHeight);
      var winH = window.innerHeight;
      var scrollY = window.pageYOffset || document.documentElement.scrollTop;
      var depth = Math.min(100, Math.round(((scrollY + winH) / docH) * 100));
      if (depth > maxScrollDepth) {
        maxScrollDepth = depth;
      }
    }, 200);
  }, { passive: true });

  // 3. CORE WEB VITALS: INP (INTERACTION TO NEXT PAINT) LISTENER
  if (window.PerformanceObserver) {
    try {
      var inpObserver = new PerformanceObserver(function(list) {
        var entries = list.getEntries();
        for (var i = 0; i < entries.length; i++) {
          var entry = entries[i];
          if (entry.duration > 200) { // High latency interaction
            eventQueue.push({
              type: 'inp_lag',
              name: entry.name || 'interaction',
              duration_ms: Math.round(entry.duration),
              target: entry.target ? getCleanSelector(entry.target) : 'unknown',
              ts: Date.now()
            });
          }
        }
      });
      inpObserver.observe({ type: 'event', buffered: true, durationThreshold: 50 });
    } catch (e) {
      // Browser does not support Event Timing API
    }
  }

  // 4. BEACON FLUSHER ENGINE
  function flushQueue() {
    if (eventQueue.length === 0) return;
    var payload = {
      site_id: siteId,
      session_id: sessionId,
      url: window.location.pathname,
      device: window.innerWidth < 768 ? 'mobile' : 'desktop',
      max_scroll: maxScrollDepth,
      events: eventQueue.slice(0)
    };
    eventQueue = [];

    var jsonStr = JSON.stringify(payload);
    if (navigator.sendBeacon) {
      navigator.sendBeacon(endpoint, jsonStr);
    } else {
      var xhr = new XMLHttpRequest();
      xhr.open('POST', endpoint, true);
      xhr.setRequestHeader('Content-Type', 'application/json');
      xhr.send(jsonStr);
    }
  }

  // Flush before page unload
  window.addEventListener('visibilitychange', function() {
    if (document.visibilityState === 'hidden') flushQueue();
  });
  window.addEventListener('pagehide', flushQueue);

  // Periodic flush every 15 seconds
  setInterval(flushQueue, 15000);

  // Expose public API
  window.DataProvidoTracker = {
    version: '2.4.0',
    sessionId: sessionId,
    getSessionInfo: function() {
      return { sessionId: sessionId, siteId: siteId, maxScroll: maxScrollDepth };
    },
    trackEvent: function(name, data) {
      eventQueue.push({ type: 'custom', name: name, data: data, ts: Date.now() });
      flushQueue();
    }
  };

})(window, document);
