(function () {
  "use strict";

  const STORAGE_KEY = "dataprovidoExcelWorkbookId";
  const state = {
    mode: "operations",
    workbookId: readSessionWorkbook(),
    lastCommand: "",
    pendingClarification: null,
    selectedClarification: "",
    selectedOperation: "",
    selectedPrompt: "",
    downloadUrl: ""
  };

  function byId(id) { return document.getElementById(id); }
  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
  function readSessionWorkbook() {
    try { return sessionStorage.getItem(STORAGE_KEY) || ""; } catch (_) { return ""; }
  }
  function setWorkbookId(workbookId) {
    state.workbookId = String(workbookId || "");
    try {
      if (state.workbookId) sessionStorage.setItem(STORAGE_KEY, state.workbookId);
      else sessionStorage.removeItem(STORAGE_KEY);
    } catch (_) {}
    return state.workbookId;
  }
  function getWorkbookId() { return state.workbookId || readSessionWorkbook(); }
  function clearWorkbookId() {
    state.downloadUrl = "";
    state.pendingClarification = null;
    return setWorkbookId("");
  }

  function clearOperationHint() {
    state.selectedOperation = "";
    state.selectedPrompt = "";
    document.querySelectorAll("[data-xw-operation].is-selected").forEach((button) => {
      button.classList.remove("is-selected");
      button.setAttribute("aria-pressed", "false");
    });
  }

  function getOperationHint() {
    const currentPrompt = byId("questionInput")?.value?.trim() || "";
    if (!state.selectedOperation || currentPrompt !== state.selectedPrompt) return "";
    return state.selectedOperation;
  }

  function setMode(mode) {
    if (mode === "ask") mode = "commerce";
    if (mode === "transform") mode = "operations";
    const nextMode = mode === "commerce" ? "commerce" : "operations";
    if (nextMode !== state.mode) clearOperationHint();
    state.mode = nextMode;
    const operations = state.mode === "operations";
    const operationButton = byId("xwOperationsMode");
    const commerceButton = byId("xwCommerceMode");
    const operationPanel = byId("xwOperationsPanel");
    const commercePanel = byId("xwCommercePanel");
    operationButton?.classList.toggle("is-active", operations);
    commerceButton?.classList.toggle("is-active", !operations);
    operationButton?.setAttribute("aria-selected", String(operations));
    commerceButton?.setAttribute("aria-selected", String(!operations));
    if (operationPanel) operationPanel.hidden = !operations;
    if (commercePanel) commercePanel.hidden = operations;

    const heading = byId("xwCommandHeading");
    const help = byId("xwCommandHelp");
    const input = byId("questionInput");
    const runText = byId("runBtnText");
    if (operations) {
      if (heading) heading.textContent = "Yapılacak Excel işlemini yazın";
      if (help) help.textContent = "Komutunuz dosyanızdaki gerçek kolonlarla kontrol edilir.";
      if (input) input.placeholder = "Örn: Brand APPLE olanları filtrele ve Revenue değerine göre büyükten küçüğe sırala…";
      if (runText) runText.textContent = "İşlemi uygula";
    } else {
      if (heading) heading.textContent = "E-ticaret sorunuzu yazın";
      if (help) help.textContent = "Analiz dosyanızı değiştirmeden gerçek kolonlarınız üzerinden hesaplanır.";
      if (input) input.placeholder = "Örn: Revenue kolonuna göre en yüksek satış yapan marka hangisi? Payını ve farkını göster…";
      if (runText) runText.textContent = "Analizi çalıştır";
    }
    document.dispatchEvent(new CustomEvent("excel-wizard-mode-change", { detail: { mode: state.mode } }));
  }

  function setPrompt(prompt, mode, operationId) {
    if (mode) setMode(mode);
    clearOperationHint();
    if (operationId) {
      state.selectedOperation = String(operationId);
      state.selectedPrompt = String(prompt || "").trim();
    }
    const input = byId("questionInput");
    if (!input) return;
    input.value = prompt || "";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.focus();
    input.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  function normalizePreview(data) {
    const preview = data?.preview || {};
    const rows = preview.rows || data?.preview_rows || data?.rows || data?.processed_rows || [];
    const columns = preview.columns || data?.column_names || Object.keys(rows[0] || {});
    return { rows, columns, totalRows: Number(preview.total_rows ?? data?.total_rows ?? rows.length) };
  }

  function renderFileState(data) {
    data = data || {};
    if (data.workbook_id) setWorkbookId(data.workbook_id);
    const preview = normalizePreview(data);
    const profile = data.profile || {};
    const rows = Number(data.total_rows ?? profile.total_rows ?? preview.totalRows ?? 0);
    const columns = Number(data.total_columns ?? profile.total_columns ?? preview.columns.length ?? 0);
    const missing = Number(profile.missing_cells || 0);
    const duplicates = Number(profile.duplicate_rows || 0);
    const totalCells = Math.max(1, rows * columns);
    const quality = rows ? Math.max(0, Math.round(100 - ((missing / totalCells) * 70) - ((duplicates / rows) * 30))) : 0;
    if (byId("xwSummaryFile")) byId("xwSummaryFile").textContent = rows ? (data.filename || "Yüklenen çalışma kitabı") : "Henüz dosya yüklenmedi";
    if (byId("xwRows")) byId("xwRows").textContent = rows ? rows.toLocaleString("tr-TR") : "—";
    if (byId("xwColumns")) byId("xwColumns").textContent = columns ? columns.toLocaleString("tr-TR") : "—";
    if (byId("xwMissing")) byId("xwMissing").textContent = rows ? missing.toLocaleString("tr-TR") : "—";
    if (byId("xwDuplicates")) byId("xwDuplicates").textContent = rows ? duplicates.toLocaleString("tr-TR") : "—";
    if (byId("xwQualityBar")) byId("xwQualityBar").style.width = quality + "%";
    if (byId("xwQualityText")) {
      byId("xwQualityText").textContent = rows
        ? `Veri kalitesi: %${quality} · ${profile.numeric_columns || 0} sayısal, ${profile.text_columns || 0} metin sütunu algılandı.`
        : "Dosya yüklendiğinde veri kalitesi burada kontrol edilir.";
    }
    if (byId("xwDropTitle")) byId("xwDropTitle").textContent = rows ? "Başka bir Excel dosyası yükleyin" : "Excel dosyanızı buraya bırakın";
    const columnList = byId("xwDetectedColumns");
    if (columnList) {
      columnList.innerHTML = preview.columns.slice(0, 7).map((column) => `<span title="${escapeHtml(column)}">${escapeHtml(column)}</span>`).join("");
      if (preview.columns.length > 7) columnList.insertAdjacentHTML("beforeend", `<span>+${preview.columns.length - 7}</span>`);
    }
  }

  function setLoading(message) {
    const section = byId("resultSection");
    const box = byId("resultBox");
    if (section) section.style.display = "block";
    if (box) {
      box.className = "result-box loading";
      box.innerHTML = `<div class="xw-loading">${escapeHtml(message || "Çalışma kitabı işleniyor…")}</div>`;
    }
  }

  function renderError(message) {
    const section = byId("resultSection");
    const box = byId("resultBox");
    if (section) section.style.display = "block";
    if (box) {
      box.className = "result-box";
      box.innerHTML = `<div class="xw-error">⚠ ${escapeHtml(message || "İşlem tamamlanamadı.")}</div>`;
    }
  }

  function renderClarification(data, originalCommand) {
    const section = byId("resultSection");
    const box = byId("resultBox");
    state.lastCommand = originalCommand || state.lastCommand || byId("questionInput")?.value?.trim() || "";
    state.pendingClarification = {
      id: String(data.clarification_id || "answer"),
      question: data.question || "Hangi kolonu kullanmalıyım?",
      role: String(data.role || data.field || "")
    };
    state.selectedClarification = "";
    const options = Array.isArray(data.options) ? data.options : [];
    const optionHtml = options.map((option, index) => {
      const item = typeof option === "string" ? { label: option, value: option } : option;
      return `<button type="button" data-xw-clarification-option data-value="${escapeHtml(item.value ?? item.label ?? "")}" aria-pressed="false"><strong>${escapeHtml(item.label ?? item.value ?? `Seçenek ${index + 1}`)}</strong>${item.description ? `<small>${escapeHtml(item.description)}</small>` : ""}</button>`;
    }).join("");
    if (section) section.style.display = "block";
    if (!box) return;
    box.className = "result-box";
    box.innerHTML = `
      <div id="xwClarificationCard" class="xw-clarification-card">
        <div class="xw-clarification-head">
          <span>?</span>
          <div><small>KOLON EŞLEŞTİRMESİ GEREKİYOR</small><strong>${escapeHtml(data.question || "Hangi kolonu kullanmalıyım?")}</strong><p>Doğru kolonu seçin veya dosyanızdaki kolon adını yazın.</p></div>
        </div>
        <div class="xw-clarification-options">${optionHtml}</div>
        <div class="xw-clarification-reply">
          <input id="xwClarificationAnswer" type="text" autocomplete="off" placeholder="Kolon adını yazabilirsiniz…" aria-label="Kolon cevabı">
          <button id="xwClarificationSubmit" type="button">Bu kolonla devam et →</button>
        </div>
        <p class="xw-clarification-foot">Bu eşleştirme yalnızca geçici çalışma kitabınız için kullanılır; kalıcı olarak saklanmaz.</p>
      </div>`;
    box.querySelectorAll("[data-xw-clarification-option]").forEach((button) => {
      button.addEventListener("click", () => {
        box.querySelectorAll("[data-xw-clarification-option]").forEach((item) => {
          item.classList.remove("is-selected");
          item.setAttribute("aria-pressed", "false");
        });
        button.classList.add("is-selected");
        button.setAttribute("aria-pressed", "true");
        state.selectedClarification = button.dataset.value || "";
        const input = byId("xwClarificationAnswer");
        if (input) input.value = state.selectedClarification;
      });
    });
    byId("xwClarificationSubmit")?.addEventListener("click", submitClarification);
    byId("xwClarificationAnswer")?.addEventListener("keydown", (event) => {
      if (event.key === "Enter") submitClarification();
    });
    box.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }

  function submitClarification() {
    const pending = state.pendingClarification;
    if (!pending) return;
    const typed = byId("xwClarificationAnswer")?.value?.trim() || "";
    const answer = typed || state.selectedClarification;
    if (!answer) {
      byId("xwClarificationAnswer")?.focus();
      return;
    }
    const clarificationAnswers = {
      clarification_id: pending.id,
      selected_value: answer,
      answer
    };
    if (pending.role) clarificationAnswers[pending.role] = answer;
    if (typeof window.runModule === "function") {
      window.runModule({ command: state.lastCommand, clarification_answers: clarificationAnswers });
    }
  }

  function metricEntries(metrics) {
    const labels = {
      total: "Toplam",
      groups: "Grup sayısı",
      leader: "Lider",
      leader_value: "Lider değeri",
      leader_share: "Lider payı",
      leader_gap: "İkinciyle fark",
      strongest_category: "En güçlü kategori",
      weakest_category: "En zayıf kategori",
      total_revenue: "Toplam ciro",
      total_orders: "Toplam sipariş",
      total_units: "Satılan adet",
      average_order_value: "Ortalama sepet",
      total_views: "PDP görüntüleme",
      total_add_to_carts: "Sepete ekleme",
      total_transactions: "Satın alma",
      view_to_purchase_rate: "View → purchase",
      largest_drop_stage: "En büyük kayıp aşaması",
      largest_drop_count: "Kayıp hacmi",
      weakest_product: "En zayıf ürün",
      at_risk_products: "Riskli ürün",
      priority_product: "İlk öncelik",
      revenue_at_risk: "Risk altındaki ciro",
      demand_metric: "Talep metriği",
      rows: "Satır",
      columns: "Sütun"
    };
    const formatValue = (label, value) => {
      if (typeof value !== "number" || !Number.isFinite(value)) return value;
      if (/(share|rate|conversion|oran|pay)/i.test(label) && Math.abs(value) <= 1) {
        return new Intl.NumberFormat("tr-TR", { style: "percent", maximumFractionDigits: 1 }).format(value);
      }
      return new Intl.NumberFormat("tr-TR", { maximumFractionDigits: 2 }).format(value);
    };
    if (Array.isArray(metrics)) {
      return metrics.slice(0, 4).map((item, index) => {
        const label = item.label || item.name || `Metrik ${index + 1}`;
        return [label, item.formatted_value ?? formatValue(label, item.value) ?? "—"];
      });
    }
    if (metrics && typeof metrics === "object") {
      let priority = [];
      if ("largest_drop_stage" in metrics) priority = ["total_views", "total_add_to_carts", "total_transactions", "view_to_purchase_rate"];
      else if ("at_risk_products" in metrics) priority = ["at_risk_products", "priority_product", "revenue_at_risk", "demand_metric"];
      else if ("strongest_category" in metrics) priority = ["strongest_category", "total_revenue", "total_orders", "average_order_value"];
      else if ("leader" in metrics) priority = ["leader", "leader_value", "leader_share", "leader_gap"];
      const entries = priority.length
        ? priority.filter((key) => key in metrics).map((key) => [key, metrics[key]])
        : Object.entries(metrics).slice(0, 4);
      return entries.map(([label, value]) => [labels[label] || label.replace(/_/g, " "), formatValue(label, value)]);
    }
    return [];
  }

  function formatPreviewCell(column, value) {
    if (typeof value !== "number" || !Number.isFinite(value)) return value ?? "—";
    if (/(share|rate|conversion|oran|pay)/i.test(column) && Math.abs(value) <= 1) {
      return new Intl.NumberFormat("tr-TR", { style: "percent", maximumFractionDigits: 1 }).format(value);
    }
    return new Intl.NumberFormat("tr-TR", { maximumFractionDigits: 2 }).format(value);
  }

  function renderExecuteResponse(data, context) {
    context = context || {};
    if (data?.status === "needs_clarification") {
      renderClarification(data, context.command);
      return { handled: true, clarification: true };
    }
    if (!data || data.status !== "success") {
      renderError(data?.error || data?.detail || "Excel işlemi tamamlanamadı.");
      return { handled: true, error: true };
    }
    const result = data.result || data;
    const preview = data.preview || {};
    const hasResultPreview = Array.isArray(result.preview_rows);
    const rows = result.preview_rows || preview.rows || data.rows || [];
    const columns = hasResultPreview && rows.length
      ? Object.keys(rows[0] || {})
      : (preview.columns || Object.keys(rows[0] || {}));
    const summary = result.summary || data.executive_summary || data.action_note || "İşlem başarıyla tamamlandı.";
    const mutation = result.is_mutation !== false && state.mode === "operations";
    const used = result.columns_used || [];
    const usedColumns = Array.isArray(used) ? used : Object.entries(used).map(([role, column]) => `${role}: ${column}`);
    const steps = Array.isArray(result.steps) ? result.steps : [];
    const metrics = metricEntries(result.metrics);
    state.lastCommand = context.command || state.lastCommand;
    state.pendingClarification = null;
    state.downloadUrl = data.download_url || (getWorkbookId() ? `/api/excel-wizard/download/${encodeURIComponent(getWorkbookId())}` : "");

    const metricHtml = metrics.length ? `<div class="xw-result-metrics">${metrics.map(([label, value]) => `<div><small>${escapeHtml(label)}</small><strong>${escapeHtml(value)}</strong></div>`).join("")}</div>` : "";
    const columnsHtml = usedColumns.length || steps.length
      ? `<div class="xw-result-meta"><span>${steps.length ? `${steps.length} adım uygulandı` : "Kullanılan kolonlar"}</span>${usedColumns.slice(0, 8).map((column) => `<b>${escapeHtml(column)}</b>`).join("")}</div>`
      : "";
    const visibleColumns = (columns.length ? columns : Object.keys(rows[0] || {})).slice(0, 12);
    const tableHtml = rows.length ? `
      <details class="xw-result-preview">
        <summary>Sonuç önizlemesini aç · ${Number(hasResultPreview ? rows.length : (preview.total_rows || rows.length)).toLocaleString("tr-TR")} satır</summary>
        <div class="xw-result-table-wrap"><table class="xw-result-table">
          <thead><tr><th>#</th>${visibleColumns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr></thead>
          <tbody>${rows.slice(0, 50).map((row, index) => `<tr><td>${index + 1}</td>${visibleColumns.map((column) => `<td>${escapeHtml(formatPreviewCell(column, row?.[column]))}</td>`).join("")}</tr>`).join("")}</tbody>
        </table></div>
      </details>` : "";

    const section = byId("resultSection");
    const box = byId("resultBox");
    if (section) section.style.display = "block";
    if (box) {
      box.className = "result-box";
      box.innerHTML = `
        <div class="xw-result-card">
          <div class="xw-result-top">
            <div class="xw-result-title"><span class="xw-result-icon">${mutation ? "✓" : "↗"}</span><div><small>${mutation ? "EXCEL İŞLEMİ TAMAMLANDI" : "E-TİCARET ANALİZİ HAZIR"}</small><strong>${escapeHtml(context.command || state.lastCommand || "İşlem sonucu")}</strong><p>${escapeHtml(summary)}</p></div></div>
            <div class="xw-result-actions"><button type="button" onclick="openExcelPreview()">Çalışma kitabı</button>${mutation && state.downloadUrl ? '<button class="primary" type="button" onclick="downloadExcel()">Excel’i indir</button>' : ""}</div>
          </div>
          ${metricHtml}${columnsHtml}${tableHtml}
        </div>`;
      box.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
    document.dispatchEvent(new CustomEvent("excel-wizard-result", { detail: { data, result, preview } }));
    return { handled: true, result, preview, rows };
  }

  function updateNetworkStatus() {
    const label = byId("xwNetworkLabel");
    if (label) label.textContent = navigator.onLine ? "Çevrimiçi" : "Çevrimdışı";
  }

  function init() {
    document.querySelectorAll("[data-xw-prompt]").forEach((button) => {
      button.setAttribute("aria-pressed", "false");
      button.addEventListener("click", () => {
        const operationId = button.dataset.xwOperation || "";
        setPrompt(button.dataset.xwPrompt, button.dataset.xwMode, operationId);
        if (operationId) {
          button.classList.add("is-selected");
          button.setAttribute("aria-pressed", "true");
        }
      });
    });
    byId("questionInput")?.addEventListener("input", (event) => {
      if (state.selectedOperation && event.target.value.trim() !== state.selectedPrompt) clearOperationHint();
    });
    const zone = byId("excelWizardDropzone");
    if (zone) {
      ["dragenter", "dragover"].forEach((name) => zone.addEventListener(name, (event) => { event.preventDefault(); zone.classList.add("is-dragover"); }));
      ["dragleave", "drop"].forEach((name) => zone.addEventListener(name, (event) => { event.preventDefault(); zone.classList.remove("is-dragover"); }));
      zone.addEventListener("drop", (event) => {
        const file = event.dataTransfer?.files?.[0];
        if (file && typeof window.handleExcelFileUpload === "function") window.handleExcelFileUpload(file);
      });
    }
    window.addEventListener("online", updateNetworkStatus);
    window.addEventListener("offline", updateNetworkStatus);
    updateNetworkStatus();
    setMode("operations");
  }

  window.ExcelWizardUI = {
    init,
    setMode,
    setPrompt,
    renderFileState,
    renderExecuteResponse,
    renderClarification,
    renderError,
    setLoading,
    setWorkbookId,
    getWorkbookId,
    clearWorkbookId,
    clearOperationHint,
    getOperationHint,
    getDownloadUrl: () => state.downloadUrl,
    getMode: () => state.mode
  };
  document.addEventListener("DOMContentLoaded", init);
})();
