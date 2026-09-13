(function () {
  "use strict";

  const state = { mode: "ask" };

  function byId(id) { return document.getElementById(id); }

  function setMode(mode) {
    state.mode = mode === "transform" ? "transform" : "ask";
    byId("xwAskMode")?.classList.toggle("is-active", state.mode === "ask");
    byId("xwTransformMode")?.classList.toggle("is-active", state.mode === "transform");
    const heading = byId("xwCommandHeading");
    const help = byId("xwCommandHelp");
    const input = byId("questionInput");
    const runText = byId("runBtnText");
    if (state.mode === "ask") {
      if (heading) heading.textContent = "Veriniz hakkında sorunuzu yazın";
      if (help) help.textContent = "Cevap verirken çalışma kitabını değiştirmeyeceğiz.";
      if (input) input.placeholder = "Örn: En yüksek ciroyu hangi marka üretiyor? Sonucu sade bir dille açıkla…";
      if (runText) runText.textContent = "Soruyu yanıtla";
    } else {
      if (heading) heading.textContent = "Excel’de yapılacak işlemi yazın";
      if (help) help.textContent = "Değişikliği önce gösterecek, ardından indirmenize izin vereceğiz.";
      if (input) input.placeholder = "Örn: Tekrar eden satırları kaldır ve ciroya göre büyükten küçüğe sırala…";
      if (runText) runText.textContent = "İşlemi uygula";
    }
    document.dispatchEvent(new CustomEvent("excel-wizard-mode-change", { detail: { mode: state.mode } }));
  }

  function setPrompt(prompt, mode) {
    if (mode) setMode(mode);
    const input = byId("questionInput");
    if (!input) return;
    input.value = prompt || "";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.focus();
    input.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  function renderFileState(data) {
    if (!data) return;
    const rows = Number(data.total_rows || 0);
    const columns = Number(data.total_columns || 0);
    const profile = data.profile || {};
    const missing = Number(profile.missing_cells || 0);
    const duplicates = Number(profile.duplicate_rows || 0);
    const totalCells = Math.max(1, rows * columns);
    const quality = rows ? Math.max(0, Math.round(100 - ((missing / totalCells) * 70) - ((duplicates / rows) * 30))) : 0;
    if (byId("xwSummaryFile")) byId("xwSummaryFile").textContent = rows ? data.filename : "Henüz dosya yüklenmedi";
    if (byId("xwRows")) byId("xwRows").textContent = rows ? rows.toLocaleString("tr-TR") : "—";
    if (byId("xwColumns")) byId("xwColumns").textContent = columns ? columns.toLocaleString("tr-TR") : "—";
    if (byId("xwMissing")) byId("xwMissing").textContent = rows ? missing.toLocaleString("tr-TR") : "—";
    if (byId("xwDuplicates")) byId("xwDuplicates").textContent = rows ? duplicates.toLocaleString("tr-TR") : "—";
    if (byId("xwQualityBar")) byId("xwQualityBar").style.width = quality + "%";
    if (byId("xwQualityText")) byId("xwQualityText").textContent = rows ? `Veri kalitesi: %${quality} · ${profile.numeric_columns || 0} sayısal, ${profile.text_columns || 0} metin sütunu algılandı.` : "Dosya yüklendiğinde veri kalitesi burada kontrol edilir.";
    if (byId("xwDropTitle")) byId("xwDropTitle").textContent = rows ? "Başka bir Excel dosyası yükleyin" : "Excel dosyanızı buraya bırakın";
  }

  function updateNetworkStatus() {
    const label = byId("xwNetworkLabel");
    if (!label) return;
    label.textContent = navigator.onLine ? "Çevrimiçi · yerel işlem hazır" : "Çevrimdışı · yerel işlem modu";
  }

  function init() {
    document.querySelectorAll("[data-xw-prompt]").forEach((button) => {
      button.addEventListener("click", () => setPrompt(button.dataset.xwPrompt, button.dataset.xwMode));
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
    setMode("ask");
  }

  window.ExcelWizardUI = { init, setMode, setPrompt, renderFileState, getMode: () => state.mode };
  document.addEventListener("DOMContentLoaded", init);
})();
