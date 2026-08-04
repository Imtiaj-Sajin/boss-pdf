// ============================================================
// boss-pdf · client
// ============================================================

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

// ---------- User chip in header ----------
(async function initUserChip() {
  const chip = document.getElementById("userChip");
  const nameEl = document.getElementById("userName");
  const titleEl = document.getElementById("userTitle");
  const avatarEl = document.getElementById("userAvatar");
  const logoutBtn = document.getElementById("logoutBtn");
  if (!chip) return;

  // Instant display from JWT, then enrich from /auth/me.
  const quick = BossAuth.quickUser();
  if (quick) {
    nameEl.textContent = quick.username;
    avatarEl.textContent = (quick.username[0] || "?").toUpperCase();
    chip.classList.remove("hidden");
  }
  logoutBtn.addEventListener("click", () => BossAuth.logout());
  try {
    const me = await BossAuth.me();
    if (me.full_name) nameEl.textContent = me.full_name;
    if (me.designation) titleEl.textContent = me.designation;
    if (me.profile_photo_url) {
      avatarEl.innerHTML = "";
      const img = document.createElement("img");
      img.src = me.profile_photo_url;
      img.alt = "";
      avatarEl.appendChild(img);
    } else {
      avatarEl.textContent = ((me.full_name || me.username || "?")[0]).toUpperCase();
    }
    chip.classList.remove("hidden");
  } catch (_) { /* token invalid — authFetch already redirected */ }
})();

// Track the last usage row id so we can bump downloads on user click.
let lastConvertUsageId = null;
let lastSplitUsageIds = []; // parallel to split results

const SECTION_COLORS = [
  "#c9a15a", // gold
  "#5fdfb0", // mint
  "#6aa7ff", // blue
  "#ff89d2", // pink
  "#f48a4d", // amber
  "#b388ff", // violet
  "#9be36b", // lime
  "#ef6f6f", // red
];

// ---------- DOM refs ----------
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const pickBtn = document.getElementById("pickBtn");

const chooserEl = document.getElementById("chooser");
const chooserName = document.getElementById("chooserName");
const chooserPages = document.getElementById("chooserPages");
const changeFileBtn = document.getElementById("changeFileBtn");

const converterEl = document.getElementById("converter");
const convName = document.getElementById("convName");
const convPages = document.getElementById("convPages");
const convSummary = document.getElementById("convSummary");
const convFrom = document.getElementById("convFrom");
const convTo = document.getElementById("convTo");
const convThumbs = document.getElementById("convThumbs");
const convertBtn = document.getElementById("convertBtn");
const convBack = document.getElementById("convBack");

const splitterEl = document.getElementById("splitter");
const splitName = document.getElementById("splitName");
const splitPages = document.getElementById("splitPages");
const sectionListEl = document.getElementById("sectionList");
const addSectionBtn = document.getElementById("addSectionBtn");
const autoSplitBtn = document.getElementById("autoSplitBtn");
const splitSummary = document.getElementById("splitSummary");
const splitThumbs = document.getElementById("splitThumbs");
const splitBtn = document.getElementById("splitBtn");
const splitBack = document.getElementById("splitBack");
const splitModeBtns = document.querySelectorAll("[data-split-mode]");
const rangesBlock = document.getElementById("rangesBlock");
const rangesToolbar = document.getElementById("rangesToolbar");
const pickPanel = document.getElementById("pickPanel");
const pickClearBtn = document.getElementById("pickClearBtn");
const pickInputEl = document.getElementById("pickInput");
const pickCountEl = document.getElementById("pickCount");
const splitModeHint = document.getElementById("splitModeHint");
const splitPreviewHint = document.getElementById("splitPreviewHint");

const statusEl = document.getElementById("status");
const fileNameEl = document.getElementById("fileName");
const fileSizeEl = document.getElementById("fileSize");
const progressBar = document.getElementById("progressBar");
const statusMsg = document.getElementById("statusMsg");

const resultEl = document.getElementById("result");
const resultMsg = document.getElementById("resultMsg");
const downloadLink = document.getElementById("downloadLink");
const againBtn = document.getElementById("againBtn");
const engineBadgeRow = document.getElementById("engineBadgeRow");
const engineBadge = document.getElementById("engineBadge");
const previewBtn = document.getElementById("previewBtn");
const logDlBtn = document.getElementById("logDlBtn");

const splitResultEl = document.getElementById("splitResult");
const splitListEl = document.getElementById("splitList");
const splitAgainBtn = document.getElementById("splitAgainBtn");

const errorEl = document.getElementById("error");
const errorMsg = document.getElementById("errorMsg");
const retryBtn = document.getElementById("retryBtn");

// ---------- State ----------
let currentFile = null;
let pageCount = 0;
let pdfDoc = null;
let pdfDocFile = null;

// ---------- Batch (folder) state ----------
// When a folder is picked, the SAME converter/splitter UI is configured on the
// first PDF; the primary button then applies that config to every file and
// downloads one ZIP. batchFiles is empty in normal single-file mode.
let batchFiles = [];
const inBatch = () => batchFiles.length > 0;
const batchLabel = () => `${batchFiles.length} PDF${batchFiles.length === 1 ? "" : "s"}`;

// ---------- Per-PDF batch config ----------
// The editor edits ONE file at a time (activeIdx). batchCfg[i] holds that file's
// spec so you can customize any PDF; files you never touch fall back to the
// shared config (file 0's). batchTool is "convert" | "split".
let batchTool = "convert";
let activeIdx = 0;
let batchCfg = [];   // per file: {pages} for convert, {sections,mode,picked} for split

const convSelected = new Set();
let sections = []; // [{ from, to, color }]

// Splitter modes: "ranges" (existing per-row inputs → many PDFs)
// or "pick" (click thumbnails → one PDF made of the picked pages, in order).
let splitMode = "ranges";
const pickedPages = new Set();

const ALL_CARDS = [chooserEl, converterEl, splitterEl, statusEl, resultEl, splitResultEl, errorEl];

// ---------- Helpers ----------
const show = el => el && el.classList.remove("hidden");
const hide = el => el && el.classList.add("hidden");
function showOnly(el) { ALL_CARDS.forEach(c => hide(c)); show(el); }

function fmtSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1024 / 1024).toFixed(2) + " MB";
}

function resetAll() {
  ALL_CARDS.forEach(hide);
  progressBar.style.width = "0%";
  progressBar.classList.remove("indeterminate");
  fileInput.value = "";
  batchFiles = [];
  currentFile = null;
  pageCount = 0;
  pdfDoc = null;
  pdfDocFile = null;
  convSelected.clear();
  sections = [];
  pickedPages.clear();
  splitMode = "ranges";
  batchCfg = [];
  activeIdx = 0;
  document.body.classList.remove("batch-mode");
  [converterEl, splitterEl].forEach(el => el && el.classList.remove("has-panel"));
  convThumbs.innerHTML = "";
  splitThumbs.innerHTML = "";
  sectionListEl.innerHTML = "";
  // Batch card + dropzone live outside ALL_CARDS — restore the entry screen.
  const batchCard = document.getElementById("batch");
  if (batchCard) hide(batchCard);
  show(dropzone);
}

function showError(message) {
  errorMsg.textContent = message || "Something went sideways.";
  showOnly(errorEl);
}

function compressPages(set) {
  const sorted = [...set].sort((a, b) => a - b);
  if (!sorted.length) return "";
  const groups = [];
  let start = sorted[0], prev = sorted[0];
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === prev + 1) { prev = sorted[i]; continue; }
    groups.push(start === prev ? `${start}` : `${start}-${prev}`);
    start = prev = sorted[i];
  }
  groups.push(start === prev ? `${start}` : `${start}-${prev}`);
  return groups.join(",");
}

// ---------- Drag/drop ----------
["dragenter", "dragover"].forEach(ev =>
  dropzone.addEventListener(ev, e => {
    e.preventDefault(); e.stopPropagation();
    dropzone.classList.add("drag");
  })
);
["dragleave", "drop"].forEach(ev =>
  dropzone.addEventListener(ev, e => {
    e.preventDefault(); e.stopPropagation();
    dropzone.classList.remove("drag");
  })
);
dropzone.addEventListener("click", () => fileInput.click());
pickBtn.addEventListener("click", e => { e.stopPropagation(); fileInput.click(); });
fileInput.addEventListener("change", () => {
  if (fileInput.files && fileInput.files[0]) handleFile(fileInput.files[0]);
});
dropzone.addEventListener("drop", e => {
  const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
  if (f) handleFile(f);
});

// ---------- Step 1: receive file ----------
async function handleFile(file) {
  if (!file.name.toLowerCase().endsWith(".pdf")) {
    showError("That doesn't look like a .pdf. The boss only does PDFs.");
    return;
  }
  if (file.size > 500 * 1024 * 1024) {
    showError("File is over 500 MB. The boss has standards.");
    return;
  }
  currentFile = file;

  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = fmtSize(file.size);
  statusMsg.textContent = "The boss is sizing up your file…";
  progressBar.style.width = "0%";
  progressBar.classList.add("indeterminate");
  showOnly(statusEl);

  try {
    const form = new FormData();
    form.append("file", file);
    const res = await BossAuth.authFetch("/api/pdf-info", { method: "POST", body: form });
    if (!res.ok) {
      const j = await res.json().catch(() => ({}));
      throw new Error(j.detail || "Could not read PDF.");
    }
    const info = await res.json();
    pageCount = info.pages || 0;
    if (pageCount < 1) throw new Error("PDF has no pages.");
    showChooser();
  } catch (err) {
    showError(err.message || String(err));
  }
}

// ============================================================
// Batch entry point — called from the folder card in index.html.
// Loads the FIRST PDF into the normal editor; the action button then applies
// that same config to every file in the folder.
// ============================================================
window.BossBatch = {
  async start(files, tool) {
    const first = files[0];
    const form = new FormData();
    form.append("file", first);
    const res = await BossAuth.authFetch("/api/pdf-info", { method: "POST", body: form });
    if (!res.ok) {
      const j = await res.json().catch(() => ({}));
      throw new Error(j.detail || "Could not read the first PDF.");
    }
    const info = await res.json();
    batchFiles = files;
    currentFile = first;
    pdfDoc = null; pdfDocFile = null;
    pageCount = info.pages || 0;
    if (pageCount < 1) throw new Error("That PDF has no pages.");
    batchTool = tool === "split" ? "split" : "convert";
    activeIdx = 0;
    batchCfg = files.map(() => null);   // null = not yet customized (uses shared)
    if (tool === "split") openSplitter();
    else openConverter();
  },
};

// Leaving a batch editor goes back to the folder card, not the chooser.
function leaveEditor() {
  if (inBatch() && window.BossBatchReturn) {
    ALL_CARDS.forEach(hide);
    window.BossBatchReturn();
  } else {
    showChooser();
  }
}

// ============================================================
// Per-PDF batch customization (sidebar of files + per-file specs)
// All of this is guarded so any failure degrades to the old
// "one config applied to all" behaviour rather than breaking the batch.
// ============================================================

function saveActiveCfg() {
  if (!inBatch()) return;
  try {
    if (batchTool === "convert") {
      batchCfg[activeIdx] = { pages: convSelected.size === pageCount ? "all" : compressPages(convSelected) };
    } else {
      batchCfg[activeIdx] = {
        sections: sections.map(s => ({ from: s.from, to: s.to })),
        mode: splitMode,
        picked: [...pickedPages],
      };
    }
  } catch (_) {}
}

function specsFromCfg(c) {
  if (!c) return [];
  if (c.mode === "pick") { const s = compressPages(new Set(c.picked || [])); return s ? [s] : []; }
  return (c.sections || []).map(s => (s.from === s.to ? `${s.from}` : `${s.from}-${s.to}`));
}
function cfgSpecString(i) {
  const c = batchCfg[i];
  if (!c) return "";
  if (batchTool === "convert") return "c:" + (c.pages || "all");
  return "s:" + (c.mode === "pick"
    ? "P:" + (c.picked || []).slice().sort((a, b) => a - b).join(",")
    : (c.sections || []).map(s => `${s.from}-${s.to}`).join("|"));
}
function isCustomized(i) {
  if (i === 0 || !batchCfg[i]) return false;
  return cfgSpecString(i) !== cfgSpecString(0);
}

async function switchBatchFile(idx) {
  if (!inBatch() || idx === activeIdx || idx < 0 || idx >= batchFiles.length) return;
  saveActiveCfg();
  activeIdx = idx;
  currentFile = batchFiles[idx];
  try { pageCount = (await ensurePdfDoc(currentFile)).numPages; } catch (_) {}
  if (batchTool === "split") openSplitter(); else openConverter();
}

function applyToAllFiles() {
  saveActiveCfg();
  const mine = batchCfg[activeIdx];
  if (!mine) return;
  batchCfg = batchFiles.map(() => JSON.parse(JSON.stringify(mine)));
  renderBatchBar();
}

function buildPerFileMap() {
  saveActiveCfg();
  const per = {};
  for (let i = 1; i < batchFiles.length; i++) {
    if (!isCustomized(i)) continue;
    per[batchFiles[i].name] = (batchTool === "convert")
      ? (batchCfg[i].pages || "all")
      : specsFromCfg(batchCfg[i]);
  }
  return per;
}

// Side panel listing every PDF in the folder, beside the editor. Clicking one
// loads it into the editor so its ranges/pages can be set individually.
function renderBatchBar() {
  const split = batchTool === "split";
  const panel = document.getElementById(split ? "splitBatchBar" : "convBatchBar");
  const other = document.getElementById(split ? "convBatchBar" : "splitBatchBar");
  const card = split ? splitterEl : converterEl;
  const otherCard = split ? converterEl : splitterEl;
  if (other) hide(other);
  if (otherCard) otherCard.classList.remove("has-panel");
  if (!panel) return;

  if (!inBatch() || batchFiles.length < 2) {
    hide(panel);
    if (card) card.classList.remove("has-panel");
    document.body.classList.remove("batch-mode");
    return;
  }
  show(panel);
  if (card) card.classList.add("has-panel");
  document.body.classList.add("batch-mode");

  const esc = s => (s || "").replace(/[&<>"]/g, "");
  const items = batchFiles.map((f, i) => {
    const custom = isCustomized(i);
    const cls = "bp-item" + (i === activeIdx ? " active" : "") + (custom ? " custom" : "");
    const sub = custom ? "custom settings"
      : (i === activeIdx ? `${pageCount} page${pageCount === 1 ? "" : "s"}` : fmtSize(f.size));
    return `<button type="button" class="${cls}" data-bidx="${i}" title="${esc(f.name)}">` +
           `<span class="idx">${i + 1}</span>` +
           `<span class="body"><span class="nm">${esc(f.name)}</span>` +
           `<span class="sub">${esc(sub)}</span></span></button>`;
  }).join("");

  panel.innerHTML =
    `<div class="bp-head">PDFs <span class="bp-count">${batchFiles.length}</span></div>` +
    `<div class="bp-list">${items}</div>` +
    `<button type="button" class="bp-apply" id="bbApply">Apply this one to all</button>`;
  panel.querySelectorAll(".bp-item").forEach(el =>
    el.addEventListener("click", () => switchBatchFile(parseInt(el.dataset.bidx, 10))));
  const ap = document.getElementById("bbApply");
  if (ap) ap.addEventListener("click", applyToAllFiles);
}

// ---------- Step 2: chooser ----------
function showChooser() {
  chooserName.textContent = currentFile.name;
  chooserPages.textContent = `${pageCount} page${pageCount === 1 ? "" : "s"} · ${fmtSize(currentFile.size)}`;
  showOnly(chooserEl);
}
document.querySelectorAll(".tool-card").forEach(btn => {
  btn.addEventListener("click", () => {
    const tool = btn.dataset.tool;
    if (tool === "convert") openConverter();
    else if (tool === "split") openSplitter();
  });
});
changeFileBtn.addEventListener("click", () => { resetAll(); });

// ---------- Thumbnails ----------
async function ensurePdfDoc(file) {
  if (pdfDoc && pdfDocFile === file) return pdfDoc;
  const buf = await file.arrayBuffer();
  pdfDoc = await pdfjsLib.getDocument({ data: buf }).promise;
  pdfDocFile = file;
  return pdfDoc;
}

async function renderThumbnails(file, container, onTileClick, decorate) {
  container.innerHTML = `<div class="thumbs-empty">Rendering pages…</div>`;
  let pdf;
  try {
    pdf = await ensurePdfDoc(file);
  } catch (e) {
    container.innerHTML = `<div class="thumbs-empty err">Couldn't render previews: ${e.message || e}</div>`;
    return [];
  }
  container.innerHTML = "";
  const tiles = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const tile = document.createElement("button");
    tile.type = "button";
    tile.className = "page-tile";
    tile.dataset.page = i;
    tile.innerHTML = `
      <div class="thumb"><div class="thumb-skel"></div></div>
      <div class="page-num">${i}</div>
    `;
    if (onTileClick) tile.addEventListener("click", () => onTileClick(i, tile));
    container.appendChild(tile);
    tiles.push(tile);
  }
  if (decorate) decorate(tiles);
  for (let i = 1; i <= pdf.numPages; i++) {
    try {
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({ scale: 0.5 });
      const canvas = document.createElement("canvas");
      const ratio = window.devicePixelRatio || 1;
      canvas.width = Math.floor(viewport.width * ratio);
      canvas.height = Math.floor(viewport.height * ratio);
      canvas.style.width = "100%";
      canvas.style.height = "100%";
      const ctx = canvas.getContext("2d");
      ctx.scale(ratio, ratio);
      await page.render({ canvasContext: ctx, viewport }).promise;
      const thumb = tiles[i - 1].querySelector(".thumb");
      thumb.innerHTML = "";
      thumb.appendChild(canvas);
    } catch (_) {}
  }
  return tiles;
}

// ============================================================
// Converter
// ============================================================
async function openConverter() {
  convName.textContent = currentFile.name;
  convPages.textContent = inBatch()
    ? `PDF ${activeIdx + 1} of ${batchLabel()} · ${pageCount} page${pageCount === 1 ? "" : "s"}`
    : `${pageCount} page${pageCount === 1 ? "" : "s"}`;
  convertBtn.textContent = inBatch() ? `Apply to all ${batchLabel()} → ZIP` : "Make it Excel";
  convSelected.clear();
  const ccfg = inBatch() ? batchCfg[activeIdx] : null;
  if (ccfg && ccfg.pages && ccfg.pages !== "all") {
    parsePagesSpec(ccfg.pages, pageCount).forEach(n => convSelected.add(n));
  } else {
    for (let i = 1; i <= pageCount; i++) convSelected.add(i);
  }
  convFrom.value = 1; convFrom.max = pageCount;
  convTo.value = pageCount; convTo.max = pageCount;
  const forceOcrEl = document.getElementById("forceOcr");
  if (forceOcrEl) forceOcrEl.checked = false;
  showOnly(converterEl);
  renderBatchBar();
  updateConvSummary();

  await renderThumbnails(currentFile, convThumbs,
    (i, tile) => {
      if (convSelected.has(i)) { convSelected.delete(i); tile.classList.remove("selected"); }
      else { convSelected.add(i); tile.classList.add("selected"); }
      tile.style.setProperty("--tile-color", SECTION_COLORS[0]);
      updateConvSummary();
    },
    (allTiles) => {
      allTiles.forEach(t => {
        t.classList.add("selected");
        t.style.setProperty("--tile-color", SECTION_COLORS[0]);
      });
    },
  );
}

function updateConvSummary() {
  const n = convSelected.size;
  if (n === 0) {
    convSummary.textContent = "No pages selected — pick at least one.";
    convSummary.classList.add("warn");
  } else if (n === pageCount) {
    convSummary.textContent = `All ${pageCount} pages selected.`;
    convSummary.classList.remove("warn");
  } else {
    convSummary.textContent = `${n} of ${pageCount} selected · pages ${compressPages(convSelected)}`;
    convSummary.classList.remove("warn");
  }
}

document.querySelectorAll("[data-conv-act]").forEach(btn => {
  btn.addEventListener("click", () => {
    const act = btn.dataset.convAct;
    if (act === "all") {
      convSelected.clear();
      for (let i = 1; i <= pageCount; i++) convSelected.add(i);
    } else if (act === "none") {
      convSelected.clear();
    } else if (act === "invert") {
      for (let i = 1; i <= pageCount; i++) {
        if (convSelected.has(i)) convSelected.delete(i);
        else convSelected.add(i);
      }
    } else if (act === "apply-range") {
      const a = parseInt(convFrom.value, 10);
      const b = parseInt(convTo.value, 10);
      if (!a || !b || a < 1 || b < 1 || a > b || b > pageCount) {
        showError(`Range must be between 1 and ${pageCount}, with From ≤ To.`);
        return;
      }
      convSelected.clear();
      for (let i = a; i <= b; i++) convSelected.add(i);
    }
    syncConvTiles();
    updateConvSummary();
  });
});

function syncConvTiles() {
  Array.from(convThumbs.children).forEach(tile => {
    const i = parseInt(tile.dataset.page, 10);
    if (convSelected.has(i)) {
      tile.classList.add("selected");
      tile.style.setProperty("--tile-color", SECTION_COLORS[0]);
    } else {
      tile.classList.remove("selected");
    }
  });
}

convBack.addEventListener("click", leaveEditor);
convertBtn.addEventListener("click", () => {
  if (convSelected.size === 0) {
    showError("The boss needs at least one page, kid.");
    return;
  }
  const spec = convSelected.size === pageCount ? "all" : compressPages(convSelected);
  const forceOcrEl = document.getElementById("forceOcr");
  const forceOcr = !!(forceOcrEl && forceOcrEl.checked);
  if (inBatch()) runConvertBatch(spec, forceOcr);
  else runConvert(currentFile, spec, forceOcr);
});

// ---- Batch: file 0 is the shared spec; per_file overrides customized PDFs ----
function runConvertBatch(_activeSpec, forceOcr) {
  saveActiveCfg();
  const shared = (batchCfg[0] && batchCfg[0].pages) || "all";
  const per = buildPerFileMap();
  const nCustom = Object.keys(per).length;
  const form = new FormData();
  batchFiles.forEach(f => form.append("files", f, f.name));
  if (shared && shared !== "all") form.append("pages", shared);
  if (nCustom) form.append("per_file", JSON.stringify(per));
  if (forceOcr) form.append("force_ocr", "true");
  const extra = nCustom ? ` · ${nCustom} customized` : (shared !== "all" ? ` · pages ${shared} each` : "");
  runBatchJob({
    url: "/api/convert-batch",
    form,
    zipName: "boss-pdf-batch.zip",
    working: `The boss is converting ${batchLabel()}${extra}…`,
    doneMsg: `${batchLabel()} converted${extra}. Your ZIP of spreadsheets is ready.`,
  });
}

function runConvert(file, pageSpec, forceOcr) {
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = fmtSize(file.size);
  progressBar.style.width = "0%";
  progressBar.classList.remove("indeterminate");
  statusMsg.textContent = forceOcr
    ? "Filing it with the boss (force OCR)…"
    : "Filing it with the boss…";
  showOnly(statusEl);

  const form = new FormData();
  form.append("file", file);
  if (pageSpec && pageSpec !== "all") form.append("pages", pageSpec);
  if (forceOcr) form.append("force_ocr", "true");

  const xhr = new XMLHttpRequest();
  xhr.open("POST", "/api/convert");
  xhr.responseType = "blob";
  BossAuth.applyAuthHeaders(xhr);

  xhr.upload.onprogress = e => {
    if (e.lengthComputable) {
      const pct = Math.round((e.loaded / e.total) * 100);
      progressBar.style.width = (pct * 0.5) + "%";
    }
  };
  xhr.upload.onload = () => {
    statusMsg.textContent = "The boss is crunching the numbers…";
    progressBar.classList.add("indeterminate");
  };
  xhr.onload = () => {
    progressBar.classList.remove("indeterminate");
    progressBar.style.width = "100%";
    if (xhr.status === 200) {
      const xlsxBlob = xhr.response;
      const ocrUsed = (xhr.getResponseHeader("X-Boss-OCR-Used") || "") === "true";
      const ocrEngine = xhr.getResponseHeader("X-Boss-OCR-Engine") || "";
      const logId = xhr.getResponseHeader("X-Boss-Log-Id") || "";
      const usageId = xhr.getResponseHeader("X-Boss-Usage-Id") || "";
      lastConvertUsageId = usageId ? parseInt(usageId, 10) : null;

      const stem = file.name.replace(/\.pdf$/i, "");
      const outName = `${stem}.xlsx`;
      const url = URL.createObjectURL(xlsxBlob);
      downloadLink.href = url;
      downloadLink.download = outName;
      downloadLink.textContent = "Download .xlsx";   // batch mode swaps this to .zip

      // engine badge
      engineBadge.classList.remove("native", "ocr");
      if (ocrUsed) {
        engineBadge.textContent = ocrEngine.replace(/-/g, " ") || "OCR";
        engineBadge.classList.add("ocr");
        resultMsg.textContent = `OCR ran via ${ocrEngine.replace(/-/g, ' ')}. Open the preview to verify text quality.`;
      } else {
        engineBadge.textContent = "native text";
        engineBadge.classList.add("native");
        resultMsg.textContent = "Native text extraction — clean.";
      }
      engineBadgeRow.classList.remove("hidden");

      // log button (always available) — fetched with auth, then opened as blob URL
      if (logId) {
        logDlBtn.dataset.logId = logId;
        logDlBtn.classList.remove("hidden");
      } else {
        logDlBtn.classList.add("hidden");
      }

      // OCR preview button (only when OCR ran)
      if (ocrUsed && logId) {
        previewBtn.dataset.logId = logId;
        previewBtn.classList.remove("hidden");
      } else {
        delete previewBtn.dataset.logId;
        previewBtn.classList.add("hidden");
      }

      showOnly(resultEl);
      downloadLink.click(); // auto-download xlsx
    } else {
      const reader = new FileReader();
      reader.onload = () => {
        let msg = "Conversion failed.";
        try { const j = JSON.parse(reader.result); if (j && j.detail) msg = j.detail; } catch (_) {}
        showError(msg);
      };
      reader.onerror = () => showError("Conversion failed.");
      reader.readAsText(xhr.response);
    }
  };
  xhr.onerror = () => showError("Network error. Is the server running?");
  xhr.send(form);
}

// ============================================================
// Splitter — section rows with from/to range inputs
// ============================================================

async function openSplitter() {
  splitName.textContent = currentFile.name;
  splitPages.textContent = inBatch()
    ? `PDF ${activeIdx + 1} of ${batchLabel()} · ${pageCount} page${pageCount === 1 ? "" : "s"}`
    : `${pageCount} page${pageCount === 1 ? "" : "s"}`;
  splitBtn.textContent = inBatch() ? `Apply to all ${batchLabel()} → ZIP` : "Slice it up";
  // Load this file's saved ranges if any, else default to ONE section 1 → last.
  const scfg = inBatch() ? batchCfg[activeIdx] : null;
  if (scfg && scfg.sections && scfg.sections.length) {
    sections = scfg.sections.map((s, i) => ({ from: s.from, to: s.to, color: SECTION_COLORS[i % SECTION_COLORS.length] }));
    pickedPages.clear();
    (scfg.picked || []).forEach(p => pickedPages.add(p));
    setSplitMode(scfg.mode || "ranges");
  } else {
    sections = [{ from: 1, to: pageCount, color: SECTION_COLORS[0] }];
    pickedPages.clear();
    setSplitMode("ranges");
  }
  showOnly(splitterEl);
  renderBatchBar();
  renderSections();
  updateSplitSummary();
  updatePickQueue();

  await renderThumbnails(currentFile, splitThumbs,
    (i, tile) => {
      // Click-to-add only meaningful in pick mode; ranges mode ignores clicks.
      if (splitMode !== "pick") return;
      if (pickedPages.has(i)) pickedPages.delete(i);
      else pickedPages.add(i);
      decorateSplitTile(tile);
      updatePickQueue();
    },
    (allTiles) => { allTiles.forEach(decorateSplitTile); },
  );
}

function setSplitMode(mode) {
  splitMode = mode;
  splitModeBtns.forEach(b =>
    b.classList.toggle("active", b.dataset.splitMode === mode));
  const pick = mode === "pick";
  rangesBlock.classList.toggle("hidden", pick);
  pickPanel.classList.toggle("hidden", !pick);
  if (splitModeHint) {
    splitModeHint.textContent = pick
      ? "Click pages in the preview to build one PDF — in click order, like 1, 5, 9, 12."
      : "Set the page range for each new PDF. Add more rows for more files.";
    if (inBatch()) {
      // Ranges are clamped per file server-side, so a folder of mixed page
      // counts still produces output instead of erroring.
      splitModeHint.textContent +=
        ` The same ranges apply to all ${batchLabel()} — pages past a shorter file's end are simply skipped.`;
    }
  }
  if (splitPreviewHint) {
    splitPreviewHint.textContent = pick
      ? "click to add/remove · hover to zoom · or type pages above"
      : "colored rings show which PDF each page goes into · hover to zoom";
  }
  syncSplitTiles();
  if (pick) updatePickQueue();
}

splitModeBtns.forEach(b => b.addEventListener("click", () => setSplitMode(b.dataset.splitMode)));
pickClearBtn.addEventListener("click", () => {
  pickedPages.clear();
  updatePickQueue();
  syncSplitTiles();
});

// Parse a free-form page spec like "1, 5, 9-12" → Set<int>. Silently drops
// out-of-range / malformed pieces so typing doesn't error mid-stroke.
function parsePagesSpec(spec, total) {
  const out = new Set();
  if (!spec || !total) return out;
  for (const raw of spec.split(/[,\s]+/)) {
    if (!raw) continue;
    const m = raw.match(/^(\d+)(?:-(\d+))?$/);
    if (!m) continue;
    const a = parseInt(m[1], 10);
    const b = m[2] ? parseInt(m[2], 10) : a;
    const lo = Math.min(a, b), hi = Math.max(a, b);
    for (let i = lo; i <= hi; i++) {
      if (i >= 1 && i <= total) out.add(i);
    }
  }
  return out;
}

function updatePickQueue({ keepInputValue = false } = {}) {
  const arr = [...pickedPages].sort((a, b) => a - b);
  const n = arr.length;
  pickCountEl.classList.toggle("warn", n === 0);
  pickCountEl.textContent = n ? `${n} page${n === 1 ? "" : "s"}` : "0 pages";
  // Don't overwrite the input while the user is mid-type; otherwise reflect
  // the canonical compressed form (so "1,2,3,4" collapses to "1-4").
  if (!keepInputValue && document.activeElement !== pickInputEl) {
    pickInputEl.value = compressPages(pickedPages);
  }
}

pickInputEl.addEventListener("input", () => {
  const next = parsePagesSpec(pickInputEl.value, pageCount);
  pickedPages.clear();
  next.forEach(n => pickedPages.add(n));
  syncSplitTiles();
  updatePickQueue({ keepInputValue: true });
});
pickInputEl.addEventListener("blur", () => {
  // Normalize on blur — "1, 2, 3, 5" becomes "1-3,5".
  pickInputEl.value = compressPages(pickedPages);
});
pickInputEl.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    pickInputEl.blur();   // triggers normalize
    pickInputEl.focus();
  }
});

function renderSections() {
  // refresh colors so they always cycle in row order
  sections.forEach((s, i) => { s.color = SECTION_COLORS[i % SECTION_COLORS.length]; });

  sectionListEl.innerHTML = "";
  sections.forEach((s, idx) => {
    const row = document.createElement("div");
    row.className = "section-row";
    row.style.setProperty("--sec-color", s.color);
    const pages = Math.max(0, s.to - s.from + 1);
    const valid = s.from >= 1 && s.to >= s.from && s.to <= pageCount;
    row.innerHTML = `
      <div class="section-marker">
        <span class="section-dot"></span>
        <span class="section-name">PDF #${idx + 1}</span>
      </div>
      <div class="section-range">
        <span class="field-label">From page</span>
        <input type="number" class="from" min="1" max="${pageCount}" value="${s.from}">
        <span class="range-arrow">→</span>
        <span class="field-label">to</span>
        <input type="number" class="to" min="1" max="${pageCount}" value="${s.to}">
      </div>
      <span class="section-count ${valid ? '' : 'bad'}">
        ${valid ? `${pages} page${pages === 1 ? "" : "s"}` : "invalid"}
      </span>
      <button class="section-remove" type="button" title="Remove" ${sections.length === 1 ? "disabled" : ""}>×</button>
    `;
    const fromInput = row.querySelector(".from");
    const toInput = row.querySelector(".to");
    fromInput.addEventListener("input", () => {
      const v = parseInt(fromInput.value, 10);
      if (!isNaN(v)) sections[idx].from = v;
      onSectionsChanged();
    });
    toInput.addEventListener("input", () => {
      const v = parseInt(toInput.value, 10);
      if (!isNaN(v)) sections[idx].to = v;
      onSectionsChanged();
    });
    row.querySelector(".section-remove").addEventListener("click", () => {
      if (sections.length === 1) return;
      sections.splice(idx, 1);
      renderSections();
      onSectionsChanged();
    });
    sectionListEl.appendChild(row);
  });
}

function onSectionsChanged() {
  // Update count badge text + thumbs, but DON'T re-render section list (that would
  // steal focus from the input the user is typing in).
  Array.from(sectionListEl.children).forEach((row, idx) => {
    const s = sections[idx];
    if (!s) return;
    const pages = Math.max(0, s.to - s.from + 1);
    const valid = s.from >= 1 && s.to >= s.from && s.to <= pageCount;
    const badge = row.querySelector(".section-count");
    if (badge) {
      badge.textContent = valid ? `${pages} page${pages === 1 ? "" : "s"}` : "invalid";
      badge.classList.toggle("bad", !valid);
    }
  });
  syncSplitTiles();
  updateSplitSummary();
}

function decorateSplitTile(tile) {
  const i = parseInt(tile.dataset.page, 10);
  let tag = tile.querySelector(".section-tag");

  if (splitMode === "pick") {
    // Single-PDF pick mode: gold ring on selected tiles, no section tag.
    if (tag) tag.remove();
    if (pickedPages.has(i)) {
      tile.classList.add("selected");
      tile.style.setProperty("--tile-color", SECTION_COLORS[0]);
    } else {
      tile.classList.remove("selected");
      tile.style.removeProperty("--tile-color");
    }
    return;
  }

  // Ranges mode (original behavior): first section that contains the page
  // wins for color/tag.
  let owner = -1;
  for (let s = 0; s < sections.length; s++) {
    const sec = sections[s];
    if (sec.from <= i && i <= sec.to) { owner = s; break; }
  }
  if (owner === -1) {
    tile.classList.remove("selected");
    tile.style.removeProperty("--tile-color");
    if (tag) tag.remove();
  } else {
    const color = sections[owner].color;
    tile.classList.add("selected");
    tile.style.setProperty("--tile-color", color);
    if (!tag) {
      tag = document.createElement("div");
      tag.className = "section-tag";
      tile.appendChild(tag);
    }
    tag.textContent = `#${owner + 1}`;
    tag.style.background = color;
  }
}
function syncSplitTiles() {
  Array.from(splitThumbs.children).forEach(tile => {
    if (tile.classList.contains("page-tile")) decorateSplitTile(tile);
  });
}

function updateSplitSummary() {
  const valid = sections.filter(s =>
    s.from >= 1 && s.to >= s.from && s.to <= pageCount
  );
  const totalPages = valid.reduce((sum, s) => sum + (s.to - s.from + 1), 0);
  const word = valid.length === 1 ? "PDF" : "PDFs";
  splitSummary.textContent = `${valid.length} ${word} · ${totalPages} of ${pageCount} pages covered`;
}

addSectionBtn.addEventListener("click", () => {
  // Heuristic for the next default range: pick up where the last one ends
  const last = sections[sections.length - 1];
  let from = 1, to = pageCount;
  if (last && last.to < pageCount) {
    from = last.to + 1;
    to = pageCount;
  }
  sections.push({ from, to, color: SECTION_COLORS[sections.length % SECTION_COLORS.length] });
  renderSections();
  onSectionsChanged();
});

autoSplitBtn.addEventListener("click", () => {
  sections = [];
  for (let i = 1; i <= pageCount; i++) {
    sections.push({ from: i, to: i, color: SECTION_COLORS[(i - 1) % SECTION_COLORS.length] });
  }
  renderSections();
  onSectionsChanged();
});

splitBack.addEventListener("click", leaveEditor);

splitBtn.addEventListener("click", () => {
  if (splitMode === "pick") {
    if (!pickedPages.size) {
      showError("Pick at least one page from the preview, kid.");
      return;
    }
    const arr = [...pickedPages].sort((a, b) => a - b);
    // One output PDF with all picked pages — e.g. "1,5,9,12".
    if (inBatch()) runSplitBatch([arr.join(",")]);
    else runSplit(currentFile, [arr.join(",")]);
    return;
  }

  // Ranges mode (unchanged)
  for (let i = 0; i < sections.length; i++) {
    const s = sections[i];
    if (!(s.from >= 1 && s.to >= s.from && s.to <= pageCount)) {
      showError(`PDF #${i + 1} has an invalid range. Use values between 1 and ${pageCount}.`);
      return;
    }
  }
  const specs = sections.map(s => s.from === s.to ? `${s.from}` : `${s.from}-${s.to}`);
  if (inBatch()) runSplitBatch(specs);
  else runSplit(currentFile, specs);
});

// ---- Batch: file 0 is the shared ranges; per_file overrides customized PDFs ----
function runSplitBatch(specs) {
  saveActiveCfg();
  const shared = (batchCfg[0] ? specsFromCfg(batchCfg[0]) : specs);
  const sharedSpecs = shared.length ? shared : specs;
  const per = buildPerFileMap();
  const nCustom = Object.keys(per).length;
  const form = new FormData();
  batchFiles.forEach(f => form.append("files", f, f.name));
  form.append("ranges", JSON.stringify(sharedSpecs));
  if (nCustom) form.append("per_file", JSON.stringify(per));
  const partWord = sharedSpecs.length === 1 ? "slice" : `${sharedSpecs.length} slices`;
  const extra = nCustom ? ` · ${nCustom} customized` : "";
  runBatchJob({
    url: "/api/split-batch",
    form,
    zipName: "boss-pdf-split-batch.zip",
    working: `The boss is slicing ${batchLabel()}…`,
    doneMsg: `${batchLabel()} sliced — ${partWord} each${extra}. Your ZIP is ready.`,
  });
}

// Shared uploader for both batch jobs: progress bar, ZIP download, result card.
// Ranges/pages that overrun a shorter PDF are clamped server-side, so a mixed
// folder still produces output instead of failing.
function runBatchJob({ url, form, zipName, working, doneMsg }) {
  fileNameEl.textContent = batchLabel();
  fileSizeEl.textContent = fmtSize(batchFiles.reduce((s, f) => s + f.size, 0));
  progressBar.style.width = "0%";
  progressBar.classList.remove("indeterminate");
  statusMsg.textContent = "Uploading the folder…";
  showOnly(statusEl);

  const xhr = new XMLHttpRequest();
  xhr.open("POST", url);
  xhr.responseType = "blob";
  BossAuth.applyAuthHeaders(xhr);

  xhr.upload.onprogress = e => {
    if (e.lengthComputable) {
      progressBar.style.width = Math.round((e.loaded / e.total) * 40) + "%";
    }
  };
  xhr.upload.onload = () => {
    statusMsg.textContent = working + " This can take a while — the boss is thorough.";
    progressBar.classList.add("indeterminate");
  };
  xhr.onload = () => {
    progressBar.classList.remove("indeterminate");
    progressBar.style.width = "100%";
    if (xhr.status === 200) {
      const urlObj = URL.createObjectURL(xhr.response);
      downloadLink.href = urlObj;
      downloadLink.download = zipName;
      downloadLink.textContent = "Download .zip";
      resultMsg.textContent = doneMsg;
      engineBadgeRow.classList.add("hidden");
      logDlBtn.classList.add("hidden");
      previewBtn.classList.add("hidden");
      showOnly(resultEl);
      downloadLink.click();
    } else {
      const reader = new FileReader();
      reader.onload = () => {
        let msg = "Batch job failed.";
        try { const j = JSON.parse(reader.result); if (j && j.detail) msg = j.detail; } catch (_) {}
        showError(msg);
      };
      reader.onerror = () => showError("Batch job failed.");
      reader.readAsText(xhr.response);
    }
  };
  xhr.onerror = () => showError("Network error. Is the server running?");
  xhr.send(form);
}

function runSplit(file, specs) {
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = fmtSize(file.size);
  progressBar.style.width = "0%";
  progressBar.classList.add("indeterminate");
  statusMsg.textContent = "The boss is slicing it up…";
  showOnly(statusEl);

  Promise.all(specs.map(spec => splitOne(file, spec)))
    .then(results => {
      progressBar.classList.remove("indeterminate");
      progressBar.style.width = "100%";
      renderSplitResults(file, results);
    })
    .catch(err => showError(err.message || String(err)));
}

function splitOne(file, spec) {
  return new Promise((resolve, reject) => {
    const form = new FormData();
    form.append("file", file);
    form.append("ranges", JSON.stringify([spec]));
    const xhr = new XMLHttpRequest();
    xhr.open("POST", "/api/split");
    xhr.responseType = "blob";
    BossAuth.applyAuthHeaders(xhr);
    xhr.onload = () => {
      if (xhr.status === 200) {
        const usageId = xhr.getResponseHeader("X-Boss-Usage-Id") || "";
        const stem = file.name.replace(/\.pdf$/i, "");
        const safeSpec = spec.replace(/,/g, "_");
        const name = `${stem}_p${safeSpec}.pdf`;
        resolve({ spec, blob: xhr.response, name, usageId: usageId ? parseInt(usageId, 10) : null });
      } else {
        const reader = new FileReader();
        reader.onload = () => {
          let msg = "Split failed.";
          try { const j = JSON.parse(reader.result); if (j && j.detail) msg = j.detail; } catch (_) {}
          reject(new Error(msg));
        };
        reader.onerror = () => reject(new Error("Split failed."));
        reader.readAsText(xhr.response);
      }
    };
    xhr.onerror = () => reject(new Error("Network error."));
    xhr.send(form);
  });
}

function renderSplitResults(originalFile, results) {
  splitListEl.innerHTML = "";
  results.forEach((r, i) => {
    const url = URL.createObjectURL(r.blob);
    const color = SECTION_COLORS[i % SECTION_COLORS.length];
    const item = document.createElement("div");
    item.className = "split-item";
    item.style.setProperty("--tile-color", color);
    item.innerHTML = `
      <div class="label">PDF #${i + 1}</div>
      <div class="pages">pages ${r.spec}</div>
      <div class="actions">
        <a class="dl" href="${url}" download="${r.name}">⬇ PDF</a>
        <button class="toxlsx" type="button">📊 Make it Excel</button>
      </div>
    `;
    item.querySelector(".dl").addEventListener("click", () => {
      if (r.usageId) {
        BossAuth.authFetch(`/api/usage/${r.usageId}/download?kind=split`,
                           { method: "POST" }).catch(() => {});
      }
    });
    item.querySelector(".toxlsx").addEventListener("click", () => {
      const f = new File([r.blob], r.name, { type: "application/pdf" });
      currentFile = f;
      pdfDoc = null; pdfDocFile = null;
      reuploadAndConvert(f);
    });
    splitListEl.appendChild(item);
  });
  showOnly(splitResultEl);
}

async function reuploadAndConvert(file) {
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = fmtSize(file.size);
  statusMsg.textContent = "Promoting this slice to the Excel department…";
  progressBar.style.width = "0%";
  progressBar.classList.add("indeterminate");
  showOnly(statusEl);
  try {
    const form = new FormData();
    form.append("file", file);
    const res = await BossAuth.authFetch("/api/pdf-info", { method: "POST", body: form });
    if (!res.ok) {
      const j = await res.json().catch(() => ({}));
      throw new Error(j.detail || "Could not read PDF.");
    }
    const info = await res.json();
    pageCount = info.pages || 1;
    openConverter();
  } catch (err) {
    showError(err.message || String(err));
  }
}

// ---------- Auth-aware log + preview buttons ----------
async function fetchBlobAuthed(url) {
  const res = await BossAuth.authFetch(url);
  if (!res.ok) {
    const j = await res.json().catch(() => ({}));
    throw new Error(j.detail || `Request failed (${res.status})`);
  }
  return await res.blob();
}

logDlBtn.addEventListener("click", async (e) => {
  e.preventDefault();
  const id = logDlBtn.dataset.logId;
  if (!id) return;
  try {
    const blob = await fetchBlobAuthed(`/api/log/${id}`);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "boss-pdf-log.txt";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  } catch (err) { showError(err.message || String(err)); }
});

previewBtn.addEventListener("click", async (e) => {
  e.preventDefault();
  const id = previewBtn.dataset.logId;
  if (!id) return;
  try {
    const blob = await fetchBlobAuthed(`/api/preview/${id}`);
    const url = URL.createObjectURL(blob);
    window.open(url, "_blank", "noopener");
    setTimeout(() => URL.revokeObjectURL(url), 60_000);
  } catch (err) { showError(err.message || String(err)); }
});

// ============================================================
// Hover-zoom popup for thumbnails (works in converter + splitter)
// ============================================================
const zoomPopup = document.getElementById("zoomPopup");
const zoomCanvas = document.getElementById("zoomCanvas");
const zoomLabel = document.getElementById("zoomLabel");
const ZOOM_DELAY_MS = 220;       // hover delay before showing
const ZOOM_MAX_W_RATIO = 0.55;   // popup max width = 55% of viewport
const ZOOM_MAX_H_RATIO = 0.85;
const ZOOM_CAP_SCALE   = 2.5;

let zoomTile = null;
let zoomReq  = 0;
let zoomTimer = null;

async function showZoom(tile) {
  if (!pdfDoc) return;
  const pageNum = parseInt(tile.dataset.page, 10);
  if (!pageNum) return;
  const reqId = ++zoomReq;
  try {
    const page = await pdfDoc.getPage(pageNum);
    if (reqId !== zoomReq) return;
    const baseVp = page.getViewport({ scale: 1 });
    const maxW = window.innerWidth  * ZOOM_MAX_W_RATIO;
    const maxH = window.innerHeight * ZOOM_MAX_H_RATIO;
    const scale = Math.min(maxW / baseVp.width, maxH / baseVp.height, ZOOM_CAP_SCALE);
    const vp = page.getViewport({ scale });
    const ratio = window.devicePixelRatio || 1;
    zoomCanvas.width  = Math.floor(vp.width  * ratio);
    zoomCanvas.height = Math.floor(vp.height * ratio);
    zoomCanvas.style.width  = vp.width  + "px";
    zoomCanvas.style.height = vp.height + "px";
    const ctx = zoomCanvas.getContext("2d");
    ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    if (reqId !== zoomReq) return;
    zoomLabel.textContent = `Page ${pageNum}`;

    // Position to the right of the tile; flip to the left if it would overflow.
    const rect = tile.getBoundingClientRect();
    const pw = vp.width + 16;   // include the 6px padding * 2 + label space
    const ph = vp.height + 32;
    let left = rect.right + 14;
    if (left + pw > window.innerWidth - 8) left = rect.left - pw - 14;
    if (left < 8) left = 8;
    let top = rect.top + (rect.height / 2) - (ph / 2);
    if (top + ph > window.innerHeight - 8) top = window.innerHeight - ph - 8;
    if (top < 8) top = 8;
    zoomPopup.style.left = left + "px";
    zoomPopup.style.top  = top  + "px";
    zoomPopup.classList.remove("hidden");
  } catch (_) { /* swallow: tile may have been removed mid-render */ }
}

function hideZoom() {
  zoomReq++;              // cancel any in-flight render
  zoomPopup.classList.add("hidden");
  zoomTile = null;
}

document.addEventListener("mouseover", (e) => {
  const tile = e.target.closest && e.target.closest(".page-tile");
  if (!tile || tile === zoomTile) return;
  zoomTile = tile;
  clearTimeout(zoomTimer);
  zoomTimer = setTimeout(() => showZoom(tile), ZOOM_DELAY_MS);
});
document.addEventListener("mouseout", (e) => {
  const tile = e.target.closest && e.target.closest(".page-tile");
  if (!tile || tile !== zoomTile) return;
  // Don't hide if the mouse just moved between the tile's own children.
  const next = e.relatedTarget;
  if (next && tile.contains(next)) return;
  clearTimeout(zoomTimer);
  hideZoom();
});
// Hide on scroll (thumbgrid scrolling moves the tile out from under the popup)
document.addEventListener("scroll", hideZoom, true);
window.addEventListener("blur", hideZoom);

// ---------- Reset ----------
againBtn.addEventListener("click", resetAll);
splitAgainBtn.addEventListener("click", resetAll);
retryBtn.addEventListener("click", resetAll);
