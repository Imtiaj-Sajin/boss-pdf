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

const convSelected = new Set();
let sections = []; // [{ from, to, color }]

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
  currentFile = null;
  pageCount = 0;
  pdfDoc = null;
  pdfDocFile = null;
  convSelected.clear();
  sections = [];
  convThumbs.innerHTML = "";
  splitThumbs.innerHTML = "";
  sectionListEl.innerHTML = "";
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
  if (file.size > 50 * 1024 * 1024) {
    showError("File is over 50 MB. The boss has standards.");
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
  convPages.textContent = `${pageCount} page${pageCount === 1 ? "" : "s"}`;
  convSelected.clear();
  for (let i = 1; i <= pageCount; i++) convSelected.add(i);
  convFrom.value = 1; convFrom.max = pageCount;
  convTo.value = pageCount; convTo.max = pageCount;
  const forceOcrEl = document.getElementById("forceOcr");
  if (forceOcrEl) forceOcrEl.checked = false;
  showOnly(converterEl);
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

convBack.addEventListener("click", showChooser);
convertBtn.addEventListener("click", () => {
  if (convSelected.size === 0) {
    showError("The boss needs at least one page, kid.");
    return;
  }
  const spec = convSelected.size === pageCount ? "all" : compressPages(convSelected);
  const forceOcrEl = document.getElementById("forceOcr");
  const forceOcr = !!(forceOcrEl && forceOcrEl.checked);
  runConvert(currentFile, spec, forceOcr);
});

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
  splitPages.textContent = `${pageCount} page${pageCount === 1 ? "" : "s"}`;
  // default: ONE section, page 1 → last page
  sections = [{ from: 1, to: pageCount, color: SECTION_COLORS[0] }];
  showOnly(splitterEl);
  renderSections();
  updateSplitSummary();

  await renderThumbnails(currentFile, splitThumbs, null, (allTiles) => {
    allTiles.forEach(decorateSplitTile);
  });
}

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
  // first section that contains the page wins for color/tag
  let owner = -1;
  for (let s = 0; s < sections.length; s++) {
    const sec = sections[s];
    if (sec.from <= i && i <= sec.to) { owner = s; break; }
  }
  let tag = tile.querySelector(".section-tag");
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

splitBack.addEventListener("click", showChooser);

splitBtn.addEventListener("click", () => {
  // validate
  for (let i = 0; i < sections.length; i++) {
    const s = sections[i];
    if (!(s.from >= 1 && s.to >= s.from && s.to <= pageCount)) {
      showError(`PDF #${i + 1} has an invalid range. Use values between 1 and ${pageCount}.`);
      return;
    }
  }
  const specs = sections.map(s => s.from === s.to ? `${s.from}` : `${s.from}-${s.to}`);
  runSplit(currentFile, specs);
});

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

// ---------- Reset ----------
againBtn.addEventListener("click", resetAll);
splitAgainBtn.addEventListener("click", resetAll);
retryBtn.addEventListener("click", resetAll);
