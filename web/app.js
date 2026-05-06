// ============================================================
// boss-pdf · client
// ============================================================

// pdf.js worker
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

// chunk colors (cycled)
const CHUNK_COLORS = [
  "#d4a64a", // gold
  "#5fdfb0", // teal
  "#6aa7ff", // blue
  "#ff89d2", // pink
  "#f48a4d", // orange
  "#b388ff", // violet
  "#9be36b", // lime
  "#ff6f6f", // red
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
const convQuick = document.getElementById("convQuick");
const convThumbs = document.getElementById("convThumbs");
const convertBtn = document.getElementById("convertBtn");
const convBack = document.getElementById("convBack");

const splitterEl = document.getElementById("splitter");
const splitName = document.getElementById("splitName");
const splitPages = document.getElementById("splitPages");
const chunkBar = document.getElementById("chunkBar");
const addChunkBtn = document.getElementById("addChunkBtn");
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

const splitResultEl = document.getElementById("splitResult");
const splitListEl = document.getElementById("splitList");
const splitAgainBtn = document.getElementById("splitAgainBtn");

const errorEl = document.getElementById("error");
const errorMsg = document.getElementById("errorMsg");
const retryBtn = document.getElementById("retryBtn");

// ---------- State ----------
let currentFile = null;
let pageCount = 0;
let pdfDoc = null;          // cached pdfjs doc for thumbnails
let pdfDocFile = null;      // file the doc was built from

const convSelected = new Set();
let chunks = [];            // [{color, pages: Set<int>}]
let activeChunkIdx = 0;

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
  chunks = [];
  activeChunkIdx = 0;
  convThumbs.innerHTML = "";
  splitThumbs.innerHTML = "";
}

function showError(message) {
  errorMsg.textContent = message || "Something went sideways.";
  showOnly(errorEl);
}

// Compress {1,2,3,5,8,9} -> "1-3,5,8-9"
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

// Parse "1-3,5,8-10" -> Set<int>, clamped to [1, pageCount]
function parsePageSpec(spec) {
  const out = new Set();
  if (!spec) return out;
  const parts = spec.split(",");
  for (const part of parts) {
    const p = part.trim();
    if (!p) continue;
    if (p.includes("-")) {
      const [a, b] = p.split("-").map(s => parseInt(s.trim(), 10));
      if (!a || !b || a < 1 || b < a || b > pageCount) return null;
      for (let i = a; i <= b; i++) out.add(i);
    } else {
      const n = parseInt(p, 10);
      if (!n || n < 1 || n > pageCount) return null;
      out.add(n);
    }
  }
  return out;
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
    showError("That doesn't look like a .pdf file. The Boss only does PDFs.");
    return;
  }
  if (file.size > 50 * 1024 * 1024) {
    showError("File is over 50 MB. The Boss has standards.");
    return;
  }
  currentFile = file;

  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = fmtSize(file.size);
  statusMsg.textContent = "The Boss is sizing up your file…";
  progressBar.style.width = "0%";
  progressBar.classList.add("indeterminate");
  showOnly(statusEl);

  try {
    const form = new FormData();
    form.append("file", file);
    const res = await fetch("/api/pdf-info", { method: "POST", body: form });
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
  chooserPages.textContent = `${pageCount} page${pageCount === 1 ? "" : "s"} • ${fmtSize(currentFile.size)}`;
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
    tile.addEventListener("click", () => onTileClick(i, tile));
    container.appendChild(tile);
    tiles.push(tile);
  }
  // decorate immediately if a state already exists (e.g. select-all on open)
  if (decorate) decorate(tiles);
  // render canvases progressively
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
    } catch (e) {
      // ignore individual page render errors
    }
  }
  return tiles;
}

// ---------- Step 3a: converter ----------
async function openConverter() {
  convName.textContent = currentFile.name;
  convPages.textContent = `${pageCount} page${pageCount === 1 ? "" : "s"}`;
  convSelected.clear();
  // default: select all
  for (let i = 1; i <= pageCount; i++) convSelected.add(i);
  convQuick.value = "";
  showOnly(converterEl);
  updateConvSummary();

  const tiles = await renderThumbnails(currentFile, convThumbs,
    (i, tile) => {
      if (convSelected.has(i)) { convSelected.delete(i); tile.classList.remove("selected"); }
      else { convSelected.add(i); tile.classList.add("selected"); }
      updateConvSummary();
    },
    (allTiles) => {
      // initial decoration: mark all as selected (gold)
      allTiles.forEach(t => {
        t.classList.add("selected");
        t.style.setProperty("--tile-color", CHUNK_COLORS[0]);
      });
    },
  );
  // ensure tile color is set even if decorate ran before tiles were appended
  tiles.forEach(t => t.style.setProperty("--tile-color", CHUNK_COLORS[0]));
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
    } else if (act === "apply-quick") {
      const set = parsePageSpec(convQuick.value.trim());
      if (!set) {
        showError(`Invalid page list. Try something like 1-3, 5, 8-10. Pages must be 1–${pageCount}.`);
        return;
      }
      convSelected.clear();
      set.forEach(p => convSelected.add(p));
    }
    syncConvTiles();
    updateConvSummary();
  });
});
convQuick.addEventListener("keydown", e => {
  if (e.key === "Enter") {
    e.preventDefault();
    document.querySelector('[data-conv-act="apply-quick"]').click();
  }
});

function syncConvTiles() {
  Array.from(convThumbs.children).forEach(tile => {
    const i = parseInt(tile.dataset.page, 10);
    if (convSelected.has(i)) tile.classList.add("selected");
    else tile.classList.remove("selected");
  });
}

convBack.addEventListener("click", showChooser);
convertBtn.addEventListener("click", () => {
  if (convSelected.size === 0) {
    showError("The Boss needs at least one page, kid.");
    return;
  }
  const spec = convSelected.size === pageCount ? "all" : compressPages(convSelected);
  runConvert(currentFile, spec);
});

function runConvert(file, pageSpec) {
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = fmtSize(file.size);
  progressBar.style.width = "0%";
  progressBar.classList.remove("indeterminate");
  statusMsg.textContent = "Uploading to The Boss…";
  showOnly(statusEl);

  const form = new FormData();
  form.append("file", file);
  if (pageSpec && pageSpec !== "all") form.append("pages", pageSpec);

  const xhr = new XMLHttpRequest();
  xhr.open("POST", "/api/convert");
  xhr.responseType = "blob";

  xhr.upload.onprogress = e => {
    if (e.lengthComputable) {
      const pct = Math.round((e.loaded / e.total) * 100);
      progressBar.style.width = (pct * 0.5) + "%";
    }
  };
  xhr.upload.onload = () => {
    statusMsg.textContent = "Crunching tables. The Boss is on it…";
    progressBar.classList.add("indeterminate");
  };
  xhr.onload = () => {
    progressBar.classList.remove("indeterminate");
    progressBar.style.width = "100%";
    if (xhr.status === 200) {
      const blob = xhr.response;
      const url = URL.createObjectURL(blob);
      const outName = file.name.replace(/\.pdf$/i, "") + ".xlsx";
      downloadLink.href = url;
      downloadLink.download = outName;
      resultMsg.textContent = pageSpec === "all"
        ? "Whole document converted. Boss is pleased."
        : `Converted pages: ${pageSpec}. Boss is pleased.`;
      showOnly(resultEl);
      downloadLink.click();
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

// ---------- Step 3b: splitter ----------
async function openSplitter() {
  splitName.textContent = currentFile.name;
  splitPages.textContent = `${pageCount} page${pageCount === 1 ? "" : "s"}`;
  // sensible default: one chunk covering the first half (or all if 1 page)
  chunks = [];
  if (pageCount === 1) {
    chunks.push(makeChunk(0, [1]));
  } else {
    const mid = Math.ceil(pageCount / 2);
    chunks.push(makeChunk(0, range(1, mid)));
    chunks.push(makeChunk(1, range(mid + 1, pageCount)));
  }
  activeChunkIdx = chunks.length - 1; // last one usually the empty/edited one
  showOnly(splitterEl);
  renderChunkBar();
  updateSplitSummary();

  await renderThumbnails(currentFile, splitThumbs,
    (i, tile) => onSplitTileClick(i, tile),
    (allTiles) => allTiles.forEach(decorateSplitTile),
  );
}

function makeChunk(idx, pages) {
  return {
    color: CHUNK_COLORS[idx % CHUNK_COLORS.length],
    pages: new Set(pages),
  };
}
function range(a, b) {
  const out = [];
  for (let i = a; i <= b; i++) out.push(i);
  return out;
}

function onSplitTileClick(i, tile) {
  if (!chunks.length) return;
  // remove from any chunk that owns it
  const owner = chunks.findIndex(c => c.pages.has(i));
  if (owner === activeChunkIdx) {
    // clicking page already in active chunk: remove
    chunks[owner].pages.delete(i);
  } else {
    if (owner !== -1) chunks[owner].pages.delete(i);
    chunks[activeChunkIdx].pages.add(i);
  }
  decorateSplitTile(tile);
  // also need to redecorate any sibling that lost ownership
  syncSplitTiles();
  renderChunkBar();
  updateSplitSummary();
}

function decorateSplitTile(tile) {
  const i = parseInt(tile.dataset.page, 10);
  const owner = chunks.findIndex(c => c.pages.has(i));
  if (owner === -1) {
    tile.classList.remove("selected");
    tile.style.removeProperty("--tile-color");
    const tag = tile.querySelector(".chunk-tag");
    if (tag) tag.remove();
  } else {
    tile.classList.add("selected");
    tile.style.setProperty("--tile-color", chunks[owner].color);
    let tag = tile.querySelector(".chunk-tag");
    if (!tag) {
      tag = document.createElement("div");
      tag.className = "chunk-tag";
      tile.appendChild(tag);
    }
    tag.textContent = `#${owner + 1}`;
    tag.style.background = chunks[owner].color;
  }
}
function syncSplitTiles() {
  Array.from(splitThumbs.children).forEach(decorateSplitTile);
}

function renderChunkBar() {
  chunkBar.innerHTML = "";
  chunks.forEach((c, idx) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chunk-chip" + (idx === activeChunkIdx ? " active" : "");
    chip.style.setProperty("--chip-color", c.color);
    const pagesText = c.pages.size === 0 ? "empty" : compressPages(c.pages);
    chip.innerHTML = `
      <span class="swatch"></span>
      <span class="chip-name">Chunk ${idx + 1}</span>
      <span class="count">${pagesText}</span>
      <span class="x" title="Remove chunk">✕</span>
    `;
    chip.addEventListener("click", e => {
      if (e.target.classList.contains("x")) {
        e.stopPropagation();
        if (chunks.length === 1) return; // keep at least one
        chunks.splice(idx, 1);
        // refresh colors so they stay in sequence
        chunks.forEach((cc, j) => { cc.color = CHUNK_COLORS[j % CHUNK_COLORS.length]; });
        if (activeChunkIdx >= chunks.length) activeChunkIdx = chunks.length - 1;
        renderChunkBar();
        syncSplitTiles();
        updateSplitSummary();
        return;
      }
      activeChunkIdx = idx;
      renderChunkBar();
    });
    chunkBar.appendChild(chip);
  });
}

function updateSplitSummary() {
  const totalAssigned = chunks.reduce((s, c) => s + c.pages.size, 0);
  const nonEmpty = chunks.filter(c => c.pages.size > 0).length;
  splitSummary.textContent = `${nonEmpty} chunk${nonEmpty === 1 ? "" : "s"} · ${totalAssigned} of ${pageCount} pages assigned`;
}

addChunkBtn.addEventListener("click", () => {
  chunks.push(makeChunk(chunks.length, []));
  activeChunkIdx = chunks.length - 1;
  renderChunkBar();
  updateSplitSummary();
});

autoSplitBtn.addEventListener("click", () => {
  chunks = [];
  for (let i = 1; i <= pageCount; i++) {
    chunks.push(makeChunk(i - 1, [i]));
  }
  activeChunkIdx = 0;
  renderChunkBar();
  syncSplitTiles();
  updateSplitSummary();
});

splitBack.addEventListener("click", showChooser);

splitBtn.addEventListener("click", () => {
  const real = chunks.filter(c => c.pages.size > 0);
  if (!real.length) {
    showError("No pages assigned to any chunk. Drop a few in.");
    return;
  }
  const specs = real.map(c => compressPages(c.pages));
  runSplit(currentFile, specs);
});

function runSplit(file, specs) {
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = fmtSize(file.size);
  progressBar.style.width = "0%";
  progressBar.classList.add("indeterminate");
  statusMsg.textContent = "Carving up the document…";
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
    xhr.onload = () => {
      if (xhr.status === 200) {
        const stem = file.name.replace(/\.pdf$/i, "");
        const safeSpec = spec.replace(/,/g, "_");
        const name = `${stem}_p${safeSpec}.pdf`;
        resolve({ spec, blob: xhr.response, name });
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
    const color = CHUNK_COLORS[i % CHUNK_COLORS.length];
    const item = document.createElement("div");
    item.className = "split-item";
    item.style.setProperty("--tile-color", color);
    item.innerHTML = `
      <span class="chunk-dot" style="background:${color}"></span>
      <div class="label">Chunk #${i + 1}</div>
      <div class="pages">pages ${r.spec}</div>
      <div class="actions">
        <a class="dl" href="${url}" download="${r.name}">⬇ PDF</a>
        <button class="toxlsx" type="button">📊 Excel-ify</button>
      </div>
    `;
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
  statusMsg.textContent = "Promoting chunk to the Excel department…";
  progressBar.style.width = "0%";
  progressBar.classList.add("indeterminate");
  showOnly(statusEl);
  try {
    const form = new FormData();
    form.append("file", file);
    const res = await fetch("/api/pdf-info", { method: "POST", body: form });
    if (!res.ok) {
      const j = await res.json().catch(() => ({}));
      throw new Error(j.detail || "Could not read split PDF.");
    }
    const info = await res.json();
    pageCount = info.pages || 1;
    openConverter();
  } catch (err) {
    showError(err.message || String(err));
  }
}

// ---------- Reset ----------
againBtn.addEventListener("click", resetAll);
splitAgainBtn.addEventListener("click", resetAll);
retryBtn.addEventListener("click", resetAll);
