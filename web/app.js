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
const convRangeBox = document.getElementById("convRangeBox");
const convCustomBox = document.getElementById("convCustomBox");
const convFrom = document.getElementById("convFrom");
const convTo = document.getElementById("convTo");
const convCustom = document.getElementById("convCustom");
const convertBtn = document.getElementById("convertBtn");
const convBack = document.getElementById("convBack");

const splitterEl = document.getElementById("splitter");
const splitName = document.getElementById("splitName");
const splitPages = document.getElementById("splitPages");
const splitRangesEl = document.getElementById("splitRanges");
const addRangeBtn = document.getElementById("addRangeBtn");
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
let currentFile = null;     // File object the user picked
let pageCount = 0;
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
}

function showError(message) {
  errorMsg.textContent = message || "Something went sideways.";
  showOnly(errorEl);
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

// ---------- Step 1: receive file, ask backend for page count ----------
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

  // Loading state
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

// ---------- Step 3a: converter ----------
function openConverter() {
  convName.textContent = currentFile.name;
  convPages.textContent = `${pageCount} page${pageCount === 1 ? "" : "s"}`;
  convFrom.max = pageCount; convFrom.value = 1;
  convTo.max = pageCount; convTo.value = pageCount;
  convCustom.value = "";
  // default mode = all
  document.querySelector('input[name="convMode"][value="all"]').checked = true;
  hide(convRangeBox); hide(convCustomBox);
  showOnly(converterEl);
}

document.querySelectorAll('input[name="convMode"]').forEach(r => {
  r.addEventListener("change", () => {
    const v = document.querySelector('input[name="convMode"]:checked').value;
    convRangeBox.classList.toggle("hidden", v !== "range");
    convCustomBox.classList.toggle("hidden", v !== "custom");
  });
});

convBack.addEventListener("click", showChooser);

convertBtn.addEventListener("click", () => {
  const mode = document.querySelector('input[name="convMode"]:checked').value;
  let pageSpec = "all";
  if (mode === "range") {
    const a = parseInt(convFrom.value, 10);
    const b = parseInt(convTo.value, 10);
    if (!a || !b || a < 1 || b < 1 || a > b || b > pageCount) {
      showError(`Range must be between 1 and ${pageCount}, with From ≤ To.`);
      return;
    }
    pageSpec = `${a}-${b}`;
  } else if (mode === "custom") {
    const raw = convCustom.value.trim();
    if (!raw) {
      showError("Tell The Boss which pages, kid.");
      return;
    }
    pageSpec = raw;
  }
  runConvert(currentFile, pageSpec);
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
function openSplitter() {
  splitName.textContent = currentFile.name;
  splitPages.textContent = `${pageCount} page${pageCount === 1 ? "" : "s"}`;
  splitRangesEl.innerHTML = "";
  // sensible defaults: one chunk covering the first half (or all if 1 page)
  if (pageCount === 1) {
    addRangeRow(1, 1);
  } else {
    const mid = Math.ceil(pageCount / 2);
    addRangeRow(1, mid);
    addRangeRow(mid + 1, pageCount);
  }
  showOnly(splitterEl);
}

function addRangeRow(from, to) {
  const idx = splitRangesEl.children.length + 1;
  const row = document.createElement("div");
  row.className = "range-row";
  row.innerHTML = `
    <span class="num">#${idx}</span>
    <label>From <input type="number" class="r-from" min="1" max="${pageCount}" value="${from || 1}" /></label>
    <label>To <input type="number" class="r-to" min="1" max="${pageCount}" value="${to || pageCount}" /></label>
    <button class="remove" type="button" aria-label="Remove chunk">remove</button>
  `;
  row.querySelector(".remove").addEventListener("click", () => {
    row.remove();
    renumberRanges();
    if (splitRangesEl.children.length === 0) addRangeRow(1, pageCount);
  });
  splitRangesEl.appendChild(row);
}

function renumberRanges() {
  Array.from(splitRangesEl.children).forEach((row, i) => {
    const n = row.querySelector(".num");
    if (n) n.textContent = `#${i + 1}`;
  });
}

addRangeBtn.addEventListener("click", () => addRangeRow(1, pageCount));
splitBack.addEventListener("click", showChooser);

splitBtn.addEventListener("click", () => {
  const rows = Array.from(splitRangesEl.querySelectorAll(".range-row"));
  if (!rows.length) { showError("Add at least one chunk."); return; }
  const specs = [];
  for (let i = 0; i < rows.length; i++) {
    const a = parseInt(rows[i].querySelector(".r-from").value, 10);
    const b = parseInt(rows[i].querySelector(".r-to").value, 10);
    if (!a || !b || a < 1 || b < 1 || a > b || b > pageCount) {
      showError(`Chunk #${i + 1}: range must be between 1 and ${pageCount}, with From ≤ To.`);
      return;
    }
    specs.push(a === b ? `${a}` : `${a}-${b}`);
  }
  runSplit(currentFile, specs);
});

function runSplit(file, specs) {
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = fmtSize(file.size);
  progressBar.style.width = "0%";
  progressBar.classList.add("indeterminate");
  statusMsg.textContent = "Carving up the document…";
  showOnly(statusEl);

  // One /api/split call per range so each chunk is individually addressable
  // in the result UI (download or promote-to-Excel).
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
        const safeSpec = spec.replace(",", "_");
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
    const item = document.createElement("div");
    item.className = "split-item";
    item.innerHTML = `
      <div class="label">Chunk #${i + 1}</div>
      <div class="pages">pages ${r.spec}</div>
      <div class="actions">
        <a class="dl" href="${url}" download="${r.name}">⬇ PDF</a>
        <button class="toxlsx" type="button">📊 Excel-ify</button>
      </div>
    `;
    item.querySelector(".toxlsx").addEventListener("click", () => {
      // Promote this split chunk into a File and run the converter on it.
      const f = new File([r.blob], r.name, { type: "application/pdf" });
      currentFile = f;
      reuploadAndConvert(f);
    });
    splitListEl.appendChild(item);
  });
  showOnly(splitResultEl);
}

async function reuploadAndConvert(file) {
  // Ask backend for page count of the new (split) file, then go to converter.
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

// ---------- Reset buttons ----------
againBtn.addEventListener("click", resetAll);
splitAgainBtn.addEventListener("click", resetAll);
retryBtn.addEventListener("click", resetAll);
