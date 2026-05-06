const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const pickBtn = document.getElementById("pickBtn");
const statusEl = document.getElementById("status");
const fileNameEl = document.getElementById("fileName");
const fileSizeEl = document.getElementById("fileSize");
const progressBar = document.getElementById("progressBar");
const statusMsg = document.getElementById("statusMsg");
const resultEl = document.getElementById("result");
const downloadLink = document.getElementById("downloadLink");
const againBtn = document.getElementById("againBtn");
const errorEl = document.getElementById("error");
const errorMsg = document.getElementById("errorMsg");
const retryBtn = document.getElementById("retryBtn");

function show(el) { el.classList.remove("hidden"); }
function hide(el) { el.classList.add("hidden"); }
function reset() {
  hide(statusEl); hide(resultEl); hide(errorEl);
  progressBar.style.width = "0%";
  progressBar.classList.remove("indeterminate");
  fileInput.value = "";
}
function fmtSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1024 / 1024).toFixed(2) + " MB";
}

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
againBtn.addEventListener("click", reset);
retryBtn.addEventListener("click", reset);

async function handleFile(file) {
  if (!file.name.toLowerCase().endsWith(".pdf")) {
    showError("That doesn't look like a .pdf file.");
    return;
  }
  if (file.size > 50 * 1024 * 1024) {
    showError("File is over 50 MB. Try a smaller PDF.");
    return;
  }

  hide(errorEl); hide(resultEl); show(statusEl);
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = fmtSize(file.size);
  progressBar.style.width = "0%";
  progressBar.classList.remove("indeterminate");
  statusMsg.textContent = "Uploading…";

  const form = new FormData();
  form.append("file", file);

  // Upload with progress, then switch to indeterminate while server converts
  try {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", "/api/convert");
    xhr.responseType = "blob";

    xhr.upload.onprogress = e => {
      if (e.lengthComputable) {
        const pct = Math.round((e.loaded / e.total) * 100);
        // cap upload at 50% — the rest is conversion
        progressBar.style.width = (pct * 0.5) + "%";
      }
    };
    xhr.upload.onload = () => {
      statusMsg.textContent = "Extracting tables…";
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
        hide(statusEl);
        show(resultEl);
        // auto-trigger save
        downloadLink.click();
      } else {
        // Try to parse JSON error from blob
        const reader = new FileReader();
        reader.onload = () => {
          let msg = "Conversion failed.";
          try {
            const j = JSON.parse(reader.result);
            if (j && j.detail) msg = j.detail;
          } catch (_) { /* ignore */ }
          showError(msg);
        };
        reader.onerror = () => showError("Conversion failed.");
        reader.readAsText(xhr.response);
      }
    };
    xhr.onerror = () => showError("Network error. Is the server running?");
    xhr.send(form);
  } catch (err) {
    showError(err.message || String(err));
  }
}

function showError(message) {
  hide(statusEl); hide(resultEl);
  errorMsg.textContent = message;
  show(errorEl);
}
