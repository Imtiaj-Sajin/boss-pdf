// ============================================================
// boss-pdf · usage dashboard
// ============================================================

// ---------- Header user chip ----------
(async function initUserChip() {
  const chip = document.getElementById("userChip");
  const nameEl = document.getElementById("userName");
  const titleEl = document.getElementById("userTitle");
  const avatarEl = document.getElementById("userAvatar");
  document.getElementById("logoutBtn").addEventListener("click", () => BossAuth.logout());
  const quick = BossAuth.quickUser();
  if (quick) { nameEl.textContent = quick.username; avatarEl.textContent = (quick.username[0] || "?").toUpperCase(); chip.classList.remove("hidden"); }
  try {
    const me = await BossAuth.me();
    if (me.full_name) nameEl.textContent = me.full_name;
    if (me.designation) titleEl.textContent = me.designation;
    if (me.profile_photo_url) {
      avatarEl.innerHTML = "";
      const img = document.createElement("img");
      img.src = me.profile_photo_url; img.alt = "";
      avatarEl.appendChild(img);
    } else {
      avatarEl.textContent = ((me.full_name || me.username || "?")[0]).toUpperCase();
    }
    chip.classList.remove("hidden");
  } catch (_) {}
})();

// ---------- Helpers ----------
function fmtMs(ms) {
  ms = +ms || 0;
  if (ms < 1000) return `${ms} ms`;
  const s = ms / 1000;
  if (s < 60) return `${s.toFixed(1)} s`;
  const m = s / 60;
  if (m < 60) return `${m.toFixed(1)} min`;
  return `${(m / 60).toFixed(1)} h`;
}
function fmtTimestamp(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(+d)) return "—";
  return d.toLocaleString();
}
function fmtRelative(iso) {
  if (!iso) return "never";
  const d = new Date(iso);
  if (isNaN(+d)) return "never";
  const diff = (Date.now() - d.getTime()) / 1000;
  if (diff < 60) return `${Math.floor(diff)}s ago`;
  if (diff < 3600) return `${Math.floor(diff / 60)}m ago`;
  if (diff < 86400) return `${Math.floor(diff / 3600)}h ago`;
  return `${Math.floor(diff / 86400)}d ago`;
}
function escapeHtml(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

// ---------- My summary ----------
async function loadMySummary() {
  try {
    const res = await BossAuth.authFetch("/api/usage/me");
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const t = await res.json();
    const meSummary = document.getElementById("meSummary");
    const stats = document.getElementById("meStats");
    const human = fmtMs(t.total_duration_ms || 0);
    meSummary.textContent =
      `${t.jobs || 0} jobs · ${t.sessions || 0} sessions · last used ${fmtRelative(t.last_used_at)} · ` +
      `${human} of compute the boss did for you.`;
    stats.innerHTML = [
      ["Files processed",    t.files_processed || 0],
      ["Pages worked",       t.pages_processed || 0],
      ["Tables → Excel",     t.tables_extracted || 0],
      ["OCR jobs",           t.ocr_jobs || 0],
      ["Split PDFs",         t.split_parts || 0],
      ["Sessions",           t.sessions || 0],
      ["Time worked",        fmtMs(t.total_duration_ms || 0)],
      ["Native downloads",   t.downloads || 0],
      ["OCR downloads",      t.ocr_downloads || 0],
      ["Split downloads",    t.split_downloads || 0],
      ["Total downloads",    t.total_downloads ||
                              (t.downloads || 0) + (t.ocr_downloads || 0) + (t.split_downloads || 0)],
    ].map(([k, v]) => `
      <div class="stat-tile">
        <div class="stat-num">${escapeHtml(v)}</div>
        <div class="stat-label">${escapeHtml(k)}</div>
      </div>
    `).join("");
  } catch (e) {
    document.getElementById("meSummary").textContent = "Couldn't load your totals.";
  }
}

// ---------- All rows ----------
let ALL_ROWS = [];
let filterRange = "all";
let filterOp = "all";
let filterText = "";

function withinRange(iso) {
  if (filterRange === "all") return true;
  if (!iso) return false;
  const d = new Date(iso); if (isNaN(+d)) return false;
  const diff = Date.now() - d.getTime();
  if (filterRange === "today") {
    const t = new Date(); t.setHours(0,0,0,0);
    return d >= t;
  }
  if (filterRange === "week")  return diff <= 7  * 86400 * 1000;
  if (filterRange === "month") return diff <= 30 * 86400 * 1000;
  return true;
}

function applyFilters(rows) {
  const q = filterText.trim().toLowerCase();
  return rows.filter(r => {
    if (filterOp !== "all" && r.operation !== filterOp) return false;
    if (!withinRange(r.used_at || r.finished_at)) return false;
    if (!q) return true;
    const blob = [r.username, r.full_name, r.designation, r.batch_name, r.email]
      .filter(Boolean).join(" ").toLowerCase();
    return blob.includes(q);
  });
}

function aggregateByUser(rows) {
  const m = new Map();
  for (const r of rows) {
    const key = r.user_id || `u${r.username || "?"}`;
    let g = m.get(key);
    if (!g) {
      g = {
        user_id: r.user_id,
        username: r.username,
        full_name: r.full_name,
        designation: r.designation,
        email: r.email,
        profile_photo_url: r.profile_photo_url,
        jobs: 0,
        files_processed: 0,
        pages_processed: 0,
        tables_extracted: 0,
        ocr_jobs: 0,
        split_parts: 0,
        downloads: 0,
        ocr_downloads: 0,
        split_downloads: 0,
        total_duration_ms: 0,
        sessions: new Set(),
        last_used: null,
        rows: [],
      };
      m.set(key, g);
    }
    g.jobs += 1;
    g.files_processed += (+r.files_processed || 0);
    g.pages_processed += (+r.pages_processed || 0);
    g.tables_extracted += (+r.tables_extracted || 0);
    g.ocr_jobs += (r.ocr_used ? 1 : 0);
    g.split_parts += (+r.split_parts || 0);
    g.downloads += (+r.downloads || 0);
    g.ocr_downloads += (+r.ocr_downloads || 0);
    g.split_downloads += (+r.split_downloads || 0);
    g.total_duration_ms += (+r.duration_ms || 0);
    if (r.session_id) g.sessions.add(r.session_id);
    const ts = r.used_at || r.finished_at;
    if (ts && (!g.last_used || new Date(ts) > new Date(g.last_used))) g.last_used = ts;
    g.rows.push(r);
  }
  // Sort: most-active first
  const list = [...m.values()];
  list.forEach(g => { g.sessions_count = g.sessions.size; });
  list.sort((a, b) => b.jobs - a.jobs);
  return list;
}

function renderUsers(groups) {
  const container = document.getElementById("usersList");
  if (!groups.length) {
    container.innerHTML = `<div class="thumbs-empty">No usage in this slice.</div>`;
    return;
  }
  container.innerHTML = groups.map((g, idx) => {
    const initial = ((g.full_name || g.username || "?")[0] || "?").toUpperCase();
    const avatar = g.profile_photo_url
      ? `<img src="${escapeHtml(g.profile_photo_url)}" alt="" />`
      : escapeHtml(initial);
    const rows = g.rows.slice().sort((a, b) =>
      new Date(b.used_at || 0) - new Date(a.used_at || 0)
    );
    return `
      <details class="user-card" ${idx === 0 ? "open" : ""}>
        <summary>
          <div class="user-summary">
            <div class="user-avatar lg">${avatar}</div>
            <div class="user-summary-meta">
              <div class="user-summary-name">
                ${escapeHtml(g.full_name || g.username || "(unknown)")}
                ${g.designation ? `<span class="user-summary-title">· ${escapeHtml(g.designation)}</span>` : ""}
              </div>
              <div class="user-summary-sub">
                @${escapeHtml(g.username || "?")} ${g.email ? `· ${escapeHtml(g.email)}` : ""}
                · last used ${fmtRelative(g.last_used)}
              </div>
            </div>
            <div class="user-summary-stats">
              <span><b>${g.jobs}</b> jobs</span>
              <span><b>${g.files_processed}</b> files</span>
              <span><b>${g.tables_extracted}</b> tables</span>
              <span><b>${g.ocr_jobs}</b> OCR</span>
              <span><b>${g.split_parts}</b> splits</span>
              <span title="native / OCR / split downloads">
                <b>${g.downloads}</b>/<b>${g.ocr_downloads}</b>/<b>${g.split_downloads}</b> dl
              </span>
              <span><b>${g.sessions_count}</b> sess</span>
              <span><b>${fmtMs(g.total_duration_ms)}</b></span>
            </div>
          </div>
        </summary>
        <div class="user-rows">
          <table class="usage-table">
            <thead>
              <tr>
                <th>When</th>
                <th>Op</th>
                <th>File</th>
                <th class="num">Pages</th>
                <th class="num">Tables</th>
                <th class="num">Splits</th>
                <th>OCR</th>
                <th class="num" title="native-text Excel">DL</th>
                <th class="num" title="OCR'd Excel">OCR DL</th>
                <th class="num" title="split PDFs">Split DL</th>
                <th>Uploaded</th>
                <th>Finished</th>
                <th class="num">Duration</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map(r => `
                <tr>
                  <td title="${escapeHtml(r.used_at || "")}">${escapeHtml(fmtRelative(r.used_at))}</td>
                  <td><span class="op-pill op-${escapeHtml(r.operation)}">${escapeHtml(r.operation)}</span></td>
                  <td class="file-cell" title="${escapeHtml(r.batch_name || "")}">${escapeHtml(r.batch_name || "—")}</td>
                  <td class="num">${escapeHtml(r.pages_processed)}</td>
                  <td class="num">${escapeHtml(r.tables_extracted)}</td>
                  <td class="num">${escapeHtml(r.split_parts)}</td>
                  <td>${r.ocr_used ? `<span class="ocr-pill">${escapeHtml(r.ocr_engine || "ocr")}</span>` : "—"}</td>
                  <td class="num">${escapeHtml(r.downloads || 0)}</td>
                  <td class="num">${escapeHtml(r.ocr_downloads || 0)}</td>
                  <td class="num">${escapeHtml(r.split_downloads || 0)}</td>
                  <td>${escapeHtml(fmtTimestamp(r.uploaded_at))}</td>
                  <td>${escapeHtml(fmtTimestamp(r.finished_at))}</td>
                  <td class="num">${escapeHtml(fmtMs(r.duration_ms))}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </details>
    `;
  }).join("");
}

function refresh() {
  const filtered = applyFilters(ALL_ROWS);
  document.getElementById("rowCount").textContent =
    `${filtered.length} of ${ALL_ROWS.length} jobs`;
  renderUsers(aggregateByUser(filtered));
}

async function loadAll() {
  try {
    const res = await BossAuth.authFetch("/api/usage");
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    ALL_ROWS = data.rows || [];
    refresh();
  } catch (e) {
    document.getElementById("usersList").innerHTML =
      `<div class="thumbs-empty err">Couldn't load usage: ${escapeHtml(e.message || e)}</div>`;
  }
}

// ---------- Filter handlers ----------
document.querySelectorAll("[data-range]").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll("[data-range]").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    filterRange = btn.dataset.range; refresh();
  });
});
document.querySelectorAll("[data-op]").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll("[data-op]").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    filterOp = btn.dataset.op; refresh();
  });
});
document.getElementById("searchBox").addEventListener("input", (e) => {
  filterText = e.target.value || ""; refresh();
});

loadMySummary();
loadAll();
