// Shared auth helper for boss-pdf — JWT in localStorage, session id per tab.
// Exposes window.BossAuth.
(function () {
  const TOKEN_KEY = "boss_pdf_access_token";
  const ME_KEY = "boss_pdf_me_cache";
  const SESSION_KEY = "boss_pdf_session_id";

  function getToken() { return localStorage.getItem(TOKEN_KEY) || ""; }
  function setToken(t) {
    if (t) localStorage.setItem(TOKEN_KEY, t);
    else localStorage.removeItem(TOKEN_KEY);
  }
  function clearAll() {
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem(ME_KEY);
    // keep session id — same browser session, same id
  }

  function getSessionId() {
    let sid = sessionStorage.getItem(SESSION_KEY);
    if (!sid) {
      sid = (crypto.randomUUID && crypto.randomUUID()) ||
            (Date.now().toString(36) + Math.random().toString(36).slice(2));
      sessionStorage.setItem(SESSION_KEY, sid);
    }
    return sid;
  }

  function decodeJwt(token) {
    try {
      const part = token.split(".")[1];
      const json = atob(part.replace(/-/g, "+").replace(/_/g, "/"));
      return JSON.parse(decodeURIComponent(escape(json)));
    } catch (_) { return null; }
  }

  function isExpired(token) {
    const p = decodeJwt(token || getToken());
    if (!p || !p.exp) return true;
    return p.exp * 1000 < Date.now();
  }

  function authHeaders(extra) {
    const h = Object.assign({}, extra || {});
    const t = getToken();
    if (t) h["Authorization"] = "Bearer " + t;
    h["X-Session-Id"] = getSessionId();
    return h;
  }

  // fetch wrapper that attaches the token and bounces to /login on 401.
  async function authFetch(url, opts) {
    opts = opts || {};
    opts.headers = authHeaders(opts.headers);
    const res = await fetch(url, opts);
    if (res.status === 401) {
      clearAll();
      window.location.href = "/login";
      throw new Error("Not authenticated.");
    }
    return res;
  }

  // XMLHttpRequest helper (used for upload progress)
  function applyAuthHeaders(xhr) {
    const t = getToken();
    if (t) xhr.setRequestHeader("Authorization", "Bearer " + t);
    xhr.setRequestHeader("X-Session-Id", getSessionId());
  }

  // Guard a page: if no/expired token, send to /login.
  function requireAuth() {
    const t = getToken();
    if (!t || isExpired(t)) {
      clearAll();
      window.location.href = "/login";
      return false;
    }
    return true;
  }

  async function login(username, password) {
    const res = await fetch("/api/auth/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, password }),
    });
    if (!res.ok) {
      const j = await res.json().catch(() => ({}));
      throw new Error(j.detail || "Login failed.");
    }
    const data = await res.json();
    setToken(data.access_token);
    return data;
  }

  function logout() {
    clearAll();
    window.location.href = "/login";
  }

  async function me({ force } = {}) {
    if (!force) {
      const cached = localStorage.getItem(ME_KEY);
      if (cached) {
        try { return JSON.parse(cached); } catch (_) {}
      }
    }
    const res = await authFetch("/api/auth/me");
    if (!res.ok) throw new Error("Could not load profile.");
    const data = await res.json();
    localStorage.setItem(ME_KEY, JSON.stringify(data));
    return data;
  }

  function quickUser() {
    const p = decodeJwt(getToken());
    return p ? { id: p.userId, username: p.username } : null;
  }

  window.BossAuth = {
    getToken, setToken, isExpired, requireAuth,
    authHeaders, authFetch, applyAuthHeaders,
    login, logout, me, quickUser, decodeJwt,
    getSessionId,
  };
})();
