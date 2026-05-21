// If already signed in with a valid token, send straight to /.
if (BossAuth.getToken() && !BossAuth.isExpired()) {
  window.location.replace("/");
}

const form = document.getElementById("loginForm");
const userEl = document.getElementById("username");
const passEl = document.getElementById("password");
const errEl = document.getElementById("loginErr");
const btn = document.getElementById("submitBtn");

function showErr(msg) {
  errEl.textContent = msg;
  errEl.classList.remove("hidden");
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  errEl.classList.add("hidden");
  btn.disabled = true;
  btn.textContent = "Checking…";
  try {
    await BossAuth.login(userEl.value.trim(), passEl.value);
    window.location.replace("/");
  } catch (err) {
    showErr(err.message || "Login failed.");
    btn.disabled = false;
    btn.textContent = "Sign in";
  }
});
