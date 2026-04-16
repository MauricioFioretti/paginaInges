// =====================
// CONFIG
// =====================
const API_BASE = "https://script.google.com/macros/s/AKfycbw--Vdhxhl2znuw_xDvAJiW7ZNyXbj4jKDwwHKG9B3VJNzbDt0jwaMBYnEo8f2GhtGT/exec";
const OAUTH_CLIENT_ID = "311839636060-a95bamqa6h8gst67tlcgo2puc3frrf9i.apps.googleusercontent.com";

const SPREADSHEET_ID = "1rJwf2PmsFyRRGoi160e5nbljfaLOY-KEniZ01M04ruc";

const OAUTH_SCOPES = [
  "openid",
  "email",
  "profile",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/userinfo.profile",
  "https://www.googleapis.com/auth/drive.metadata.readonly",
  "https://www.googleapis.com/auth/spreadsheets.readonly"
].join(" ");

const LS_OAUTH = "english_study_oauth_token_v1";
const LS_OAUTH_EMAIL = "english_study_oauth_email_v1";
const LS_CACHE = "english_study_cache_v1";
const LS_PENDING = "english_study_pending_v1";

const REQUIRED_COLUMNS = [
  "id","type","level","group_name","topic","text","translation","notes",
  "known","starred","created_at","updated_at","owner_email"
];

const SAMPLE_BULK = `phrase|A1|1|daily|How are you?|¿Cómo estás?|saludo básico
phrase|A1|1|daily|I am learning English.|Estoy aprendiendo inglés.|presente simple
word|Core|1|tech|server|servidor|infraestructura
phrase|Tech|1|programming|I need to update the server.|Necesito actualizar el servidor.|frase técnica`;

const headerSyncPill = document.getElementById("syncPill");
const btnConnect = document.getElementById("btnConnect");
const btnRefresh = document.getElementById("btnRefresh");
const accountPill = document.getElementById("accountPill");

const viewMode = document.getElementById("viewMode");
const levelFilter = document.getElementById("levelFilter");
const groupFilter = document.getElementById("groupFilter");
const searchInput = document.getElementById("searchInput");
const btnClearSearch = document.getElementById("btnClearSearch");
const statusFilter = document.getElementById("statusFilter");
const topicFilter = document.getElementById("topicFilter");
const sortMode = document.getElementById("sortMode");

const statTotal = document.getElementById("statTotal");
const statKnown = document.getElementById("statKnown");
const statUnknown = document.getElementById("statUnknown");
const statStarred = document.getElementById("statStarred");
const progressText = document.getElementById("progressText");
const progressFill = document.getElementById("progressFill");

const itemType = document.getElementById("itemType");
const itemLevel = document.getElementById("itemLevel");
const itemGroup = document.getElementById("itemGroup");
const itemTopic = document.getElementById("itemTopic");
const itemText = document.getElementById("itemText");
const itemTranslation = document.getElementById("itemTranslation");
const itemNotes = document.getElementById("itemNotes");
const btnAddItem = document.getElementById("btnAddItem");

const bulkInput = document.getElementById("bulkInput");
const btnImportBulk = document.getElementById("btnImportBulk");
const btnLoadSample = document.getElementById("btnLoadSample");

const btnExportProgress = document.getElementById("btnExportProgress");
const btnExportItems = document.getElementById("btnExportItems");
const btnMarkVisibleKnown = document.getElementById("btnMarkVisibleKnown");
const btnMarkVisibleUnknown = document.getElementById("btnMarkVisibleUnknown");
const btnToggleTranslations = document.getElementById("btnToggleTranslations");
const btnOnlyUnknown = document.getElementById("btnOnlyUnknown");
const visibleCount = document.getElementById("visibleCount");
const studyList = document.getElementById("studyList");
const toastRoot = document.getElementById("toastRoot");

let items = [];
let remoteMeta = { updatedAt: 0 };
let tokenClient = null;
let oauthAccessToken = "";
let oauthExpiresAt = 0;
let saveTimer = null;
let saving = false;
let connectInFlight = null;
let showTranslations = true;
let localVersion = 0;

// =====================
// HELPERS
// =====================
function uuid() {
  return "id_" + Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function nowIso() {
  return new Date().toISOString();
}

function normalizeStr(value) {
  return (value ?? "").toString().trim();
}

function normalizeItem(raw = {}) {
  return {
    id: normalizeStr(raw.id) || uuid(),
    type: normalizeStr(raw.type).toLowerCase() === "word" ? "word" : "phrase",
    level: normalizeStr(raw.level) || "Core",
    group_name: normalizeStr(raw.group_name || raw.group || "1"),
    topic: normalizeStr(raw.topic) || "general",
    text: normalizeStr(raw.text),
    translation: normalizeStr(raw.translation),
    notes: normalizeStr(raw.notes),
    known: toBool(raw.known),
    starred: toBool(raw.starred),
    created_at: normalizeStr(raw.created_at) || nowIso(),
    updated_at: normalizeStr(raw.updated_at) || nowIso(),
    owner_email: normalizeStr(raw.owner_email || "")
  };
}

function toBool(v) {
  if (typeof v === "boolean") return v;
  const s = normalizeStr(v).toLowerCase();
  return s === "true" || s === "1" || s === "si" || s === "sí" || s === "x";
}

function dedupItems(arr) {
  const seen = new Set();
  const out = [];
  for (const raw of arr || []) {
    const it = normalizeItem(raw);
    if (!it.text) continue;
    const key = it.id || (it.type + "|" + it.level + "|" + it.group_name + "|" + it.text.toLowerCase());
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(it);
  }
  return out;
}

function toast(msg, type = "ok", small = "") {
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.innerHTML = `<div>${escapeHtml(msg)}</div>${small ? `<div class="small">${escapeHtml(small)}</div>` : ""}`;
  toastRoot.appendChild(el);
  setTimeout(() => {
    el.style.opacity = "0";
    el.style.transform = "translateY(8px)";
    el.style.transition = "all .2s ease";
  }, 2400);
  setTimeout(() => el.remove(), 2800);
}

function escapeHtml(s) {
  return (s ?? "").toString()
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function setSync(state, text) {
  headerSyncPill.classList.remove("ok", "saving", "offline", "err");
  if (state) headerSyncPill.classList.add(state);
  headerSyncPill.querySelector(".sync-text").textContent = text;
}

function isOnline() {
  return navigator.onLine !== false;
}

function saveCache() {
  try {
    localStorage.setItem(LS_CACHE, JSON.stringify({ items, meta: remoteMeta, ts: Date.now() }));
  } catch {}
}

function loadCache() {
  try {
    const raw = localStorage.getItem(LS_CACHE);
    const parsed = raw ? JSON.parse(raw) : null;
    if (!Array.isArray(parsed?.items)) return null;
    return parsed;
  } catch {
    return null;
  }
}

function setPending(arr) {
  try {
    localStorage.setItem(LS_PENDING, JSON.stringify({ items: arr, ts: Date.now() }));
  } catch {}
}

function loadPending() {
  try {
    const raw = localStorage.getItem(LS_PENDING);
    const parsed = raw ? JSON.parse(raw) : null;
    return Array.isArray(parsed?.items) ? parsed : null;
  } catch {
    return null;
  }
}

function clearPending() {
  try {
    localStorage.removeItem(LS_PENDING);
  } catch {}
}

function loadStoredOAuth() {
  try {
    const raw = localStorage.getItem(LS_OAUTH);
    const parsed = raw ? JSON.parse(raw) : null;
    if (!parsed?.access_token || !parsed?.expires_at) return null;
    return { access_token: parsed.access_token, expires_at: Number(parsed.expires_at) };
  } catch {
    return null;
  }
}

function saveStoredOAuth(access_token, expires_at) {
  try {
    localStorage.setItem(LS_OAUTH, JSON.stringify({ access_token, expires_at }));
  } catch {}
}

function clearStoredOAuth() {
  try {
    localStorage.removeItem(LS_OAUTH);
  } catch {}
}

function loadStoredOAuthEmail() {
  try {
    return String(localStorage.getItem(LS_OAUTH_EMAIL) || "").trim().toLowerCase();
  } catch {
    return "";
  }
}

function saveStoredOAuthEmail(email) {
  try {
    localStorage.setItem(LS_OAUTH_EMAIL, (email || "").toString());
  } catch {}
}

function clearStoredOAuthEmail() {
  try {
    localStorage.removeItem(LS_OAUTH_EMAIL);
  } catch {}
}

function isTokenValid() {
  return !!oauthAccessToken && Date.now() < (oauthExpiresAt - 10000);
}

function setAccountUI(email) {
  const e = normalizeStr(email);
  if (!e) {
    accountPill.classList.add("hidden");
    accountPill.textContent = "";
    btnConnect.textContent = "Conectar";
    btnConnect.dataset.mode = "connect";
    return;
  }
  accountPill.classList.remove("hidden");
  accountPill.textContent = e;
  btnConnect.textContent = "Cambiar cuenta";
  btnConnect.dataset.mode = "switch";
}

// =====================
// OAUTH
// =====================
function initOAuth() {
  if (!window.google?.accounts?.oauth2?.initTokenClient) {
    throw new Error("GIS no está cargado");
  }

  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: OAUTH_CLIENT_ID,
    scope: OAUTH_SCOPES,
    include_granted_scopes: true,
    use_fedcm_for_prompt: true,
    callback: () => {}
  });
}

function requestAccessToken({ prompt, hint } = {}) {
  return new Promise((resolve, reject) => {
    if (!tokenClient) return reject(new Error("OAuth no inicializado"));
    let done = false;

    const timer = setTimeout(() => {
      if (done) return;
      done = true;
      reject(new Error("popup_timeout_or_closed"));
    }, 45000);

    tokenClient.callback = (resp) => {
      if (done) return;
      done = true;
      clearTimeout(timer);

      if (!resp || resp.error) {
        const err = String(resp?.error || "oauth_error");
        const sub = String(resp?.error_subtype || "");
        const msg = (err + (sub ? `:${sub}` : "")).toLowerCase();
        const e = new Error(err);
        e.isCanceled = msg.includes("popup_closed") ||
          msg.includes("popup_closed_by_user") ||
          msg.includes("access_denied") ||
          msg.includes("user_cancel") ||
          msg.includes("interaction_required");
        return reject(e);
      }

      const accessToken = resp.access_token;
      const expiresIn = Number(resp.expires_in || 3600);
      const expiresAt = Date.now() + (expiresIn * 1000);

      oauthAccessToken = accessToken;
      oauthExpiresAt = expiresAt;
      saveStoredOAuth(accessToken, expiresAt);
      resolve({ access_token: accessToken, expires_at: expiresAt });
    };

    const req = {};
    if (prompt !== undefined) req.prompt = prompt;
    if (hint && hint.includes("@")) req.hint = hint;

    try {
      tokenClient.requestAccessToken(req);
    } catch (e) {
      clearTimeout(timer);
      reject(e);
    }
  });
}

async function ensureOAuthToken(allowInteractive = false, interactivePrompt = "consent") {
  if (isTokenValid()) return oauthAccessToken;

  const stored = loadStoredOAuth();
  if (stored?.access_token && Date.now() < (stored.expires_at - 10000)) {
    oauthAccessToken = stored.access_token;
    oauthExpiresAt = Number(stored.expires_at);
    return oauthAccessToken;
  }

  const hintEmail = loadStoredOAuthEmail();

  if (!allowInteractive && !hintEmail) {
    throw new Error("TOKEN_NEEDS_INTERACTIVE");
  }

  try {
    await requestAccessToken({ prompt: "", hint: hintEmail || undefined });
    if (isTokenValid()) return oauthAccessToken;
  } catch (e) {
    if (!allowInteractive) throw new Error("TOKEN_NEEDS_INTERACTIVE");
  }

  await requestAccessToken({ prompt: interactivePrompt ?? "consent", hint: hintEmail || undefined });

  if (!isTokenValid()) throw new Error("TOKEN_NEEDS_INTERACTIVE");
  return oauthAccessToken;
}

// =====================
// API
// =====================
async function apiPost_(payload) {
  let r, text;

  try {
    r = await fetch(API_BASE, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload || {}),
      cache: "no-store",
      redirect: "follow"
    });
  } catch (e) {
    return { ok: false, error: "network_error", detail: String(e?.message || e) };
  }

  try {
    text = await r.text();
  } catch (e) {
    return { ok: false, error: "read_error", detail: String(e?.message || e) };
  }

  if (!r.ok) {
    return { ok: false, error: "http_error", status: r.status, detail: (text || "").slice(0, 800) };
  }

  try {
    return JSON.parse(text);
  } catch {
    return { ok: false, error: "non_json", detail: (text || "").slice(0, 800) };
  }
}

async function apiCall(mode, payload = {}, opts = {}) {
  const allowInteractive = !!opts.allowInteractive;
  let token = await ensureOAuthToken(allowInteractive, opts.interactivePrompt || "consent");

  const body = { mode, access_token: token, ...(payload || {}) };
  let data = await apiPost_(body);

  if (!data?.ok && (data?.error === "missing_scope" || data?.error === "auth_required")) {
    token = await ensureOAuthToken(true, "consent");
    body.access_token = token;
    data = await apiPost_(body);
  }

  return data || { ok: false, error: "empty_response" };
}

async function verifyBackendAccessOrThrow(allowInteractive) {
  const data = await apiCall("whoami", {}, { allowInteractive });
  if (!data?.ok) {
    const msg = (data?.error || "no_access") + (data?.detail ? ` | ${data.detail}` : "");
    throw new Error(msg);
  }
  return data;
}

// =====================
// DATA FLOW
// =====================
function buildFilters() {
  const levels = [...new Set(items.map(it => it.level).filter(Boolean))].sort();
  const groups = [...new Set(items.map(it => it.group_name).filter(Boolean))].sort((a,b) => String(a).localeCompare(String(b), undefined, { numeric:true }));
  const topics = [...new Set(items.map(it => it.topic).filter(Boolean))].sort();

  fillSelect(levelFilter, levels, "Todos");
  fillSelect(groupFilter, groups, "Todos");
  fillSelect(topicFilter, topics, "Todos");
}

function fillSelect(select, values, firstLabel) {
  const current = select.value;
  select.innerHTML = "";
  const first = document.createElement("option");
  first.value = "";
  first.textContent = firstLabel;
  select.appendChild(first);

  values.forEach(v => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    select.appendChild(opt);
  });

  if ([...select.options].some(o => o.value === current)) {
    select.value = current;
  }
}

function getVisibleItems() {
  const query = normalizeStr(searchInput.value).toLowerCase();
  const status = statusFilter.value;
  const mode = viewMode.value;
  const level = levelFilter.value;
  const group = groupFilter.value;
  const topic = topicFilter.value;
  const sort = sortMode.value;

  let out = items.filter(it => {
    if (mode !== "all") {
      if (mode === "phrases" && it.type !== "phrase") return false;
      if (mode === "words" && it.type !== "word") return false;
    }
    if (level && it.level !== level) return false;
    if (group && it.group_name !== group) return false;
    if (topic && it.topic !== topic) return false;
    if (status === "known" && !it.known) return false;
    if (status === "unknown" && it.known) return false;
    if (status === "starred" && !it.starred) return false;

    if (query) {
      const hay = [
        it.text, it.translation, it.notes, it.level, it.group_name, it.topic, it.type
      ].join(" ").toLowerCase();
      if (!hay.includes(query)) return false;
    }
    return true;
  });

  if (sort === "alpha") {
    out.sort((a, b) => a.text.localeCompare(b.text, undefined, { sensitivity: "base" }));
  } else if (sort === "knownFirst") {
    out.sort((a, b) => Number(b.known) - Number(a.known) || a.text.localeCompare(b.text));
  } else if (sort === "unknownFirst") {
    out.sort((a, b) => Number(a.known) - Number(b.known) || a.text.localeCompare(b.text));
  } else {
    out.sort((a, b) => {
      const lvl = a.level.localeCompare(b.level, undefined, { sensitivity: "base" });
      if (lvl !== 0) return lvl;
      const grp = String(a.group_name).localeCompare(String(b.group_name), undefined, { numeric: true });
      if (grp !== 0) return grp;
      return a.text.localeCompare(b.text, undefined, { sensitivity: "base" });
    });
  }

  return out;
}

function renderSummary(visible) {
  const total = items.length;
  const known = items.filter(it => it.known).length;
  const starred = items.filter(it => it.starred).length;
  const unknown = total - known;
  const pct = total ? Math.round((known / total) * 100) : 0;

  statTotal.textContent = String(total);
  statKnown.textContent = String(known);
  statUnknown.textContent = String(unknown);
  statStarred.textContent = String(starred);
  progressText.textContent = pct + "%";
  progressFill.style.width = pct + "%";
  visibleCount.textContent = `${visible.length} visibles`;
}

function render() {
  buildFilters();
  const visible = getVisibleItems();
  renderSummary(visible);

  studyList.innerHTML = "";

  if (!visible.length) {
    const empty = document.createElement("div");
    empty.className = "study-item";
    empty.innerHTML = `<div class="item-text">No hay items para mostrar.</div><div class="item-notes">Probá cambiando filtros o agregando contenido.</div>`;
    studyList.appendChild(empty);
    return;
  }

  visible.forEach(it => {
    const card = document.createElement("article");
    card.className = "study-item";

    card.innerHTML = `
      <div class="study-item-top">
        <div class="item-main">
          <div class="item-badges">
            <span class="badge">${escapeHtml(it.type)}</span>
            <span class="badge">${escapeHtml(it.level)}</span>
            <span class="badge">Grupo ${escapeHtml(it.group_name)}</span>
            <span class="badge">${escapeHtml(it.topic || "general")}</span>
          </div>
          <div class="item-text">${escapeHtml(it.text)}</div>
          <div class="item-translation ${showTranslations ? "" : "hidden"}">${escapeHtml(it.translation || "")}</div>
          ${it.notes ? `<div class="item-notes">${escapeHtml(it.notes)}</div>` : ``}
        </div>
        <div class="item-actions">
          <label class="toggle-wrap"><input type="checkbox" data-action="known" data-id="${escapeHtml(it.id)}" ${it.known ? "checked" : ""}> Ya sé</label>
          <button type="button" class="btn-secondary" data-action="star" data-id="${escapeHtml(it.id)}">${it.starred ? "★ Favorita" : "☆ Favorita"}</button>
          <button type="button" class="btn-secondary" data-action="delete" data-id="${escapeHtml(it.id)}">Eliminar</button>
        </div>
      </div>
    `;

    studyList.appendChild(card);
  });
}

function updateItem(id, patch = {}) {
  const idx = items.findIndex(it => it.id === id);
  if (idx === -1) return;

  items[idx] = normalizeItem({
    ...items[idx],
    ...patch,
    updated_at: nowIso()
  });

  localVersion++;
  saveCache();
  render();
  scheduleSave("Cambios guardados");
}

function addOneItem(raw) {
  const it = normalizeItem(raw);
  if (!it.text) {
    toast("Falta el texto principal", "warn");
    return false;
  }

  const exists = items.some(x =>
    x.type === it.type &&
    x.level.toLowerCase() === it.level.toLowerCase() &&
    x.group_name.toLowerCase() === it.group_name.toLowerCase() &&
    x.text.toLowerCase() === it.text.toLowerCase()
  );

  if (exists) {
    toast("Ese item ya existe", "warn", it.text);
    return false;
  }

  items.push(it);
  items = dedupItems(items);
  localVersion++;
  saveCache();
  render();
  scheduleSave("Item agregado");
  return true;
}

function deleteItem(id) {
  const current = items.find(it => it.id === id);
  if (!current) return;
  if (!confirm(`¿Eliminar "${current.text}"?`)) return;

  items = items.filter(it => it.id !== id);
  localVersion++;
  saveCache();
  render();
  scheduleSave("Item eliminado");
}

function parseBulkLine(line) {
  const parts = line.split("|").map(x => x.trim());
  if (parts.length < 6) return null;

  return {
    type: parts[0] || "phrase",
    level: parts[1] || "Core",
    group_name: parts[2] || "1",
    topic: parts[3] || "general",
    text: parts[4] || "",
    translation: parts[5] || "",
    notes: parts[6] || ""
  };
}

function exportJson(filename, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 500);
}

// =====================
// SAVE / SYNC
// =====================
async function scheduleSave(reason = "") {
  saveCache();

  if (!isOnline()) {
    setSync("offline", "Sin conexión — guardado local");
    setPending(items);
    return;
  }

  if (!isTokenValid()) {
    try {
      await ensureOAuthToken(false);
    } catch {}
  }

  if (!isTokenValid()) {
    setSync("offline", "Necesita Conectar");
    setPending(items);
    btnRefresh.classList.remove("hidden");
    return;
  }

  setSync("saving", "Guardando…");
  clearTimeout(saveTimer);

  saveTimer = setTimeout(async () => {
    if (saving) return;
    saving = true;

    try {
      const startedVersion = localVersion;
      const payloadItems = items.map(normalizeItem);
      if (!payloadItems.length) {
        setPending(payloadItems);
        setSync("offline", "No se guardó (vacío)");
        return;
      }

      let saved = await apiCall("set", { items: payloadItems, expectedUpdatedAt: Number(remoteMeta.updatedAt || 0) }, { allowInteractive: false });

      if (!saved?.ok && saved?.error === "conflict") {
        const remoteItems = Array.isArray(saved?.items) ? saved.items : [];
        const merged = mergeRemoteWithLocal(remoteItems, items);
        const saved2 = await apiCall("set", { items: merged, expectedUpdatedAt: Number(saved?.meta?.updatedAt || 0) }, { allowInteractive: false });
        if (!saved2?.ok) throw new Error(saved2?.error || "set_failed");
        saved = saved2;
        items = dedupItems(merged);
      }

      if (!saved?.ok) throw new Error(saved?.error || "set_failed");
      remoteMeta = { updatedAt: Number(saved?.meta?.updatedAt || 0) };

      if (localVersion !== startedVersion) {
        setPending(items);
        saving = false;
        scheduleSave("");
        return;
      }

      clearPending();
      saveCache();
      render();
      setSync("ok", "Guardado ✅");
      btnRefresh.classList.add("hidden");
      if (reason) toast("Guardado ✅", "ok", reason);
    } catch (e) {
      setPending(items);
      setSync("offline", "No se pudo guardar");
      btnRefresh.classList.remove("hidden");
      toast("No se pudo guardar", "err", "Quedó pendiente.");
    } finally {
      saving = false;
    }
  }, 650);
}

function mergeRemoteWithLocal(remoteItems, localItems) {
  const map = new Map();

  for (const raw of remoteItems || []) {
    const it = normalizeItem(raw);
    if (!it.text) continue;
    map.set(it.id, it);
  }

  for (const raw of localItems || []) {
    const it = normalizeItem(raw);
    if (!it.text) continue;
    map.set(it.id, it);
  }

  return dedupItems([...map.values()]);
}

async function refreshFromRemote(showToast = false, opts = { skipEnsureToken: false }) {
  if (!isOnline()) {
    setSync("offline", "Sin conexión — usando cache");
    return;
  }

  if (!opts?.skipEnsureToken) {
    try {
      await ensureOAuthToken(false);
    } catch {}
  }

  if (!isTokenValid()) {
    setSync("offline", "Necesita Conectar");
    btnRefresh.classList.remove("hidden");
    return;
  }

  try {
    const resp = await apiCall("get", {}, { allowInteractive: false });
    if (!resp?.ok) throw new Error(resp?.error || "get_failed");

    items = dedupItems(Array.isArray(resp.items) ? resp.items : []);
    remoteMeta = { updatedAt: Number(resp?.meta?.updatedAt || 0) };
    saveCache();
    render();
    setSync("ok", "Listo ✅");
    btnRefresh.classList.add("hidden");
    if (showToast) toast("Datos actualizados", "ok");
  } catch {
    setSync("offline", "No se pudo cargar");
    btnRefresh.classList.remove("hidden");
  }
}

async function trySyncPending() {
  const pending = loadPending();
  if (!pending?.items) {
    await refreshFromRemote(false);
    return;
  }

  if (!isOnline()) {
    setSync("offline", "Sin conexión — guardado local");
    return;
  }

  try {
    await ensureOAuthToken(false);
  } catch {}

  if (!isTokenValid()) {
    setSync("offline", "Necesita Conectar");
    btnRefresh.classList.remove("hidden");
    return;
  }

  setSync("saving", "Sincronizando…");

  try {
    const before = await apiCall("get", {}, { allowInteractive: false });
    if (!before?.ok) throw new Error(before?.error || "get_failed");

    const merged = mergeRemoteWithLocal(before.items || [], pending.items || []);
    const saved = await apiCall("set", { items: merged, expectedUpdatedAt: Number(before?.meta?.updatedAt || 0) }, { allowInteractive: false });
    if (!saved?.ok) throw new Error(saved?.error || "set_failed");

    items = dedupItems(merged);
    remoteMeta = { updatedAt: Number(saved?.meta?.updatedAt || 0) };
    clearPending();
    saveCache();
    render();
    setSync("ok", "Sincronizado ✅");
    btnRefresh.classList.add("hidden");
  } catch {
    setSync("offline", "Sincronización pendiente");
    btnRefresh.classList.remove("hidden");
  }
}

// =====================
// CONNECT FLOW
// =====================
async function runConnectFlow({ interactive, prompt } = { interactive: false, prompt: "consent" }) {
  if (connectInFlight) return connectInFlight;

  connectInFlight = (async () => {
    try {
      setSync("saving", interactive ? "Conectando…" : "Reconectando…");

      try {
        await ensureOAuthToken(!!interactive, prompt || "consent");
      } catch (e) {
        if (e?.isCanceled) {
          if (isTokenValid()) setSync("ok", "Listo ✅");
          else {
            setSync("offline", "Necesita Conectar");
            btnRefresh.classList.remove("hidden");
          }
          return { ok: false, canceled: true };
        }
        throw e;
      }

      const who = await verifyBackendAccessOrThrow(!!interactive);
      const email = normalizeStr(who?.email);
      if (email) saveStoredOAuthEmail(email);
      setAccountUI(email);

      btnRefresh.classList.add("hidden");
      await refreshFromRemote(true, { skipEnsureToken: true });
      return { ok: true };
    } catch (e) {
      const msg = String(e?.message || e || "");
      if (msg === "TOKEN_NEEDS_INTERACTIVE") {
        setSync("offline", "Necesita Conectar");
        btnRefresh.classList.remove("hidden");
        return { ok: false, needsInteractive: true };
      }

      setSync("offline", "Necesita Conectar");
      btnRefresh.classList.remove("hidden");
      return { ok: false, error: msg };
    } finally {
      connectInFlight = null;
    }
  })();

  return connectInFlight;
}

async function reconnectAndRefresh() {
  return await runConnectFlow({ interactive: false, prompt: "" });
}

// =====================
// EVENTS
// =====================
btnAddItem.addEventListener("click", () => {
  const ok = addOneItem({
    type: itemType.value,
    level: itemLevel.value || "Core",
    group_name: itemGroup.value || "1",
    topic: itemTopic.value || "general",
    text: itemText.value,
    translation: itemTranslation.value,
    notes: itemNotes.value
  });

  if (!ok) return;

  itemText.value = "";
  itemTranslation.value = "";
  itemNotes.value = "";
  itemText.focus();
});

btnImportBulk.addEventListener("click", () => {
  const lines = bulkInput.value.split("\n").map(x => x.trim()).filter(Boolean);
  if (!lines.length) {
    toast("Pegá un listado primero", "warn");
    return;
  }

  let added = 0;
  for (const line of lines) {
    const parsed = parseBulkLine(line);
    if (!parsed) continue;
    if (addOneItem(parsed)) added++;
  }

  toast("Importación terminada", "ok", `${added} items agregados`);
  bulkInput.value = "";
});

btnLoadSample.addEventListener("click", () => {
  bulkInput.value = SAMPLE_BULK;
});

btnExportProgress.addEventListener("click", () => {
  const data = items.map(it => ({
    id: it.id,
    text: it.text,
    known: it.known,
    starred: it.starred,
    updated_at: it.updated_at
  }));
  exportJson("english-study-progress.json", data);
});

btnExportItems.addEventListener("click", () => {
  exportJson("english-study-items.json", items);
});

btnMarkVisibleKnown.addEventListener("click", () => {
  const visible = getVisibleItems();
  visible.forEach(it => {
    const idx = items.findIndex(x => x.id === it.id);
    if (idx !== -1) items[idx].known = true;
  });
  localVersion++;
  saveCache();
  render();
  scheduleSave("Visibles marcados como ya sé");
});

btnMarkVisibleUnknown.addEventListener("click", () => {
  const visible = getVisibleItems();
  visible.forEach(it => {
    const idx = items.findIndex(x => x.id === it.id);
    if (idx !== -1) items[idx].known = false;
  });
  localVersion++;
  saveCache();
  render();
  scheduleSave("Visibles marcados como no sé");
});

btnToggleTranslations.addEventListener("click", () => {
  showTranslations = !showTranslations;
  render();
});

btnOnlyUnknown.addEventListener("click", () => {
  statusFilter.value = statusFilter.value === "unknown" ? "" : "unknown";
  render();
});

[viewMode, levelFilter, groupFilter, statusFilter, topicFilter, sortMode].forEach(el => {
  el.addEventListener("change", render);
});

searchInput.addEventListener("input", () => {
  btnClearSearch.classList.toggle("hidden", !searchInput.value);
  render();
});

btnClearSearch.addEventListener("click", () => {
  searchInput.value = "";
  btnClearSearch.classList.add("hidden");
  render();
});

studyList.addEventListener("click", (e) => {
  const btn = e.target.closest("[data-action]");
  if (!btn) return;

  const action = btn.dataset.action;
  const id = btn.dataset.id;

  if (action === "star") {
    const current = items.find(it => it.id === id);
    updateItem(id, { starred: !current?.starred });
    return;
  }

  if (action === "delete") {
    deleteItem(id);
  }
});

studyList.addEventListener("change", (e) => {
  const target = e.target;
  if (!(target instanceof HTMLInputElement)) return;
  if (target.dataset.action !== "known") return;

  updateItem(target.dataset.id, { known: target.checked });
});

btnConnect.addEventListener("click", async () => {
  if (btnConnect.dataset.mode === "switch") {
    const prevStored = loadStoredOAuth();
    const prevEmail = loadStoredOAuthEmail();
    const prevRuntimeToken = oauthAccessToken;
    const prevRuntimeExp = oauthExpiresAt;

    clearStoredOAuth();
    clearStoredOAuthEmail();
    oauthAccessToken = "";
    oauthExpiresAt = 0;

    const res = await runConnectFlow({ interactive: true, prompt: "select_account" });

    if (res?.canceled) {
      if (prevStored?.access_token && prevStored?.expires_at) {
        saveStoredOAuth(prevStored.access_token, prevStored.expires_at);
      }
      if (prevEmail) saveStoredOAuthEmail(prevEmail);
      oauthAccessToken = prevRuntimeToken || "";
      oauthExpiresAt = prevRuntimeExp || 0;
      setAccountUI(prevEmail || "");
      return;
    }

    if (!res?.ok) toast("No se pudo cambiar cuenta", "err");
    return;
  }

  const res = await runConnectFlow({ interactive: true, prompt: "consent" });
  if (!res?.ok && !res?.canceled) toast("No se pudo conectar", "err");
});

btnRefresh.addEventListener("click", async () => {
  await reconnectAndRefresh();
});

window.addEventListener("online", () => {
  toast("Volvió la conexión", "ok", "Intentando sincronizar");
  trySyncPending();
});

window.addEventListener("offline", () => {
  setSync("offline", "Sin conexión — guardado local");
});

setInterval(async () => {
  try {
    if (document.visibilityState !== "visible") return;
    if (!oauthAccessToken) return;
    if (Date.now() < (oauthExpiresAt - 120000)) return;
    await ensureOAuthToken(false);

    if (isTokenValid() && headerSyncPill.querySelector(".sync-text")?.textContent?.includes("Necesita Conectar")) {
      await reconnectAndRefresh();
    }
  } catch {}
}, 20000);

// =====================
// INIT
// =====================
window.addEventListener("load", async () => {
  try {
    initOAuth();

    const stored = loadStoredOAuth();
    if (stored?.access_token && Date.now() < (stored.expires_at - 10000)) {
      oauthAccessToken = stored.access_token;
      oauthExpiresAt = stored.expires_at;
      setAccountUI(loadStoredOAuthEmail());
    } else {
      setAccountUI(loadStoredOAuthEmail());
    }
  } catch {}

  const cached = loadCache();
  if (cached?.items) {
    items = dedupItems(cached.items);
    remoteMeta = cached.meta?.updatedAt ? { updatedAt: cached.meta.updatedAt } : { updatedAt: 0 };
    render();
    setSync(isOnline() ? "saving" : "offline", isOnline() ? "Cargando… (cache)" : "Sin conexión — usando cache");
  } else {
    setSync(isOnline() ? "saving" : "offline", isOnline() ? "Cargando…" : "Sin conexión");
  }

  const pending = loadPending();
  if (pending?.items) {
    items = dedupItems(pending.items);
    render();
    if (!isOnline()) {
      setSync("offline", "Sin conexión — cambios pendientes");
    } else {
      await trySyncPending();
    }
    return;
  }

  if (isOnline()) {
    const emailHint = loadStoredOAuthEmail();
    const stored = loadStoredOAuth();

    if (emailHint || (stored?.access_token && stored?.expires_at)) {
      await reconnectAndRefresh();
    } else {
      setSync("offline", "Necesita Conectar");
      btnRefresh.classList.remove("hidden");
    }
  } else {
    setSync("offline", "Sin conexión");
  }

  bulkInput.value = SAMPLE_BULK;
  render();
});
