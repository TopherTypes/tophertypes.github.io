/**
 * Meeting Notes single-page app.
 * Responsibilities:
 * - Manage local state (IndexedDB)
 * - Render UI and handle events
 * - Sync with Google Drive app data
 */

/* ================================
   CONFIG
=================================== */

// Google OAuth client credentials and API config.
// TODO: paste from Google Cloud Console
const CLIENT_ID = "15768027919-fs339ovijr4ueh0hkn77974bmbq8d9m1.apps.googleusercontent.com";
const API_KEY = "AIzaSyCxjOFISZK_OMVHN22OSdLf5CaLAeC9yDk";

// Google Drive discovery & scopes (appData is private to the user/app).
const DISCOVERY_DOC = "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest";
const SCOPES = "https://www.googleapis.com/auth/drive.appdata";
const DRIVE_FILENAME = "meeting-notes.v1.json";
const STANDARD_TEMPLATE_ID = "tpl_standard";
const ONE_TO_ONE_TEMPLATE_ID = "tpl_1on1";

// IndexedDB schema identifiers.
const IDB_NAME = "meeting-notes-db";
const IDB_VERSION = 1;
const IDB_STORE = "kv";

/* ================================
   GLOBAL STATE
=================================== */

let gapiInited = false;
let gisInited = false;
let tokenClient = null;
let driveReady = false;

let db = null;            // in-memory working DB
let currentMeetingId = null;
let actionsFilters = { ownerId: null, topicId: "", status: "" };
let meetingCalendarView = "week";
let meetingCalendarAnchor = new Date();
let meetingView = "setup";
// Tracks the meeting currently being edited in the lightbox, if any.
let meetingEditId = null;
// Tracks top-level module selection for the modular UI shell.
let activeModule = "meetings";
// Tracks the active tab within the Meetings module.
let meetingModuleTab = "meeting";
// Filters for the Tasks module list view.
let taskFilters = { status: "", priority: "" };

let syncInProgress = false;
let hasUnsyncedChanges = false;
let lastSyncAt = null;
let lastRemoteModifiedTime = null;
const itemEditState = new Map();
let personViewId = null;
let personEditorState = { isNew: false, draft: null, error: "" };
// Tracks draft state for the person creation lightbox flow.
let personCreateState = { draft: null, error: "" };

/* ================================
   TASKS MODULE CONFIG
=================================== */

const TASK_STATUS_LABELS = {
  todo: "To do",
  in_progress: "In progress",
  blocked: "Blocked",
  done: "Done"
};

const TASK_PRIORITY_LABELS = {
  low: "Low",
  medium: "Medium",
  high: "High"
};

// Sort weight maps for consistent task ordering.
const TASK_PRIORITY_ORDER = { high: 0, medium: 1, low: 2 };
const TASK_STATUS_ORDER = { todo: 0, in_progress: 1, blocked: 2, done: 3 };

/* ================================
   UTIL
=================================== */

/** @returns {string} ISO-8601 timestamp for consistent audit fields. */
function nowIso() { return new Date().toISOString(); }
/** @returns {string} Unique ID with a stable prefix for entity types. */
function uid(prefix) {
  // crypto.randomUUID is supported in modern browsers
  return `${prefix}_${crypto.randomUUID()}`;
}
/** @returns {HTMLElement|null} Convenience DOM accessor by id. */
function byId(id){ return document.getElementById(id); }

/**
 * Opens a lightbox modal and optionally focuses a field within it.
 * @param {string} lightboxId DOM id of the lightbox container.
 * @param {string} [focusId] Optional field id to focus on open.
 */
function openLightbox(lightboxId, focusId = "") {
  const lightbox = byId(lightboxId);
  if (!lightbox) return;
  lightbox.classList.add("is-visible");
  lightbox.setAttribute("aria-hidden", "false");
  if (focusId) {
    byId(focusId)?.focus();
  }
}

/**
 * Closes a lightbox modal by id.
 * @param {string} lightboxId DOM id of the lightbox container.
 */
function closeLightbox(lightboxId) {
  const lightbox = byId(lightboxId);
  if (!lightbox) return;
  lightbox.classList.remove("is-visible");
  lightbox.setAttribute("aria-hidden", "true");
}

/**
 * Closes any visible lightbox modals to keep keyboard escape behavior consistent.
 */
function closeVisibleLightboxes() {
  document.querySelectorAll(".lightbox.is-visible").forEach(lightbox => {
    lightbox.classList.remove("is-visible");
    lightbox.setAttribute("aria-hidden", "true");
  });
}

/** Updates the network status pill based on navigator connectivity. */
function setNetStatus() {
  const el = byId("net_status");
  const online = navigator.onLine;
  el.textContent = online ? "Online" : "Offline";
  el.style.borderColor = online ? "rgba(125,255,106,0.35)" : "rgba(255,211,106,0.35)";
  el.style.background = online ? "rgba(125,255,106,0.10)" : "rgba(255,211,106,0.10)";
  el.style.color = "var(--text)";
}

/** Sets sync status text and colors for the header pill. */
function setSyncStatus(text, kind="neutral") {
  const el = byId("sync_status");
  el.textContent = text;
  const map = {
    neutral: ["rgba(255,255,255,0.10)","rgba(255,255,255,0.03)"],
    ok: ["rgba(125,255,106,0.35)","rgba(125,255,106,0.10)"],
    warn: ["rgba(255,211,106,0.35)","rgba(255,211,106,0.10)"],
    bad: ["rgba(255,106,106,0.35)","rgba(255,106,106,0.10)"],
    accent: ["rgba(106,169,255,0.35)","rgba(106,169,255,0.12)"],
  };
  const [border, bg] = map[kind] || map.neutral;
  el.style.borderColor = border;
  el.style.background = bg;
  el.style.color = "var(--text)";
}

function escapeHtml(s){
  return (s ?? "").replace(/[&<>"']/g, (c) => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"
  }[c]));
}

/**
 * Renders required field indicators.
 * @param {boolean} required Whether a field is required.
 * @param {string} key Optional key to target later updates.
 */
function fieldTag(required, key = "") {
  if (!key) {
    return required
      ? `<span class="field-tag field-tag--required">Required</span>`
      : "";
  }
  return `<span class="field-tag field-tag--required${required ? "" : " is-hidden"}" data-required-tag="${key}">Required</span>`;
}

/**
 * Normalizes a task due date (YYYY-MM-DD) for display.
 * @param {string} dueDate ISO-like date string.
 * @returns {string} Friendly label for the due date.
 */
function formatTaskDueDate(dueDate) {
  if (!dueDate) return "No due date";
  const parsed = new Date(`${dueDate}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return dueDate;
  return parsed.toLocaleDateString();
}

/**
 * Formats an ISO timestamp for a datetime-local input value.
 * @param {string} isoDate ISO-8601 date/time string.
 * @returns {string} Localized value for datetime-local inputs.
 */
function toLocalDateTimeValue(isoDate) {
  if (!isoDate) return "";
  const parsed = new Date(isoDate);
  if (Number.isNaN(parsed.getTime())) return "";
  return new Date(parsed.getTime() - parsed.getTimezoneOffset() * 60000)
    .toISOString()
    .slice(0, 16);
}

function createBuiltinTemplates() {
  const updatedAt = nowIso();
  return [
    {
      id: STANDARD_TEMPLATE_ID,
      name: "Standard",
      updatedAt,
      sections: [
        { key: "info",     label: "Information", requires: [] },
        { key: "question", label: "Questions",   requires: [] },
        { key: "decision", label: "Decisions",   requires: [] },
        { key: "action",   label: "Actions",     requires: ["ownerId", "status"] }
      ]
    },
    {
      id: ONE_TO_ONE_TEMPLATE_ID,
      name: "1:1",
      updatedAt,
      sections: [
        { key: "info",     label: "Notes",         requires: [] },
        { key: "decision", label: "Decisions",     requires: [] },
        { key: "action",   label: "Actions",       requires: ["ownerId", "status"] },
        { key: "question", label: "Follow-ups",    requires: ["updateTargets"] }
      ]
    }
  ];
}

function findPersonByName(name, people = alive(db?.people || [])) {
  const needle = name.trim().toLowerCase();
  if (!needle) return null;
  return (people || []).find(p => p.name.toLowerCase() === needle) || null;
}

function ensurePeopleEmptyState(listEl) {
  if (!listEl) return;
  const hasSelected = listEl.querySelector("[data-selected-person]");
  const empty = listEl.querySelector("[data-empty]");
  if (!hasSelected && !empty) {
    const msg = document.createElement("div");
    msg.className = "muted";
    msg.setAttribute("data-empty", "true");
    msg.textContent = "No people selected yet.";
    listEl.appendChild(msg);
  } else if (hasSelected && empty) {
    empty.remove();
  }
}

function createPersonDraft(person) {
  return {
    name: person?.name || "",
    email: person?.email || "",
    organisation: person?.organisation || "",
    jobTitle: person?.jobTitle || "",
  };
}

function validatePersonDraft(draft, personId = null) {
  const errs = [];
  if (!draft.name.trim()) errs.push("Name is required.");
  if (!draft.email.trim()) errs.push("Email is required.");
  if (!draft.organisation.trim()) errs.push("Organisation is required.");
  const nameExists = alive(db.people).some(p => p.id !== personId && p.name.toLowerCase() === draft.name.trim().toLowerCase());
  if (nameExists) errs.push("Name must be unique.");
  return errs;
}

/**
 * Creates an editable task draft from a task record.
 * @param {object} task Task entity.
 * @returns {object} Draft fields for editing.
 */
function createTaskDraft(task) {
  return {
    title: task?.title || "",
    notes: task?.notes || "",
    dueDate: task?.dueDate || "",
    priority: task?.priority || "medium",
    status: task?.status || "todo"
  };
}

/**
 * Validates a task draft and returns a list of user-facing errors.
 * @param {object} draft Task draft input.
 * @returns {string[]} Validation error messages.
 */
function validateTaskDraft(draft) {
  const errs = [];
  if (!draft.title.trim()) errs.push("Task title is required.");
  return errs;
}

function wirePeoplePickers(container, people) {
  const peopleMap = new Map(people.map(p => [p.name.toLowerCase(), p]));
  container.querySelectorAll("[data-people-picker]").forEach(picker => {
    const input = picker.querySelector("[data-people-input]");
    const addBtn = picker.querySelector("[data-add-person]");
    const list = picker.querySelector("[data-selected-list]");
    ensurePeopleEmptyState(list);

    const addPersonFromInput = () => {
      const value = input?.value.trim() || "";
      if (!value) return;
      const person = peopleMap.get(value.toLowerCase());
      if (!person) {
        alert("Choose a person from the list.");
        return;
      }
      if (list?.querySelector(`[data-selected-person="${person.id}"]`)) {
        input.value = "";
        return;
      }
      const pill = document.createElement("span");
      pill.className = "person-pill";
      pill.setAttribute("data-selected-person", person.id);

      const name = document.createElement("span");
      name.textContent = person.name;

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "person-pill__remove";
      removeBtn.setAttribute("aria-label", `Remove ${person.name}`);
      removeBtn.textContent = "×";
      removeBtn.addEventListener("click", () => {
        pill.remove();
        ensurePeopleEmptyState(list);
      });

      pill.appendChild(name);
      pill.appendChild(removeBtn);
      list.appendChild(pill);
      input.value = "";
      ensurePeopleEmptyState(list);
    };

    addBtn?.addEventListener("click", addPersonFromInput);
    input?.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        addPersonFromInput();
      }
    });
  });
}

function fmtDateTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString(undefined, { year:"numeric", month:"short", day:"2-digit", hour:"2-digit", minute:"2-digit" });
}

function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleDateString(undefined, { year:"numeric", month:"short", day:"2-digit" });
}

function copyToClipboard(text) {
  return navigator.clipboard.writeText(text);
}

function debounce(fn, delay=200) {
  let t = null;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), delay);
  };
}

function normalizeDate(d) {
  const out = new Date(d);
  out.setHours(0, 0, 0, 0);
  return out;
}

function dateKeyFromDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function dateKeyFromIso(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return null;
  return dateKeyFromDate(d);
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function startOfWeek(date) {
  const d = normalizeDate(date);
  const day = (d.getDay() + 6) % 7; // Monday = 0
  return addDays(d, -day);
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function formatTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
}

/* ================================
   DEFAULT DB / TEMPLATES
=================================== */

function makeDefaultDb() {
  // starter people/groups/topics empty
  return {
    schemaVersion: 1,
    updatedAt: nowIso(),
    settings: {
      defaultOwnerName: "",
      updatedAt: nowIso(),
    },
    templates: createBuiltinTemplates(),
    people: [],
    groups: [],
    topics: [],
    meetings: [],
    items: [],
    tasks: []
  };
}

/* ================================
   INDEXEDDB (simple KV store)
=================================== */

function idbOpen() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbGet(key) {
  const dbi = await idbOpen();
  return new Promise((resolve, reject) => {
    const tx = dbi.transaction(IDB_STORE, "readonly");
    const store = tx.objectStore(IDB_STORE);
    const req = store.get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbSet(key, val) {
  const dbi = await idbOpen();
  return new Promise((resolve, reject) => {
    const tx = dbi.transaction(IDB_STORE, "readwrite");
    const store = tx.objectStore(IDB_STORE);
    const req = store.put(val, key);
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error);
  });
}

/* ================================
   GOOGLE API (Drive appDataFolder)
=================================== */

// Called by script tag onload in index.html
window.gapiLoaded = function gapiLoaded() {
  gapi.load("client", async () => {
    await gapi.client.init({
      apiKey: API_KEY,
      discoveryDocs: [DISCOVERY_DOC],
    });
    gapiInited = true;
    maybeEnableAuth();
  });
};

// Called by script tag onload in index.html
window.gisLoaded = function gisLoaded() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: "" // set later
  });
  gisInited = true;
  maybeEnableAuth();
};

function maybeEnableAuth() {
  const authBtn = byId("auth_btn");
  const signoutBtn = byId("signout_btn");
  const syncBtn = byId("sync_btn");

  if (!(gapiInited && gisInited)) return;

  authBtn.disabled = false;
  signoutBtn.disabled = false;
  syncBtn.disabled = false;

  // Default state before sign-in
  updateAuthUi();
}

function updateAuthUi() {
  const gapiClient = window.gapi?.client;
  const token = gapiClient?.getToken?.() || null;
  driveReady = !!token;

  const authBtn = byId("auth_btn");
  const signoutBtn = byId("signout_btn");

  authBtn.textContent = driveReady ? "Re-authorize" : "Sign in";
  signoutBtn.style.display = driveReady ? "inline-block" : "none";

  setSyncStatus(driveReady
    ? (hasUnsyncedChanges ? "Unsynced changes" : (lastSyncAt ? `Synced ${fmtDateTime(lastSyncAt)}` : "Drive ready"))
    : "Local only",
    driveReady ? (hasUnsyncedChanges ? "warn" : "ok") : "neutral"
  );
}

async function handleAuthClick() {
  if (!tokenClient) return;
  if (!window.gapi || !window.google) {
    alert("Google APIs are unavailable. Check your network or ad blocker settings.");
    return;
  }

  tokenClient.callback = async (resp) => {
    if (resp.error) {
      console.error(resp);
      alert("Sign-in failed. Check console for details.");
      return;
    }
    // attach token to gapi
    window.gapi?.client?.setToken?.(resp);

    updateAuthUi();

    // Optional: if we have unsynced changes, prompt to sync
    if (hasUnsyncedChanges && confirm("You have local changes. Sync now?")) {
      await syncNow();
    }
  };

  // prompt consent first time, silent refresh after
  const gapiClient = window.gapi?.client;
  const token = gapiClient?.getToken?.() || null;
  if (token === null) {
    tokenClient.requestAccessToken({ prompt: "consent" });
  } else {
    tokenClient.requestAccessToken({ prompt: "" });
  }
}

function handleSignoutClick() {
  const gapiClient = window.gapi?.client;
  if (!gapiClient || !window.google) {
    driveReady = false;
    updateAuthUi();
    return;
  }
  const token = gapiClient.getToken();
  if (token) {
    window.google.accounts.oauth2.revoke(token.access_token);
    gapiClient.setToken("");
  }
  driveReady = false;
  updateAuthUi();
}

async function findDriveFileId() {
  const res = await gapi.client.drive.files.list({
    spaces: "appDataFolder",
    q: `name='${DRIVE_FILENAME.replace(/'/g, "\\'")}' and trashed=false`,
    fields: "files(id, name, modifiedTime)",
    pageSize: 10
  });
  const files = res.result.files || [];
  if (!files.length) return null;
  const f = files[0];
  lastRemoteModifiedTime = f.modifiedTime || null;
  return f.id;
}

async function createDriveFile(initialJson) {
  const metadata = {
    name: DRIVE_FILENAME,
    parents: ["appDataFolder"],
    mimeType: "application/json"
  };

  const boundary = "-------314159265358979323846";
  const delimiter = "\r\n--" + boundary + "\r\n";
  const closeDelim = "\r\n--" + boundary + "--";

  const multipartRequestBody =
    delimiter +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata) +
    delimiter +
    "Content-Type: application/json\r\n\r\n" +
    JSON.stringify(initialJson) +
    closeDelim;

  const res = await gapi.client.request({
    path: "/upload/drive/v3/files?uploadType=multipart&fields=id,name,modifiedTime",
    method: "POST",
    headers: { "Content-Type": 'multipart/related; boundary="' + boundary + '"' },
    body: multipartRequestBody
  });

  lastRemoteModifiedTime = res.result.modifiedTime || null;
  return res.result.id;
}

async function getDriveFileMeta(fileId) {
  const res = await gapi.client.drive.files.get({
    fileId,
    fields: "id,name,modifiedTime,size"
  });
  lastRemoteModifiedTime = res.result.modifiedTime || null;
  return res.result;
}

async function downloadDriveJson(fileId) {
  const res = await gapi.client.request({
    path: `/drive/v3/files/${fileId}`,
    method: "GET",
    params: { alt: "media" }
  });
  // gapi may return string body
  if (typeof res.body === "string") return JSON.parse(res.body);
  return res.result;
}

async function uploadDriveJson(fileId, jsonObj) {
  const res = await gapi.client.request({
    path: `/upload/drive/v3/files/${fileId}`,
    method: "PATCH",
    params: { uploadType: "media" },
    headers: { "Content-Type": "application/json; charset=UTF-8" },
    body: JSON.stringify(jsonObj)
  });
  // refresh modifiedTime
  await getDriveFileMeta(fileId);
  return res;
}

async function ensureDriveFile() {
  let fileId = await findDriveFileId();
  if (!fileId) {
    fileId = await createDriveFile(db ?? makeDefaultDb());
  }
  return fileId;
}

/* ================================
   MERGE LOGIC
=================================== */

function isoNewer(a, b) {
  // returns true if a > b
  const da = Date.parse(a || 0);
  const db = Date.parse(b || 0);
  return da > db;
}

function indexById(arr) {
  const m = new Map();
  for (const x of (arr || [])) m.set(x.id, x);
  return m;
}

function mergeUpdateStatus(a = {}, b = {}) {
  const out = { ...a };
  for (const [personId, stB] of Object.entries(b)) {
    const stA = out[personId];
    if (!stA) {
      out[personId] = stB;
    } else {
      // if either updated true, keep true with latest updatedAt
      const updatedA = !!stA.updated;
      const updatedB = !!stB.updated;
      if (updatedA || updatedB) {
        const best = (!updatedA) ? stB : (!updatedB) ? stA
          : isoNewer(stA.updatedAt, stB.updatedAt) ? stA : stB;
        out[personId] = { ...best, updated: true };
      } else {
        // both false -> keep latest "updatedAt" if present, else keep A
        out[personId] = isoNewer(stB.updatedAt, stA.updatedAt) ? stB : stA;
      }
    }
  }
  return out;
}

function mergeRecord(a, b) {
  // handle deletes/tombstones
  const aDel = !!a?.deleted;
  const bDel = !!b?.deleted;

  if (a && !b) return a;
  if (b && !a) return b;

  // choose base by updatedAt
  const base = isoNewer(a.updatedAt, b.updatedAt) ? a : b;
  const other = base === a ? b : a;

  const merged = { ...base };

  // merge special fields for items
  if (merged.type === "item" || merged.kind === "item" || ("updateStatus" in merged) || ("updateTargets" in merged)) {
    merged.updateStatus = mergeUpdateStatus(a.updateStatus, b.updateStatus);
  }

  // merge updateTargets as union (safe)
  if (Array.isArray(a.updateTargets) || Array.isArray(b.updateTargets)) {
    const s = new Set([...(a.updateTargets || []), ...(b.updateTargets || [])]);
    merged.updateTargets = Array.from(s);
  }

  // If one side deleted and has later updatedAt, keep deleted.
  if (aDel || bDel) {
    const delWinner = isoNewer(a.updatedAt, b.updatedAt) ? a : b;
    if (delWinner.deleted) merged.deleted = true;
  }

  // preserve missing fields from other (non-overwriting)
  for (const [k, v] of Object.entries(other)) {
    if (!(k in merged)) merged[k] = v;
  }

  return merged;
}

function mergeCollections(localArr, remoteArr) {
  const L = indexById(localArr);
  const R = indexById(remoteArr);

  const ids = new Set([...L.keys(), ...R.keys()]);
  const out = [];

  for (const id of ids) {
    const a = L.get(id);
    const b = R.get(id);
    const merged = mergeRecord(a, b);
    out.push(merged);
  }

  // keep non-deleted first for nicer output (optional)
  out.sort((x, y) => (x.deleted === y.deleted) ? 0 : (x.deleted ? 1 : -1));
  return out;
}

function normalizeBuiltinTemplates(store) {
  if (!store) return false;
  const builtinTemplates = createBuiltinTemplates();
  const templates = alive(store.templates || []);
  const meetings = alive(store.meetings || []);
  let changed = false;

  const updatedTemplates = [...templates];

  for (const builtin of builtinTemplates) {
    const matches = updatedTemplates.filter(t => t.name === builtin.name);
    if (!matches.length) {
      updatedTemplates.push(builtin);
      changed = true;
      continue;
    }

    let winner = matches.find(t => t.id === builtin.id) || matches[0];
    for (const candidate of matches) {
      if (isoNewer(candidate.updatedAt, winner.updatedAt)) winner = candidate;
    }

    if (winner.id !== builtin.id) {
      const oldId = winner.id;
      winner.id = builtin.id;
      meetings.forEach(m => {
        if (m.templateId === oldId) m.templateId = builtin.id;
      });
      changed = true;
    }

    if (!Array.isArray(winner.sections) || !winner.sections.length) {
      winner.sections = builtin.sections;
      changed = true;
    }

    for (const dup of matches) {
      if (dup === winner) continue;
      const oldId = dup.id;
      meetings.forEach(m => {
        if (m.templateId === oldId) m.templateId = winner.id;
      });
      const idx = updatedTemplates.indexOf(dup);
      if (idx !== -1) updatedTemplates.splice(idx, 1);
      changed = true;
    }
  }

  if (changed) {
    store.templates = updatedTemplates;
  }
  return changed;
}

function mergeDb(localDb, remoteDb) {
  // schema guard
  const l = localDb || makeDefaultDb();
  const r = remoteDb || makeDefaultDb();
  const lSettings = l.settings || { defaultOwnerName: "", updatedAt: l.updatedAt || nowIso() };
  const rSettings = r.settings || { defaultOwnerName: "", updatedAt: r.updatedAt || nowIso() };

  const merged = {
    schemaVersion: 1,
    updatedAt: nowIso(),
    settings: mergeRecord(lSettings, rSettings),
    templates: mergeCollections(l.templates || [], r.templates || []),
    people: mergeCollections(l.people || [], r.people || []),
    groups: mergeCollections(l.groups || [], r.groups || []),
    topics: mergeCollections(l.topics || [], r.topics || []),
    meetings: mergeCollections(l.meetings || [], r.meetings || []),
    items: mergeCollections(l.items || [], r.items || []),
    tasks: mergeCollections(l.tasks || [], r.tasks || []),
  };

  if (normalizeBuiltinTemplates(merged)) {
    merged.updatedAt = nowIso();
  }

  return merged;
}

/* ================================
   PERSISTENCE (local)
=================================== */

async function loadLocal() {
  const stored = await idbGet("db");
  const meta = await idbGet("meta");

  if (stored) {
    db = stored;
  } else {
    db = makeDefaultDb();
    await idbSet("db", db);
  }

  if (!db.settings) {
    db.settings = { defaultOwnerName: "", updatedAt: nowIso() };
  }
  if (!db.tasks) {
    db.tasks = [];
  }

  if (normalizeBuiltinTemplates(db)) {
    markDirty();
    await saveLocal();
  }

  if (meta) {
    currentMeetingId = meta.currentMeetingId || null;
    hasUnsyncedChanges = !!meta.hasUnsyncedChanges;
    lastSyncAt = meta.lastSyncAt || null;
    lastRemoteModifiedTime = meta.lastRemoteModifiedTime || null;
    actionsFilters = {
      ownerId: meta.actionsFilters?.ownerId ?? null,
      topicId: meta.actionsFilters?.topicId || "",
      status: meta.actionsFilters?.status || "",
    };
    meetingCalendarView = meta.meetingCalendarView || "week";
    meetingCalendarAnchor = meta.meetingCalendarAnchor ? new Date(meta.meetingCalendarAnchor) : new Date();
    activeModule = meta.activeModule || "meetings";
    meetingModuleTab = meta.meetingModuleTab || "meeting";
    taskFilters = {
      status: meta.taskFilters?.status || "",
      priority: meta.taskFilters?.priority || "",
    };
    if (Number.isNaN(meetingCalendarAnchor.getTime())) {
      meetingCalendarAnchor = new Date();
    }
  } else {
    await saveMeta();
  }
}

async function saveLocal() {
  db.updatedAt = nowIso();
  await idbSet("db", db);
  await saveMeta();
}

async function saveMeta() {
  await idbSet("meta", {
    currentMeetingId,
    hasUnsyncedChanges,
    lastSyncAt,
    lastRemoteModifiedTime,
    actionsFilters,
    meetingCalendarView,
    meetingCalendarAnchor: meetingCalendarAnchor?.toISOString?.() || nowIso(),
    activeModule,
    meetingModuleTab,
    taskFilters
  });
  updateAuthUi();
}

function markDirty() {
  hasUnsyncedChanges = true;
  updateAuthUi();
  saveMeta().catch(console.error);
}

function markClean() {
  hasUnsyncedChanges = false;
  updateAuthUi();
  saveMeta().catch(console.error);
}

/* ================================
   SYNC
=================================== */

async function syncNow() {
  if (!driveReady) {
    alert("Sign in first.");
    return;
  }
  if (!navigator.onLine) {
    alert("You appear to be offline. Sync will work when you're online.");
    return;
  }
  if (syncInProgress) return;

  syncInProgress = true;
  setSyncStatus("Syncing…", "accent");
  byId("sync_btn").disabled = true;

  try {
    const fileId = await ensureDriveFile();

    // Download remote
    const remote = await downloadDriveJson(fileId);
    await getDriveFileMeta(fileId);

    // Merge
    const merged = mergeDb(db, remote);

    // Upload merged
    await uploadDriveJson(fileId, merged);

    // Save merged locally
    db = merged;
    lastSyncAt = nowIso();
    await saveLocal();
    await saveMeta();
    markClean();

    renderAll();
    setSyncStatus(`Synced ${fmtDateTime(lastSyncAt)}`, "ok");
  } catch (e) {
    console.error(e);
    setSyncStatus("Sync failed", "bad");
    alert("Sync failed. Check console for details.");
  } finally {
    syncInProgress = false;
    byId("sync_btn").disabled = false;
    updateAuthUi();
  }
}

/* ================================
   DB HELPERS
=================================== */

function alive(arr) {
  return (arr || []).filter(x => !x.deleted);
}

function getPerson(id) {
  return alive(db.people).find(p => p.id === id) || null;
}
function getGroup(id) {
  return alive(db.groups).find(g => g.id === id) || null;
}
function getTopic(id) {
  return alive(db.topics).find(t => t.id === id) || null;
}
function getMeeting(id) {
  return alive(db.meetings).find(m => m.id === id) || null;
}
function getTemplate(id) {
  return alive(db.templates).find(t => t.id === id) || null;
}
function getItem(id) {
  return alive(db.items).find(i => i.id === id) || null;
}
function getTask(id) {
  return alive(db.tasks).find(t => t.id === id) || null;
}

function ensureTopic(name) {
  const existing = alive(db.topics).find(t => t.name.toLowerCase() === name.toLowerCase());
  if (existing) return existing.id;
  const t = { id: uid("topic"), name, updatedAt: nowIso() };
  db.topics.push(t);
  return t.id;
}

function ensurePerson(name) {
  const existing = alive(db.people).find(p => p.name.toLowerCase() === name.toLowerCase());
  if (existing) return existing.id;
  const p = { id: uid("person"), name, email: "", organisation: "", jobTitle: "", updatedAt: nowIso() };
  db.people.push(p);
  return p.id;
}

function expandTargets(selectedPeopleIds, selectedGroupIds) {
  const s = new Set(selectedPeopleIds || []);
  for (const gid of (selectedGroupIds || [])) {
    const g = getGroup(gid);
    if (!g) continue;
    for (const pid of (g.memberIds || [])) s.add(pid);
  }
  return Array.from(s);
}

/**
 * Builds the 1:1 counterpart context for a meeting, if configured.
 * @param {object} meeting Meeting record.
 * @returns {object|null} Counterpart context with type, label, and person ids.
 */
function getOneToOneContext(meeting) {
  if (!meeting) return null;
  const personId = meeting.oneToOnePersonId || "";
  const groupId = meeting.oneToOneGroupId || "";
  if (personId) {
    const person = getPerson(personId);
    if (!person) return null;
    return {
      type: "person",
      label: person.name,
      personIds: [person.id],
    };
  }
  if (groupId) {
    const group = getGroup(groupId);
    if (!group) return null;
    const memberIds = expandTargets([], [group.id]);
    return {
      type: "group",
      label: group.name,
      personIds: memberIds,
    };
  }
  return null;
}

/**
 * Returns items linked to any of the provided person identifiers.
 * @param {string[]} personIds People to match against owner/update targets.
 * @returns {object[]} Meeting items linked to the target people.
 */
function getItemsLinkedToPeople(personIds) {
  const targetSet = new Set(personIds || []);
  if (!targetSet.size) return [];
  return alive(db.items).filter(item => {
    const matchesOwner = item.ownerId && targetSet.has(item.ownerId);
    const matchesUpdateTargets = (item.updateTargets || []).some(pid => targetSet.has(pid));
    return matchesOwner || matchesUpdateTargets;
  });
}

/* ================================
   UI RENDERING
=================================== */

/**
 * Toggles the top-level module shell in the interface.
 * @param {string} name Module name to activate.
 */
function setActiveModule(name) {
  activeModule = name;
  document.querySelectorAll(".module-tab").forEach(btn => {
    btn.classList.toggle("is-active", btn.dataset.module === name);
  });
  document.querySelectorAll(".module-panel").forEach(panel => panel.classList.remove("is-active"));
  const target = byId(`module_${name}`);
  if (target) {
    target.classList.add("is-active");
  }
  if (name === "meetings") {
    setMeetingModuleTab(meetingModuleTab);
  }
  saveMeta().catch(console.error);
}

/**
 * Toggles the active tab within the Meetings module.
 * @param {string} name Meeting module tab name to activate.
 */
function setMeetingModuleTab(name) {
  meetingModuleTab = name;
  const container = byId("module_meetings");
  if (!container) return;
  container.querySelectorAll(".module-subtab").forEach(btn => {
    btn.classList.toggle("is-active", btn.dataset.moduleTab === name);
  });
  container.querySelectorAll(".module-section").forEach(section => {
    section.classList.toggle("is-active", section.id === `module_tab_${name}`);
  });
  if (name === "meeting") {
    setMeetingView(meetingView);
  }
  saveMeta().catch(console.error);
}

function setMeetingView(view) {
  meetingView = view;
  document.querySelectorAll("[data-meeting-view]").forEach(btn => {
    const isActive = btn.getAttribute("data-meeting-view") === view;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-selected", isActive ? "true" : "false");
  });
  document.querySelectorAll("[data-meeting-view-panel]").forEach(panel => {
    const isActive = panel.getAttribute("data-meeting-view-panel") === view;
    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
  });
}

function renderSelectOptions(select, options, {placeholder=null} = {}) {
  select.innerHTML = "";
  if (placeholder) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = placeholder;
    select.appendChild(opt);
  }
  for (const o of options) {
    const opt = document.createElement("option");
    opt.value = o.value;
    opt.textContent = o.label;
    select.appendChild(opt);
  }
}

function renderTemplates() {
  const tplSel = byId("meeting_template");
  const tplList = byId("templates_list");

  const templates = alive(db.templates);
  renderSelectOptions(tplSel, templates.map(t => ({ value:t.id, label:t.name })));

  tplList.innerHTML = templates.map(t => `
    <div class="item">
      <div class="item__top">
        <div><strong>${escapeHtml(t.name)}</strong></div>
        <div class="badges">
          <span class="badge">${escapeHtml(t.id)}</span>
        </div>
      </div>
      <div class="item__meta">
        ${t.sections.map(s => `<span class="badge badge--accent">${escapeHtml(s.label)}</span>`).join("")}
      </div>
    </div>
  `).join("");
}

function renderTopics() {
  const topicSel = byId("meeting_topic");
  const topicsSel = byId("topics_topic");

  const topics = alive(db.topics).sort((a,b)=>a.name.localeCompare(b.name));
  const opts = topics.map(t => ({ value:t.id, label:t.name }));

  renderSelectOptions(topicSel, opts, { placeholder: topics.length ? null : "No topics yet — add one" });
  renderSelectOptions(topicsSel, opts, { placeholder: topics.length ? "Choose a topic…" : "No topics yet" });
}

function renderPeopleSelects() {
  const updatesSel = byId("updates_person");
  if (!updatesSel) return;
  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));
  renderSelectOptions(updatesSel, people.map(p => ({ value:p.id, label:p.name })), { placeholder: people.length ? "Choose a person…" : "No people yet" });
}

/**
 * Populates the 1:1 counterpart selects in the meeting lightbox.
 */
function renderMeetingCounterpartSelects() {
  const personSel = byId("meeting_one_to_one_person");
  const groupSel = byId("meeting_one_to_one_group");
  if (!personSel || !groupSel) return;

  const currentPerson = personSel.value;
  const currentGroup = groupSel.value;
  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));
  const groups = alive(db.groups).sort((a,b)=>a.name.localeCompare(b.name));

  renderSelectOptions(
    personSel,
    people.map(p => ({ value: p.id, label: p.name })),
    { placeholder: people.length ? "Choose a person…" : "No people yet" }
  );
  renderSelectOptions(
    groupSel,
    groups.map(g => ({ value: g.id, label: g.name })),
    { placeholder: groups.length ? "Choose a group…" : "No groups yet" }
  );
  if (currentPerson) personSel.value = currentPerson;
  if (currentGroup) groupSel.value = currentGroup;
}

function renderPeopleManager() {
  const list = byId("people_list");
  const editor = byId("person_editor");
  const ownedList = byId("person_owned_updates");
  const targetList = byId("person_target_updates");
  if (!list || !editor || !ownedList || !targetList) return;

  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));
  if (!personEditorState.draft && personViewId) {
    const selected = getPerson(personViewId);
    if (selected) {
      personEditorState = { isNew: false, draft: createPersonDraft(selected), error: "" };
    }
  }

  list.innerHTML = people.map(p => {
    const meta = [
      p.email ? escapeHtml(p.email) : null,
      p.organisation ? escapeHtml(p.organisation) : null,
      p.jobTitle ? escapeHtml(p.jobTitle) : null,
    ].filter(Boolean).join(" • ");
    return `
      <div class="item ${p.id === personViewId ? "item--selected" : ""}">
        <div class="item__top">
          <div>
            <strong>${escapeHtml(p.name)}</strong>
            <div class="muted">${meta || "No details added yet."}</div>
          </div>
          <div class="badges"><span class="badge">${escapeHtml(p.id)}</span></div>
        </div>
        <div class="item__actions">
          <button class="smallbtn" data-person-select="${escapeHtml(p.id)}">View</button>
          <button class="smallbtn smallbtn--danger" data-del-person="${escapeHtml(p.id)}">Delete</button>
        </div>
      </div>
    `;
  }).join("") || `<div class="muted">No people yet. Create one to get started.</div>`;

  list.querySelectorAll("[data-person-select]").forEach(btn => {
    btn.addEventListener("click", () => {
      const id = btn.getAttribute("data-person-select");
      personViewId = id;
      const person = getPerson(id);
      personEditorState = { isNew: false, draft: createPersonDraft(person), error: "" };
      renderPeopleManager();
    });
  });

  list.querySelectorAll("[data-del-person]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const id = btn.getAttribute("data-del-person");
      const p = getPerson(id);
      if (!p) return;
      if (!confirm(`Delete ${p.name}? This won’t erase history but removes them from lists.`)) return;
      p.deleted = true;
      p.updatedAt = nowIso();
      if (personViewId === id) {
        personViewId = null;
        personEditorState = { isNew: false, draft: null, error: "" };
      }
      markDirty();
      await saveLocal();
      renderAll();
    });
  });

  if (!personEditorState.draft) {
    editor.innerHTML = `<div class="muted">Select a person to view or edit their details.</div>`;
    ownedList.innerHTML = `<div class="muted">Select a person to see owned updates.</div>`;
    targetList.innerHTML = `<div class="muted">Select a person to see update targets.</div>`;
    return;
  }

  const draft = personEditorState.draft;
  editor.innerHTML = `
    <div class="formrow">
      <label>Name ${fieldTag(true)}</label>
      <input id="person_name" type="text" value="${escapeHtml(draft.name)}" />
    </div>
    <div class="formrow">
      <label>Email ${fieldTag(true)}</label>
      <input id="person_email" type="email" value="${escapeHtml(draft.email)}" />
    </div>
    <div class="formrow">
      <label>Organisation ${fieldTag(true)}</label>
      <input id="person_org" type="text" value="${escapeHtml(draft.organisation)}" />
    </div>
    <div class="formrow">
      <label>Job title ${fieldTag(false)}</label>
      <input id="person_title" type="text" value="${escapeHtml(draft.jobTitle)}" />
    </div>
    ${personEditorState.error ? `<div class="item__error">${escapeHtml(personEditorState.error)}</div>` : ""}
  `;

  if (personEditorState.isNew) {
    ownedList.innerHTML = `<div class="muted">Save the person to see owned updates.</div>`;
    targetList.innerHTML = `<div class="muted">Save the person to see update targets.</div>`;
    return;
  }

  const activePerson = personViewId ? getPerson(personViewId) : null;
  const owned = activePerson ? alive(db.items).filter(it => it.ownerId === activePerson.id) : [];
  owned.sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));

  const targetItems = activePerson
    ? alive(db.items).filter(it => (it.updateTargets || []).includes(activePerson.id))
    : [];
  const pendingTargets = targetItems.filter(it => !it.updateStatus?.[activePerson.id]?.updated);
  const completedTargets = targetItems.filter(it => it.updateStatus?.[activePerson.id]?.updated);
  pendingTargets.sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));
  completedTargets.sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));

  ownedList.innerHTML = owned.map(it => renderItemCard(it)).join("") || `<div class="muted">No owned updates yet.</div>`;
  targetList.innerHTML = `
    <div class="sectioncard">
      <div class="sectionhead">
        <h3>Pending</h3>
        <div class="muted">${pendingTargets.length} item(s)</div>
      </div>
      <div class="sectionbox sectionbox--compact">
        <div class="list">
          ${pendingTargets.map(it => renderItemCard(it)).join("") || `<div class="muted">Nothing pending.</div>`}
        </div>
      </div>
    </div>
    <div class="sectioncard">
      <div class="sectionhead">
        <h3>Updated</h3>
        <div class="muted">${completedTargets.length} item(s)</div>
      </div>
      <div class="sectionbox sectionbox--compact">
        <div class="list">
          ${completedTargets.map(it => renderItemCard(it)).join("") || `<div class="muted">No updates marked yet.</div>`}
        </div>
      </div>
    </div>
  `;

  wireItemButtons(ownedList);
  wireItemButtons(targetList);
}

function getDefaultActionsOwnerId() {
  const name = db.settings?.defaultOwnerName?.trim() || "";
  if (!name) return "";
  const match = alive(db.people).find(p => p.name.toLowerCase() === name.toLowerCase());
  return match ? match.id : "";
}

function renderActionsFiltersOptions() {
  const ownerSel = byId("actions_owner");
  const topicSel = byId("actions_topic");

  if (!ownerSel || !topicSel) return;

  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));
  renderSelectOptions(
    ownerSel,
    people.map(p => ({ value:p.id, label:p.name })),
    { placeholder: "All people" }
  );

  const topics = alive(db.topics).sort((a,b)=>a.name.localeCompare(b.name));
  renderSelectOptions(
    topicSel,
    topics.map(t => ({ value:t.id, label:t.name })),
    { placeholder: "All topics" }
  );
}

function renderGroups() {
  const list = byId("groups_list");
  const groups = alive(db.groups).sort((a,b)=>a.name.localeCompare(b.name));
  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));

  list.innerHTML = groups.map(g => {
    const members = (g.memberIds || []).map(pid => getPerson(pid)?.name).filter(Boolean);
    return `
      <div class="item">
        <div class="item__top">
          <div>
            <strong>${escapeHtml(g.name)}</strong>
            <div class="muted">${members.length ? escapeHtml(members.join(", ")) : "No members yet"}</div>
          </div>
          <div class="badges"><span class="badge">${escapeHtml(g.id)}</span></div>
        </div>

        <div class="picker" style="margin-top:10px">
          <div class="pickcol">
            <h4>Members</h4>
            <div class="picklist">
              ${people.map(p => {
                const checked = (g.memberIds || []).includes(p.id) ? "checked" : "";
                return `
                  <label class="checkline">
                    <input type="checkbox" data-group="${escapeHtml(g.id)}" data-member="${escapeHtml(p.id)}" ${checked} />
                    ${escapeHtml(p.name)}
                  </label>
                `;
              }).join("")}
            </div>
          </div>
        </div>

        <div class="item__actions">
          <button class="smallbtn smallbtn--danger" data-del-group="${escapeHtml(g.id)}">Delete group</button>
        </div>
      </div>
    `;
  }).join("");

  list.querySelectorAll("input[type=checkbox][data-group]").forEach(cb => {
    cb.addEventListener("change", async () => {
      const gid = cb.getAttribute("data-group");
      const pid = cb.getAttribute("data-member");
      const g = getGroup(gid);
      if (!g) return;
      g.memberIds = g.memberIds || [];
      if (cb.checked) {
        if (!g.memberIds.includes(pid)) g.memberIds.push(pid);
      } else {
        g.memberIds = g.memberIds.filter(x => x !== pid);
      }
      g.updatedAt = nowIso();
      markDirty();
      await saveLocal();
      renderAll();
    });
  });

  list.querySelectorAll("[data-del-group]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const gid = btn.getAttribute("data-del-group");
      const g = getGroup(gid);
      if (!g) return;
      if (!confirm(`Delete group "${g.name}"?`)) return;
      g.deleted = true;
      g.updatedAt = nowIso();
      markDirty();
      await saveLocal();
      renderAll();
    });
  });
}

function renderCurrentMeetingHeader() {
  const label = byId("current_meeting_label");
  const area = byId("meeting_work_area");

  const meeting = currentMeetingId ? getMeeting(currentMeetingId) : null;
  if (!meeting) {
    label.textContent = "None selected.";
    area.innerHTML = `<h2>Meeting notes</h2><div class="muted">Create or open a meeting to start taking notes.</div>`;
    return;
  }

  const topic = getTopic(meeting.topicId);
  const tpl = getTemplate(meeting.templateId);
  const oneToOneContext = tpl?.id === ONE_TO_ONE_TEMPLATE_ID ? getOneToOneContext(meeting) : null;
  const oneToOneLabel = oneToOneContext ? ` • 1:1 with ${escapeHtml(oneToOneContext.label)}` : "";

  label.innerHTML = `
    <div><strong>${escapeHtml(meeting.title || "(Untitled meeting)")}</strong></div>
    <div class="muted">${escapeHtml(tpl?.name || "Template")} • ${escapeHtml(topic?.name || "No topic")} • ${fmtDateTime(meeting.datetime)}${oneToOneLabel}</div>
  `;

  const oneToOneUpdatesCard = oneToOneContext ? renderOneToOneUpdatesCard(meeting, oneToOneContext) : "";

  area.innerHTML = `
    <h2>Meeting notes</h2>
    <div class="muted">Template: <strong>${escapeHtml(tpl?.name || "")}</strong> • Topic: <strong>${escapeHtml(topic?.name || "")}</strong></div>
    ${oneToOneUpdatesCard}
    <div class="sectioncard scratchpad-card">
      <div class="sectionhead">
        <h3>Meeting scratchpad</h3>
        <div class="muted">Capture rich text during the meeting, then move items into the structured sections below.</div>
      </div>
      <div
        class="scratchpad-field"
        data-scratchpad
        contenteditable="true"
        role="textbox"
        aria-multiline="true"
        data-placeholder="Type meeting notes here..."
      ></div>
    </div>
    <div id="sections_container"></div>
  `;

  renderMeetingSections(meeting, tpl);
  wireMeetingScratchpad(meeting);
  wireOneToOneUpdatesSection(meeting);
}

function wireMeetingScratchpad(meeting) {
  const field = document.querySelector("[data-scratchpad]");
  if (!field) return;
  field.innerHTML = meeting.scratchpadHtml || "";

  const persistScratchpad = debounce(() => {
    if (!currentMeetingId || meeting.id !== currentMeetingId) return;
    meeting.scratchpadHtml = field.innerHTML;
    meeting.updatedAt = nowIso();
    markDirty();
    saveLocal();
  }, 250);

  field.addEventListener("input", persistScratchpad);
  field.addEventListener("blur", () => {
    persistScratchpad();
  });
}

/**
 * Builds the 1:1 updates card markup for the meeting notes view.
 * @param {object} meeting Meeting record.
 * @param {object} context 1:1 context with label and person ids.
 * @returns {string} HTML string for the updates card.
 */
function renderOneToOneUpdatesCard(meeting, context) {
  const items = getItemsLinkedToPeople(context.personIds)
    .sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));
  const memberNames = context.personIds.map(pid => getPerson(pid)?.name).filter(Boolean);
  const subtitle = context.type === "group"
    ? `${escapeHtml(context.label)} • ${escapeHtml(memberNames.join(", ") || "No members yet")}`
    : escapeHtml(context.label);
  const listHtml = items.map(it => renderItemCard(it)).join("")
    || `<div class="muted">No linked items yet.</div>`;
  const itemIds = items.map(it => it.id);
  const pendingCount = items.filter(it => (it.updateTargets || []).some(pid => context.personIds.includes(pid) && !it.updateStatus?.[pid]?.updated)).length;

  return `
    <div class="sectioncard" data-one-to-one-section data-target-ids='${escapeHtml(JSON.stringify(context.personIds))}' data-item-ids='${escapeHtml(JSON.stringify(itemIds))}'>
      <div class="sectionhead">
        <div>
          <h3>1:1 updates</h3>
          <div class="muted">${subtitle}</div>
        </div>
        <div class="muted">${items.length} linked item(s) • ${pendingCount} pending update(s)</div>
      </div>
      <div class="sectionbox sectionbox--compact">
        <div class="list">
          ${listHtml}
        </div>
      </div>
      <div class="row row--space">
        <div class="muted">Marks updates for ${escapeHtml(context.label)} across all linked items listed above.</div>
        <button class="btn btn--primary" type="button" data-one-to-one-mark ${itemIds.length ? "" : "disabled"}>Mark all listed as updated</button>
      </div>
    </div>
  `;
}

/**
 * Wires the 1:1 updates card button in the meeting notes view.
 * @param {object} meeting Active meeting record for stamp metadata.
 */
function wireOneToOneUpdatesSection(meeting) {
  const section = document.querySelector("[data-one-to-one-section]");
  if (!section) return;
  const button = section.querySelector("[data-one-to-one-mark]");
  if (!button) return;

  button.addEventListener("click", async () => {
    const targetIds = JSON.parse(section.dataset.targetIds || "[]");
    const itemIds = JSON.parse(section.dataset.itemIds || "[]");
    if (!targetIds.length || !itemIds.length) {
      alert("No linked items to mark as updated.");
      return;
    }

    const targetLabel = targetIds.map(pid => getPerson(pid)?.name).filter(Boolean).join(", ");
    if (!confirm(`Mark ${itemIds.length} item(s) as updated for ${targetLabel || "this counterpart"}?`)) return;

    const stamp = nowIso();

    for (const id of itemIds) {
      const it = getItem(id);
      if (!it) continue;
      it.updateTargets = it.updateTargets || [];
      it.updateStatus = it.updateStatus || {};
      for (const pid of targetIds) {
        if (!it.updateTargets.includes(pid)) it.updateTargets.push(pid);
        it.updateStatus[pid] = {
          updated: true,
          updatedAt: stamp,
          meetingId: meeting?.id || null
        };
      }
      it.updatedAt = stamp;
    }

    markDirty();
    await saveLocal();
    renderAll();
  });
}

function renderMeetingSections(meeting, tpl) {
  const container = byId("sections_container");
  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));
  const groups = alive(db.groups).sort((a,b)=>a.name.localeCompare(b.name));

  const sections = tpl?.sections || [];
  const items = alive(db.items).filter(i => i.meetingId === meeting.id);

  container.innerHTML = `
    <div class="sectioncard sectioncard--entry">
      <div class="sectionhead">
        <h3>Add a note</h3>
        <div class="muted">Choose a type to set the required fields.</div>
      </div>
      <div class="sectionbox sectionbox--compact field-table" data-entry-form>
        <div class="section-form">
          <div class="section-form__col">
            <div class="formrow">
              <label>Type ${fieldTag(true)}</label>
              <select data-field="section">
                ${sections.map(sec => `<option value="${escapeHtml(sec.key)}">${escapeHtml(sec.label)}</option>`).join("")}
              </select>
            </div>

            <div class="formrow">
              <label>Text ${fieldTag(true)}</label>
              <textarea data-field="text" placeholder="Type quickly…"></textarea>
            </div>

            <div class="grid2">
              <div class="formrow">
                <label>Owner ${fieldTag(false, "ownerId")}</label>
                <input data-field="ownerName" list="owner_list_entry" type="text" placeholder="Type to search…" ${people.length ? "" : "disabled"} />
                <datalist id="owner_list_entry">
                  ${people.map(p => `<option value="${escapeHtml(p.name)}"></option>`).join("")}
                </datalist>
              </div>

              <div class="formrow">
                <label>Status ${fieldTag(false, "status")}</label>
                <select data-field="status">
                  <option value="">— None —</option>
                  <option value="open">Open</option>
                  <option value="in_progress">In progress</option>
                  <option value="blocked">Blocked</option>
                  <option value="done">Done</option>
                </select>
              </div>
            </div>

            <div class="grid2">
              <div class="formrow">
                <label>Due date</label>
                <input data-field="dueDate" type="date" />
              </div>

              <div class="formrow">
                <label>Link</label>
                <input data-field="link" type="url" placeholder="https://…" />
              </div>
            </div>
          </div>

          <div class="section-form__col">
            <div class="formrow">
              <label>People to update ${fieldTag(false, "updateTargets")}</label>
              <div class="update-targets" data-people-picker>
                <div class="people-select__controls">
                  <input data-people-input type="text" list="people_list_entry" placeholder="Type a name to add…" ${people.length ? "" : "disabled"} />
                  <datalist id="people_list_entry">
                    ${people.map(p => `<option value="${escapeHtml(p.name)}"></option>`).join("")}
                  </datalist>
                  <button class="btn btn--ghost" type="button" data-add-person ${people.length ? "" : "disabled"}>Add person</button>
                </div>
                <div class="people-select__selected" data-selected-list>
                  <div class="muted" data-empty="true">No people selected yet.</div>
                </div>
                <div class="update-targets__groups">
                  <div class="muted">Groups</div>
                  <div class="picklist">
                    ${groups.length ? groups.map(g => `
                      <label class="checkline">
                        <input type="checkbox" data-target-group="${escapeHtml(g.id)}" />
                        ${escapeHtml(g.name)}
                      </label>
                    `).join("") : `<div class="muted">No groups yet.</div>`}
                  </div>
                </div>
              </div>
            </div>

            <div class="row row--space">
              <div class="muted">Tip: use groups for “team”, “stakeholders”, etc.</div>
              <button class="btn btn--primary" data-add-note>Add note</button>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div class="sections-grid sections-grid--list">
      ${sections.map(sec => {
        const secItems = items.filter(i => i.section === sec.key);
        return `
          <section class="sectioncard">
            <div class="sectionhead">
              <h3>${escapeHtml(sec.label)}</h3>
              <div class="muted">${secItems.length} item(s)</div>
            </div>
            <div class="sectionbox sectionbox--compact">
              <div class="list">
                ${secItems.map(it => renderItemCard(it)).join("") || `<div class="muted">Nothing yet.</div>`}
              </div>
            </div>
          </section>
        `;
      }).join("")}
    </div>
  `;

  wirePeoplePickers(container, people);

  const entryForm = container.querySelector("[data-entry-form]");
  const typeSelect = entryForm?.querySelector("select[data-field=section]");
  if (entryForm && typeSelect) {
    const updateEntryRequirements = () => {
      const sectionKey = typeSelect.value;
      const secDef = (tpl?.sections || []).find(s => s.key === sectionKey);
      const requires = new Set(secDef?.requires || []);
      const requiredMap = {
        ownerId: requires.has("ownerId"),
        status: requires.has("status"),
        updateTargets: requires.has("updateTargets"),
      };

      Object.entries(requiredMap).forEach(([key, required]) => {
        const tag = entryForm.querySelector(`[data-required-tag="${key}"]`);
        if (tag) {
          tag.classList.toggle("is-hidden", !required);
        }
      });
    };

    updateEntryRequirements();
    typeSelect.addEventListener("change", updateEntryRequirements);
  }

  // wire add buttons & item actions
  container.querySelectorAll("[data-add-note]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const box = btn.closest(".sectionbox");
      const sectionKey = box.querySelector("select[data-field=section]")?.value || "";
      const text = box.querySelector("textarea[data-field=text]").value.trim();
      const ownerName = box.querySelector("input[data-field=ownerName]")?.value.trim() || "";
      const ownerMatch = ownerName ? findPersonByName(ownerName, people) : null;
      const ownerId = ownerMatch ? ownerMatch.id : null;
      const status = box.querySelector("select[data-field=status]").value || null;
      const dueDate = box.querySelector("input[data-field=dueDate]").value || null;
      const link = box.querySelector("input[data-field=link]").value.trim() || null;

      // selected targets
      const selectedPeople = Array.from(box.querySelectorAll("[data-selected-person]"))
        .map(x => x.getAttribute("data-selected-person"));
      const selectedGroups = Array.from(box.querySelectorAll("input[data-target-group]:checked"))
        .map(x => x.getAttribute("data-target-group"));
      const expandedTargets = expandTargets(selectedPeople, selectedGroups);

      const secDef = (tpl?.sections || []).find(s => s.key === sectionKey);
      const req = new Set(secDef?.requires || []);
      const errs = [];

      if (!text) errs.push("Text is required.");
      if (ownerName && !ownerMatch) errs.push("Owner must match a person from the list.");
      if (req.has("ownerId") && !ownerId) errs.push("Owner is required for this section.");
      if (req.has("status") && !status) errs.push("Status is required for this section.");
      if (req.has("updateTargets") && expandedTargets.length === 0) errs.push("At least one update target is required for this section.");

      if (errs.length) {
        alert(errs.join("\n"));
        return;
      }

      const item = {
        id: uid("item"),
        kind: "item",
        meetingId: meeting.id,
        topicId: meeting.topicId,
        section: sectionKey,
        text,
        ownerId,
        status,
        dueDate,
        link,
        updateTargets: expandedTargets,
        updateStatus: expandedTargets.reduce((acc, pid) => {
          acc[pid] = { updated: false };
          return acc;
        }, {}),
        createdAt: nowIso(),
        updatedAt: nowIso(),
      };

      db.items.push(item);
      markDirty();
      await saveLocal();

      // clear inputs
      box.querySelector("textarea[data-field=text]").value = "";
      box.querySelector("input[data-field=ownerName]").value = "";
      box.querySelector("select[data-field=status]").value = "";
      box.querySelector("input[data-field=dueDate]").value = "";
      box.querySelector("input[data-field=link]").value = "";
      box.querySelectorAll("input[type=checkbox]").forEach(c => c.checked = false);
      box.querySelectorAll("[data-selected-person]").forEach(pill => pill.remove());
      box.querySelectorAll("[data-selected-list]").forEach(list => ensurePeopleEmptyState(list));

      renderAll();
    });
  });

  wireItemButtons(container);
}

function renderItemCard(it) {
  const meeting = getMeeting(it.meetingId);
  const topic = getTopic(it.topicId);
  const owner = it.ownerId ? getPerson(it.ownerId) : null;
  const editState = itemEditState.get(it.id);
  const isEditing = !!editState;
  const tpl = meeting ? getTemplate(meeting.templateId) : null;
  const secDef = (tpl?.sections || []).find(s => s.key === it.section);
  const requires = new Set(secDef?.requires || []);
  const ownerRequired = requires.has("ownerId");
  const statusRequired = requires.has("status");
  const sectionLabel = secDef?.label || it.section;

  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));

  const draft = editState?.draft || {
    text: it.text || "",
    status: it.status || "",
    dueDate: it.dueDate || "",
    link: it.link || "",
    ownerName: owner?.name || "",
  };
  const editError = editState?.error || "";

  const status = it.status || "";
  const statusBadge = status
    ? (status === "done" ? `badge--ok` : status === "blocked" ? `badge--danger` : `badge--warn`)
    : "";

  const targets = (it.updateTargets || []).map(pid => getPerson(pid)?.name).filter(Boolean);
  const updatedCount = Object.values(it.updateStatus || {}).filter(s => s.updated).length;
  const totalTargets = targets.length;

  const updBadge = totalTargets
    ? (updatedCount === totalTargets ? `<span class="badge badge--ok">All updated</span>` : `<span class="badge badge--warn">${updatedCount}/${totalTargets} updated</span>`)
    : "";

  const linkBadge = it.link ? `<span class="badge badge--accent">Link</span>` : "";

  if (isEditing) {
    return `
      <div class="item item--editing" data-item="${escapeHtml(it.id)}" data-editing="true">
        <div class="item__top">
          <div class="badges">
            <span class="badge badge--accent">${escapeHtml(sectionLabel)}</span>
            ${status ? `<span class="badge ${statusBadge}">${escapeHtml(status.replace("_"," "))}</span>` : ""}
            ${it.dueDate ? `<span class="badge">${escapeHtml(it.dueDate)}</span>` : ""}
            ${updBadge}
            ${linkBadge}
          </div>
          <div class="muted">${meeting ? fmtDateTime(meeting.datetime) : ""}</div>
        </div>

        <div class="item__edit">
          <div class="formrow">
            <label>Text ${fieldTag(true)}</label>
            <textarea data-edit-field="text" placeholder="Update text…">${escapeHtml(draft.text)}</textarea>
          </div>

          <div class="grid2">
            <div class="formrow">
              <label>Owner ${fieldTag(ownerRequired)}</label>
              <input data-edit-field="ownerName" list="owner_edit_${escapeHtml(it.id)}" type="text" placeholder="Type to search…" value="${escapeHtml(draft.ownerName)}" ${people.length ? "" : "disabled"} />
              <datalist id="owner_edit_${escapeHtml(it.id)}">
                ${people.map(p => `<option value="${escapeHtml(p.name)}"></option>`).join("")}
              </datalist>
            </div>

            <div class="formrow">
              <label>Status ${fieldTag(statusRequired)}</label>
              <select data-edit-field="status">
                <option value="">— None —</option>
                <option value="open"${draft.status === "open" ? " selected" : ""}>Open</option>
                <option value="in_progress"${draft.status === "in_progress" ? " selected" : ""}>In progress</option>
                <option value="blocked"${draft.status === "blocked" ? " selected" : ""}>Blocked</option>
                <option value="done"${draft.status === "done" ? " selected" : ""}>Done</option>
              </select>
            </div>
          </div>

          <div class="grid2">
            <div class="formrow">
              <label>Due date ${fieldTag(false)}</label>
              <input data-edit-field="dueDate" type="date" value="${escapeHtml(draft.dueDate)}" />
            </div>

            <div class="formrow">
              <label>Link ${fieldTag(false)}</label>
              <input data-edit-field="link" type="url" placeholder="https://…" value="${escapeHtml(draft.link)}" />
            </div>
          </div>
          ${editError ? `<div class="item__error">${escapeHtml(editError)}</div>` : ""}
        </div>

        <div class="item__actions">
          <button class="smallbtn" data-save-item="${escapeHtml(it.id)}">Save</button>
          <button class="smallbtn" data-cancel-item="${escapeHtml(it.id)}">Cancel</button>
        </div>
      </div>
    `;
  }

  return `
    <div class="item" data-item="${escapeHtml(it.id)}">
      <div class="item__top">
        <div class="badges">
          <span class="badge badge--accent">${escapeHtml(sectionLabel)}</span>
          ${status ? `<span class="badge ${statusBadge}">${escapeHtml(status.replace("_"," "))}</span>` : ""}
          ${it.dueDate ? `<span class="badge">${escapeHtml(it.dueDate)}</span>` : ""}
          ${updBadge}
          ${linkBadge}
        </div>
        <div class="muted">${meeting ? fmtDateTime(meeting.datetime) : ""}</div>
      </div>

      <div class="item__text">${escapeHtml(it.text)}</div>

      <div class="item__meta">
        ${topic ? `<span>Topic: <strong>${escapeHtml(topic.name)}</strong></span>` : ""}
        ${owner ? `<span>Owner: <strong>${escapeHtml(owner.name)}</strong></span>` : ""}
        ${targets.length ? `<span>Update: <strong>${escapeHtml(targets.join(", "))}</strong></span>` : ""}
      </div>

      <div class="item__actions">
        <button class="smallbtn" data-edit-item="${escapeHtml(it.id)}">Edit</button>
        <button class="smallbtn" data-open-meeting="${escapeHtml(it.meetingId)}">Open meeting</button>
        ${it.link ? `<button class="smallbtn" data-open-link="${escapeHtml(it.id)}">Open link</button>` : ""}
        <button class="smallbtn smallbtn--danger" data-del-item="${escapeHtml(it.id)}">Delete</button>
      </div>
    </div>
  `;
}

function wireItemButtons(rootEl) {
  rootEl.querySelectorAll("[data-edit-item]").forEach(btn => {
    btn.addEventListener("click", () => {
      const id = btn.getAttribute("data-edit-item");
      const it = getItem(id);
      if (!it) return;
      itemEditState.set(id, {
        draft: {
          text: it.text || "",
          status: it.status || "",
          dueDate: it.dueDate || "",
          link: it.link || "",
          ownerName: getPerson(it.ownerId)?.name || "",
        },
        error: "",
      });
      renderAll();
    });
  });

  rootEl.querySelectorAll("[data-save-item]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const id = btn.getAttribute("data-save-item");
      const it = getItem(id);
      if (!it) return;
      const card = btn.closest(".item");
      if (!card) return;

      const text = card.querySelector("[data-edit-field=text]")?.value.trim() || "";
      const ownerName = card.querySelector("[data-edit-field=ownerName]")?.value.trim() || "";
      const ownerMatch = ownerName ? findPersonByName(ownerName) : null;
      const ownerId = ownerMatch ? ownerMatch.id : null;
      const status = card.querySelector("[data-edit-field=status]")?.value || null;
      const dueDate = card.querySelector("[data-edit-field=dueDate]")?.value || null;
      const link = card.querySelector("[data-edit-field=link]")?.value.trim() || null;

      const meeting = getMeeting(it.meetingId);
      const tpl = meeting ? getTemplate(meeting.templateId) : null;
      const secDef = (tpl?.sections || []).find(s => s.key === it.section);
      const req = new Set(secDef?.requires || []);
      const errs = [];

      if (!text) errs.push("Text is required.");
      if (ownerName && !ownerMatch) errs.push("Owner must match a person from the list.");
      if (req.has("ownerId") && !ownerId) errs.push("Owner is required for this section.");
      if (req.has("status") && !status) errs.push("Status is required for this section.");
      if (req.has("updateTargets") && (it.updateTargets || []).length === 0) {
        errs.push("At least one update target is required for this section.");
      }

      if (errs.length) {
        itemEditState.set(id, {
          draft: {
            text,
            status: status || "",
            dueDate: dueDate || "",
            link: link || "",
            ownerName: ownerName || "",
          },
          error: errs.join(" "),
        });
        renderAll();
        return;
      }

      it.text = text;
      it.ownerId = ownerId;
      it.status = status;
      it.dueDate = dueDate;
      it.link = link;
      it.updatedAt = nowIso();

      itemEditState.delete(id);
      markDirty();
      await saveLocal();
      renderAll();
    });
  });

  rootEl.querySelectorAll("[data-cancel-item]").forEach(btn => {
    btn.addEventListener("click", () => {
      const id = btn.getAttribute("data-cancel-item");
      itemEditState.delete(id);
      renderAll();
    });
  });

  rootEl.querySelectorAll("[data-del-item]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const id = btn.getAttribute("data-del-item");
      const it = getItem(id);
      if (!it) return;
      if (!confirm("Delete this item?")) return;
      it.deleted = true;
      it.updatedAt = nowIso();
      itemEditState.delete(id);
      markDirty();
      await saveLocal();
      renderAll();
    });
  });

  rootEl.querySelectorAll("[data-open-meeting]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const mid = btn.getAttribute("data-open-meeting");
      const m = getMeeting(mid);
      if (!m) return;
      currentMeetingId = mid;
      await saveMeta();
      setMeetingView("notes");
      setActiveModule("meetings");
      setMeetingModuleTab("meeting");
      renderAll();
    });
  });

  rootEl.querySelectorAll("[data-meeting-notes]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const mid = btn.getAttribute("data-meeting-notes");
      const m = getMeeting(mid);
      if (!m) return;
      currentMeetingId = mid;
      await saveMeta();
      setMeetingView("notes");
      setActiveModule("meetings");
      setMeetingModuleTab("meeting");
      renderAll();
    });
  });

  rootEl.querySelectorAll("[data-meeting-edit]").forEach(btn => {
    btn.addEventListener("click", () => {
      const mid = btn.getAttribute("data-meeting-edit");
      openMeetingEditLightbox(mid);
    });
  });

  rootEl.querySelectorAll("[data-open-link]").forEach(btn => {
    btn.addEventListener("click", () => {
      const id = btn.getAttribute("data-open-link");
      const it = getItem(id);
      if (!it?.link) return;
      window.open(it.link, "_blank", "noopener,noreferrer");
    });
  });
}

function renderQuickSearch() {
  const q = byId("quick_search").value.trim().toLowerCase();
  const out = byId("quick_search_results");
  if (!q) {
    out.innerHTML = `<div class="muted">Type to search across all items.</div>`;
    return;
  }
  const matches = alive(db.items)
    .filter(it => (it.text || "").toLowerCase().includes(q))
    .sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt))
    .slice(0, 20);

  out.innerHTML = matches.map(it => renderItemCard(it)).join("") || `<div class="muted">No matches.</div>`;
  wireItemButtons(out);
}

function renderMeetingCalendar() {
  const calendar = byId("meeting_calendar");
  const rangeEl = byId("calendar_range");
  if (!calendar || !rangeEl) return;

  const view = meetingCalendarView || "week";
  const anchor = meetingCalendarAnchor || new Date();

  document.querySelectorAll("[data-calendar-view]").forEach(btn => {
    btn.classList.toggle("is-active", btn.getAttribute("data-calendar-view") === view);
  });

  const meetings = alive(db.meetings)
    .filter(m => m.datetime)
    .sort((a,b)=>Date.parse(a.datetime)-Date.parse(b.datetime));

  const meetingsByDay = new Map();
  meetings.forEach(m => {
    const key = dateKeyFromIso(m.datetime);
    if (!key) return;
    if (!meetingsByDay.has(key)) meetingsByDay.set(key, []);
    meetingsByDay.get(key).push(m);
  });

  meetingsByDay.forEach(list => {
    list.sort((a,b)=>Date.parse(a.datetime)-Date.parse(b.datetime));
  });

  const weekdayFormatter = new Intl.DateTimeFormat(undefined, { weekday: "short" });
  const dayFormatter = new Intl.DateTimeFormat(undefined, { day: "numeric" });
  const monthFormatter = new Intl.DateTimeFormat(undefined, { month: "long", year: "numeric" });
  const rangeFormatter = new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric" });

  const renderDay = (d, isOutside = false) => {
    const key = dateKeyFromDate(d);
    const items = meetingsByDay.get(key) || [];
    const list = items.map(m => {
      const topic = getTopic(m.topicId)?.name || "No topic";
      const title = m.title || "Untitled meeting";
      const time = formatTime(m.datetime) || "Time TBD";
      return `
        <div class="calendar-meeting">
          <div class="calendar-meeting__title">${escapeHtml(title)}</div>
          <div class="calendar-meeting__meta">${escapeHtml(topic)} • ${escapeHtml(time)}</div>
          <div class="calendar-meeting__actions">
            <button class="smallbtn" type="button" data-meeting-edit="${escapeHtml(m.id)}">Edit</button>
            <button class="smallbtn smallbtn--primary" type="button" data-meeting-notes="${escapeHtml(m.id)}">Notes</button>
          </div>
        </div>
      `;
    }).join("");

    return `
      <div class="calendar-day ${isOutside ? "is-outside" : ""}">
        <div class="calendar-day__header">
          <span>${weekdayFormatter.format(d)}</span>
          <span class="calendar-day__date">${dayFormatter.format(d)}</span>
        </div>
        ${list || `<div class="calendar-day__empty">No meetings</div>`}
      </div>
    `;
  };

  if (view === "month") {
    const monthStart = startOfMonth(anchor);
    const gridStart = addDays(monthStart, -monthStart.getDay());
    const days = Array.from({ length: 42 }, (_, i) => addDays(gridStart, i));
    rangeEl.textContent = monthFormatter.format(monthStart);
    calendar.innerHTML = `
      <div class="calendar-grid calendar-grid--month">
        ${days.map(d => renderDay(d, d.getMonth() !== monthStart.getMonth())).join("")}
      </div>
    `;
  } else {
    const weekStart = startOfWeek(anchor);
    const days = Array.from({ length: 5 }, (_, i) => addDays(weekStart, i));
    const startLabel = rangeFormatter.format(days[0]);
    const endLabel = rangeFormatter.format(days[days.length - 1]);
    rangeEl.textContent = `Week of ${startLabel} – ${endLabel}`;
    calendar.innerHTML = `
      <div class="calendar-grid calendar-grid--week">
        ${days.map(d => renderDay(d)).join("")}
      </div>
    `;
  }

  wireItemButtons(calendar);
}

function setMeetingCalendarView(view) {
  meetingCalendarView = view;
  saveMeta().catch(console.error);
  renderMeetingCalendar();
}

function shiftMeetingCalendar(direction) {
  if (meetingCalendarView === "month") {
    const anchor = new Date(meetingCalendarAnchor);
    anchor.setDate(1);
    anchor.setMonth(anchor.getMonth() + direction);
    meetingCalendarAnchor = anchor;
  } else {
    meetingCalendarAnchor = addDays(meetingCalendarAnchor, direction * 7);
  }
  saveMeta().catch(console.error);
  renderMeetingCalendar();
}

function resetMeetingCalendarToToday() {
  meetingCalendarAnchor = new Date();
  saveMeta().catch(console.error);
  renderMeetingCalendar();
}

function renderUpdates() {
  const personId = byId("updates_person").value || "";
  const filter = byId("updates_filter").value.trim().toLowerCase();
  const list = byId("updates_list");
  const count = byId("updates_count");

  if (!personId) {
    list.innerHTML = `<div class="muted">Choose a person.</div>`;
    count.textContent = "";
    return;
  }

  const pending = alive(db.items).filter(it => {
    if (!it.updateTargets || !it.updateTargets.includes(personId)) return false;
    const st = it.updateStatus?.[personId];
    if (st?.updated) return false;
    return true;
  }).filter(it => {
    if (!filter) return true;
    const topic = getTopic(it.topicId)?.name || "";
    const meeting = getMeeting(it.meetingId);
    const title = meeting?.title || "";
    return (it.text || "").toLowerCase().includes(filter)
      || topic.toLowerCase().includes(filter)
      || title.toLowerCase().includes(filter);
  }).sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));

  count.textContent = `${pending.length} pending update(s)`;

  list.innerHTML = pending.map(it => renderItemCard(it)).join("") || `<div class="muted">Nothing pending for this person. ✨</div>`;
  wireItemButtons(list);

  // stash for mark-all operation
  list.dataset.pendingIds = JSON.stringify(pending.map(x => x.id));
}

function renderActionsDashboard() {
  const ownerSel = byId("actions_owner");
  const topicSel = byId("actions_topic");
  const statusSel = byId("actions_status");
  const list = byId("actions_list");
  const count = byId("actions_count");

  if (!ownerSel || !topicSel || !statusSel || !list) return;

  if (actionsFilters.ownerId === null) {
    actionsFilters.ownerId = getDefaultActionsOwnerId();
    saveMeta().catch(console.error);
  }

  ownerSel.value = actionsFilters.ownerId || "";
  topicSel.value = actionsFilters.topicId || "";
  statusSel.value = actionsFilters.status || "";

  const ownerId = ownerSel.value || "";
  const topicId = topicSel.value || "";
  const status = statusSel.value || "";

  const matches = alive(db.items).filter(it => it.section === "action")
    .filter(it => !ownerId || it.ownerId === ownerId)
    .filter(it => !topicId || it.topicId === topicId)
    .filter(it => !status || ((it.status || "open") === status))
    .sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));

  if (count) count.textContent = `${matches.length} action item(s)`;

  list.innerHTML = matches.map(it => renderItemCard(it)).join("")
    || `<div class="muted">No actions match these filters.</div>`;
  wireItemButtons(list);
}

function renderTopicOverview() {
  const topicId = byId("topics_topic").value || "";
  const focus = byId("topics_focus").value || "overview";
  const out = byId("topic_output");

  if (!topicId) {
    out.innerHTML = `<div class="muted">Choose a topic.</div>`;
    return;
  }

  const topic = getTopic(topicId);
  const items = alive(db.items).filter(it => it.topicId === topicId);

  const byType = (secKey) => items.filter(it => it.section === secKey);

  const actionsOpen = items.filter(it => it.section === "action" && (it.status || "open") !== "done");
  const decisions = byType("decision");
  const questions = byType("question");
  const info = byType("info");

  const renderBucket = (title, arr) => `
    <div class="item">
      <div class="item__top">
        <div><strong>${escapeHtml(title)}</strong></div>
        <div class="badges"><span class="badge">${arr.length}</span></div>
      </div>
      <div class="list">
        ${arr.slice(0, 50).map(it => renderItemCard(it)).join("") || `<div class="muted">None.</div>`}
      </div>
    </div>
  `;

  const blocks = [];
  if (focus === "overview") {
    blocks.push(renderBucket("Open actions", actionsOpen));
    blocks.push(renderBucket("Decisions", decisions));
    blocks.push(renderBucket("Questions", questions));
    blocks.push(renderBucket("Information", info));
  } else if (focus === "actions") blocks.push(renderBucket("Open actions", actionsOpen));
  else if (focus === "decisions") blocks.push(renderBucket("Decisions", decisions));
  else if (focus === "questions") blocks.push(renderBucket("Questions", questions));
  else if (focus === "info") blocks.push(renderBucket("Information", info));

  out.innerHTML = `
    <div class="item">
      <div class="item__top">
        <div>
          <strong>${escapeHtml(topic?.name || "Topic")}</strong>
          <div class="muted">${items.length} item(s) across all meetings</div>
        </div>
        <div class="badges"><span class="badge badge--accent">${escapeHtml(topicId)}</span></div>
      </div>
    </div>
    ${blocks.join("")}
  `;

  wireItemButtons(out);
}

function renderSearch() {
  const q = byId("search_query").value.trim().toLowerCase();
  const type = byId("search_type").value || "all";
  const out = byId("search_results");

  if (!q) {
    out.innerHTML = `<div class="muted">Type to search. Results will show here.</div>`;
    return;
  }

  const matches = alive(db.items).filter(it => {
    const inText = (it.text || "").toLowerCase().includes(q);
    const topic = getTopic(it.topicId)?.name?.toLowerCase() || "";
    const meeting = getMeeting(it.meetingId);
    const title = meeting?.title?.toLowerCase() || "";
    const inMeta = topic.includes(q) || title.includes(q);

    const typeOk = type === "all" ? true : it.section === type;
    return typeOk && (inText || inMeta);
  }).sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt)).slice(0, 80);

  out.innerHTML = matches.map(it => renderItemCard(it)).join("") || `<div class="muted">No matches.</div>`;
  wireItemButtons(out);
}

/**
 * Renders the Tasks module list view.
 */
function renderTasks() {
  const list = byId("tasks_list");
  const countEl = byId("tasks_count");
  if (!list || !countEl) return;

  const statusFilter = taskFilters.status || "";
  const priorityFilter = taskFilters.priority || "";

  const statusSelect = byId("task_filter_status");
  const prioritySelect = byId("task_filter_priority");
  if (statusSelect && document.activeElement !== statusSelect) statusSelect.value = statusFilter;
  if (prioritySelect && document.activeElement !== prioritySelect) prioritySelect.value = priorityFilter;

  const filtered = alive(db.tasks).filter(task => {
    const statusOk = statusFilter ? task.status === statusFilter : true;
    const priorityOk = priorityFilter ? task.priority === priorityFilter : true;
    return statusOk && priorityOk;
  });

  // Sort by due date (ascending), then priority, then status, then updated time.
  const sorted = filtered.sort((a, b) => {
    const aDue = a.dueDate ? new Date(`${a.dueDate}T00:00:00`).getTime() : Number.POSITIVE_INFINITY;
    const bDue = b.dueDate ? new Date(`${b.dueDate}T00:00:00`).getTime() : Number.POSITIVE_INFINITY;
    if (aDue !== bDue) return aDue - bDue;

    const aPriority = TASK_PRIORITY_ORDER[a.priority] ?? TASK_PRIORITY_ORDER.medium;
    const bPriority = TASK_PRIORITY_ORDER[b.priority] ?? TASK_PRIORITY_ORDER.medium;
    if (aPriority !== bPriority) return aPriority - bPriority;

    const aStatus = TASK_STATUS_ORDER[a.status] ?? TASK_STATUS_ORDER.todo;
    const bStatus = TASK_STATUS_ORDER[b.status] ?? TASK_STATUS_ORDER.todo;
    if (aStatus !== bStatus) return aStatus - bStatus;

    return Date.parse(b.updatedAt) - Date.parse(a.updatedAt);
  });

  countEl.textContent = `${filtered.length} task${filtered.length === 1 ? "" : "s"} shown`;

  if (!sorted.length) {
    list.innerHTML = `<div class="muted">No tasks yet. Add your first task to get started.</div>`;
    return;
  }

  list.innerHTML = sorted.map(task => renderTaskCard(task)).join("");
  wireTaskList(list);
}

/**
 * Builds a task card for the Tasks module list.
 * @param {object} task Task record.
 * @returns {string} Task list markup.
 */
function renderTaskCard(task) {
  const statusKey = task.status || "todo";
  const priorityKey = task.priority || "medium";
  const statusLabel = TASK_STATUS_LABELS[statusKey] || statusKey;
  const priorityLabel = TASK_PRIORITY_LABELS[priorityKey] || priorityKey;
  const dueLabel = formatTaskDueDate(task.dueDate);

  // Status and priority color classes for visual scanning.
  const statusBadgeClass = `status-pill status-pill--${statusKey}`;
  const priorityBadgeClass = `priority-pill priority-pill--${priorityKey}`;

  return `
    <div class="task-card" data-task-id="${task.id}">
      <div class="task-card__header">
        <div>
          <div class="task-card__title">${escapeHtml(task.title)}</div>
          <div class="task-card__meta">Updated ${escapeHtml(fmtDateTime(task.updatedAt))}</div>
        </div>
        <div class="task-card__badges">
          <span class="badge ${statusBadgeClass}">${escapeHtml(statusLabel)}</span>
          <span class="badge ${priorityBadgeClass}">${escapeHtml(priorityLabel)}</span>
          <span class="badge">Due: ${escapeHtml(dueLabel)}</span>
        </div>
      </div>
      <div class="task-card__body">
        <div class="task-field">
          <label>Title</label>
          <input type="text" value="${escapeHtml(task.title)}" readonly />
        </div>
        <div class="task-field">
          <label>Notes</label>
          <textarea readonly>${escapeHtml(task.notes || "No notes yet.")}</textarea>
        </div>
        <div class="task-field-grid">
          <div class="task-field">
            <label>Due date</label>
            <input type="text" value="${escapeHtml(dueLabel)}" readonly />
          </div>
          <div class="task-field">
            <label>Priority</label>
            <div class="task-pill ${priorityBadgeClass}">${escapeHtml(priorityLabel)}</div>
          </div>
          <div class="task-field">
            <label>Status</label>
            <select class="task-status task-status--${statusKey}" data-task-field="status">
              <option value="todo" ${statusKey === "todo" ? "selected" : ""}>To do</option>
              <option value="in_progress" ${statusKey === "in_progress" ? "selected" : ""}>In progress</option>
              <option value="blocked" ${statusKey === "blocked" ? "selected" : ""}>Blocked</option>
              <option value="done" ${statusKey === "done" ? "selected" : ""}>Done</option>
            </select>
          </div>
        </div>
      </div>
      <div class="task-card__actions">
        <button class="smallbtn smallbtn--danger" data-task-action="delete">Delete</button>
      </div>
    </div>
  `;
}

/**
 * Wires click handlers for task cards.
 * @param {HTMLElement} container Task list container.
 */
function wireTaskList(container) {
  // Handle status updates inline so only the status remains editable.
  container.querySelectorAll("[data-task-field='status']").forEach(select => {
    select.addEventListener("change", async () => {
      const card = select.closest("[data-task-id]");
      if (!card) return;
      const taskId = card.getAttribute("data-task-id");
      const task = getTask(taskId);
      if (!task) return;

      task.status = select.value || "todo";
      task.updatedAt = nowIso();
      select.className = `task-status task-status--${task.status}`;
      markDirty();
      await saveLocal();
      renderTasks();
    });
  });

  container.querySelectorAll("[data-task-action='delete']").forEach(btn => {
    btn.addEventListener("click", async () => {
      const card = btn.closest("[data-task-id]");
      if (!card) return;
      const taskId = card.getAttribute("data-task-id");
      const task = getTask(taskId);
      if (!task) return;

      const confirmDelete = confirm("Delete this task? This cannot be undone.");
      if (!confirmDelete) return;
      task.deleted = true;
      task.updatedAt = nowIso();
      markDirty();
      await saveLocal();
      renderTasks();
    });
  });
}

function renderAll() {
  renderTemplates();
  renderTopics();
  renderPeopleSelects();
  renderMeetingCounterpartSelects();
  renderPeopleManager();
  renderGroups();
  renderActionsFiltersOptions();
  renderCurrentMeetingHeader();

  // update overview selects might have changed
  renderUpdates();
  renderActionsDashboard();
  renderTopicOverview();
  renderSearch();
  renderQuickSearch();
  renderMeetingCalendar();
  renderTasks();

  const defaultOwnerInput = byId("default_owner_name");
  if (defaultOwnerInput && document.activeElement !== defaultOwnerInput) {
    defaultOwnerInput.value = db.settings?.defaultOwnerName || "";
  }
  syncMeetingOneToOneFields();
}

/* ================================
   ACTIONS
=================================== */

/**
 * Updates the meeting lightbox to reflect whether 1:1 counterpart fields are required.
 */
function syncMeetingOneToOneFields() {
  const templateSelect = byId("meeting_template");
  const container = byId("meeting_one_to_one_fields");
  if (!templateSelect || !container) return;
  const isOneToOne = templateSelect.value === ONE_TO_ONE_TEMPLATE_ID;
  container.hidden = !isOneToOne;
  const tag = container.querySelector('[data-required-tag="oneToOneTarget"]');
  if (tag) {
    tag.classList.toggle("is-hidden", !isOneToOne);
  }
  if (!isOneToOne) {
    const personSelect = byId("meeting_one_to_one_person");
    const groupSelect = byId("meeting_one_to_one_group");
    if (personSelect) personSelect.value = "";
    if (groupSelect) groupSelect.value = "";
  }
}

/**
 * Opens the task creation lightbox.
 */
function openTaskLightbox() {
  openLightbox("task_lightbox", "task_title");
}

/**
 * Closes the task creation lightbox and clears draft errors.
 */
function closeTaskLightbox() {
  closeLightbox("task_lightbox");
}

/**
 * Clears the meeting creation form inputs so each draft starts clean.
 */
function clearMeetingForm() {
  const titleInput = byId("meeting_title");
  const datetimeInput = byId("meeting_datetime");
  const personSelect = byId("meeting_one_to_one_person");
  const groupSelect = byId("meeting_one_to_one_group");
  if (titleInput) titleInput.value = "";
  if (datetimeInput) {
    datetimeInput.value = toLocalDateTimeValue(nowIso());
  }
  if (personSelect) personSelect.value = "";
  if (groupSelect) groupSelect.value = "";
  syncMeetingOneToOneFields();
}

/**
 * Opens the meeting creation lightbox and resets draft-specific inputs.
 */
function openMeetingLightbox() {
  // Ensure the lightbox is reset to creation mode with a fresh draft.
  meetingEditId = null;
  setMeetingLightboxMode("create");
  clearMeetingForm();
  openLightbox("meeting_lightbox", "meeting_title");
}

/**
 * Updates the meeting lightbox title and primary button label by mode.
 * @param {"create"|"edit"} mode Lightbox mode.
 */
function setMeetingLightboxMode(mode) {
  const titleEl = byId("meeting_lightbox_title");
  const actionBtn = byId("create_meeting_btn");
  const deleteBtn = byId("meeting_delete_btn");
  if (titleEl) {
    titleEl.textContent = mode === "edit" ? "Edit meeting" : "Create a meeting";
  }
  if (actionBtn) {
    actionBtn.textContent = mode === "edit" ? "Save changes" : "Create meeting";
  }
  if (deleteBtn) {
    // Only show the destructive delete action when editing an existing meeting.
    deleteBtn.hidden = mode !== "edit";
  }
}

/**
 * Loads a meeting's details into the lightbox form controls.
 * @param {object} meeting Meeting record to edit.
 */
function populateMeetingForm(meeting) {
  const templateSelect = byId("meeting_template");
  const topicSelect = byId("meeting_topic");
  const titleInput = byId("meeting_title");
  const datetimeInput = byId("meeting_datetime");
  const personSelect = byId("meeting_one_to_one_person");
  const groupSelect = byId("meeting_one_to_one_group");

  if (templateSelect) templateSelect.value = meeting.templateId || "";
  if (topicSelect) topicSelect.value = meeting.topicId || "";
  if (titleInput) titleInput.value = meeting.title || "";
  if (datetimeInput) datetimeInput.value = toLocalDateTimeValue(meeting.datetime);
  if (personSelect) personSelect.value = meeting.oneToOnePersonId || "";
  if (groupSelect) groupSelect.value = meeting.oneToOneGroupId || "";
  syncMeetingOneToOneFields();
}

/**
 * Opens the meeting lightbox in edit mode for the selected meeting.
 * @param {string} meetingId Meeting identifier to edit.
 */
function openMeetingEditLightbox(meetingId) {
  const meeting = getMeeting(meetingId);
  if (!meeting) return;
  meetingEditId = meeting.id;
  setMeetingLightboxMode("edit");
  populateMeetingForm(meeting);
  openLightbox("meeting_lightbox", "meeting_title");
}

/**
 * Opens the topic creation lightbox and resets validation state.
 */
function openTopicLightbox() {
  const input = byId("topic_name_input");
  if (input) input.value = "";
  setLightboxError("topic_lightbox_error", "");
  openLightbox("topic_lightbox", "topic_name_input");
}

/**
 * Opens the group creation lightbox and resets validation state.
 */
function openGroupLightbox() {
  const input = byId("group_name_input");
  if (input) input.value = "";
  setLightboxError("group_lightbox_error", "");
  openLightbox("group_lightbox", "group_name_input");
}

/**
 * Opens the person creation lightbox and resets draft and validation state.
 */
function openPersonLightbox() {
  personCreateState = { draft: createPersonDraft(null), error: "" };
  const nameInput = byId("person_create_name");
  const emailInput = byId("person_create_email");
  const orgInput = byId("person_create_org");
  const titleInput = byId("person_create_title");

  if (nameInput) nameInput.value = personCreateState.draft.name;
  if (emailInput) emailInput.value = personCreateState.draft.email;
  if (orgInput) orgInput.value = personCreateState.draft.organisation;
  if (titleInput) titleInput.value = personCreateState.draft.jobTitle;

  setLightboxError("person_lightbox_error", "");
  openLightbox("person_lightbox", "person_create_name");
}

/**
 * Updates a lightbox error container with the provided message.
 * @param {string} errorId DOM id of the error container.
 * @param {string} message Error message to display.
 */
function setLightboxError(errorId, message) {
  const errorEl = byId(errorId);
  if (!errorEl) return;
  errorEl.textContent = message;
  errorEl.hidden = !message;
}

/**
 * Reads the task creation form fields into a draft object.
 * @returns {object} Task draft from the form.
 */
function readTaskFormDraft() {
  return {
    title: byId("task_title")?.value || "",
    notes: byId("task_notes")?.value || "",
    dueDate: byId("task_due")?.value || "",
    priority: byId("task_priority")?.value || "medium",
    status: byId("task_status")?.value || "todo"
  };
}

/**
 * Resets the task creation form to its defaults.
 */
function clearTaskForm() {
  const defaults = createTaskDraft(null);
  const titleInput = byId("task_title");
  const notesInput = byId("task_notes");
  const dueInput = byId("task_due");
  const prioritySelect = byId("task_priority");
  const statusSelect = byId("task_status");

  if (titleInput) titleInput.value = defaults.title;
  if (notesInput) notesInput.value = defaults.notes;
  if (dueInput) dueInput.value = defaults.dueDate;
  if (prioritySelect) prioritySelect.value = defaults.priority;
  if (statusSelect) statusSelect.value = defaults.status;
}

/**
 * Reads the topic creation input, validates, and persists a new topic.
 */
async function addTopicFromLightbox() {
  const name = byId("topic_name_input")?.value.trim() || "";
  if (!name) {
    setLightboxError("topic_lightbox_error", "Topic name is required.");
    return;
  }

  const topicId = ensureTopic(name);
  markDirty();
  await saveLocal();
  renderAll();
  const meetingTopic = byId("meeting_topic");
  if (meetingTopic) {
    meetingTopic.value = topicId;
  }
  closeLightbox("topic_lightbox");
}

/**
 * Reads the group creation input, validates, and persists a new group.
 */
async function addGroupFromLightbox() {
  const name = byId("group_name_input")?.value.trim() || "";
  if (!name) {
    setLightboxError("group_lightbox_error", "Group name is required.");
    return;
  }

  const exists = alive(db.groups).some(group => group.name.toLowerCase() === name.toLowerCase());
  if (exists) {
    setLightboxError("group_lightbox_error", "A group with this name already exists.");
    return;
  }

  db.groups.push({ id: uid("group"), name, memberIds: [], updatedAt: nowIso() });
  markDirty();
  await saveLocal();
  renderAll();
  closeLightbox("group_lightbox");
}

/**
 * Reads the person creation inputs, validates, and persists a new person.
 */
async function addPersonFromLightbox() {
  const name = byId("person_create_name")?.value.trim() || "";
  const email = byId("person_create_email")?.value.trim() || "";
  const organisation = byId("person_create_org")?.value.trim() || "";
  const jobTitle = byId("person_create_title")?.value.trim() || "";
  const draft = { name, email, organisation, jobTitle };
  const errs = validatePersonDraft(draft, null);

  if (errs.length) {
    setLightboxError("person_lightbox_error", errs.join(" "));
    return;
  }

  const person = {
    id: uid("person"),
    ...draft,
    updatedAt: nowIso(),
  };
  db.people.push(person);
  personViewId = person.id;
  personEditorState = { isNew: false, draft: createPersonDraft(person), error: "" };
  markDirty();
  await saveLocal();
  renderAll();
  closeLightbox("person_lightbox");
}

/**
 * Adds a new task to the local database.
 */
async function addTask() {
  const draft = readTaskFormDraft();
  const errs = validateTaskDraft(draft);
  if (errs.length) {
    alert(errs.join(" "));
    return;
  }

  const now = nowIso();
  const task = {
    id: uid("task"),
    title: draft.title.trim(),
    notes: draft.notes.trim(),
    dueDate: draft.dueDate,
    priority: draft.priority,
    status: draft.status,
    createdAt: now,
    updatedAt: now
  };

  db.tasks.push(task);
  clearTaskForm();
  closeTaskLightbox();
  markDirty();
  await saveLocal();
  renderTasks();
}

async function saveMeetingFromLightbox() {
  const templateId = byId("meeting_template").value;
  const topicId = byId("meeting_topic").value || null;
  const oneToOnePersonId = byId("meeting_one_to_one_person")?.value || "";
  const oneToOneGroupId = byId("meeting_one_to_one_group")?.value || "";
  const isOneToOneTemplate = templateId === ONE_TO_ONE_TEMPLATE_ID;

  if (!templateId) { alert("Choose a template."); return; }
  if (!topicId) { alert("Choose or add a topic."); return; }
  if (isOneToOneTemplate) {
    if (!oneToOnePersonId && !oneToOneGroupId) {
      alert("Select a person or group for the 1:1 meeting.");
      return;
    }
    if (oneToOnePersonId && oneToOneGroupId) {
      alert("Choose either a person or a group, not both.");
      return;
    }
  }

  const title = byId("meeting_title").value.trim() || "";
  const dt = byId("meeting_datetime").value;
  const datetime = dt ? new Date(dt).toISOString() : nowIso();

  let meeting = null;

  if (meetingEditId) {
    meeting = getMeeting(meetingEditId);
    if (!meeting) {
      alert("Meeting not found.");
      meetingEditId = null;
      return;
    }
    meeting.templateId = templateId;
    meeting.topicId = topicId;
    meeting.title = title;
    meeting.datetime = datetime;
    meeting.oneToOnePersonId = isOneToOneTemplate ? (oneToOnePersonId || null) : null;
    meeting.oneToOneGroupId = isOneToOneTemplate ? (oneToOneGroupId || null) : null;
    meeting.updatedAt = nowIso();
  } else {
    meeting = {
      id: uid("meeting"),
      templateId,
      topicId,
      title,
      datetime,
      oneToOnePersonId: isOneToOneTemplate ? (oneToOnePersonId || null) : null,
      oneToOneGroupId: isOneToOneTemplate ? (oneToOneGroupId || null) : null,
      createdAt: nowIso(),
      updatedAt: nowIso(),
    };
    db.meetings.push(meeting);
  }

  currentMeetingId = meeting.id;
  setMeetingView("notes");
  meetingEditId = null;
  clearMeetingForm();
  closeLightbox("meeting_lightbox");

  markDirty();
  await saveLocal();
  await saveMeta();
  renderAll();
}

/**
 * Deletes a meeting by id and optionally purges its items.
 * @param {string} meetingId Meeting identifier to delete.
 * @returns {Promise<boolean>} Whether a meeting was deleted.
 */
async function deleteMeetingById(meetingId) {
  if (!meetingId) return false;
  const meeting = getMeeting(meetingId);
  if (!meeting) return false;

  const confirmDelete = confirm("Delete this meeting? Items will remain but become orphaned unless you delete them too.");
  if (!confirmDelete) return false;

  meeting.deleted = true;
  meeting.updatedAt = nowIso();

  // Optionally also delete items for the meeting to reduce orphaned records.
  if (confirm("Also delete all items from this meeting? (Recommended)")) {
    for (const it of alive(db.items).filter(i => i.meetingId === meeting.id)) {
      it.deleted = true;
      it.updatedAt = nowIso();
    }
  }

  if (currentMeetingId === meeting.id) {
    currentMeetingId = null;
  }

  markDirty();
  await saveLocal();
  await saveMeta();
  renderAll();
  return true;
}

/**
 * Deletes the currently open meeting in the Notes view.
 */
async function deleteCurrentMeeting() {
  await deleteMeetingById(currentMeetingId);
}

/**
 * Deletes the meeting currently loaded in the edit lightbox.
 */
async function deleteMeetingFromLightbox() {
  if (!meetingEditId) {
    alert("No meeting selected.");
    return;
  }
  const deleted = await deleteMeetingById(meetingEditId);
  if (!deleted) return;

  meetingEditId = null;
  clearMeetingForm();
  closeLightbox("meeting_lightbox");
}

function buildMeetingSummary(meetingId) {
  const m = getMeeting(meetingId);
  if (!m) return "No meeting selected.";

  const tpl = getTemplate(m.templateId);
  const topic = getTopic(m.topicId);
  const items = alive(db.items).filter(i => i.meetingId === m.id);

  const lines = [];
  lines.push(`Meeting: ${m.title || "(Untitled)"}`);
  lines.push(`Date: ${fmtDateTime(m.datetime)}`);
  lines.push(`Template: ${tpl?.name || ""}`);
  lines.push(`Topic: ${topic?.name || ""}`);
  if (tpl?.id === ONE_TO_ONE_TEMPLATE_ID) {
    const context = getOneToOneContext(m);
    if (context) {
      lines.push(`1:1 counterpart: ${context.label}`);
    }
  }
  lines.push("");

  const secOrder = tpl?.sections || [];
  for (const sec of secOrder) {
    const secItems = items.filter(i => i.section === sec.key);
    lines.push(`${sec.label}:`);
    if (!secItems.length) {
      lines.push(`- (none)`);
      lines.push("");
      continue;
    }
    for (const it of secItems) {
      const owner = it.ownerId ? getPerson(it.ownerId)?.name : null;
      const due = it.dueDate ? ` due ${it.dueDate}` : "";
      const status = it.status ? ` [${it.status}]` : "";
      const targets = (it.updateTargets || []).map(pid => getPerson(pid)?.name).filter(Boolean);
      const targetTxt = targets.length ? ` (update: ${targets.join(", ")})` : "";
      lines.push(`- ${it.text}${status}${owner ? ` (owner: ${owner})` : ""}${due}${targetTxt}`);
    }
    lines.push("");
  }

  return lines.join("\n");
}

async function copyMeetingSummary() {
  if (!currentMeetingId) { alert("No meeting selected."); return; }
  const text = buildMeetingSummary(currentMeetingId);
  await copyToClipboard(text);
  alert("Meeting summary copied.");
}

async function markAllShownUpdates() {
  const personId = byId("updates_person").value || "";
  if (!personId) return;

  const list = byId("updates_list");
  const pendingIds = JSON.parse(list.dataset.pendingIds || "[]");
  if (!pendingIds.length) { alert("No pending updates shown."); return; }

  const person = getPerson(personId);
  if (!confirm(`Mark ${pendingIds.length} item(s) as updated for ${person?.name || "this person"}?`)) return;

  const stamp = nowIso();
  const meetingIdForStamp = currentMeetingId || null;

  for (const id of pendingIds) {
    const it = getItem(id);
    if (!it) continue;
    it.updateStatus = it.updateStatus || {};
    it.updateStatus[personId] = {
      updated: true,
      updatedAt: stamp,
      meetingId: meetingIdForStamp
    };
    it.updatedAt = stamp;
  }

  markDirty();
  await saveLocal();
  renderAll();
}

async function copyPendingUpdatesText() {
  const personId = byId("updates_person").value || "";
  if (!personId) return;
  const person = getPerson(personId);

  const pending = alive(db.items).filter(it => {
    if (!it.updateTargets || !it.updateTargets.includes(personId)) return false;
    const st = it.updateStatus?.[personId];
    return !st?.updated;
  }).sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));

  const lines = [];
  lines.push(`Updates for: ${person?.name || personId}`);
  lines.push(`Generated: ${fmtDateTime(nowIso())}`);
  lines.push("");

  for (const it of pending) {
    const topic = getTopic(it.topicId)?.name || "No topic";
    const meeting = getMeeting(it.meetingId);
    const when = meeting ? fmtDateTime(meeting.datetime) : "";
    lines.push(`- [${topic}] ${it.text} (${it.section}${when ? ` • ${when}` : ""})`);
  }

  await copyToClipboard(lines.join("\n"));
  alert("Updates copied.");
}

async function downloadJsonBackup() {
  const blob = new Blob([JSON.stringify(db, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `meeting-notes-backup_${new Date().toISOString().replace(/[:.]/g,"-")}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function importJsonBackup(file) {
  const text = await file.text();
  const incoming = JSON.parse(text);

  // Merge incoming into local
  db = mergeDb(db, incoming);

  markDirty();
  await saveLocal();
  renderAll();
  alert("Import complete (merged). You can Sync to Drive if you want.");
}

/* ================================
   WIRING
=================================== */

function wireModules() {
  document.querySelectorAll(".module-tab").forEach(btn => {
    btn.addEventListener("click", () => setActiveModule(btn.dataset.module));
  });
  document.querySelectorAll(".module-subtab").forEach(btn => {
    btn.addEventListener("click", () => setMeetingModuleTab(btn.dataset.moduleTab));
  });
}

function wireTasksControls() {
  // Task lightbox triggers.
  byId("task_open_modal_btn")?.addEventListener("click", openTaskLightbox);
  byId("task_add_btn")?.addEventListener("click", addTask);
  byId("task_clear_btn")?.addEventListener("click", clearTaskForm);
  byId("task_filter_status")?.addEventListener("change", (event) => {
    taskFilters.status = event.target.value;
    saveMeta().catch(console.error);
    renderTasks();
  });
  byId("task_filter_priority")?.addEventListener("change", (event) => {
    taskFilters.priority = event.target.value;
    saveMeta().catch(console.error);
    renderTasks();
  });
}

/**
 * Wires shared lightbox dismissal controls for all modal flows.
 */
function wireLightboxControls() {
  const lightboxIds = [
    "task_lightbox",
    "meeting_lightbox",
    "topic_lightbox",
    "person_lightbox",
    "group_lightbox"
  ];

  lightboxIds.forEach((id) => {
    byId(id)?.addEventListener("click", (event) => {
      if (event.target?.matches("[data-lightbox-close]")) {
        closeLightbox(id);
      }
    });
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeVisibleLightboxes();
    }
  });
}

function wireTopButtons() {
  byId("auth_btn").addEventListener("click", handleAuthClick);
  byId("signout_btn").addEventListener("click", handleSignoutClick);
  byId("sync_btn").addEventListener("click", syncNow);
}

function wireMeetingControls() {
  byId("meeting_open_lightbox_btn")?.addEventListener("click", openMeetingLightbox);
  byId("create_meeting_btn").addEventListener("click", saveMeetingFromLightbox);
  byId("delete_meeting_btn").addEventListener("click", deleteCurrentMeeting);
  // Delete the meeting directly from the edit lightbox.
  byId("meeting_delete_btn")?.addEventListener("click", deleteMeetingFromLightbox);
  byId("download_meeting_summary_btn").addEventListener("click", copyMeetingSummary);

  byId("add_topic_btn").addEventListener("click", openTopicLightbox);
  byId("topic_add_btn").addEventListener("click", addTopicFromLightbox);

  byId("quick_search").addEventListener("input", debounce(renderQuickSearch, 150));

  const templateSelect = byId("meeting_template");
  const personSelect = byId("meeting_one_to_one_person");
  const groupSelect = byId("meeting_one_to_one_group");

  if (templateSelect) {
    templateSelect.addEventListener("change", () => {
      // Toggle the 1:1 counterpart controls when the template changes.
      syncMeetingOneToOneFields();
    });
  }
  if (personSelect) {
    personSelect.addEventListener("change", () => {
      // Ensure only one counterpart type is chosen at a time.
      if (personSelect.value && groupSelect) {
        groupSelect.value = "";
      }
    });
  }
  if (groupSelect) {
    groupSelect.addEventListener("change", () => {
      // Ensure only one counterpart type is chosen at a time.
      if (groupSelect.value && personSelect) {
        personSelect.value = "";
      }
    });
  }

  document.querySelectorAll("[data-calendar-view]").forEach(btn => {
    btn.addEventListener("click", () => {
      const view = btn.getAttribute("data-calendar-view");
      if (!view) return;
      setMeetingCalendarView(view);
    });
  });

  document.querySelectorAll("[data-calendar-nav]").forEach(btn => {
    btn.addEventListener("click", () => {
      const action = btn.getAttribute("data-calendar-nav");
      if (action === "prev") shiftMeetingCalendar(-1);
      if (action === "next") shiftMeetingCalendar(1);
      if (action === "today") resetMeetingCalendarToToday();
    });
  });

  document.querySelectorAll("[data-meeting-view]").forEach(btn => {
    btn.addEventListener("click", () => {
      const view = btn.getAttribute("data-meeting-view");
      if (!view) return;
      setMeetingView(view);
    });
  });
}

function wireUpdatesControls() {
  byId("updates_person").addEventListener("change", renderUpdates);
  byId("updates_filter").addEventListener("input", debounce(renderUpdates, 150));
  byId("mark_updates_btn").addEventListener("click", markAllShownUpdates);
  byId("copy_updates_btn").addEventListener("click", copyPendingUpdatesText);
}

function wireActionsControls() {
  const updateFilters = () => {
    actionsFilters = {
      ownerId: byId("actions_owner").value || "",
      topicId: byId("actions_topic").value || "",
      status: byId("actions_status").value || "",
    };
    saveMeta().catch(console.error);
    renderActionsDashboard();
  };

  byId("actions_owner").addEventListener("change", updateFilters);
  byId("actions_topic").addEventListener("change", updateFilters);
  byId("actions_status").addEventListener("change", updateFilters);
}

function wireTopicControls() {
  byId("topics_topic").addEventListener("change", renderTopicOverview);
  byId("topics_focus").addEventListener("change", renderTopicOverview);
}

function wireSearchControls() {
  byId("search_query").addEventListener("input", debounce(renderSearch, 150));
  byId("search_type").addEventListener("change", renderSearch);
}

/**
 * Wires group creation controls housed in the People module.
 * Keeps group management near people records for easier maintenance.
 */
function wireGroupControls() {
  byId("group_open_lightbox_btn")?.addEventListener("click", openGroupLightbox);
  byId("group_add_btn")?.addEventListener("click", addGroupFromLightbox);
}

function wireSettingsControls() {
  const downloadJsonBtn = byId("download_json_btn");
  if (downloadJsonBtn) {
    downloadJsonBtn.addEventListener("click", downloadJsonBackup);
  }

  const importJsonFile = byId("import_json_file");
  if (importJsonFile) {
    importJsonFile.addEventListener("change", async (e) => {
      const f = e.target.files?.[0];
      if (!f) return;
      await importJsonBackup(f);
      e.target.value = "";
    });
  }

  const defaultOwnerInput = byId("default_owner_name");
  if (defaultOwnerInput) {
    defaultOwnerInput.addEventListener("input", debounce(async () => {
      db.settings = db.settings || { defaultOwnerName: "", updatedAt: nowIso() };
      db.settings.defaultOwnerName = defaultOwnerInput.value.trim();
      db.settings.updatedAt = nowIso();
      markDirty();
      await saveLocal();
      renderActionsDashboard();
    }, 200));
  }
}

function wirePeopleControls() {
  byId("person_new_btn").addEventListener("click", openPersonLightbox);
  byId("person_create_btn").addEventListener("click", addPersonFromLightbox);

  byId("person_cancel_btn").addEventListener("click", () => {
    if (personViewId) {
      const person = getPerson(personViewId);
      personEditorState = { isNew: false, draft: createPersonDraft(person), error: "" };
    } else {
      personEditorState = { isNew: false, draft: null, error: "" };
    }
    renderPeopleManager();
  });

  byId("person_save_btn").addEventListener("click", async () => {
    if (!personEditorState.draft) return;
    const name = byId("person_name")?.value.trim() || "";
    const email = byId("person_email")?.value.trim() || "";
    const organisation = byId("person_org")?.value.trim() || "";
    const jobTitle = byId("person_title")?.value.trim() || "";
    const draft = { name, email, organisation, jobTitle };
    const errs = validatePersonDraft(draft, personEditorState.isNew ? null : personViewId);
    if (errs.length) {
      personEditorState = { ...personEditorState, draft, error: errs.join(" ") };
      renderPeopleManager();
      return;
    }

    if (personEditorState.isNew) {
      const person = {
        id: uid("person"),
        ...draft,
        updatedAt: nowIso(),
      };
      db.people.push(person);
      personViewId = person.id;
    } else if (personViewId) {
      const person = getPerson(personViewId);
      if (!person) return;
      person.name = draft.name;
      person.email = draft.email;
      person.organisation = draft.organisation;
      person.jobTitle = draft.jobTitle;
      person.updatedAt = nowIso();
    }

    personEditorState = { isNew: false, draft: createPersonDraft(getPerson(personViewId)), error: "" };
    markDirty();
    await saveLocal();
    renderAll();
  });
}

/* ================================
   INIT
=================================== */

async function init() {
  // basic UI wiring
  wireModules();
  wireTopButtons();
  wireMeetingControls();
  wireUpdatesControls();
  wireActionsControls();
  wireTopicControls();
  wireSearchControls();
  wireGroupControls();
  wireSettingsControls();
  wirePeopleControls();
  wireTasksControls();
  wireLightboxControls();

  // buttons disabled until APIs ready
  byId("auth_btn").disabled = true;
  byId("sync_btn").disabled = true;
  byId("signout_btn").style.display = "none";

  // network status
  setNetStatus();
  window.addEventListener("online", () => { setNetStatus(); updateAuthUi(); });
  window.addEventListener("offline", () => { setNetStatus(); updateAuthUi(); });

  // load local DB
  await loadLocal();

  // set default meeting datetime input to now
  const dt = new Date();
  const localISO = new Date(dt.getTime() - dt.getTimezoneOffset()*60000).toISOString().slice(0,16);
  byId("meeting_datetime").value = localISO;

  renderAll();
  setActiveModule(activeModule);
  setMeetingView(meetingView);
  updateAuthUi();
}

document.addEventListener("DOMContentLoaded", init);
