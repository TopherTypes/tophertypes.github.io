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
// Canonical name for the built-in default person entity.
const DEFAULT_PERSON_NAME = "Me";
// Default project color used when none is provided.
const DEFAULT_PROJECT_COLOR = "#7c9bff";

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
// Tracks if the Notes view has been opened from the calendar view.
let meetingNotesEnabled = false;
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
// Auto-sync cadence (30 seconds) for background refresh.
const AUTO_SYNC_INTERVAL_MS = 30 * 1000;
// Holds the interval ID for periodic auto-sync scheduling.
let autoSyncTimerId = null;
const itemEditState = new Map();
let personViewId = null;
let personEditorState = { isNew: false, draft: null, error: "" };
// Tracks draft state for the person creation lightbox flow.
let personCreateState = { draft: null, error: "" };
// Tracks the active project selection in the Projects module.
let projectViewId = null;
// Stores draft edits for the active project details panel.
let projectEditorState = { projectId: null, draft: null, error: "" };

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

// Status mapping helpers to keep action items and tasks aligned across modules.
const ACTION_STATUS_TO_TASK_STATUS = {
  open: "todo",
  in_progress: "in_progress",
  blocked: "blocked",
  done: "done",
};

const TASK_STATUS_TO_ACTION_STATUS = {
  todo: "open",
  in_progress: "in_progress",
  blocked: "blocked",
  done: "done",
};

/* ================================
   TASK LINKING HELPERS
=================================== */

/**
 * Normalizes an action item status value to its task-module equivalent.
 * @param {string} status Action item status.
 * @returns {string} Task status key for rendering/filtering.
 */
function mapActionStatusToTaskStatus(status) {
  return ACTION_STATUS_TO_TASK_STATUS[status] || "todo";
}

/**
 * Normalizes a task status value back to the action item status vocabulary.
 * @param {string} status Task status.
 * @returns {string} Action status key for persistence.
 */
function mapTaskStatusToActionStatus(status) {
  return TASK_STATUS_TO_ACTION_STATUS[status] || "open";
}

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
 * Creates the canonical default person record used across modules.
 * @returns {object} Default person entity with locked fields.
 */
function createDefaultPersonRecord() {
  return {
    id: uid("person"),
    name: DEFAULT_PERSON_NAME,
    email: "",
    organisation: "",
    jobTitle: "",
    isDefault: true,
    updatedAt: nowIso(),
  };
}

/**
 * Builds the default settings object for the app.
 * @returns {object} Default settings structure.
 */
function buildDefaultSettings() {
  return {
    sync: {
      frequencyMinutes: 30,
      autoSyncOnLaunch: true,
      autoSyncOnReconnect: true,
    },
    updatedAt: nowIso(),
  };
}

/**
 * Ensures settings exist and are normalized to the latest schema.
 * @param {object} targetDb Database instance to normalize.
 * @returns {boolean} Whether any updates were applied.
 */
function normalizeSettings(targetDb) {
  if (!targetDb) return false;
  const defaults = buildDefaultSettings();
  let changed = false;

  if (!targetDb.settings || typeof targetDb.settings !== "object") {
    targetDb.settings = defaults;
    return true;
  }

  if (!targetDb.settings.sync || typeof targetDb.settings.sync !== "object") {
    targetDb.settings.sync = { ...defaults.sync };
    changed = true;
  } else {
    if (typeof targetDb.settings.sync.frequencyMinutes !== "number") {
      targetDb.settings.sync.frequencyMinutes = defaults.sync.frequencyMinutes;
      changed = true;
    }
    if (typeof targetDb.settings.sync.autoSyncOnLaunch !== "boolean") {
      targetDb.settings.sync.autoSyncOnLaunch = defaults.sync.autoSyncOnLaunch;
      changed = true;
    }
    if (typeof targetDb.settings.sync.autoSyncOnReconnect !== "boolean") {
      targetDb.settings.sync.autoSyncOnReconnect = defaults.sync.autoSyncOnReconnect;
      changed = true;
    }
  }

  if (!targetDb.settings.updatedAt) {
    targetDb.settings.updatedAt = nowIso();
    changed = true;
  }

  if (changed) {
    targetDb.settings.updatedAt = nowIso();
  }

  return changed;
}

/**
 * Reads a nested settings value from the active database.
 * @param {string} path Dot-separated path for the setting key.
 * @returns {unknown} Setting value or undefined if missing.
 */
function getSettingValue(path) {
  const keys = path.split(".");
  let cursor = db?.settings;
  for (const key of keys) {
    if (!cursor || typeof cursor !== "object") return undefined;
    cursor = cursor[key];
  }
  return cursor;
}

/**
 * Updates a nested settings value on the database, creating objects as needed.
 * @param {string} path Dot-separated path for the setting key.
 * @param {unknown} value New value to persist.
 * @returns {boolean} Whether the value changed.
 */
function setSettingValue(path, value) {
  const keys = path.split(".");
  let cursor = db.settings;
  for (let i = 0; i < keys.length - 1; i += 1) {
    const key = keys[i];
    if (!cursor[key] || typeof cursor[key] !== "object") {
      cursor[key] = {};
    }
    cursor = cursor[key];
  }
  const lastKey = keys[keys.length - 1];
  if (cursor[lastKey] === value) return false;
  cursor[lastKey] = value;
  return true;
}

/**
 * Normalizes values coming from settings inputs for safe persistence.
 * @param {string} path Dot-separated path for the setting key.
 * @param {unknown} rawValue Raw UI value.
 * @returns {unknown} Cleaned value for storage.
 */
function coerceSettingValue(path, rawValue) {
  if (path === "sync.frequencyMinutes") {
    const asNumber = Number(rawValue);
    if (Number.isNaN(asNumber) || asNumber < 0) {
      return buildDefaultSettings().sync.frequencyMinutes;
    }
    return asNumber;
  }
  return rawValue;
}

/**
 * Reads a settings value from a form control element.
 * @param {HTMLElement} control Form control with data-setting metadata.
 * @returns {unknown} Parsed control value.
 */
function readSettingControlValue(control) {
  const type = control.dataset.settingType || control.type;
  if (type === "checkbox") {
    return control.checked;
  }
  if (type === "number") {
    return Number(control.value);
  }
  return control.value;
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
    ownerId: task?.ownerId || getDefaultPersonId() || "",
    dueDate: task?.dueDate || "",
    priority: task?.priority || "medium",
    status: task?.status || "todo",
    link: task?.link || "",
    updateTargets: Array.isArray(task?.updateTargets) ? task.updateTargets : [],
    updateStatus: task?.updateStatus || {},
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

/**
 * Creates an editable project draft from a project record.
 * @param {object} project Project record from the database.
 * @returns {object} Draft fields for editing.
 */
function createProjectDraft(project) {
  return {
    name: project?.name || "",
    ownerId: project?.ownerId || getDefaultPersonId() || "",
    startDate: project?.startDate || "",
    endDate: project?.endDate || "",
    color: project?.color || DEFAULT_PROJECT_COLOR,
  };
}

/**
 * Validates a project draft and returns a list of user-facing errors.
 * @param {object} draft Project draft input.
 * @returns {string[]} Validation error messages.
 */
function validateProjectDraft(draft) {
  const errs = [];
  if (!draft.name.trim()) errs.push("Project name is required.");
  if (draft.startDate && draft.endDate) {
    const start = new Date(`${draft.startDate}T00:00:00`);
    const end = new Date(`${draft.endDate}T00:00:00`);
    if (!Number.isNaN(start.getTime()) && !Number.isNaN(end.getTime()) && start > end) {
      errs.push("Project end date must be after the start date.");
    }
  }
  return errs;
}

/**
 * Normalizes a hex color string for project styling.
 * @param {string} value Raw color string from user input.
 * @param {string} fallback Fallback color value if invalid.
 * @returns {string} Sanitized hex color string.
 */
function normalizeHexColor(value, fallback = DEFAULT_PROJECT_COLOR) {
  const trimmed = (value || "").trim();
  if (!trimmed) return fallback;
  const hex = trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
  if (/^#([0-9a-fA-F]{6}|[0-9a-fA-F]{3})$/.test(hex)) {
    return hex.toLowerCase();
  }
  return fallback;
}

/**
 * Converts a hex color string into an RGB tuple.
 * @param {string} hex Hex color string.
 * @returns {{r:number,g:number,b:number}|null} RGB values or null when invalid.
 */
function hexToRgb(hex) {
  const normalized = normalizeHexColor(hex);
  const shorthand = normalized.length === 4;
  const value = normalized.replace("#", "");
  const fullValue = shorthand
    ? value.split("").map(ch => ch + ch).join("")
    : value;
  const intVal = Number.parseInt(fullValue, 16);
  if (Number.isNaN(intVal)) return null;
  return {
    r: (intVal >> 16) & 255,
    g: (intVal >> 8) & 255,
    b: intVal & 255,
  };
}

/**
 * Builds RGBA variants of a project color for calendar styling.
 * @param {string} hex Hex color string.
 * @returns {{border:string, background:string}} RGBA colors for UI surfaces.
 */
function buildProjectColorPalette(hex) {
  const rgb = hexToRgb(hex);
  if (!rgb) {
    return {
      border: "rgba(124,155,255,0.35)",
      background: "rgba(124,155,255,0.16)",
    };
  }
  return {
    border: `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.5)`,
    background: `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.18)`,
  };
}

/**
 * Builds a default update-status map for a set of target person ids.
 * @param {string[]} targetIds Person identifiers to seed in the status map.
 * @param {object} [existingStatus={}] Existing updateStatus map to preserve.
 * @returns {object} Normalized updateStatus object keyed by person id.
 */
function buildUpdateStatusForTargets(targetIds, existingStatus = {}) {
  const out = {};
  (targetIds || []).forEach(pid => {
    const prior = existingStatus?.[pid];
    out[pid] = prior ? { updated: !!prior.updated, updatedAt: prior.updatedAt } : { updated: false };
  });
  return out;
}

/**
 * Normalizes a comma-separated list of people names into ids.
 * @param {string} rawValue Input string from the UI.
 * @param {object[]} people People list for name lookup.
 * @returns {{ids: string[], missing: string[]}} Parsed ids + missing names.
 */
function parseTargetNames(rawValue, people) {
  const names = rawValue
    .split(",")
    .map(name => name.trim())
    .filter(Boolean);
  const ids = [];
  const missing = [];
  names.forEach(name => {
    const match = findPersonByName(name, people);
    if (match) {
      ids.push(match.id);
    } else {
      missing.push(name);
    }
  });
  return { ids, missing };
}

/**
 * Formats update target ids into a comma-separated list of names.
 * @param {string[]} targetIds Person ids to display.
 * @returns {string} Human-readable target list.
 */
function formatTargetNames(targetIds) {
  const names = (targetIds || [])
    .map(pid => getPerson(pid)?.name)
    .filter(Boolean);
  return names.join(", ");
}

/**
 * Summarizes update completion for a record.
 * @param {object} record Task or action item record with update targets.
 * @returns {string} Update completion label.
 */
function buildUpdateProgressLabel(record) {
  const targets = record?.updateTargets || [];
  if (!targets.length) return "No update targets";
  const updatedCount = Object.values(record.updateStatus || {}).filter(st => st.updated).length;
  return `${updatedCount}/${targets.length} updated`;
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

/**
 * Determines the most relevant date for an action item timeline entry.
 * Prefers due dates when available, otherwise falls back to updated timestamps.
 * @param {object} action Action item record.
 * @returns {Date} Normalized Date object for sorting.
 */
function getActionTimelineDate(action) {
  if (action?.dueDate) {
    const due = new Date(`${action.dueDate}T00:00:00`);
    if (!Number.isNaN(due.getTime())) return due;
  }
  const updated = new Date(action?.updatedAt || "");
  if (!Number.isNaN(updated.getTime())) return updated;
  return new Date();
}

/**
 * Formats the timeline label for an action item.
 * @param {object} action Action item record.
 * @returns {string} Human-readable label for the timeline entry.
 */
function formatActionTimelineLabel(action) {
  if (action?.dueDate) {
    return `Due ${fmtDate(`${action.dueDate}T00:00:00`)}`;
  }
  if (action?.updatedAt) {
    return `Updated ${fmtDate(action.updatedAt)}`;
  }
  return "No date";
}

/* ================================
   DEFAULT DB / TEMPLATES
=================================== */

function makeDefaultDb() {
  // Seed with the required default person record so every module has a shared entity.
  const defaultPerson = createDefaultPersonRecord();
  return {
    schemaVersion: 1,
    updatedAt: nowIso(),
    settings: buildDefaultSettings(),
    templates: createBuiltinTemplates(),
    people: [defaultPerson],
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
      await syncNow({ silent: false, trigger: "post-auth" });
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

/**
 * Builds a stable, comparable signature for an entity by sorting keys and omitting volatile fields.
 * @param {object} entity Entity record to sign.
 * @param {Set<string>} ignoreKeys Fields excluded from the signature.
 * @returns {string} Stable JSON signature for equality comparisons.
 */
function buildEntitySignature(entity, ignoreKeys) {
  const normalizeValue = (value) => {
    if (Array.isArray(value)) {
      return value.map(normalizeValue);
    }
    if (value && typeof value === "object") {
      const sortedKeys = Object.keys(value).sort();
      const normalized = {};
      sortedKeys.forEach(key => {
        normalized[key] = normalizeValue(value[key]);
      });
      return normalized;
    }
    return value;
  };

  const base = {};
  Object.keys(entity || {})
    .filter(key => !ignoreKeys.has(key))
    .sort()
    .forEach(key => {
      base[key] = normalizeValue(entity[key]);
    });
  return JSON.stringify(base);
}

/**
 * Dedupe a collection by collapsing entities with identical content signatures.
 * @param {object[]} records Collection to dedupe.
 * @param {object} options Dedupe options.
 * @param {string[]} options.ignoreKeys Keys to exclude from signature comparisons.
 * @returns {{ records: object[], idMap: Map<string, string>, changed: boolean }} Dedupe result.
 */
function dedupeCollectionBySignature(records, { ignoreKeys = [] } = {}) {
  const signatureMap = new Map();
  const indexMap = new Map();
  const idMap = new Map();
  const updated = [];
  let changed = false;
  const ignoreSet = new Set(ignoreKeys);

  (records || []).forEach(record => {
    if (!record || record.deleted) {
      updated.push(record);
      return;
    }
    const signature = buildEntitySignature(record, ignoreSet);
    const existing = signatureMap.get(signature);
    if (!existing) {
      signatureMap.set(signature, record);
      indexMap.set(record.id, updated.length);
      updated.push(record);
      return;
    }

    const winner = isoNewer(record.updatedAt, existing.updatedAt) ? record : existing;
    const loser = winner === existing ? record : existing;
    idMap.set(loser.id, winner.id);
    changed = true;

    if (winner !== existing) {
      const index = indexMap.get(existing.id);
      if (index !== undefined) {
        updated[index] = winner;
        indexMap.delete(existing.id);
        indexMap.set(winner.id, index);
      }
      signatureMap.set(signature, winner);
    }
  });

  return { records: updated, idMap, changed };
}

/**
 * Enforces a single canonical default person, remapping any duplicates to one id.
 * @param {object[]} people People collection to normalize.
 * @returns {{ people: object[], idMap: Map<string, string>, changed: boolean, defaultPerson: object|null }}
 */
function enforceSingleDefaultPerson(people) {
  const alivePeople = (people || []).filter(p => p && !p.deleted);
  const candidates = alivePeople.filter(person =>
    person.isDefault || (person.name || "").trim().toLowerCase() === DEFAULT_PERSON_NAME.toLowerCase()
  );
  const idMap = new Map();
  let changed = false;

  if (!candidates.length) {
    return { people, idMap, changed, defaultPerson: null };
  }

  let winner = candidates[0];
  candidates.forEach(candidate => {
    if (candidate.isDefault && !winner.isDefault) {
      winner = candidate;
      return;
    }
    if (candidate.isDefault === winner.isDefault && isoNewer(candidate.updatedAt, winner.updatedAt)) {
      winner = candidate;
    }
  });

  const filtered = (people || []).filter(person => {
    if (!person || person.deleted) return true;
    const isDefaultCandidate = candidates.includes(person);
    if (!isDefaultCandidate) return true;
    if (person.id === winner.id) return true;
    idMap.set(person.id, winner.id);
    changed = true;
    return false;
  });

  let winnerChanged = false;
  if (winner.name !== DEFAULT_PERSON_NAME) {
    winner.name = DEFAULT_PERSON_NAME;
    winnerChanged = true;
  }
  if (winner.email) {
    winner.email = "";
    winnerChanged = true;
  }
  if (winner.organisation) {
    winner.organisation = "";
    winnerChanged = true;
  }
  if (winner.jobTitle) {
    winner.jobTitle = "";
    winnerChanged = true;
  }
  if (!winner.isDefault) {
    winner.isDefault = true;
    winnerChanged = true;
  }
  if (winner.deleted) {
    winner.deleted = false;
    winnerChanged = true;
  }
  if (winnerChanged) {
    winner.updatedAt = nowIso();
    changed = true;
  }

  return { people: filtered, idMap, changed, defaultPerson: winner };
}

/**
 * Remaps update-status entries when person ids are deduped.
 * @param {object} updateStatus Update-status map keyed by person id.
 * @param {Map<string, string>} idMap Person id remapping.
 * @returns {object} Remapped update-status object.
 */
function remapUpdateStatus(updateStatus, idMap) {
  if (!updateStatus || typeof updateStatus !== "object") return updateStatus;
  const out = {};
  Object.entries(updateStatus).forEach(([personId, status]) => {
    const mappedId = idMap.get(personId) || personId;
    const existing = out[mappedId];
    if (!existing) {
      out[mappedId] = status;
      return;
    }
    if (!existing.updated && status.updated) {
      out[mappedId] = status;
      return;
    }
    if (existing.updated && status.updated && isoNewer(status.updatedAt, existing.updatedAt)) {
      out[mappedId] = status;
    }
  });
  return out;
}

/**
 * Applies a person id remap across every collection that references people.
 * @param {object} targetDb Database instance to update.
 * @param {Map<string, string>} idMap Person id remapping table.
 * @returns {boolean} Whether any updates were applied.
 */
function remapPersonReferences(targetDb, idMap) {
  if (!targetDb || !idMap.size) return false;
  let changed = false;

  const remapId = (value) => idMap.get(value) || value;
  const remapArray = (values) => Array.from(new Set((values || []).map(remapId)));

  (targetDb.items || []).forEach(item => {
    if (!item) return;
    let itemChanged = false;
    const mappedOwner = item.ownerId ? remapId(item.ownerId) : item.ownerId;
    if (mappedOwner !== item.ownerId) {
      item.ownerId = mappedOwner;
      itemChanged = true;
    }
    const mappedTargets = remapArray(item.updateTargets);
    if (JSON.stringify(mappedTargets) !== JSON.stringify(item.updateTargets || [])) {
      item.updateTargets = mappedTargets;
      itemChanged = true;
    }
    const mappedStatus = remapUpdateStatus(item.updateStatus, idMap);
    if (mappedStatus !== item.updateStatus) {
      item.updateStatus = mappedStatus;
      itemChanged = true;
    }
    if (itemChanged) {
      item.updatedAt = nowIso();
      changed = true;
    }
  });

  (targetDb.tasks || []).forEach(task => {
    if (!task) return;
    let taskChanged = false;
    const mappedOwner = task.ownerId ? remapId(task.ownerId) : task.ownerId;
    if (mappedOwner !== task.ownerId) {
      task.ownerId = mappedOwner;
      taskChanged = true;
    }
    const mappedTargets = remapArray(task.updateTargets);
    if (JSON.stringify(mappedTargets) !== JSON.stringify(task.updateTargets || [])) {
      task.updateTargets = mappedTargets;
      taskChanged = true;
    }
    const mappedStatus = remapUpdateStatus(task.updateStatus, idMap);
    if (mappedStatus !== task.updateStatus) {
      task.updateStatus = mappedStatus;
      taskChanged = true;
    }
    if (taskChanged) {
      task.updatedAt = nowIso();
      changed = true;
    }
  });

  (targetDb.groups || []).forEach(group => {
    if (!group) return;
    const mappedMembers = remapArray(group.memberIds);
    if (JSON.stringify(mappedMembers) !== JSON.stringify(group.memberIds || [])) {
      group.memberIds = mappedMembers;
      group.updatedAt = nowIso();
      changed = true;
    }
  });

  (targetDb.meetings || []).forEach(meeting => {
    if (!meeting) return;
    let meetingChanged = false;
    const mappedPerson = meeting.oneToOnePersonId ? remapId(meeting.oneToOnePersonId) : meeting.oneToOnePersonId;
    if (mappedPerson !== meeting.oneToOnePersonId) {
      meeting.oneToOnePersonId = mappedPerson;
      meetingChanged = true;
    }
    if (meetingChanged) {
      meeting.updatedAt = nowIso();
      changed = true;
    }
  });

  return changed;
}

/**
 * Applies entity id remaps across collections that reference topics, meetings, templates, or groups.
 * @param {object} targetDb Database instance to update.
 * @param {object} maps Map bundle keyed by entity type.
 * @returns {boolean} Whether any updates were applied.
 */
function remapEntityReferences(targetDb, maps) {
  if (!targetDb) return false;
  let changed = false;

  const remapId = (map, value) => (map && value ? (map.get(value) || value) : value);

  (targetDb.meetings || []).forEach(meeting => {
    if (!meeting) return;
    let meetingChanged = false;
    const mappedTopic = remapId(maps.topics, meeting.topicId);
    const mappedTemplate = remapId(maps.templates, meeting.templateId);
    const mappedGroup = remapId(maps.groups, meeting.oneToOneGroupId);
    if (mappedTopic !== meeting.topicId) {
      meeting.topicId = mappedTopic;
      meetingChanged = true;
    }
    if (mappedTemplate !== meeting.templateId) {
      meeting.templateId = mappedTemplate;
      meetingChanged = true;
    }
    if (mappedGroup !== meeting.oneToOneGroupId) {
      meeting.oneToOneGroupId = mappedGroup;
      meetingChanged = true;
    }
    if (meetingChanged) {
      meeting.updatedAt = nowIso();
      changed = true;
    }
  });

  (targetDb.items || []).forEach(item => {
    if (!item) return;
    let itemChanged = false;
    const mappedTopic = remapId(maps.topics, item.topicId);
    const mappedMeeting = remapId(maps.meetings, item.meetingId);
    if (mappedTopic !== item.topicId) {
      item.topicId = mappedTopic;
      itemChanged = true;
    }
    if (mappedMeeting !== item.meetingId) {
      item.meetingId = mappedMeeting;
      itemChanged = true;
    }
    if (itemChanged) {
      item.updatedAt = nowIso();
      changed = true;
    }
  });

  (targetDb.tasks || []).forEach(task => {
    if (!task) return;
    let taskChanged = false;
    const mappedTopic = remapId(maps.topics, task.topicId);
    const mappedMeeting = remapId(maps.meetings, task.meetingId);
    if (mappedTopic !== task.topicId) {
      task.topicId = mappedTopic;
      taskChanged = true;
    }
    if (mappedMeeting !== task.meetingId) {
      task.meetingId = mappedMeeting;
      taskChanged = true;
    }
    if (taskChanged) {
      task.updatedAt = nowIso();
      changed = true;
    }
  });

  return changed;
}

/**
 * Dedupe merged entities and remap references to guarantee unique records.
 * @param {object} targetDb Database instance to sanitize.
 * @returns {boolean} Whether any updates were applied.
 */
function normalizeDuplicateEntities(targetDb) {
  if (!targetDb) return false;
  let changed = false;
  const ignoreKeys = ["id", "updatedAt", "createdAt", "deleted"];

  const defaultResult = enforceSingleDefaultPerson(targetDb.people || []);
  if (defaultResult.changed) {
    targetDb.people = defaultResult.people;
    changed = true;
  }

  const peopleDeduped = dedupeCollectionBySignature(targetDb.people || [], { ignoreKeys });
  if (peopleDeduped.changed) {
    targetDb.people = peopleDeduped.records;
    changed = true;
  }

  const groupDeduped = dedupeCollectionBySignature(targetDb.groups || [], { ignoreKeys });
  if (groupDeduped.changed) {
    targetDb.groups = groupDeduped.records;
    changed = true;
  }

  const topicDeduped = dedupeCollectionBySignature(targetDb.topics || [], { ignoreKeys });
  if (topicDeduped.changed) {
    targetDb.topics = topicDeduped.records;
    changed = true;
  }

  const templateDeduped = dedupeCollectionBySignature(targetDb.templates || [], { ignoreKeys });
  if (templateDeduped.changed) {
    targetDb.templates = templateDeduped.records;
    changed = true;
  }

  const meetingDeduped = dedupeCollectionBySignature(targetDb.meetings || [], { ignoreKeys });
  if (meetingDeduped.changed) {
    targetDb.meetings = meetingDeduped.records;
    changed = true;
  }

  const itemDeduped = dedupeCollectionBySignature(targetDb.items || [], { ignoreKeys });
  if (itemDeduped.changed) {
    targetDb.items = itemDeduped.records;
    changed = true;
  }

  const taskDeduped = dedupeCollectionBySignature(targetDb.tasks || [], { ignoreKeys });
  if (taskDeduped.changed) {
    targetDb.tasks = taskDeduped.records;
    changed = true;
  }

  const personMap = new Map([
    ...defaultResult.idMap.entries(),
    ...peopleDeduped.idMap.entries(),
  ]);
  if (remapPersonReferences(targetDb, personMap)) {
    changed = true;
  }

  if (remapEntityReferences(targetDb, {
    topics: topicDeduped.idMap,
    templates: templateDeduped.idMap,
    meetings: meetingDeduped.idMap,
    groups: groupDeduped.idMap,
  })) {
    changed = true;
  }

  return changed;
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
  const lSettings = l.settings || buildDefaultSettings();
  const rSettings = r.settings || buildDefaultSettings();

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

  if (normalizeDuplicateEntities(merged)) {
    merged.updatedAt = nowIso();
  }

  if (normalizeBuiltinTemplates(merged)) {
    merged.updatedAt = nowIso();
  }

  if (normalizeSettings(merged)) {
    merged.updatedAt = nowIso();
  }

  // Guarantee the default person exists after merges for consistent cross-module linking.
  const defaultResult = ensureDefaultPersonInDb(merged);
  if (defaultResult.changed) {
    merged.updatedAt = nowIso();
  }

  if (normalizeProjectData(merged)) {
    merged.updatedAt = nowIso();
  }

  if (normalizeTaskAndActionData(merged)) {
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

  if (normalizeSettings(db)) {
    markDirty();
    await saveLocal();
  }
  if (!db.tasks) {
    db.tasks = [];
  }

  let defaultPersonResult = ensureDefaultPersonInDb(db);

  if (normalizeBuiltinTemplates(db)) {
    markDirty();
    await saveLocal();
  }

  if (defaultPersonResult.changed) {
    markDirty();
    await saveLocal();
  }

  if (normalizeProjectData(db)) {
    markDirty();
    await saveLocal();
  }

  if (normalizeTaskAndActionData(db)) {
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

/**
 * Determines whether a silent auto-sync can run without user prompts.
 * @returns {boolean} True when auth, connectivity, and sync state allow auto-sync.
 */
function canAutoSync() {
  return driveReady && navigator.onLine && !syncInProgress;
}

/**
 * Requests a silent auto-sync when local changes exist and conditions allow it.
 * @param {string} trigger Describes the action that initiated the auto-sync request.
 */
function requestAutoSync(trigger) {
  if (!hasUnsyncedChanges) return;
  if (!canAutoSync()) return;
  syncNow({ silent: true, trigger }).catch((error) => {
    console.error("Auto-sync failed.", { trigger, error });
  });
}

/**
 * Starts or restarts the periodic auto-sync timer.
 */
function startAutoSyncTimer() {
  if (autoSyncTimerId) {
    clearInterval(autoSyncTimerId);
  }
  autoSyncTimerId = setInterval(() => {
    requestAutoSync("interval");
  }, AUTO_SYNC_INTERVAL_MS);
}

/**
 * Deep-clones a database snapshot to avoid mutating the original input.
 * @param {object} sourceDb Database snapshot to clone.
 * @returns {object} Cloned database snapshot.
 */
function cloneDbSnapshot(sourceDb) {
  if (!sourceDb) return makeDefaultDb();
  if (typeof structuredClone === "function") {
    return structuredClone(sourceDb);
  }
  return JSON.parse(JSON.stringify(sourceDb));
}

/**
 * Safely parses a timestamp value for sync comparison.
 * @param {string|null|undefined} value Timestamp string.
 * @returns {number} Milliseconds since epoch (0 when invalid).
 */
function parseSyncTimestamp(value) {
  const parsed = Date.parse(value || 0);
  return Number.isNaN(parsed) ? 0 : parsed;
}

/**
 * Selects the winning snapshot for the first-ever sync based on recency.
 * @param {object} localDb Local database snapshot.
 * @param {object} remoteDb Remote database snapshot.
 * @param {string|null} remoteModifiedTime Drive-modified timestamp fallback.
 * @returns {{ source: "local"|"remote", snapshot: object }} Winner + snapshot.
 */
function pickInitialSyncSnapshot(localDb, remoteDb, remoteModifiedTime) {
  const localTimestamp = parseSyncTimestamp(localDb?.updatedAt);
  const remoteTimestamp = Math.max(
    parseSyncTimestamp(remoteDb?.updatedAt),
    parseSyncTimestamp(remoteModifiedTime)
  );
  const useRemote = remoteTimestamp > localTimestamp;
  return {
    source: useRemote ? "remote" : "local",
    snapshot: useRemote ? remoteDb : localDb,
  };
}

/**
 * Normalizes a snapshot after selection to keep schema invariants intact.
 * @param {object} snapshot Selected snapshot to normalize.
 * @returns {object} Normalized snapshot ready for persistence.
 */
function normalizeSyncSnapshot(snapshot) {
  const normalized = cloneDbSnapshot(snapshot);
  let changed = false;

  if (normalizeDuplicateEntities(normalized)) {
    changed = true;
  }
  if (normalizeBuiltinTemplates(normalized)) {
    changed = true;
  }
  if (normalizeSettings(normalized)) {
    changed = true;
  }
  const defaultResult = ensureDefaultPersonInDb(normalized);
  if (defaultResult.changed) {
    changed = true;
  }
  if (normalizeProjectData(normalized)) {
    changed = true;
  }
  if (normalizeTaskAndActionData(normalized)) {
    changed = true;
  }

  if (changed) {
    normalized.updatedAt = nowIso();
  }

  return normalized;
}

/**
 * Runs a sync cycle with optional silent mode for auto-sync.
 * @param {{ silent?: boolean, trigger?: string }} [options] Optional sync controls.
 */
async function syncNow(options = {}) {
  const { silent = false, trigger = "manual" } = options || {};
  if (!driveReady) {
    if (!silent) {
      alert("Sign in first.");
    }
    return;
  }
  if (!navigator.onLine) {
    if (!silent) {
      alert("You appear to be offline. Sync will work when you're online.");
    }
    return;
  }
  if (syncInProgress) return;

  syncInProgress = true;
  setSyncStatus(silent ? "Auto-syncing…" : "Syncing…", "accent");
  byId("sync_btn").disabled = true;

  try {
    const fileId = await ensureDriveFile();

    // Download remote
    const remote = await downloadDriveJson(fileId);
    await getDriveFileMeta(fileId);

    // For the very first sync, keep only the most recent snapshot (local or remote).
    const isInitialSync = !lastSyncAt;
    let merged = null;

    if (isInitialSync) {
      const { source, snapshot } = pickInitialSyncSnapshot(db, remote, lastRemoteModifiedTime);
      merged = normalizeSyncSnapshot(snapshot);
      console.info("Initial sync snapshot selected.", { source });
    } else {
      // Merge once both sides have already been aligned by an initial sync.
      merged = mergeDb(db, remote);
    }

    // Upload merged
    await uploadDriveJson(fileId, merged);

    // Save merged locally
    db = merged;
    lastSyncAt = nowIso();
    await saveLocal();
    await saveMeta();
    markClean();

    renderAll();
    setSyncStatus(
      `${silent ? "Auto-synced" : "Synced"} ${fmtDateTime(lastSyncAt)}`,
      "ok"
    );
  } catch (e) {
    console.error(e);
    setSyncStatus(silent ? "Auto-sync failed" : "Sync failed", silent ? "warn" : "bad");
    if (!silent) {
      alert("Sync failed. Check console for details.");
    }
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

/**
 * Determines whether a person is the built-in default entity.
 * @param {object|null} person Person record to check.
 * @returns {boolean} Whether the person is the default "Me" entity.
 */
function isDefaultPerson(person) {
  if (!person) return false;
  if (person.isDefault) return true;
  return (person.name || "").trim().toLowerCase() === DEFAULT_PERSON_NAME.toLowerCase();
}

/**
 * Ensures the default person exists in the provided database and is not deleted.
 * @param {object} targetDb Database instance to normalize.
 * @returns {{ person: object, changed: boolean }} Normalized person and change flag.
 */
function ensureDefaultPersonInDb(targetDb) {
  if (!targetDb) return { person: null, changed: false };
  if (!targetDb.people) targetDb.people = [];

  const people = targetDb.people;
  const flagged = people.find(p => p.isDefault);
  const named = people.find(p => (p.name || "").trim().toLowerCase() === DEFAULT_PERSON_NAME.toLowerCase());
  const person = flagged || named;

  if (!person) {
    const created = createDefaultPersonRecord();
    people.push(created);
    return { person: created, changed: true };
  }

  let changed = false;
  if (person.name !== DEFAULT_PERSON_NAME) {
    person.name = DEFAULT_PERSON_NAME;
    changed = true;
  }
  if (person.email) {
    person.email = "";
    changed = true;
  }
  if (person.organisation) {
    person.organisation = "";
    changed = true;
  }
  if (person.jobTitle) {
    person.jobTitle = "";
    changed = true;
  }
  if (person.deleted) {
    person.deleted = false;
    changed = true;
  }
  if (!person.isDefault) {
    person.isDefault = true;
    changed = true;
  }
  if (changed) {
    person.updatedAt = nowIso();
  }

  return { person, changed };
}

/**
 * Retrieves the default person identifier from a provided database instance.
 * @param {object} targetDb Database instance to scan.
 * @returns {string} Default person id, or empty string if unavailable.
 */
function getDefaultPersonIdFromDb(targetDb) {
  const people = alive(targetDb?.people || []);
  const defaultPerson = people.find(p => p.isDefault)
    || people.find(p => (p.name || "").trim().toLowerCase() === DEFAULT_PERSON_NAME.toLowerCase());
  return defaultPerson?.id || "";
}

/**
 * Ensures project records include the latest schema fields and valid owners.
 * @param {object} targetDb Database instance to normalize.
 * @returns {boolean} Whether any records were changed.
 */
function normalizeProjectData(targetDb) {
  if (!targetDb) return false;
  let changed = false;
  if (!Array.isArray(targetDb.topics)) {
    targetDb.topics = [];
    return true;
  }

  const defaultOwnerId = getDefaultPersonIdFromDb(targetDb);
  const people = alive(targetDb.people || []);

  alive(targetDb.topics).forEach(topic => {
    if (!topic) return;
    let topicChanged = false;

    if (!topic.name) {
      topic.name = "Untitled project";
      topicChanged = true;
    }
    if (!("ownerId" in topic) || !topic.ownerId) {
      topic.ownerId = defaultOwnerId || "";
      topicChanged = true;
    } else if (!people.some(person => person.id === topic.ownerId)) {
      topic.ownerId = defaultOwnerId || "";
      topicChanged = true;
    }
    if (!("startDate" in topic)) {
      topic.startDate = "";
      topicChanged = true;
    }
    if (!("endDate" in topic)) {
      topic.endDate = "";
      topicChanged = true;
    }
    if (!("color" in topic) || !topic.color) {
      topic.color = DEFAULT_PROJECT_COLOR;
      topicChanged = true;
    }
    if (!topic.createdAt) {
      topic.createdAt = topic.updatedAt || nowIso();
      topicChanged = true;
    }
    if (!topic.updatedAt) {
      topic.updatedAt = nowIso();
      topicChanged = true;
    }

    if (topicChanged) {
      topic.updatedAt = nowIso();
      changed = true;
    }
  });

  return changed;
}

/**
 * Ensures tasks and meeting action items share the same schema fields.
 * @param {object} targetDb Database instance to normalize.
 * @returns {boolean} Whether any records were changed.
 */
function normalizeTaskAndActionData(targetDb) {
  if (!targetDb) return false;
  let changed = false;
  const defaultOwnerId = getDefaultPersonIdFromDb(targetDb);

  const normalizeCommonFields = (record, defaults) => {
    Object.entries(defaults).forEach(([key, value]) => {
      if (!(key in record)) {
        record[key] = value;
        changed = true;
      }
    });
  };

  const normalizeUpdateTracking = (record) => {
    if (!Array.isArray(record.updateTargets)) {
      record.updateTargets = [];
      changed = true;
    }
    const cleanedTargets = record.updateTargets.filter(Boolean);
    if (cleanedTargets.length !== record.updateTargets.length) {
      record.updateTargets = cleanedTargets;
      changed = true;
    }
    const existingStatus = (record.updateStatus && typeof record.updateStatus === "object") ? record.updateStatus : {};
    const rebuiltStatus = buildUpdateStatusForTargets(record.updateTargets, existingStatus);
    if (JSON.stringify(existingStatus) !== JSON.stringify(rebuiltStatus)) {
      record.updateStatus = rebuiltStatus;
      changed = true;
    }
  };

  alive(targetDb.tasks || []).forEach(task => {
    normalizeCommonFields(task, {
      kind: "task",
      title: task.title || task.text || "",
      text: task.text || task.title || "",
      notes: task.notes || "",
      ownerId: task.ownerId || defaultOwnerId || null,
      dueDate: task.dueDate || "",
      priority: task.priority || "medium",
      status: task.status || "todo",
      link: task.link || "",
      meetingId: task.meetingId || null,
      topicId: task.topicId || null,
      section: task.section || "task",
      createdAt: task.createdAt || task.updatedAt || nowIso(),
      updatedAt: task.updatedAt || nowIso(),
    });
    normalizeUpdateTracking(task);
  });

  alive(targetDb.items || []).forEach(item => {
    normalizeCommonFields(item, {
      kind: "item",
      title: item.title || item.text || "",
      text: item.text || item.title || "",
      notes: item.notes || "",
      ownerId: item.ownerId || null,
      dueDate: item.dueDate || "",
      priority: item.priority || "medium",
      status: item.status || "",
      link: item.link || "",
      meetingId: item.meetingId || null,
      topicId: item.topicId || null,
      section: item.section || "action",
      createdAt: item.createdAt || item.updatedAt || nowIso(),
      updatedAt: item.updatedAt || nowIso(),
    });
    normalizeUpdateTracking(item);
  });

  return changed;
}

/**
 * Retrieves the default person record from the active database.
 * @returns {object|null} Default person record, if available.
 */
function getDefaultPerson() {
  const people = alive(db.people);
  return people.find(p => p.isDefault) || people.find(p => (p.name || "").trim().toLowerCase() === DEFAULT_PERSON_NAME.toLowerCase()) || null;
}

/**
 * Retrieves the default person identifier for linking across modules.
 * @returns {string} Default person id, or empty string if unavailable.
 */
function getDefaultPersonId() {
  return getDefaultPerson()?.id || "";
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
  const now = nowIso();
  const t = {
    id: uid("topic"),
    name,
    ownerId: getDefaultPersonId() || "",
    startDate: "",
    endDate: "",
    color: DEFAULT_PROJECT_COLOR,
    createdAt: now,
    updatedAt: now
  };
  db.topics.push(t);
  return t.id;
}

function ensurePerson(name) {
  const existing = alive(db.people).find(p => p.name.toLowerCase() === name.toLowerCase());
  if (existing) return existing.id;
  if (name.trim().toLowerCase() === DEFAULT_PERSON_NAME.toLowerCase()) {
    return getDefaultPersonId();
  }
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
  if (view === "notes" && !meetingNotesEnabled) {
    alert("Select a meeting from the calendar to open notes.");
    return;
  }
  meetingView = view;
  if (view === "setup") {
    // Reset the notes gate when returning to the schedule view.
    meetingNotesEnabled = false;
  }
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
  if (!tplSel) return;
  const templates = alive(db.templates);
  renderSelectOptions(tplSel, templates.map(t => ({ value:t.id, label:t.name })));
}

/**
 * Syncs settings form controls with the persisted settings values.
 */
function renderSettings() {
  document.querySelectorAll("[data-setting-path]").forEach(control => {
    if (document.activeElement === control) return;
    const path = control.dataset.settingPath;
    const currentValue = getSettingValue(path);
    if (typeof currentValue === "undefined") return;
    const type = control.dataset.settingType || control.type;
    if (type === "checkbox") {
      control.checked = Boolean(currentValue);
    } else {
      control.value = String(currentValue);
    }
  });
}

/**
 * Renders project (topic) selectors across the meetings and updates modules.
 */
function renderTopics() {
  const topicSel = byId("meeting_topic");
  const topicsSel = byId("topics_topic");
  const updatesProjectSel = byId("updates_project");
  const projectSelect = byId("project_select");
  const updateTopicSelect = byId("update_topic");

  const topics = alive(db.topics).sort((a,b)=>a.name.localeCompare(b.name));
  const opts = topics.map(t => ({ value:t.id, label:t.name }));

  renderSelectOptions(topicSel, opts, { placeholder: topics.length ? null : "No projects yet — add one" });
  renderSelectOptions(topicsSel, opts, { placeholder: topics.length ? "Choose a project…" : "No projects yet" });
  renderSelectOptions(updatesProjectSel, opts, { placeholder: "All projects" });
  renderSelectOptions(projectSelect, opts, { placeholder: topics.length ? "Choose a project…" : "No projects yet" });
  renderSelectOptions(updateTopicSelect, opts, { placeholder: topics.length ? "No project" : "No projects yet" });
}

function renderPeopleSelects() {
  const updatesSel = byId("updates_person");
  if (!updatesSel) return;
  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));
  renderSelectOptions(updatesSel, people.map(p => ({ value:p.id, label:p.name })), { placeholder: people.length ? "Choose a person…" : "No people yet" });
}

/**
 * Populates task creation owner and update target selectors.
 */
function renderTaskLightboxOptions() {
  const ownerSelect = byId("task_owner");
  const targetsList = byId("task_targets_list");
  if (!ownerSelect && !targetsList) return;

  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));

  if (ownerSelect) {
    const currentValue = ownerSelect.value || getDefaultPersonId();
    renderSelectOptions(
      ownerSelect,
      people.map(p => ({ value: p.id, label: p.name })),
      { placeholder: people.length ? "Unassigned" : "No people yet" }
    );
    if (currentValue) {
      ownerSelect.value = currentValue;
    }
  }

  if (targetsList) {
    targetsList.innerHTML = people.map(p => `<option value="${escapeHtml(p.name)}"></option>`).join("");
  }
}

/**
 * Populates the ad-hoc update lightbox selectors and target list.
 */
function renderUpdateLightboxOptions() {
  const ownerSelect = byId("update_owner");
  const targetsList = byId("update_targets_list");
  if (!ownerSelect && !targetsList) return;

  const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));

  if (ownerSelect) {
    const currentValue = ownerSelect.value || getDefaultPersonId();
    renderSelectOptions(
      ownerSelect,
      people.map(p => ({ value: p.id, label: p.name })),
      { placeholder: people.length ? "Unassigned" : "No people yet" }
    );
    if (currentValue) {
      ownerSelect.value = currentValue;
    }
  }

  if (targetsList) {
    targetsList.innerHTML = people.map(p => `<option value="${escapeHtml(p.name)}"></option>`).join("");
  }
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
    const isDefault = isDefaultPerson(p);
    const meta = [
      p.email ? escapeHtml(p.email) : null,
      p.organisation ? escapeHtml(p.organisation) : null,
      p.jobTitle ? escapeHtml(p.jobTitle) : null,
    ].filter(Boolean).join(" • ");
    const badgeLabel = isDefault ? "Default" : p.id;
    return `
      <div class="item ${p.id === personViewId ? "item--selected" : ""}">
        <div class="item__top">
          <div>
            <strong>${escapeHtml(p.name)}</strong>
            <div class="muted">${meta || "No details added yet."}</div>
          </div>
          <div class="badges"><span class="badge">${escapeHtml(badgeLabel)}</span></div>
        </div>
        <div class="item__actions">
          <button class="smallbtn" data-person-select="${escapeHtml(p.id)}">View</button>
          <button class="smallbtn smallbtn--danger" data-del-person="${escapeHtml(p.id)}" ${isDefault ? "disabled" : ""} title="${isDefault ? "The default person cannot be deleted." : "Delete this person"}">Delete</button>
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
      if (isDefaultPerson(p)) {
        alert("The default person cannot be deleted.");
        return;
      }
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
    const saveBtn = byId("person_save_btn");
    if (saveBtn) {
      saveBtn.disabled = true;
      saveBtn.title = "Select a person to edit.";
    }
    editor.innerHTML = `<div class="muted">Select a person to view or edit their details.</div>`;
    ownedList.innerHTML = `<div class="muted">Select a person to see owned updates.</div>`;
    targetList.innerHTML = `<div class="muted">Select a person to see update targets.</div>`;
    return;
  }

  const draft = personEditorState.draft;
  const activePerson = personViewId ? getPerson(personViewId) : null;
  const isDefault = activePerson ? isDefaultPerson(activePerson) : false;
  const lockHint = isDefault ? `<div class="muted">Default person details are locked.</div>` : "";
  const lockAttr = isDefault ? "disabled" : "";
  const saveBtn = byId("person_save_btn");
  if (saveBtn) {
    saveBtn.disabled = isDefault;
    saveBtn.title = isDefault ? "Default person details are locked." : "Save changes";
  }
  editor.innerHTML = `
    <div class="formrow">
      <label>Name ${fieldTag(true)}</label>
      <input id="person_name" type="text" value="${escapeHtml(draft.name)}" ${lockAttr} />
    </div>
    <div class="formrow">
      <label>Email ${fieldTag(true)}</label>
      <input id="person_email" type="email" value="${escapeHtml(draft.email)}" ${lockAttr} />
    </div>
    <div class="formrow">
      <label>Organisation ${fieldTag(true)}</label>
      <input id="person_org" type="text" value="${escapeHtml(draft.organisation)}" ${lockAttr} />
    </div>
    <div class="formrow">
      <label>Job title ${fieldTag(false)}</label>
      <input id="person_title" type="text" value="${escapeHtml(draft.jobTitle)}" ${lockAttr} />
    </div>
    ${lockHint}
    ${personEditorState.error ? `<div class="item__error">${escapeHtml(personEditorState.error)}</div>` : ""}
  `;

  if (personEditorState.isNew) {
    ownedList.innerHTML = `<div class="muted">Save the person to see owned updates.</div>`;
    targetList.innerHTML = `<div class="muted">Save the person to see update targets.</div>`;
    return;
  }

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
    { placeholder: "All projects" }
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
    <div class="muted">${escapeHtml(tpl?.name || "Template")} • ${escapeHtml(topic?.name || "No project")} • ${fmtDateTime(meeting.datetime)}${oneToOneLabel}</div>
  `;

  const oneToOneUpdatesCard = oneToOneContext ? renderOneToOneUpdatesCard(meeting, oneToOneContext) : "";

  area.innerHTML = `
    <h2>Meeting notes</h2>
    <div class="muted">Template: <strong>${escapeHtml(tpl?.name || "")}</strong> • Project: <strong>${escapeHtml(topic?.name || "")}</strong></div>
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
  const linkedItems = getItemsLinkedToPeople(context.personIds)
    .sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));
  // Only surface items that still need updates for the 1:1 counterpart.
  const pendingItems = linkedItems.filter(item => {
    const targets = new Set([item.ownerId, ...(item.updateTargets || [])].filter(Boolean));
    return context.personIds.some(personId => {
      if (!targets.has(personId)) return false;
      return !item.updateStatus?.[personId]?.updated;
    });
  });
  const memberNames = context.personIds.map(pid => getPerson(pid)?.name).filter(Boolean);
  const subtitle = context.type === "group"
    ? `${escapeHtml(context.label)} • ${escapeHtml(memberNames.join(", ") || "No members yet")}`
    : escapeHtml(context.label);
  const listHtml = pendingItems.map(it => renderItemCard(it)).join("")
    || `<div class="muted">No pending updates for this counterpart yet.</div>`;
  const itemIds = pendingItems.map(it => it.id);
  const pendingCount = pendingItems.length;

  return `
    <div class="sectioncard" data-one-to-one-section data-target-ids='${escapeHtml(JSON.stringify(context.personIds))}' data-item-ids='${escapeHtml(JSON.stringify(itemIds))}'>
      <div class="sectionhead">
        <div>
          <h3>1:1 updates</h3>
          <div class="muted">${subtitle}</div>
        </div>
        <div class="muted">${linkedItems.length} linked item(s) • ${pendingCount} pending update(s)</div>
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

            <div class="formrow">
              <label>Notes</label>
              <textarea data-field="notes" placeholder="Add optional context…"></textarea>
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
                <label>Priority</label>
                <select data-field="priority">
                  <option value="low">Low</option>
                  <option value="medium" selected>Medium</option>
                  <option value="high">High</option>
                </select>
              </div>

              <div class="formrow">
                <label>Due date</label>
                <input data-field="dueDate" type="date" />
              </div>
            </div>

            <div class="formrow">
              <label>Link</label>
              <input data-field="link" type="url" placeholder="https://…" />
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
      const notes = box.querySelector("textarea[data-field=notes]").value.trim();
      const ownerName = box.querySelector("input[data-field=ownerName]")?.value.trim() || "";
      const ownerMatch = ownerName ? findPersonByName(ownerName, people) : null;
      const ownerId = ownerMatch ? ownerMatch.id : null;
      const status = box.querySelector("select[data-field=status]").value || null;
      const priority = box.querySelector("select[data-field=priority]").value || "medium";
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
        title: text,
        notes,
        ownerId,
        status,
        dueDate,
        priority,
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
      box.querySelector("textarea[data-field=notes]").value = "";
      box.querySelector("input[data-field=ownerName]").value = "";
      box.querySelector("select[data-field=status]").value = "";
      box.querySelector("select[data-field=priority]").value = "medium";
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
    notes: it.notes || "",
    status: it.status || "",
    dueDate: it.dueDate || "",
    link: it.link || "",
    ownerName: owner?.name || "",
    priority: it.priority || "medium",
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
  const priorityBadge = it.priority ? `<span class="badge">Priority: ${escapeHtml(TASK_PRIORITY_LABELS[it.priority] || it.priority)}</span>` : "";

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

          <div class="formrow">
            <label>Notes</label>
            <textarea data-edit-field="notes" placeholder="Add optional notes…">${escapeHtml(draft.notes)}</textarea>
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
              <label>Priority</label>
              <select data-edit-field="priority">
                <option value="low"${draft.priority === "low" ? " selected" : ""}>Low</option>
                <option value="medium"${draft.priority === "medium" ? " selected" : ""}>Medium</option>
                <option value="high"${draft.priority === "high" ? " selected" : ""}>High</option>
              </select>
            </div>

            <div class="formrow">
              <label>Due date ${fieldTag(false)}</label>
              <input data-edit-field="dueDate" type="date" value="${escapeHtml(draft.dueDate)}" />
            </div>
          </div>

          <div class="formrow">
            <label>Link ${fieldTag(false)}</label>
            <input data-edit-field="link" type="url" placeholder="https://…" value="${escapeHtml(draft.link)}" />
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
          ${priorityBadge}
          ${updBadge}
          ${linkBadge}
        </div>
        <div class="muted">${meeting ? fmtDateTime(meeting.datetime) : ""}</div>
      </div>

      <div class="item__text">${escapeHtml(it.text)}</div>
      ${it.notes ? `<div class="item__notes">${escapeHtml(it.notes)}</div>` : ""}

      <div class="item__meta">
        ${topic ? `<span>Project: <strong>${escapeHtml(topic.name)}</strong></span>` : ""}
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
          notes: it.notes || "",
          status: it.status || "",
          dueDate: it.dueDate || "",
          link: it.link || "",
          ownerName: getPerson(it.ownerId)?.name || "",
          priority: it.priority || "medium",
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
      const notes = card.querySelector("[data-edit-field=notes]")?.value.trim() || "";
      const ownerName = card.querySelector("[data-edit-field=ownerName]")?.value.trim() || "";
      const ownerMatch = ownerName ? findPersonByName(ownerName) : null;
      const ownerId = ownerMatch ? ownerMatch.id : null;
      const status = card.querySelector("[data-edit-field=status]")?.value || null;
      const priority = card.querySelector("[data-edit-field=priority]")?.value || "medium";
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
            notes,
            status: status || "",
            dueDate: dueDate || "",
            link: link || "",
            ownerName: ownerName || "",
            priority,
          },
          error: errs.join(" "),
        });
        renderAll();
        return;
      }

      it.text = text;
      it.title = text;
      it.notes = notes;
      it.ownerId = ownerId;
      it.status = status;
      it.priority = priority;
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
      // Keep the user in the schedule view when opening from outside the calendar.
      meetingNotesEnabled = false;
      setMeetingView("setup");
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
      // Notes access is intentionally unlocked only from the calendar view.
      meetingNotesEnabled = true;
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

/**
 * Builds the calendar metadata lines for a meeting card.
 * @param {object} meeting Meeting record for the calendar.
 * @returns {string[]} Metadata lines with project, time, and 1:1 counterpart when applicable.
 */
function buildCalendarMeetingMetaLines(meeting) {
  const topic = getTopic(meeting.topicId)?.name || "No project";
  const time = formatTime(meeting.datetime) || "Time TBD";
  const template = getTemplate(meeting.templateId);
  const oneToOneContext = template?.id === ONE_TO_ONE_TEMPLATE_ID ? getOneToOneContext(meeting) : null;
  const counterpartLabel = oneToOneContext ? `1:1 with ${oneToOneContext.label}` : "";
  return [
    `Project: ${topic}`,
    `Time: ${time}`,
    counterpartLabel,
  ].filter(Boolean);
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
      const title = m.title || "Untitled meeting";
      const metaLines = buildCalendarMeetingMetaLines(m);
      const projectColor = normalizeHexColor(getTopic(m.topicId)?.color || DEFAULT_PROJECT_COLOR);
      const palette = buildProjectColorPalette(projectColor);
      return `
        <div class="calendar-meeting" style="--project-color:${escapeHtml(projectColor)}; --project-color-border:${escapeHtml(palette.border)}; --project-color-bg:${escapeHtml(palette.background)};">
          <div class="calendar-meeting__title">${escapeHtml(title)}</div>
          ${metaLines.map(line => `<div class="calendar-meeting__meta-line">${escapeHtml(line)}</div>`).join("")}
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
  // Project filter narrows updates to a specific project (topic).
  const projectId = byId("updates_project")?.value || "";
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
    if (projectId && it.topicId !== projectId) return false;
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
    actionsFilters.ownerId = getDefaultPersonId();
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

/**
 * Reads the current project details form values from the DOM.
 * @returns {object} Project draft fields.
 */
function readProjectDetailsForm() {
  return {
    name: byId("project_name")?.value || "",
    ownerId: byId("project_owner")?.value || "",
    startDate: byId("project_start")?.value || "",
    endDate: byId("project_end")?.value || "",
    color: byId("project_color")?.value || DEFAULT_PROJECT_COLOR,
  };
}

/**
 * Updates the project details error banner based on the latest validation output.
 * @param {string} message Validation error message to display.
 */
function setProjectDetailsError(message) {
  const errorEl = byId("project_details_error");
  if (!errorEl) return;
  if (!message) {
    errorEl.textContent = "";
    errorEl.hidden = true;
    return;
  }
  errorEl.textContent = message;
  errorEl.hidden = false;
}

/**
 * Renders the Projects module, including summary, details, timeline, decisions, and meetings.
 */
function renderProjectsModule() {
  const select = byId("project_select");
  const summary = byId("project_summary");
  const form = byId("project_details_form");
  const timeline = byId("project_actions_timeline");
  const decisions = byId("project_decision_log");
  const meetingsList = byId("project_meetings_list");

  if (!select || !summary || !form || !timeline || !decisions || !meetingsList) return;

  // Keep the selected project stable as options refresh.
  if (!projectViewId || !getTopic(projectViewId)) {
    projectViewId = select.value || "";
  }
  if (document.activeElement !== select) {
    select.value = projectViewId || "";
  }

  const project = projectViewId ? getTopic(projectViewId) : null;
  const saveBtn = byId("project_save_btn");
  const resetBtn = byId("project_reset_btn");

  if (!project) {
    if (saveBtn) saveBtn.disabled = true;
    if (resetBtn) resetBtn.disabled = true;
    summary.textContent = "Choose a project to view details.";
    form.innerHTML = `<div class="muted">Select a project to edit its details.</div>`;
    timeline.innerHTML = `<div class="muted">No project selected.</div>`;
    decisions.innerHTML = `<div class="muted">No project selected.</div>`;
    meetingsList.innerHTML = `<div class="muted">No project selected.</div>`;
    setProjectDetailsError("");
    return;
  }

  if (projectEditorState.projectId !== project.id || !projectEditorState.draft) {
    projectEditorState = { projectId: project.id, draft: createProjectDraft(project), error: "" };
  }

  if (saveBtn) saveBtn.disabled = false;
  if (resetBtn) resetBtn.disabled = false;

  const owner = project.ownerId ? getPerson(project.ownerId) : null;
  const summaryBits = [
    `Owner: ${owner?.name || "Unassigned"}`,
    project.startDate ? `Start: ${fmtDate(`${project.startDate}T00:00:00`)}` : "Start: Not set",
    project.endDate ? `End: ${fmtDate(`${project.endDate}T00:00:00`)}` : "End: Not set",
  ];
  summary.textContent = summaryBits.join(" • ");

  const draft = projectEditorState.draft;
  form.innerHTML = `
    <div class="formrow">
      <label>Project name ${fieldTag(true)}</label>
      <input id="project_name" type="text" value="${escapeHtml(draft.name)}" />
    </div>
    <div class="formrow">
      <label>Owner ${fieldTag(true)}</label>
      <select id="project_owner"></select>
    </div>
    <div class="grid2">
      <div class="formrow">
        <label>Project color</label>
        <input id="project_color" type="color" value="${escapeHtml(draft.color || DEFAULT_PROJECT_COLOR)}" />
      </div>
      <div class="formrow">
        <label>Start date</label>
        <input id="project_start" type="date" value="${escapeHtml(draft.startDate)}" />
      </div>
      <div class="formrow">
        <label>End date</label>
        <input id="project_end" type="date" value="${escapeHtml(draft.endDate)}" />
      </div>
    </div>
  `;

  const ownerSelect = byId("project_owner");
  if (ownerSelect) {
    const people = alive(db.people).sort((a,b)=>a.name.localeCompare(b.name));
    renderSelectOptions(
      ownerSelect,
      people.map(person => ({ value: person.id, label: person.name })),
      { placeholder: people.length ? "Choose an owner…" : "No people yet" }
    );
    ownerSelect.value = draft.ownerId || "";
  }

  setProjectDetailsError(projectEditorState.error);

  const syncDraftFromForm = () => {
    projectEditorState = { ...projectEditorState, draft: readProjectDetailsForm(), error: "" };
    setProjectDetailsError("");
  };

  byId("project_name")?.addEventListener("input", syncDraftFromForm);
  byId("project_owner")?.addEventListener("change", syncDraftFromForm);
  byId("project_color")?.addEventListener("input", syncDraftFromForm);
  byId("project_start")?.addEventListener("change", syncDraftFromForm);
  byId("project_end")?.addEventListener("change", syncDraftFromForm);

  const actionItems = alive(db.items)
    .filter(it => it.topicId === project.id && it.section === "action")
    .sort((a, b) => getActionTimelineDate(a) - getActionTimelineDate(b));

  timeline.innerHTML = actionItems.map(action => {
    const ownerName = action.ownerId ? getPerson(action.ownerId)?.name : "Unassigned";
    const meeting = action.meetingId ? getMeeting(action.meetingId) : null;
    const status = action.status || "open";
    const meetingLabel = meeting
      ? `${meeting.title || "Meeting"} • ${fmtDateTime(meeting.datetime)}`
      : "No meeting linked";

    return `
      <div class="timeline-item">
        <div class="timeline-item__date">${escapeHtml(formatActionTimelineLabel(action))}</div>
        <div class="timeline-item__card">
          <div class="item__top">
            <div><strong>${escapeHtml(action.text || "Untitled action")}</strong></div>
            <div class="badges">
              <span class="badge badge--accent">${escapeHtml(status)}</span>
              <span class="badge">${escapeHtml(ownerName || "Unassigned")}</span>
            </div>
          </div>
          ${action.notes ? `<div class="item__notes">${escapeHtml(action.notes)}</div>` : ""}
          <div class="timeline-item__meta">${escapeHtml(meetingLabel)}</div>
          ${meeting ? `<button class="smallbtn" data-project-open-meeting="${escapeHtml(meeting.id)}">Open meeting</button>` : ""}
        </div>
      </div>
    `;
  }).join("") || `<div class="muted">No action items yet.</div>`;

  const decisionItems = alive(db.items)
    .filter(it => it.topicId === project.id && it.section === "decision")
    .sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));
  decisions.innerHTML = decisionItems.map(it => renderItemCard(it)).join("")
    || `<div class="muted">No decisions logged yet.</div>`;
  wireItemButtons(decisions);

  const meetings = alive(db.meetings)
    .filter(m => m.topicId === project.id)
    .sort((a,b)=>Date.parse(b.datetime)-Date.parse(a.datetime));

  meetingsList.innerHTML = meetings.map(meeting => `
    <div class="item">
      <div class="item__top">
        <div>
          <strong>${escapeHtml(meeting.title || "Untitled meeting")}</strong>
          <div class="muted">${fmtDateTime(meeting.datetime)}</div>
        </div>
        <div class="badges">
          <span class="badge">${escapeHtml(meeting.id)}</span>
        </div>
      </div>
      <div class="item__actions">
        <button class="smallbtn" data-project-open-meeting="${escapeHtml(meeting.id)}">Open meeting</button>
      </div>
    </div>
  `).join("") || `<div class="muted">No meetings linked to this project yet.</div>`;

  timeline.querySelectorAll("[data-project-open-meeting]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const meetingId = btn.getAttribute("data-project-open-meeting");
      if (!meetingId) return;
      currentMeetingId = meetingId;
      await saveMeta();
      setActiveModule("meetings");
      setMeetingModuleTab("meeting");
      // Keep notes gated to calendar navigation.
      meetingNotesEnabled = false;
      setMeetingView("setup");
      renderAll();
    });
  });

  meetingsList.querySelectorAll("[data-project-open-meeting]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const meetingId = btn.getAttribute("data-project-open-meeting");
      if (!meetingId) return;
      currentMeetingId = meetingId;
      await saveMeta();
      setActiveModule("meetings");
      setMeetingModuleTab("meeting");
      // Keep notes gated to calendar navigation.
      meetingNotesEnabled = false;
      setMeetingView("setup");
      renderAll();
    });
  });
}

function renderTopicOverview() {
  const topicId = byId("topics_topic").value || "";
  const focus = byId("topics_focus").value || "overview";
  const out = byId("topic_output");

  if (!topicId) {
    out.innerHTML = `<div class="muted">Choose a project.</div>`;
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
          <strong>${escapeHtml(topic?.name || "Project")}</strong>
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
 * Builds a notes string for action items linked into the Tasks module.
 * @param {object} item Action item record.
 * @returns {string} Multi-line notes string for display.
 */
function buildActionTaskNotes(item) {
  const meeting = getMeeting(item.meetingId);
  const topic = getTopic(item.topicId);
  const targets = (item.updateTargets || [])
    .map(pid => getPerson(pid)?.name)
    .filter(Boolean);

  const lines = [];
  if (meeting) lines.push(`Meeting: ${meeting.title || "(Untitled)"}`);
  if (topic) lines.push(`Project: ${topic.name}`);
  if (targets.length) lines.push(`Update targets: ${targets.join(", ")}`);
  return lines.join("\n");
}

/**
 * Creates the combined task list, including linked action items owned by "Me".
 * @returns {object[]} Normalized task views for rendering.
 */
function buildTaskViews() {
  const tasks = alive(db.tasks).map(task => ({
    id: task.id,
    sourceType: "task",
    title: task.title,
    text: task.text || task.title || "",
    notes: task.notes || "",
    ownerId: task.ownerId || "",
    dueDate: task.dueDate || "",
    priority: task.priority || "medium",
    status: task.status || "todo",
    link: task.link || "",
    updateTargets: Array.isArray(task.updateTargets) ? task.updateTargets : [],
    updateStatus: task.updateStatus || {},
    updatedAt: task.updatedAt,
    createdAt: task.createdAt,
  }));

  const meId = getDefaultPersonId();
  const actionItems = meId
    ? alive(db.items).filter(item => item.section === "action" && item.ownerId === meId)
    : [];

  const linkedActions = actionItems.map(item => ({
    id: item.id,
    sourceType: "item",
    title: item.title || item.text || "Untitled action",
    text: item.text || item.title || "",
    notes: item.notes || "",
    ownerId: item.ownerId || "",
    dueDate: item.dueDate || "",
    priority: item.priority || "medium",
    status: mapActionStatusToTaskStatus(item.status),
    link: item.link || "",
    updateTargets: Array.isArray(item.updateTargets) ? item.updateTargets : [],
    updateStatus: item.updateStatus || {},
    context: buildActionTaskNotes(item),
    updatedAt: item.updatedAt,
    createdAt: item.createdAt,
    meetingId: item.meetingId,
    topicId: item.topicId,
  }));

  return [...tasks, ...linkedActions];
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

  const taskViews = buildTaskViews();
  const filtered = taskViews.filter(task => {
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
    list.innerHTML = `<div class="muted">No tasks or linked action items yet. Add your first task to get started.</div>`;
    return;
  }

  list.innerHTML = `
    <div class="task-table-wrapper">
      <div class="task-table">
        <div class="task-row task-row--header">
          <div>Task</div>
          <div>Notes</div>
          <div>Status</div>
          <div>Priority</div>
          <div>Due date</div>
          <div>Actions</div>
        </div>
        <div class="task-table__body">
          ${sorted.map(task => renderTaskCard(task)).join("")}
        </div>
      </div>
    </div>
  `;
  wireTaskList(list);
}

/**
 * Builds a task row for the Tasks module list.
 * @param {object} task Task record.
 * @returns {string} Task row markup.
 */
function renderTaskCard(task) {
  const statusKey = task.status || "todo";
  const priorityKey = task.priority || "medium";
  const isLinkedAction = task.sourceType === "item";
  const meeting = task.meetingId ? getMeeting(task.meetingId) : null;
  const topic = task.topicId ? getTopic(task.topicId) : null;
  const sourceMeta = isLinkedAction ? "Linked action" : "Personal task";
  const contextMeta = isLinkedAction
    ? ""
    : [
      meeting ? `Meeting: ${meeting.title || "Untitled"}` : null,
      topic ? `Project: ${topic.name}` : null,
    ].filter(Boolean).join(" • ");
  // Hidden columns remain in data, but are intentionally not shown in the condensed layout.

  return `
    <div class="task-row" data-task-id="${task.id}" data-task-source="${escapeHtml(task.sourceType || "task")}">
      <div class="task-cell">
        <input class="task-input" type="text" value="${escapeHtml(task.title || "")}" data-task-field="title" />
        <div class="task-row__meta">
          <span class="task-row__meta-label">${escapeHtml(sourceMeta)}</span>
          ${contextMeta ? `<span class="task-row__meta-context">${escapeHtml(contextMeta)}</span>` : ""}
          ${task.context ? `<span class="task-row__meta-context">${escapeHtml(task.context)}</span>` : ""}
        </div>
      </div>
      <div class="task-cell">
        <textarea class="task-input" data-task-field="notes" placeholder="Add supporting notes...">${escapeHtml(task.notes || "")}</textarea>
      </div>
      <div class="task-cell">
        <select class="task-status task-status--${statusKey}" data-task-field="status">
          <option value="todo"${statusKey === "todo" ? " selected" : ""}>To do</option>
          <option value="in_progress"${statusKey === "in_progress" ? " selected" : ""}>In progress</option>
          <option value="blocked"${statusKey === "blocked" ? " selected" : ""}>Blocked</option>
          <option value="done"${statusKey === "done" ? " selected" : ""}>Done</option>
        </select>
      </div>
      <div class="task-cell">
        <select class="task-priority task-priority--${priorityKey}" data-task-field="priority">
          <option value="low"${priorityKey === "low" ? " selected" : ""}>Low</option>
          <option value="medium"${priorityKey === "medium" ? " selected" : ""}>Medium</option>
          <option value="high"${priorityKey === "high" ? " selected" : ""}>High</option>
        </select>
      </div>
      <div class="task-cell">
        <input class="task-input" type="date" value="${escapeHtml(task.dueDate || "")}" data-task-field="dueDate" />
      </div>
      <div class="task-cell task-cell--actions">
        ${isLinkedAction && task.meetingId ? `<button class="smallbtn" data-task-action="open-meeting" data-task-meeting="${escapeHtml(task.meetingId)}">Open meeting</button>` : ""}
        <button class="smallbtn smallbtn--danger" data-task-action="delete">Delete</button>
      </div>
    </div>
  `;
}

/**
 * Wires click handlers for task rows.
 * @param {HTMLElement} container Task list container.
 */
function wireTaskList(container) {
  /**
   * Updates a task or linked action item record, then schedules persistence.
   * @param {string} taskId Task or item id.
   * @param {string} sourceType "task" or "item".
   * @param {object} updates Key/value field updates.
   */
  const applyTaskUpdates = (taskId, sourceType, updates) => {
    if (sourceType === "item") {
      const item = getItem(taskId);
      if (!item) return;
      if (typeof updates.title === "string") {
        item.title = updates.title;
        item.text = updates.title;
      }
      if (typeof updates.notes === "string") item.notes = updates.notes;
      if (typeof updates.ownerId === "string") item.ownerId = updates.ownerId || null;
      if (typeof updates.dueDate === "string") item.dueDate = updates.dueDate;
      if (typeof updates.priority === "string") item.priority = updates.priority;
      if (typeof updates.status === "string") item.status = mapTaskStatusToActionStatus(updates.status);
      if (typeof updates.link === "string") item.link = updates.link;
      if (Array.isArray(updates.updateTargets)) {
        item.updateTargets = updates.updateTargets;
        item.updateStatus = buildUpdateStatusForTargets(updates.updateTargets, item.updateStatus);
      }
      item.updatedAt = nowIso();
      return;
    }

    const task = getTask(taskId);
    if (!task) return;
    if (typeof updates.title === "string") {
      task.title = updates.title;
      task.text = updates.title;
    }
    if (typeof updates.notes === "string") task.notes = updates.notes;
    if (typeof updates.ownerId === "string") task.ownerId = updates.ownerId || null;
    if (typeof updates.dueDate === "string") task.dueDate = updates.dueDate;
    if (typeof updates.priority === "string") task.priority = updates.priority;
    if (typeof updates.status === "string") task.status = updates.status;
    if (typeof updates.link === "string") task.link = updates.link;
    if (Array.isArray(updates.updateTargets)) {
      task.updateTargets = updates.updateTargets;
      task.updateStatus = buildUpdateStatusForTargets(updates.updateTargets, task.updateStatus);
    }
    task.updatedAt = nowIso();
  };

  /**
   * Schedules a throttled persistence write to keep typing responsive.
   */
  const persistTaskChanges = debounce(async () => {
    markDirty();
    await saveLocal();
  }, 350);

  /**
   * Updates card-level UI elements after edits without forcing a full re-render.
   * @param {HTMLElement} card Task card element.
   * @param {object} updates Updated field values.
   */
  const syncTaskCardDisplay = (card, updates) => {
    if (updates.status !== undefined) {
      const statusKey = updates.status || "todo";
      const statusSelect = card.querySelector("[data-task-field='status']");
      if (statusSelect) statusSelect.className = `task-status task-status--${statusKey}`;
    }

    if (updates.priority !== undefined) {
      const priorityKey = updates.priority || "medium";
      const prioritySelect = card.querySelector("[data-task-field='priority']");
      if (prioritySelect) prioritySelect.className = `task-priority task-priority--${priorityKey}`;
    }

    if (updates.updateTargets !== undefined) {
      const progress = card.querySelector("[data-task-display='update-progress']");
      if (progress) {
        const updateStatus = updates.updateStatus || buildUpdateStatusForTargets(updates.updateTargets);
        progress.textContent = buildUpdateProgressLabel({ updateTargets: updates.updateTargets, updateStatus });
      }
    }
  };

  // Inline edits for all editable task fields (title, notes, owner, status, priority, due date, link, updates).
  container.querySelectorAll("[data-task-field]").forEach(field => {
    const fieldName = field.getAttribute("data-task-field");
    if (!fieldName) return;

    const eventName = fieldName === "updateTargets"
      ? "change"
      : (field.tagName === "SELECT" ? "change" : "input");
    field.addEventListener(eventName, () => {
      const card = field.closest("[data-task-id]");
      if (!card) return;
      const taskId = card.getAttribute("data-task-id");
      const sourceType = card.getAttribute("data-task-source") || "task";

      let value = field.value ?? "";

      if (fieldName === "updateTargets") {
        const people = alive(db.people);
        const { ids, missing } = parseTargetNames(value, people);
        if (missing.length) {
          alert(`Unknown people: ${missing.join(", ")}.`);
          field.value = field.getAttribute("data-task-targets") || "";
          return;
        }
        const formatted = formatTargetNames(ids);
        field.value = formatted;
        field.setAttribute("data-task-targets", formatted);
        applyTaskUpdates(taskId, sourceType, { updateTargets: ids });
        syncTaskCardDisplay(card, { updateTargets: ids, updateStatus: buildUpdateStatusForTargets(ids) });
        persistTaskChanges();
        return;
      }

      applyTaskUpdates(taskId, sourceType, { [fieldName]: value });
      syncTaskCardDisplay(card, { [fieldName]: value });
      persistTaskChanges();
    });
  });

  container.querySelectorAll("[data-task-action='delete']").forEach(btn => {
    btn.addEventListener("click", async () => {
      const card = btn.closest("[data-task-id]");
      if (!card) return;
      const taskId = card.getAttribute("data-task-id");
      const sourceType = card.getAttribute("data-task-source") || "task";

      if (sourceType === "item") {
        const item = getItem(taskId);
        if (!item) return;
        const confirmDelete = confirm("Delete this linked action item? It will be removed from meeting notes.");
        if (!confirmDelete) return;
        item.deleted = true;
        item.updatedAt = nowIso();
      } else {
        const task = getTask(taskId);
        if (!task) return;
        const confirmDelete = confirm("Delete this task? This cannot be undone.");
        if (!confirmDelete) return;
        task.deleted = true;
        task.updatedAt = nowIso();
      }
      markDirty();
      await saveLocal();
      renderTasks();
    });
  });

  container.querySelectorAll("[data-task-action='open-meeting']").forEach(btn => {
    btn.addEventListener("click", async () => {
      const meetingId = btn.getAttribute("data-task-meeting");
      const meeting = getMeeting(meetingId);
      if (!meeting) return;
      currentMeetingId = meetingId;
      await saveMeta();
      // Keep notes gated to calendar navigation.
      meetingNotesEnabled = false;
      setMeetingView("setup");
      setActiveModule("meetings");
      setMeetingModuleTab("meeting");
      renderAll();
    });
  });

  container.querySelectorAll("[data-task-action='open-link']").forEach(btn => {
    btn.addEventListener("click", () => {
      const link = btn.getAttribute("data-task-link");
      if (!link) return;
      window.open(link, "_blank", "noopener,noreferrer");
    });
  });
}

function renderAll() {
  renderTemplates();
  renderTopics();
  renderPeopleSelects();
  renderTaskLightboxOptions();
  renderUpdateLightboxOptions();
  renderMeetingCounterpartSelects();
  renderPeopleManager();
  renderGroups();
  renderActionsFiltersOptions();
  renderCurrentMeetingHeader();
  renderSettings();

  // update overview selects might have changed
  renderUpdates();
  renderActionsDashboard();
  renderTopicOverview();
  renderProjectsModule();
  renderSearch();
  renderQuickSearch();
  renderMeetingCalendar();
  renderTasks();
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
  const topicSelect = byId("meeting_topic");
  const addProjectBtn = byId("add_topic_btn");
  if (!templateSelect || !container) return;
  const isOneToOne = templateSelect.value === ONE_TO_ONE_TEMPLATE_ID;
  container.hidden = !isOneToOne;
  // Align ARIA state so non-1:1 meetings do not expose counterpart fields to assistive tech.
  container.setAttribute("aria-hidden", (!isOneToOne).toString());
  const tag = container.querySelector('[data-required-tag="oneToOneTarget"]');
  if (tag) {
    tag.classList.toggle("is-hidden", !isOneToOne);
  }
  const personSelect = byId("meeting_one_to_one_person");
  const groupSelect = byId("meeting_one_to_one_group");
  if (personSelect) personSelect.disabled = !isOneToOne;
  if (groupSelect) groupSelect.disabled = !isOneToOne;
  if (!isOneToOne) {
    if (personSelect) personSelect.value = "";
    if (groupSelect) groupSelect.value = "";
  }
  if (topicSelect) {
    // Enforce the 1:1 rule: project selection is cleared and locked.
    if (isOneToOne) {
      topicSelect.value = "";
    }
    topicSelect.disabled = isOneToOne;
  }
  if (addProjectBtn) {
    addProjectBtn.disabled = isOneToOne;
  }
}

/**
 * Opens the task creation lightbox.
 */
function openTaskLightbox() {
  renderTaskLightboxOptions();
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
  // Default the meeting type to a non-1:1 template so counterpart fields stay hidden initially.
  const templateSelect = byId("meeting_template");
  const titleInput = byId("meeting_title");
  const datetimeInput = byId("meeting_datetime");
  const personSelect = byId("meeting_one_to_one_person");
  const groupSelect = byId("meeting_one_to_one_group");
  if (templateSelect) {
    const hasStandard = Array.from(templateSelect.options).some(opt => opt.value === STANDARD_TEMPLATE_ID);
    if (hasStandard) {
      templateSelect.value = STANDARD_TEMPLATE_ID;
    } else if (templateSelect.value === ONE_TO_ONE_TEMPLATE_ID) {
      const fallback = Array.from(templateSelect.options).find(opt => opt.value && opt.value !== ONE_TO_ONE_TEMPLATE_ID);
      templateSelect.value = fallback?.value || templateSelect.value;
    }
  }
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
  // Only surface counterpart selections when the meeting uses the 1:1 template.
  const isOneToOneTemplate = meeting.templateId === ONE_TO_ONE_TEMPLATE_ID;

  if (templateSelect) templateSelect.value = meeting.templateId || "";
  if (topicSelect) topicSelect.value = isOneToOneTemplate ? "" : (meeting.topicId || "");
  if (titleInput) titleInput.value = meeting.title || "";
  if (datetimeInput) datetimeInput.value = toLocalDateTimeValue(meeting.datetime);
  if (personSelect) personSelect.value = isOneToOneTemplate ? (meeting.oneToOnePersonId || "") : "";
  if (groupSelect) groupSelect.value = isOneToOneTemplate ? (meeting.oneToOneGroupId || "") : "";
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
 * Opens the project creation lightbox and resets validation state.
 */
function openTopicLightbox() {
  const input = byId("topic_name_input");
  const colorInput = byId("topic_color_input");
  if (input) input.value = "";
  if (colorInput) colorInput.value = DEFAULT_PROJECT_COLOR;
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
    ownerId: byId("task_owner")?.value || "",
    dueDate: byId("task_due")?.value || "",
    priority: byId("task_priority")?.value || "medium",
    status: byId("task_status")?.value || "todo",
    link: byId("task_link")?.value?.trim() || "",
    updateTargetsRaw: byId("task_update_targets")?.value || "",
  };
}

/**
 * Resets the task creation form to its defaults.
 */
function clearTaskForm() {
  const defaults = createTaskDraft(null);
  const titleInput = byId("task_title");
  const notesInput = byId("task_notes");
  const ownerSelect = byId("task_owner");
  const dueInput = byId("task_due");
  const prioritySelect = byId("task_priority");
  const statusSelect = byId("task_status");
  const linkInput = byId("task_link");
  const updateTargetsInput = byId("task_update_targets");

  if (titleInput) titleInput.value = defaults.title;
  if (notesInput) notesInput.value = defaults.notes;
  if (ownerSelect) ownerSelect.value = defaults.ownerId || "";
  if (dueInput) dueInput.value = defaults.dueDate;
  if (prioritySelect) prioritySelect.value = defaults.priority;
  if (statusSelect) statusSelect.value = defaults.status;
  if (linkInput) linkInput.value = defaults.link || "";
  if (updateTargetsInput) updateTargetsInput.value = formatTargetNames(defaults.updateTargets);
}

/**
 * Reads the ad-hoc update lightbox inputs into a draft object.
 * @returns {object} Update draft from the form.
 */
function readUpdateFormDraft() {
  return {
    title: byId("update_title")?.value || "",
    notes: byId("update_notes")?.value || "",
    ownerId: byId("update_owner")?.value || "",
    topicId: byId("update_topic")?.value || "",
    status: byId("update_status")?.value || "",
    priority: byId("update_priority")?.value || "medium",
    dueDate: byId("update_due")?.value || "",
    link: byId("update_link")?.value?.trim() || "",
    updateTargetsRaw: byId("update_targets")?.value || "",
  };
}

/**
 * Clears the ad-hoc update lightbox inputs to their defaults.
 */
function clearUpdateForm() {
  const titleInput = byId("update_title");
  const notesInput = byId("update_notes");
  const ownerSelect = byId("update_owner");
  const topicSelect = byId("update_topic");
  const statusSelect = byId("update_status");
  const prioritySelect = byId("update_priority");
  const dueInput = byId("update_due");
  const linkInput = byId("update_link");
  const targetsInput = byId("update_targets");

  if (titleInput) titleInput.value = "";
  if (notesInput) notesInput.value = "";
  if (ownerSelect) ownerSelect.value = getDefaultPersonId() || "";
  if (topicSelect) topicSelect.value = "";
  if (statusSelect) statusSelect.value = "";
  if (prioritySelect) prioritySelect.value = "medium";
  if (dueInput) dueInput.value = "";
  if (linkInput) linkInput.value = "";
  if (targetsInput) targetsInput.value = "";
  setLightboxError("update_lightbox_error", "");
}

/**
 * Opens the ad-hoc update lightbox with fresh options and defaults.
 */
function openUpdateLightbox() {
  renderUpdateLightboxOptions();
  clearUpdateForm();
  openLightbox("update_lightbox", "update_title");
}

/**
 * Reads the project creation input, validates, and persists a new project.
 */
async function addTopicFromLightbox() {
  const name = byId("topic_name_input")?.value.trim() || "";
  const colorInput = byId("topic_color_input")?.value || "";
  if (!name) {
    setLightboxError("topic_lightbox_error", "Project name is required.");
    return;
  }

  const topicId = ensureTopic(name);
  const topic = getTopic(topicId);
  if (topic) {
    topic.color = normalizeHexColor(colorInput);
    topic.updatedAt = nowIso();
  }
  markDirty();
  await saveLocal();
  renderAll();
  const meetingTopic = byId("meeting_topic");
  if (meetingTopic) {
    meetingTopic.value = topicId;
  }
  projectViewId = topicId;
  projectEditorState = { projectId: topicId, draft: createProjectDraft(getTopic(topicId)), error: "" };
  closeLightbox("topic_lightbox");
  // Auto-sync after a user creates a project.
  requestAutoSync("create-project");
}

/**
 * Saves edits from the Projects module into the selected project record.
 */
async function saveProjectDetails() {
  if (!projectViewId) return;
  const project = getTopic(projectViewId);
  if (!project) return;

  const draft = readProjectDetailsForm();
  const errs = validateProjectDraft(draft);
  if (errs.length) {
    projectEditorState = { ...projectEditorState, draft, error: errs.join(" ") };
    setProjectDetailsError(projectEditorState.error);
    return;
  }

  project.name = draft.name.trim();
  project.ownerId = draft.ownerId || "";
  project.startDate = draft.startDate;
  project.endDate = draft.endDate;
  project.color = normalizeHexColor(draft.color);
  project.updatedAt = nowIso();

  projectEditorState = { projectId: project.id, draft: createProjectDraft(project), error: "" };
  markDirty();
  await saveLocal();
  renderAll();
  // Auto-sync after a user saves project details.
  requestAutoSync("save-project");
}

/**
 * Resets the project detail form to the last saved state.
 */
function resetProjectDetails() {
  if (!projectViewId) return;
  const project = getTopic(projectViewId);
  if (!project) return;
  projectEditorState = { projectId: project.id, draft: createProjectDraft(project), error: "" };
  setProjectDetailsError("");
  renderProjectsModule();
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
  // Auto-sync after a user creates a group.
  requestAutoSync("create-group");
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
  // Auto-sync after a user creates a person.
  requestAutoSync("create-person");
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

  const { ids: updateTargets, missing } = parseTargetNames(draft.updateTargetsRaw, alive(db.people));
  if (missing.length) {
    alert(`Unknown people: ${missing.join(", ")}.`);
    return;
  }

  const now = nowIso();
  const task = {
    id: uid("task"),
    kind: "task",
    title: draft.title.trim(),
    text: draft.title.trim(),
    notes: draft.notes.trim(),
    ownerId: draft.ownerId || getDefaultPersonId() || null,
    dueDate: draft.dueDate,
    priority: draft.priority,
    status: draft.status,
    link: draft.link.trim(),
    updateTargets,
    updateStatus: buildUpdateStatusForTargets(updateTargets),
    meetingId: null,
    topicId: null,
    section: "task",
    createdAt: now,
    updatedAt: now
  };

  db.tasks.push(task);
  clearTaskForm();
  closeTaskLightbox();
  markDirty();
  await saveLocal();
  renderTasks();
  // Auto-sync after a user creates a task.
  requestAutoSync("create-task");
}

/**
 * Adds an ad-hoc update (not tied to a meeting) into the items collection.
 */
async function addAdHocUpdate() {
  const draft = readUpdateFormDraft();
  const errs = [];

  if (!draft.title.trim()) errs.push("Update summary is required.");

  const { ids: updateTargets, missing } = parseTargetNames(draft.updateTargetsRaw, alive(db.people));
  if (missing.length) {
    errs.push(`Unknown people: ${missing.join(", ")}.`);
  }
  if (!updateTargets.length) {
    errs.push("At least one update target is required.");
  }

  if (errs.length) {
    setLightboxError("update_lightbox_error", errs.join(" "));
    return;
  }

  const now = nowIso();
  const item = {
    id: uid("item"),
    kind: "item",
    meetingId: null,
    topicId: draft.topicId || null,
    section: "update",
    text: draft.title.trim(),
    title: draft.title.trim(),
    notes: draft.notes.trim(),
    ownerId: draft.ownerId || null,
    status: draft.status || null,
    dueDate: draft.dueDate || null,
    priority: draft.priority || "medium",
    link: draft.link.trim(),
    updateTargets,
    updateStatus: buildUpdateStatusForTargets(updateTargets),
    createdAt: now,
    updatedAt: now,
  };

  db.items.push(item);
  clearUpdateForm();
  closeLightbox("update_lightbox");
  markDirty();
  await saveLocal();
  renderAll();
  // Auto-sync after a user captures an ad-hoc update.
  requestAutoSync("create-ad-hoc-update");
}

async function saveMeetingFromLightbox() {
  const templateId = byId("meeting_template").value;
  const topicId = byId("meeting_topic").value || null;
  const oneToOnePersonId = byId("meeting_one_to_one_person")?.value || "";
  const oneToOneGroupId = byId("meeting_one_to_one_group")?.value || "";
  const isOneToOneTemplate = templateId === ONE_TO_ONE_TEMPLATE_ID;

  if (!templateId) { alert("Choose a template."); return; }
  if (!isOneToOneTemplate && !topicId) { alert("Choose or add a project."); return; }
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
    meeting.topicId = isOneToOneTemplate ? null : topicId;
    meeting.title = title;
    meeting.datetime = datetime;
    meeting.oneToOnePersonId = isOneToOneTemplate ? (oneToOnePersonId || null) : null;
    meeting.oneToOneGroupId = isOneToOneTemplate ? (oneToOneGroupId || null) : null;
    meeting.updatedAt = nowIso();
  } else {
    meeting = {
      id: uid("meeting"),
      templateId,
      topicId: isOneToOneTemplate ? null : topicId,
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
  // Keep the schedule view active after creating or editing meetings.
  setMeetingView("setup");
  meetingEditId = null;
  clearMeetingForm();
  closeLightbox("meeting_lightbox");

  markDirty();
  await saveLocal();
  await saveMeta();
  renderAll();
  // Auto-sync after a user saves a meeting.
  requestAutoSync("save-meeting");
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
  lines.push(`Project: ${topic?.name || ""}`);
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
  const projectId = byId("updates_project")?.value || "";
  if (!personId) return;
  const person = getPerson(personId);

  const pending = alive(db.items).filter(it => {
    if (!it.updateTargets || !it.updateTargets.includes(personId)) return false;
    const st = it.updateStatus?.[personId];
    if (projectId && it.topicId !== projectId) return false;
    return !st?.updated;
  }).sort((a,b)=>Date.parse(b.updatedAt)-Date.parse(a.updatedAt));

  const lines = [];
  lines.push(`Updates for: ${person?.name || personId}`);
  lines.push(`Generated: ${fmtDateTime(nowIso())}`);
  lines.push("");

  for (const it of pending) {
    const topic = getTopic(it.topicId)?.name || "No project";
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
  document.querySelectorAll(".module-tab[data-module]").forEach(btn => {
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
 * Wires the ad-hoc update lightbox controls.
 */
function wireUpdateControls() {
  byId("ad_hoc_update_btn")?.addEventListener("click", openUpdateLightbox);
  byId("update_add_btn")?.addEventListener("click", addAdHocUpdate);
  byId("update_clear_btn")?.addEventListener("click", clearUpdateForm);
}

/**
 * Wires shared lightbox dismissal controls for all modal flows.
 */
function wireLightboxControls() {
  const lightboxIds = [
    "task_lightbox",
    "meeting_lightbox",
    "update_lightbox",
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
  byId("sync_btn").addEventListener("click", () => {
    syncNow({ silent: false, trigger: "manual" }).catch(console.error);
  });
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
  byId("updates_project").addEventListener("change", renderUpdates);
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
  // Project management lives in the former "Topics" module; reuse the same lightbox workflow.
  byId("project_open_lightbox_btn")?.addEventListener("click", openTopicLightbox);
  byId("topics_topic").addEventListener("change", renderTopicOverview);
  byId("topics_focus").addEventListener("change", renderTopicOverview);
}

/**
 * Wires the Projects module controls for selection and editing.
 */
function wireProjectControls() {
  const projectSelect = byId("project_select");
  projectSelect?.addEventListener("change", () => {
    projectViewId = projectSelect.value || "";
    projectEditorState = projectViewId
      ? { projectId: projectViewId, draft: createProjectDraft(getTopic(projectViewId)), error: "" }
      : { projectId: null, draft: null, error: "" };
    renderProjectsModule();
  });

  byId("project_module_new_btn")?.addEventListener("click", openTopicLightbox);
  byId("project_save_btn")?.addEventListener("click", saveProjectDetails);
  byId("project_reset_btn")?.addEventListener("click", resetProjectDetails);
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

  // Generic wiring for settings controls to keep expansion easy as new fields are added.
  document.querySelectorAll("[data-setting-path]").forEach(control => {
    const path = control.dataset.settingPath;
    const eventName = control.dataset.settingType === "checkbox" || control.type === "checkbox" ? "change" : "input";
    const handler = async () => {
      db.settings = db.settings || buildDefaultSettings();
      const rawValue = readSettingControlValue(control);
      const value = coerceSettingValue(path, rawValue);
      if (setSettingValue(path, value)) {
        db.settings.updatedAt = nowIso();
        markDirty();
        await saveLocal();
      }
    };
    const wrappedHandler = eventName === "input" ? debounce(handler, 200) : handler;
    control.addEventListener(eventName, wrappedHandler);
  });
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
    const activePerson = personViewId ? getPerson(personViewId) : null;
    if (activePerson && isDefaultPerson(activePerson)) {
      alert("The default person details are locked and cannot be edited.");
      return;
    }
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
    // Auto-sync after a user saves person details.
    requestAutoSync("save-person");
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
  wireProjectControls();
  wireSearchControls();
  wireGroupControls();
  wireSettingsControls();
  wirePeopleControls();
  wireTasksControls();
  wireUpdateControls();
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
  // Begin periodic auto-sync checks once the app is ready.
  startAutoSyncTimer();
}

document.addEventListener("DOMContentLoaded", init);
