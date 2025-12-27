/* Work Order Viewer (CSV/XLSX) - GitHub Pages friendly
   - Tabulator for table, column move, persistence
   - PapaParse for CSV
   - SheetJS for XLSX
   - Luxon for date parsing/formatting (MM/DD/YYYY display)
*/

const { DateTime } = luxon;

// ====== Storage Keys ======
const STORAGE = {
  tabulatorPersistenceID: "wo_viewer_tabulator_v1",
  sidePanelHidden: "wo_viewer_side_hidden_v1",
  sidePanelWidth: "wo_viewer_side_width_v1",
  lastDateField: "wo_viewer_date_field_v1",
};

// ====== Column mapping ======
// We normalize incoming headers and map them to our internal field names.
const WANTED_COLUMNS = [
  { field: "work_order", title: "Work Order" },
  { field: "description", title: "Description" },
  { field: "status", title: "Status" },
  { field: "type", title: "Type" },
  { field: "department", title: "Department" },
  { field: "equipment", title: "Equipment" },
  { field: "equipment_description", title: "Equipment Description" },
  { field: "sched_start", title: "Sched. Start Date" },
  { field: "orig_due", title: "Original PM Due Date" },
  { field: "sched_end", title: "Sched. End Date" },
  { field: "assigned_to", title: "Assigned To" },
];

// Synonyms you might see in exports (case/spacing/punctuation varies)
const HEADER_SYNONYMS = new Map([
  // Work Order
  ["workorder", "work_order"],
  ["work_order", "work_order"],
  ["wo", "work_order"],
  ["work order", "work_order"],

  // Description
  ["description", "description"],
  ["wo description", "description"],

  // Status / Type / Dept
  ["status", "status"],
  ["type", "type"],
  ["department", "department"],

  // Equipment
  ["equipment", "equipment"],
  ["equipmentid", "equipment"],
  ["equipment id", "equipment"],

  ["equipmentdescription", "equipment_description"],
  ["equipment description", "equipment_description"],

  // Dates
  ["sched. start date", "sched_start"],
  ["sched start date", "sched_start"],
  ["scheduled start date", "sched_start"],
  ["scheduled start", "sched_start"],
  ["sched_start", "sched_start"],

  ["original pm due date", "orig_due"],
  ["original pm due", "orig_due"],
  ["pm due date", "orig_due"],
  ["orig_due", "orig_due"],

  ["sched. end date", "sched_end"],
  ["sched end date", "sched_end"],
  ["scheduled end date", "sched_end"],
  ["scheduled end", "sched_end"],
  ["sched_end", "sched_end"],

  // Assigned To
  ["assigned to", "assigned_to"],
  ["assignedto", "assigned_to"],
  ["assigned_to", "assigned_to"],
  ["assignee", "assigned_to"],
]);

// ====== DOM ======
const el = {
  fileInput: document.getElementById("fileInput"),
  clearBtn: document.getElementById("clearBtn"),
  resetLayoutBtn: document.getElementById("resetLayoutBtn"),
  statusMsg: document.getElementById("statusMsg"),

  searchInput: document.getElementById("searchInput"),
  dateField: document.getElementById("dateField"),
  dateFrom: document.getElementById("dateFrom"),
  dateTo: document.getElementById("dateTo"),
  assigneeFilter: document.getElementById("assigneeFilter"),

  table: document.getElementById("table"),
  sidePanel: document.getElementById("sidePanel"),
  toggleSideBtn: document.getElementById("toggleSideBtn"),
  divider: document.getElementById("divider"),

  countsSub: document.getElementById("countsSub"),
  countsMeta: document.getElementById("countsMeta"),
  countsList: document.getElementById("countsList"),
};

// ====== State ======
let table = null;
let rawRows = [];   // normalized rows (only wanted fields)
let filteredRows = []; // based on Tabulator current filters

// ====== Helpers ======
function setStatus(msg, isError = false) {
  el.statusMsg.textContent = msg || "";
  el.statusMsg.style.color = isError ? "#ff9aa0" : "";
}

function normalizeHeader(h) {
  return String(h ?? "")
    .trim()
    .toLowerCase()
    .replace(/\uFEFF/g, "")           // strip BOM
    .replace(/[_\-]+/g, " ")
    .replace(/\s+/g, " ")
    .replace(/[.]/g, ".")            // keep dots (we also match without)
    .trim();
}

function headerToField(header) {
  const h = normalizeHeader(header);
  const hNoDots = h.replace(/[.]/g, "");
  const direct = HEADER_SYNONYMS.get(h);
  if (direct) return direct;
  const noDots = HEADER_SYNONYMS.get(hNoDots);
  if (noDots) return noDots;
  return null;
}

function toISODateMaybe(value) {
  // Returns ISO yyyy-MM-dd or "" if not parseable.
  if (value === null || value === undefined || value === "") return "";

  // If SheetJS gives a Date object:
  if (value instanceof Date && !isNaN(value.valueOf())) {
    return DateTime.fromJSDate(value).toISODate();
  }

  // If SheetJS gives number (Excel serial date):
  if (typeof value === "number" && isFinite(value)) {
    // Excel date serial: use SheetJS to parse by converting to JS date
    // But safer: DateTime from Excel epoch (1899-12-30)
    const dt = DateTime.fromISO("1899-12-30").plus({ days: Math.floor(value) });
    return dt.isValid ? dt.toISODate() : "";
  }

  const s = String(value).trim();
  if (!s) return "";

  // Common formats
  const candidates = [
    DateTime.fromFormat(s, "M/d/yyyy"),
    DateTime.fromFormat(s, "MM/dd/yyyy"),
    DateTime.fromFormat(s, "M/d/yy"),
    DateTime.fromFormat(s, "MM/dd/yy"),
    DateTime.fromISO(s),
    DateTime.fromRFC2822(s),
  ];

  for (const c of candidates) {
    if (c.isValid) return c.toISODate();
  }

  // Try Luxon auto parse (last resort)
  const auto = DateTime.fromJSDate(new Date(s));
  if (auto.isValid) return auto.toISODate();

  return "";
}

function formatMMDDYYYYFromISO(iso) {
  if (!iso) return "";
  const dt = DateTime.fromISO(iso);
  return dt.isValid ? dt.toFormat("MM/dd/yyyy") : "";
}

function uniqSorted(arr) {
  return Array.from(new Set(arr)).filter(Boolean).sort((a, b) => a.localeCompare(b));
}

// ====== Build Tabulator ======
function buildTable() {
  const columns = WANTED_COLUMNS.map(col => {
    const isDate = ["sched_start", "orig_due", "sched_end"].includes(col.field);

    return {
      title: col.title,
      field: col.field,
      headerFilter: true,
      headerFilterPlaceholder: "filter…",
      widthGrow: col.field === "description" ? 3 : 1,
      formatter: isDate ? (cell) => {
        const iso = cell.getValue();
        return formatMMDDYYYYFromISO(iso);
      } : undefined,
      sorter: isDate ? (a, b) => (a || "").localeCompare(b || "") : "string",
    };
  });

  table = new Tabulator("#table", {
    data: [],
    columns,
    layout: "fitColumns",
    height: "100%",

    movableColumns: true,

    persistence: {
      sort: true,
      filter: true,
      columns: true,
    },
    persistenceID: STORAGE.tabulatorPersistenceID,
    persistenceMode: "local",

    placeholder: "Upload a CSV or XLSX to begin.",
    reactiveData: false,

    // Keep counts synced when the table changes
    dataFiltered: function(filters, rows) {
      filteredRows = rows.map(r => r.getData());
      refreshCounts(filteredRows);
      refreshAssigneeDropdown(filteredRows);
    },
  });
}

function resetColumnLayout() {
  // Clear Tabulator persistence only (leaves other app storage alone)
  const prefix = "tabulator-" + STORAGE.tabulatorPersistenceID;
  for (let i = localStorage.length - 1; i >= 0; i--) {
    const k = localStorage.key(i);
    if (k && k.startsWith(prefix)) localStorage.removeItem(k);
  }
  setStatus("Saved column layout cleared. Refreshing…");
  // Rebuild table to apply reset
  table?.destroy();
  buildTable();
  if (rawRows.length) table.setData(rawRows);
}

// ====== File parsing ======
async function handleFile(file) {
  setStatus("");
  if (!file) return;

  const name = file.name.toLowerCase();

  try {
    if (name.endsWith(".csv")) {
      const text = await file.text();
      const parsed = Papa.parse(text, {
        header: true,
        skipEmptyLines: true,
        dynamicTyping: false,
        // auto delimiter detect usually works; still accepts comma/semicolon
      });

      if (parsed.errors?.length) {
        console.warn(parsed.errors);
        setStatus("CSV parsed with warnings. If data is missing, check delimiter/headers.", true);
      }

      const normalized = normalizeRowsFromObjects(parsed.data);
      loadData(normalized, file.name);
      return;
    }

    if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const first = wb.SheetNames[0];
      const ws = wb.Sheets[first];

      // sheet_to_json with raw false helps get readable values
      const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
      const normalized = normalizeRowsFromObjects(json);
      loadData(normalized, file.name);
      return;
    }

    setStatus("Unsupported file type. Please upload a .csv, .xlsx, or .xls", true);
  } catch (err) {
    console.error(err);
    setStatus("Failed to read file: " + (err?.message || String(err)), true);
  }
}

function normalizeRowsFromObjects(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];

  // Build a mapping from incoming headers -> our internal fields
  const sample = rows[0];
  const incomingHeaders = Object.keys(sample);

  const headerMap = new Map(); // incoming header -> internal field
  for (const h of incomingHeaders) {
    const field = headerToField(h);
    if (field) headerMap.set(h, field);
  }

  // If we didn't match anything, show a better error
  const matchedFields = new Set(Array.from(headerMap.values()));
  if (matchedFields.size === 0) {
    setStatus(
      "No recognizable headers found. Make sure your export includes columns like 'Work Order', 'Sched. Start Date', 'Assigned To', etc.",
      true
    );
  }

  // Normalize each row into only wanted columns
  const normalized = rows.map((r) => {
    const out = {};
    for (const col of WANTED_COLUMNS) out[col.field] = "";

    for (const [incomingHeader, internalField] of headerMap.entries()) {
      let v = r[incomingHeader];

      // Date fields -> ISO for filtering/sorting, display as MM/DD/YYYY
      if (["sched_start", "orig_due", "sched_end"].includes(internalField)) {
        out[internalField] = toISODateMaybe(v);
      } else {
        out[internalField] = String(v ?? "").trim();
      }
    }
    return out;
  });

  // Remove completely empty rows (common in exports)
  return normalized.filter(r => {
    return (r.work_order || r.description || r.equipment || r.assigned_to ||
      r.sched_start || r.orig_due || r.sched_end);
  });
}

function loadData(rows, filename) {
  rawRows = rows || [];
  if (!rawRows.length) {
    table.setData([]);
    refreshCounts([]);
    refreshAssigneeDropdown([]);
    setStatus("Loaded file, but no rows matched your selected columns. Check the headers/export.", true);
    return;
  }

  table.setData(rawRows);
  setStatus(`Loaded ${rawRows.length} rows from ${filename}`);

  // After data loads, apply current UI filters
  applyFiltersFromUI();
}

// ====== Filtering ======
function applyFiltersFromUI() {
  if (!table) return;

  const q = el.searchInput.value.trim().toLowerCase();
  const assignee = el.assigneeFilter.value;
  const dateField = el.dateField.value; // sched_start | orig_due | sched_end
  const from = el.dateFrom.value ? DateTime.fromISO(el.dateFrom.value).toISODate() : "";
  const to = el.dateTo.value ? DateTime.fromISO(el.dateTo.value).toISODate() : "";

  // Clear all filters first
  table.clearFilter(true);

  // Global search (custom filter across fields)
  if (q) {
    table.addFilter((data) => {
      const hay = [
        data.work_order,
        data.description,
        data.status,
        data.type,
        data.department,
        data.equipment,
        data.equipment_description,
        data.assigned_to,
        formatMMDDYYYYFromISO(data.sched_start),
        formatMMDDYYYYFromISO(data.orig_due),
        formatMMDDYYYYFromISO(data.sched_end),
      ].join(" ").toLowerCase();

      return hay.includes(q);
    });
  }

  if (assignee) {
    table.addFilter("assigned_to", "=", assignee);
  }

  if (from || to) {
    table.addFilter((data) => {
      const d = data[dateField] || "";
      if (!d) return false;
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    });
  }

  // save preferred date field
  localStorage.setItem(STORAGE.lastDateField, dateField);

  // Trigger Tabulator to recompute filteredRows (counts update in dataFiltered callback)
  table.redraw(true);
}

// ====== Counts (employee x day) ======
function refreshCounts(rows) {
  const dateField = el.dateField.value;
  const total = rows.length;

  // Build counts[assignee][date] = n
  const counts = new Map();

  for (const r of rows) {
    const who = (r.assigned_to || "").trim() || "(Unassigned)";
    const iso = (r[dateField] || "").trim();
    if (!iso) continue;

    if (!counts.has(who)) counts.set(who, new Map());
    const map = counts.get(who);
    map.set(iso, (map.get(iso) || 0) + 1);
  }

  // Render
  const people = Array.from(counts.keys()).sort((a, b) => a.localeCompare(b));
  const dateFrom = el.dateFrom.value ? DateTime.fromISO(el.dateFrom.value).toFormat("MM/dd/yyyy") : "—";
  const dateTo = el.dateTo.value ? DateTime.fromISO(el.dateTo.value).toFormat("MM/dd/yyyy") : "—";

  el.countsMeta.textContent =
    `Rows shown: ${total}\n` +
    `Counting by: ${labelForDateField(dateField)}\n` +
    `Range: ${dateFrom} → ${dateTo}`;

  el.countsList.innerHTML = "";

  if (people.length === 0) {
    el.countsList.innerHTML = `<div class="countRow"><div class="countName">No dated rows to count</div></div>`;
    return;
  }

  for (const person of people) {
    const dateMap = counts.get(person);
    const dates = Array.from(dateMap.keys()).sort((a, b) => a.localeCompare(b));
    const totalForPerson = dates.reduce((sum, d) => sum + (dateMap.get(d) || 0), 0);

    const row = document.createElement("div");
    row.className = "countRow";

    const top = document.createElement("div");
    top.className = "countRowTop";

    const name = document.createElement("div");
    name.className = "countName";
    name.textContent = person;

    const totalEl = document.createElement("div");
    totalEl.className = "countTotal";
    totalEl.textContent = totalForPerson;

    top.appendChild(name);
    top.appendChild(totalEl);

    const datesWrap = document.createElement("div");
    datesWrap.className = "countDates";

    for (const iso of dates) {
      const pill = document.createElement("div");
      pill.className = "datePill";

      const d = document.createElement("div");
      d.className = "d";
      d.textContent = formatMMDDYYYYFromISO(iso);

      const n = document.createElement("div");
      n.className = "n";
      n.textContent = dateMap.get(iso);

      pill.appendChild(d);
      pill.appendChild(n);
      datesWrap.appendChild(pill);
    }

    row.appendChild(top);
    row.appendChild(datesWrap);
    el.countsList.appendChild(row);
  }
}

function labelForDateField(v) {
  if (v === "sched_start") return "Sched. Start Date";
  if (v === "orig_due") return "Original PM Due Date";
  if (v === "sched_end") return "Sched. End Date";
  return v;
}

// ====== Assignee dropdown ======
function refreshAssigneeDropdown(rows) {
  const names = uniqSorted(rows.map(r => (r.assigned_to || "").trim()).filter(Boolean));

  // keep selection if possible
  const current = el.assigneeFilter.value;

  el.assigneeFilter.innerHTML = `<option value="">All</option>` + names.map(n => {
    const safe = n.replace(/"/g, "&quot;");
    return `<option value="${safe}">${n}</option>`;
  }).join("");

  if (current && names.includes(current)) {
    el.assigneeFilter.value = current;
  } else {
    el.assigneeFilter.value = "";
  }
}

// ====== Side panel toggle + resizing ======
function setSideHidden(hidden) {
  el.sidePanel.classList.toggle("hidden", hidden);
  localStorage.setItem(STORAGE.sidePanelHidden, hidden ? "1" : "0");
}

function loadSidePrefs() {
  const hidden = localStorage.getItem(STORAGE.sidePanelHidden) === "1";
  setSideHidden(hidden);

  const width = parseInt(localStorage.getItem(STORAGE.sidePanelWidth) || "", 10);
  if (Number.isFinite(width) && width >= 260 && width <= window.innerWidth * 0.55) {
    el.sidePanel.style.width = width + "px";
  }
}

function initDividerDrag() {
  let dragging = false;

  el.divider.addEventListener("mousedown", () => {
    dragging = true;
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
  });

  window.addEventListener("mouseup", () => {
    if (!dragging) return;
    dragging = false;
    document.body.style.cursor = "";
    document.body.style.userSelect = "";
    const w = parseInt(getComputedStyle(el.sidePanel).width, 10);
    if (Number.isFinite(w)) localStorage.setItem(STORAGE.sidePanelWidth, String(w));
  });

  window.addEventListener("mousemove", (e) => {
    if (!dragging) return;
    if (el.sidePanel.classList.contains("hidden")) return;

    // Set side panel width based on mouse X from right edge
    const viewportW = window.innerWidth;
    const newW = Math.min(Math.max(viewportW - e.clientX, 260), viewportW * 0.55);
    el.sidePanel.style.width = newW + "px";
    table?.redraw(true);
  });
}

// ====== Events ======
function wireEvents() {
  el.fileInput.addEventListener("change", (e) => {
    const file = e.target.files?.[0];
    handleFile(file);
    // allow re-upload same file by clearing value
    e.target.value = "";
  });

  el.clearBtn.addEventListener("click", () => {
    rawRows = [];
    filteredRows = [];
    table.setData([]);
    refreshCounts([]);
    refreshAssigneeDropdown([]);
    setStatus("Cleared.");
  });

  el.resetLayoutBtn.addEventListener("click", resetColumnLayout);

  el.searchInput.addEventListener("input", debounce(applyFiltersFromUI, 150));
  el.assigneeFilter.addEventListener("change", applyFiltersFromUI);
  el.dateFrom.addEventListener("change", applyFiltersFromUI);
  el.dateTo.addEventListener("change", applyFiltersFromUI);
  el.dateField.addEventListener("change", () => {
    applyFiltersFromUI();
    // counts refresh happens via dataFiltered, but if same filtered set, force refresh:
    refreshCounts(filteredRows);
  });

  el.toggleSideBtn.addEventListener("click", () => {
    const hidden = el.sidePanel.classList.contains("hidden");
    setSideHidden(!hidden);
    table?.redraw(true);
  });
}

function debounce(fn, wait) {
  let t = null;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), wait);
  };
}

// ====== Init ======
function initDefaults() {
  // Default date field preference
  const savedDateField = localStorage.getItem(STORAGE.lastDateField);
  if (savedDateField && ["sched_start","orig_due","sched_end"].includes(savedDateField)) {
    el.dateField.value = savedDateField;
  } else {
    el.dateField.value = "sched_start";
  }

  // Optional: default date range to today..today+7? (leave empty unless you want it)
  // el.dateFrom.value = DateTime.now().toISODate();
  // el.dateTo.value = DateTime.now().plus({days:7}).toISODate();

  loadSidePrefs();
}

window.addEventListener("DOMContentLoaded", () => {
  buildTable();
  wireEvents();
  initDefaults();
  initDividerDrag();
  setStatus("Upload a CSV or XLSX to load work orders.");
});
