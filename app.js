/* Work Order Viewer
   - Upload CSV
   - Filter by keyword, date range, assigned/status/type/department
   - Sort by clicking headers
   - Drag headers to reorder; persists in localStorage
   - Daily counts per employee; date column selectable
   - Dates displayed as MM/DD/YYYY
*/

const STORAGE_KEY_COL_ORDER = "wo_col_order_v1";

const CANON_COLUMNS = [
  "Work Order",
  "Description",
  "Status",
  "Type",
  "Department",
  "Equipment",
  "Equipment Description",
  "Scheduled start date",
  "Original pm due date",
  "Scheduled end date",
  "Assigned to"
];

// Some CSVs have slightly different header text; we’ll match loosely.
const COLUMN_ALIASES = {
  "Work Order": ["work order", "wo", "workorder", "work_order", "work order #", "work order number"],
  "Description": ["description", "wo description", "work order description", "details"],
  "Status": ["status", "wo status"],
  "Type": ["type", "wo type", "work type"],
  "Department": ["department", "dept"],
  "Equipment": ["equipment", "asset", "asset number", "equipment id"],
  "Equipment Description": ["equipment description", "asset description", "equip description", "equipment desc"],
  "Scheduled start date": ["scheduled start date", "scheduled start", "sched start", "start date", "scheduled start"],
  "Original pm due date": ["original pm due date", "pm due date", "original due date", "due date"],
  "Scheduled end date": ["scheduled end date", "scheduled end", "sched end", "end date"],
  "Assigned to": ["assigned to", "assigned", "assignee", "owner", "assigned_to"]
};

const els = {
  fileInput: document.getElementById("fileInput"),
  btnResetLayout: document.getElementById("btnResetLayout"),
  btnClearData: document.getElementById("btnClearData"),

  qSearch: document.getElementById("qSearch"),
  assignedFilter: document.getElementById("assignedFilter"),
  statusFilter: document.getElementById("statusFilter"),
  typeFilter: document.getElementById("typeFilter"),
  deptFilter: document.getElementById("deptFilter"),

  dateFilterColumn: document.getElementById("dateFilterColumn"),
  dateFrom: document.getElementById("dateFrom"),
  dateTo: document.getElementById("dateTo"),

  btnClearFilters: document.getElementById("btnClearFilters"),
  rowCount: document.getElementById("rowCount"),

  countDateColumn: document.getElementById("countDateColumn"),
  countsWrap: document.getElementById("countsWrap"),

  table: document.getElementById("dataTable"),
  thead: document.querySelector("#dataTable thead"),
  tbody: document.querySelector("#dataTable tbody"),
  colDebug: document.getElementById("colDebug"),
};

let rawRows = [];        // parsed rows with original headers
let rows = [];           // normalized rows with canonical columns only
let visibleColumns = []; // current ordered canonical columns
let sortState = { col: null, dir: "asc" }; // dir: asc|desc

// ---------- Utilities ----------
function normalizeHeader(s) {
  if (s == null) return "";
  return String(s)
    .replace(/^\uFEFF/, "")  // remove BOM
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function pickHeaderMap(headers) {
  // Map canonical -> actual header in CSV, using aliases and fuzzy matching
  const normToActual = new Map();
  headers.forEach(h => normToActual.set(normalizeHeader(h), h));

  const headerMap = {};

  for (const canon of CANON_COLUMNS) {
    const targets = [canon, ...(COLUMN_ALIASES[canon] || [])];
    let found = null;

    // 1) exact normalized match
    for (const t of targets) {
      const key = normalizeHeader(t);
      if (normToActual.has(key)) {
        found = normToActual.get(key);
        break;
      }
    }

    // 2) contains match (helps with things like "Scheduled start date (local)" etc.)
    if (!found) {
      const targetNorms = targets.map(normalizeHeader);
      for (const [norm, actual] of normToActual.entries()) {
        if (targetNorms.some(tn => norm.includes(tn))) {
          found = actual;
          break;
        }
      }
    }

    headerMap[canon] = found; // may be null if missing
  }

  return headerMap;
}

function parseDateFlexible(v) {
  if (v == null) return null;
  const s = String(v).trim();
  if (!s) return null;

  // Try ISO (YYYY-MM-DD or YYYY-MM-DDTHH:mm...)
  const iso = new Date(s);
  if (!Number.isNaN(iso.getTime())) return iso;

  // Try MM/DD/YYYY or M/D/YYYY (optionally with time)
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+.*)?$/);
  if (m) {
    let mm = parseInt(m[1], 10);
    let dd = parseInt(m[2], 10);
    let yy = parseInt(m[3], 10);
    if (yy < 100) yy += 2000;
    const d = new Date(yy, mm - 1, dd);
    if (!Number.isNaN(d.getTime())) return d;
  }

  return null;
}

function formatMMDDYYYY(dateObj) {
  if (!dateObj) return "";
  const mm = String(dateObj.getMonth() + 1).padStart(2, "0");
  const dd = String(dateObj.getDate()).padStart(2, "0");
  const yy = dateObj.getFullYear();
  return `${mm}/${dd}/${yy}`;
}

function toDateInputValue(dateObj) {
  // YYYY-MM-DD for <input type=date>
  if (!dateObj) return "";
  const mm = String(dateObj.getMonth() + 1).padStart(2, "0");
  const dd = String(dateObj.getDate()).padStart(2, "0");
  const yy = dateObj.getFullYear();
  return `${yy}-${mm}-${dd}`;
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function uniqueSorted(values) {
  const set = new Set(values.filter(v => String(v ?? "").trim() !== ""));
  return Array.from(set).sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
}

function loadColumnOrder() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_COL_ORDER);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return null;

    // Must contain only known columns; ignore unknowns
    const filtered = parsed.filter(c => CANON_COLUMNS.includes(c));
    // Add any missing columns at the end (in default order)
    for (const c of CANON_COLUMNS) if (!filtered.includes(c)) filtered.push(c);
    return filtered;
  } catch {
    return null;
  }
}

function saveColumnOrder(order) {
  localStorage.setItem(STORAGE_KEY_COL_ORDER, JSON.stringify(order));
}

// ---------- Parsing ----------
function parseCsvFile(file) {
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    dynamicTyping: false,
    transformHeader: h => h, // keep original header text; we normalize later
    complete: (results) => {
      if (results.errors && results.errors.length) {
        console.warn("CSV parse errors:", results.errors);
      }
      rawRows = results.data || [];
      onDataLoaded(results.meta?.fields || []);
    }
  });
}

function onDataLoaded(headers) {
  if (!headers || headers.length === 0) {
    // Sometimes PapaParse doesn’t give meta fields when headers are weird; derive from first row
    headers = rawRows.length ? Object.keys(rawRows[0]) : [];
  }

  // Debug: show loaded headers
  els.colDebug.textContent = headers.length
    ? headers.join(" | ")
    : "No headers detected. Make sure your CSV has a header row.";

  const headerMap = pickHeaderMap(headers);

  // Normalize into canonical rows only
  rows = rawRows.map(r => {
    const obj = {};
    for (const canon of CANON_COLUMNS) {
      const actual = headerMap[canon];
      const value = actual ? r[actual] : "";
      obj[canon] = (value ?? "").toString().trim();
    }

    // Parse dates into hidden fields for filtering/sorting
    obj.__date = {
      "Scheduled start date": parseDateFlexible(obj["Scheduled start date"]),
      "Original pm due date": parseDateFlexible(obj["Original pm due date"]),
      "Scheduled end date": parseDateFlexible(obj["Scheduled end date"]),
    };

    return obj;
  });

  // Set visible columns from saved order (or default)
  visibleColumns = loadColumnOrder() || [...CANON_COLUMNS];

  // Build filter dropdown values
  rebuildFilterOptions();

  // Render everything
  renderTable();
  renderCounts();
}

// ---------- Filters ----------
function rebuildFilterOptions() {
  const assigned = uniqueSorted(rows.map(r => r["Assigned to"]));
  const status = uniqueSorted(rows.map(r => r["Status"]));
  const type = uniqueSorted(rows.map(r => r["Type"]));
  const dept = uniqueSorted(rows.map(r => r["Department"]));

  fillSelect(els.assignedFilter, assigned, "All");
  fillSelect(els.statusFilter, status, "All");
  fillSelect(els.typeFilter, type, "All");
  fillSelect(els.deptFilter, dept, "All");
}

function fillSelect(sel, items, allLabel) {
  const current = sel.value;
  sel.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "";
  optAll.textContent = allLabel;
  sel.appendChild(optAll);

  for (const it of items) {
    const opt = document.createElement("option");
    opt.value = it;
    opt.textContent = it;
    sel.appendChild(opt);
  }

  // keep selection if still present
  if ([...sel.options].some(o => o.value === current)) {
    sel.value = current;
  }
}

function getFilteredRows() {
  const q = els.qSearch.value.trim().toLowerCase();
  const assigned = els.assignedFilter.value;
  const status = els.statusFilter.value;
  const type = els.typeFilter.value;
  const dept = els.deptFilter.value;

  const dateCol = els.dateFilterColumn.value;
  const from = els.dateFrom.value ? new Date(els.dateFrom.value + "T00:00:00") : null;
  const to = els.dateTo.value ? new Date(els.dateTo.value + "T23:59:59") : null;

  return rows.filter(r => {
    if (assigned && r["Assigned to"] !== assigned) return false;
    if (status && r["Status"] !== status) return false;
    if (type && r["Type"] !== type) return false;
    if (dept && r["Department"] !== dept) return false;

    if (q) {
      // Search across visible canonical columns (not the hidden date fields)
      const hay = CANON_COLUMNS.map(c => (r[c] ?? "").toString().toLowerCase()).join(" | ");
      if (!hay.includes(q)) return false;
    }

    if (from || to) {
      const d = r.__date?.[dateCol] || null;
      if (!d) return false;
      if (from && d < from) return false;
      if (to && d > to) return false;
    }

    return true;
  });
}

// ---------- Table Rendering ----------
function renderTable() {
  const filtered = getFilteredRows();

  // apply sorting
  const sorted = applySort(filtered);

  // header
  els.thead.innerHTML = "";
  const tr = document.createElement("tr");

  visibleColumns.forEach(col => {
    const th = document.createElement("th");
    th.textContent = headerLabel(col);
    th.dataset.col = col;
    th.title = "Click to sort • Drag to reorder";

    // sort indicator
    if (sortState.col === col) {
      th.textContent = headerLabel(col) + (sortState.dir === "asc" ? " ▲" : " ▼");
    }

    th.addEventListener("click", () => toggleSort(col));
    tr.appendChild(th);
  });

  els.thead.appendChild(tr);

  // enable drag reorder on header row
  enableHeaderDrag(tr);

  // body
  els.tbody.innerHTML = "";
  const frag = document.createDocumentFragment();

  for (const r of sorted) {
    const rowEl = document.createElement("tr");

    for (const col of visibleColumns) {
      const td = document.createElement("td");

      if (isDateColumn(col)) {
        const d = r.__date?.[col] || null;
        td.textContent = d ? formatMMDDYYYY(d) : (r[col] || "");
      } else {
        td.innerHTML = escapeHtml(r[col] || "");
      }

      rowEl.appendChild(td);
    }

    frag.appendChild(rowEl);
  }

  els.tbody.appendChild(frag);
  els.rowCount.textContent = `${sorted.length} rows`;

  // counts should reflect current filters
  renderCounts();
}

function headerLabel(col) {
  return col; // you can customize labels here if you want shorter text
}

function isDateColumn(col) {
  return col === "Scheduled start date" || col === "Original pm due date" || col === "Scheduled end date";
}

function toggleSort(col) {
  if (sortState.col === col) {
    sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
  } else {
    sortState.col = col;
    sortState.dir = "asc";
  }
  renderTable();
}

function applySort(list) {
  const { col, dir } = sortState;
  if (!col) return list;

  const mult = dir === "asc" ? 1 : -1;
  const copy = [...list];

  copy.sort((a, b) => {
    let av, bv;

    if (isDateColumn(col)) {
      av = a.__date?.[col]?.getTime?.() ?? null;
      bv = b.__date?.[col]?.getTime?.() ?? null;
      if (av == null && bv == null) return 0;
      if (av == null) return 1 * mult;  // null dates go to bottom
      if (bv == null) return -1 * mult;
      return (av - bv) * mult;
    }

    av = (a[col] ?? "").toString();
    bv = (b[col] ?? "").toString();
    return av.localeCompare(bv, undefined, { numeric: true, sensitivity: "base" }) * mult;
  });

  return copy;
}

function enableHeaderDrag(headerRowEl) {
  // Destroy/recreate safe: Sortable attaches to the element and persists.
  // We’ll just create once if not present.
  if (headerRowEl.__sortableAttached) return;
  headerRowEl.__sortableAttached = true;

  Sortable.create(headerRowEl, {
    animation: 120,
    ghostClass: "dragging",
    onEnd: () => {
      const newOrder = [...headerRowEl.querySelectorAll("th")].map(th => th.dataset.col);
      visibleColumns = newOrder;
      saveColumnOrder(visibleColumns);
      renderTable(); // re-render to ensure body matches
    }
  });
}

// ---------- Counts ----------
function renderCounts() {
  const filtered = getFilteredRows();
  const dateCol = els.countDateColumn.value;

  // group by day then by employee
  const dayMap = new Map(); // dayStr -> Map(employee -> count)

  for (const r of filtered) {
    const assignee = (r["Assigned to"] || "").trim() || "(Unassigned)";
    const d = r.__date?.[dateCol] || null;
    const dayStr = d ? formatMMDDYYYY(d) : "(No date)";

    if (!dayMap.has(dayStr)) dayMap.set(dayStr, new Map());
    const empMap = dayMap.get(dayStr);

    empMap.set(assignee, (empMap.get(assignee) || 0) + 1);
  }

  // Sort days (real dates first, then no date)
  const days = Array.from(dayMap.keys());
  days.sort((a, b) => {
    const da = parseDateFlexible(a);
    const db = parseDateFlexible(b);
    if (a === "(No date)") return 1;
    if (b === "(No date)") return -1;
    if (da && db) return da - db;
    return a.localeCompare(b);
  });

  if (!rows.length) {
    els.countsWrap.innerHTML = `<div class="emptyState">Upload a CSV to see counts.</div>`;
    return;
  }

  if (days.length === 0) {
    els.countsWrap.innerHTML = `<div class="emptyState">No matching rows for current filters.</div>`;
    return;
  }

  // Build a compact table of day + top-level totals + expand-like detail (simple layout)
  // We'll show each day with employees under it.
  let html = `<table class="countsTable">
    <thead>
      <tr>
        <th style="width: 140px;">Day</th>
        <th>Employee</th>
        <th style="width: 90px;">Count</th>
      </tr>
    </thead>
    <tbody>
  `;

  for (const day of days) {
    const empMap = dayMap.get(day);
    const empEntries = Array.from(empMap.entries())
      .sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));

    // Day total
    const total = empEntries.reduce((sum, [, c]) => sum + c, 0);

    // First row: day + total badge, first employee
    for (let i = 0; i < empEntries.length; i++) {
      const [emp, cnt] = empEntries[i];
      if (i === 0) {
        html += `
          <tr>
            <td>
              <div>${escapeHtml(day)}</div>
              <div class="badge" style="margin-top:6px;">Total: ${total}</div>
            </td>
            <td>${escapeHtml(emp)}</td>
            <td><strong>${cnt}</strong></td>
          </tr>
        `;
      } else {
        html += `
          <tr>
            <td></td>
            <td>${escapeHtml(emp)}</td>
            <td><strong>${cnt}</strong></td>
          </tr>
        `;
      }
    }
  }

  html += `</tbody></table>`;
  els.countsWrap.innerHTML = html;
}

// ---------- Events ----------
els.fileInput.addEventListener("change", (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  if (!file.name.toLowerCase().endsWith(".csv")) {
    alert("Please upload a .csv file.");
    return;
  }

  // Reset sort on new load
  sortState = { col: null, dir: "asc" };

  parseCsvFile(file);
});

[
  els.qSearch,
  els.assignedFilter,
  els.statusFilter,
  els.typeFilter,
  els.deptFilter,
  els.dateFilterColumn,
  els.dateFrom,
  els.dateTo,
  els.countDateColumn
].forEach(el => el.addEventListener("input", renderTable));

els.btnClearFilters.addEventListener("click", () => {
  els.qSearch.value = "";
  els.assignedFilter.value = "";
  els.statusFilter.value = "";
  els.typeFilter.value = "";
  els.deptFilter.value = "";
  els.dateFrom.value = "";
  els.dateTo.value = "";
  els.dateFilterColumn.value = "Scheduled start date";
  els.countDateColumn.value = "Scheduled start date";
  renderTable();
});

els.btnResetLayout.addEventListener("click", () => {
  localStorage.removeItem(STORAGE_KEY_COL_ORDER);
  visibleColumns = [...CANON_COLUMNS];
  renderTable();
});

els.btnClearData.addEventListener("click", () => {
  rawRows = [];
  rows = [];
  els.thead.innerHTML = "";
  els.tbody.innerHTML = "";
  els.rowCount.textContent = "0 rows";
  els.colDebug.textContent = "No file loaded.";
  els.countsWrap.innerHTML = `<div class="emptyState">Upload a CSV to see counts.</div>`;

  // reset filters dropdowns
  ["assignedFilter","statusFilter","typeFilter","deptFilter"].forEach(id => {
    const sel = els[id];
    sel.innerHTML = `<option value="">All</option>`;
    sel.value = "";
  });

  els.fileInput.value = "";
});

// On first load, use saved column order even before a file is loaded
visibleColumns = loadColumnOrder() || [...CANON_COLUMNS];
