/* Work Order Viewer (client-side)
   - Upload CSV/XLSX
   - Keep only desired columns
   - Filter by keyword / employee / status / type / dept / date range
   - Sort by clicking header
   - Drag headers to rearrange columns (saved in localStorage)
   - Summary pivot: counts per Assigned To per day (Sched. Start Date)
   - Dates displayed as MM/DD/YYYY
*/

const STORAGE_KEY = "wo_viewer_column_order_v1";

const WANTED = [
  { key: "work_order", label: "Work Order" },
  { key: "description", label: "Description" },
  { key: "status", label: "Status" },
  { key: "type", label: "Type" },
  { key: "department", label: "Department" },
  { key: "equipment", label: "Equipment" },
  { key: "equipment_description", label: "Equipment Description" },
  { key: "sched_start_date", label: "Scheduled start date" },
  { key: "original_pm_due_date", label: "Original pm due date" },
  { key: "sched_end_date", label: "Scheduled end date" },
  { key: "assigned_to", label: "Assigned to" },
];

// Header normalization helpers
function normHeader(h) {
  return String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^\w\s.]/g, ""); // keep dots because your sample uses "Sched. Start Date"
}
function makeHeaderMap(headers) {
  // Map file headers → our keys (handles variations)
  const map = {};
  const normalized = headers.map(h => ({ raw: h, n: normHeader(h) }));

  function pick(key, candidates) {
    for (const c of candidates) {
      const hit = normalized.find(x => x.n === c);
      if (hit) { map[key] = hit.raw; return; }
    }
    // fallback: contains
    for (const c of candidates) {
      const hit = normalized.find(x => x.n.includes(c));
      if (hit) { map[key] = hit.raw; return; }
    }
  }

  pick("work_order", ["work order", "workorder", "wo", "work_order"]);
  pick("description", ["description", "wo description", "work order description"]);
  pick("status", ["status"]);
  pick("type", ["type", "work type"]);
  pick("department", ["department", "dept"]);
  pick("equipment", ["equipment", "equip"]);
  pick("equipment_description", ["equipment description", "equip description"]);
  pick("sched_start_date", ["sched. start date", "scheduled start date", "sched start date", "start date"]);
  pick("original_pm_due_date", ["original pm due date", "pm due date", "original due date", "due date"]);
  pick("sched_end_date", ["sched. end date", "scheduled end date", "sched end date", "end date"]);
  pick("assigned_to", ["assigned to", "assigned_to", "assignee", "assigned"]);

  return map;
}

// Date formatting/parsing
const fmtUS = new Intl.DateTimeFormat("en-US", { year: "numeric", month: "2-digit", day: "2-digit" });

function parseDateLoose(val) {
  if (val == null || val === "") return null;

  // If already a Date
  if (val instanceof Date && !isNaN(val)) return val;

  // If it's an Excel serial number (common in XLSX conversions)
  if (typeof val === "number" && isFinite(val)) {
    // Excel's epoch starts 1899-12-30 for most systems
    const ms = Math.round((val - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d) ? null : d;
  }

  const s = String(val).trim();
  if (!s) return null;

  // If it's ISO-like: 2025-12-27
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    const d = new Date(`${iso[1]}-${iso[2]}-${iso[3]}T00:00:00`);
    return isNaN(d) ? null : d;
  }

  // If it's MM/DD/YYYY or M/D/YYYY
  const us = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (us) {
    const mm = Number(us[1]);
    const dd = Number(us[2]);
    let yyyy = Number(us[3]);
    if (yyyy < 100) yyyy += 2000;
    const d = new Date(yyyy, mm - 1, dd);
    return isNaN(d) ? null : d;
  }

  // Last resort
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function toUSDate(d) {
  if (!d) return "";
  return fmtUS.format(d); // MM/DD/YYYY
}

function toISODateOnly(d) {
  if (!d) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function clampDateRange(d, from, to) {
  if (!d) return false;
  if (from && d < from) return false;
  if (to) {
    const end = new Date(to.getFullYear(), to.getMonth(), to.getDate(), 23, 59, 59, 999);
    if (d > end) return false;
  }
  return true;
}

// CSV parsing (handles quotes/commas reasonably well)
function parseCSV(text) {
  const rows = [];
  let row = [];
  let cur = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"' && inQuotes && next === '"') {
      cur += '"';
      i++;
      continue;
    }
    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && (ch === ",")) {
      row.push(cur);
      cur = "";
      continue;
    }
    if (!inQuotes && (ch === "\n")) {
      row.push(cur);
      rows.push(row);
      row = [];
      cur = "";
      continue;
    }
    if (!inQuotes && ch === "\r") continue;

    cur += ch;
  }
  row.push(cur);
  rows.push(row);

  // Remove blank trailing lines
  while (rows.length && rows[rows.length - 1].every(x => String(x).trim() === "")) rows.pop();

  const headers = rows[0] || [];
  const dataRows = rows.slice(1);

  const objs = dataRows.map(r => {
    const o = {};
    headers.forEach((h, idx) => o[h] = r[idx] ?? "");
    return o;
  });

  return { headers, rows: objs };
}

// App state
let allRows = [];
let filteredRows = [];
let headerMap = {};
let columnOrder = loadColumnOrder();

let sortState = { key: null, dir: "asc" }; // asc/desc

// Elements
const el = {
  fileInput: document.getElementById("fileInput"),
  fileStatus: document.getElementById("fileStatus"),
  resetLayoutBtn: document.getElementById("resetLayoutBtn"),

  keywordInput: document.getElementById("keywordInput"),
  assignedSelect: document.getElementById("assignedSelect"),
  statusSelect: document.getElementById("statusSelect"),
  typeSelect: document.getElementById("typeSelect"),
  deptSelect: document.getElementById("deptSelect"),
  dateFrom: document.getElementById("dateFrom"),
  dateTo: document.getElementById("dateTo"),
  hideUnassigned: document.getElementById("hideUnassigned"),
  useCurrentFiltersForCounts: document.getElementById("useCurrentFiltersForCounts"),

  rowCount: document.getElementById("rowCount"),
  dataTable: document.getElementById("dataTable"),
  summaryTable: document.getElementById("summaryTable"),
  exportSummaryBtn: document.getElementById("exportSummaryBtn"),
};

// Init
wireEvents();
renderEmpty();

function wireEvents() {
  el.fileInput.addEventListener("change", onFileChosen);
  el.resetLayoutBtn.addEventListener("click", () => {
    localStorage.removeItem(STORAGE_KEY);
    columnOrder = loadColumnOrder(true);
    renderAll();
  });

  const filterInputs = [
    el.keywordInput, el.assignedSelect, el.statusSelect, el.typeSelect, el.deptSelect,
    el.dateFrom, el.dateTo, el.hideUnassigned, el.useCurrentFiltersForCounts
  ];
  filterInputs.forEach(x => x.addEventListener("input", applyFilters));
  filterInputs.forEach(x => x.addEventListener("change", applyFilters));

  el.exportSummaryBtn.addEventListener("click", exportSummaryCSV);
}

function loadColumnOrder(forceDefault = false) {
  if (forceDefault) return WANTED.map(x => x.key);
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return WANTED.map(x => x.key);
    const parsed = JSON.parse(raw);
    // keep only known keys, and append any missing
    const known = new Set(WANTED.map(x => x.key));
    const cleaned = parsed.filter(k => known.has(k));
    for (const k of WANTED.map(x => x.key)) if (!cleaned.includes(k)) cleaned.push(k);
    return cleaned;
  } catch {
    return WANTED.map(x => x.key);
  }
}

function saveColumnOrder() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(columnOrder));
}

async function onFileChosen(e) {
  const file = e.target.files?.[0];
  if (!file) return;

  el.fileStatus.textContent = `Loading: ${file.name}...`;

  try {
    const ext = file.name.toLowerCase().split(".").pop();
    if (ext === "csv") {
      const text = await file.text();
      const parsed = parseCSV(text);
      setData(parsed.headers, parsed.rows, file.name);
    } else if (ext === "xlsx" || ext === "xls") {
      if (!window.XLSX) throw new Error("XLSX library not loaded.");
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const firstSheet = wb.SheetNames[0];
      const ws = wb.Sheets[firstSheet];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      const headers = json.length ? Object.keys(json[0]) : [];
      setData(headers, json, file.name);
    } else {
      throw new Error("Unsupported file type. Please upload a .csv or .xlsx.");
    }
  } catch (err) {
    console.error(err);
    el.fileStatus.textContent = `Error: ${err.message}`;
    allRows = [];
    filteredRows = [];
    renderAll();
  } finally {
    el.fileInput.value = "";
  }
}

function setData(headers, rows, filename) {
  headerMap = makeHeaderMap(headers);

  // Keep only desired columns (mapped)
  allRows = rows.map(r => {
    const o = {};
    for (const col of WANTED) {
      const src = headerMap[col.key];
      o[col.key] = src ? r[src] : "";
    }

    // Parse/normalize date fields so filtering + display are consistent
    o.sched_start_date__d = parseDateLoose(o.sched_start_date);
    o.original_pm_due_date__d = parseDateLoose(o.original_pm_due_date);
    o.sched_end_date__d = parseDateLoose(o.sched_end_date);

    // Clean some common string fields
    o.assigned_to = String(o.assigned_to ?? "").trim();
    o.status = String(o.status ?? "").trim();
    o.type = String(o.type ?? "").trim();
    o.department = String(o.department ?? "").trim();

    return o;
  });

  el.fileStatus.textContent = `Loaded: ${filename} (${allRows.length} rows)`;

  // Populate filter dropdowns from data
  populateSelect(el.assignedSelect, uniqueNonEmpty(allRows.map(r => r.assigned_to)));
  populateSelect(el.statusSelect, uniqueNonEmpty(allRows.map(r => r.status)));
  populateSelect(el.typeSelect, uniqueNonEmpty(allRows.map(r => r.type)));
  populateSelect(el.deptSelect, uniqueNonEmpty(allRows.map(r => r.department)));

  // Set date pickers to min/max of scheduled start date if available
  const schedDates = allRows.map(r => r.sched_start_date__d).filter(Boolean).sort((a,b)=>a-b);
  if (schedDates.length) {
    el.dateFrom.value = toISODateOnly(schedDates[0]);
    el.dateTo.value = toISODateOnly(schedDates[schedDates.length - 1]);
  } else {
    el.dateFrom.value = "";
    el.dateTo.value = "";
  }

  // Reset sort
  sortState = { key: null, dir: "asc" };

  applyFilters();
}

function uniqueNonEmpty(arr) {
  const set = new Set();
  for (const v of arr) {
    const s = String(v ?? "").trim();
    if (s) set.add(s);
  }
  return Array.from(set).sort((a,b)=>a.localeCompare(b));
}

function populateSelect(selectEl, values) {
  const keepFirst = selectEl.options[0];
  selectEl.innerHTML = "";
  selectEl.appendChild(keepFirst);
  for (const v of values) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    selectEl.appendChild(opt);
  }
}

function applyFilters() {
  const kw = el.keywordInput.value.trim().toLowerCase();
  const assigned = el.assignedSelect.value;
  const status = el.statusSelect.value;
  const type = el.typeSelect.value;
  const dept = el.deptSelect.value;
  const hideUnassigned = el.hideUnassigned.checked;

  const from = el.dateFrom.value ? new Date(el.dateFrom.value + "T00:00:00") : null;
  const to = el.dateTo.value ? new Date(el.dateTo.value + "T00:00:00") : null;

  filteredRows = allRows.filter(r => {
    if (assigned && r.assigned_to !== assigned) return false;
    if (status && r.status !== status) return false;
    if (type && r.type !== type) return false;
    if (dept && r.department !== dept) return false;
    if (hideUnassigned && !r.assigned_to) return false;

    // Date filter uses Scheduled start date
    if ((from || to) && !clampDateRange(r.sched_start_date__d, from, to)) return false;

    if (kw) {
      const hay = [
        r.work_order, r.description, r.status, r.type, r.department,
        r.equipment, r.equipment_description, r.assigned_to
      ].map(x => String(x ?? "").toLowerCase()).join(" | ");
      if (!hay.includes(kw)) return false;
    }
    return true;
  });

  // Apply sorting (if set)
  if (sortState.key) sortRows(filteredRows, sortState.key, sortState.dir);

  renderAll();
}

function sortRows(rows, key, dir) {
  const mul = dir === "asc" ? 1 : -1;
  rows.sort((a,b) => {
    // date keys
    if (key === "sched_start_date") return mul * cmpDate(a.sched_start_date__d, b.sched_start_date__d);
    if (key === "original_pm_due_date") return mul * cmpDate(a.original_pm_due_date__d, b.original_pm_due_date__d);
    if (key === "sched_end_date") return mul * cmpDate(a.sched_end_date__d, b.sched_end_date__d);

    const av = String(a[key] ?? "");
    const bv = String(b[key] ?? "");
    // numeric-ish
    const an = Number(av);
    const bn = Number(bv);
    if (av !== "" && bv !== "" && isFinite(an) && isFinite(bn)) return mul * (an - bn);

    return mul * av.localeCompare(bv, undefined, { numeric: true, sensitivity: "base" });
  });
}

function cmpDate(a, b) {
  const ta = a ? a.getTime() : -Infinity;
  const tb = b ? b.getTime() : -Infinity;
  return ta - tb;
}

function renderAll() {
  el.rowCount.textContent = String(filteredRows.length);

  renderTable();
  renderSummary();
}

function renderEmpty() {
  el.dataTable.querySelector("thead").innerHTML = "";
  el.dataTable.querySelector("tbody").innerHTML = "";
  el.summaryTable.querySelector("thead").innerHTML = "";
  el.summaryTable.querySelector("tbody").innerHTML = "";
  el.rowCount.textContent = "0";
}

function renderTable() {
  const thead = el.dataTable.querySelector("thead");
  const tbody = el.dataTable.querySelector("tbody");

  // header
  const tr = document.createElement("tr");
  for (const key of columnOrder) {
    const col = WANTED.find(c => c.key === key);
    const th = document.createElement("th");
    th.dataset.key = key;
    th.draggable = true;

    const inner = document.createElement("div");
    inner.className = "th-inner";

    const handle = document.createElement("span");
    handle.className = "th-handle";
    handle.title = "Drag to reorder";

    const title = document.createElement("span");
    title.className = "th-title";
    title.textContent = col?.label ?? key;

    const sort = document.createElement("span");
    sort.className = "th-sort";
    if (sortState.key === key) sort.textContent = sortState.dir === "asc" ? "▲" : "▼";
    else sort.textContent = "";

    inner.appendChild(handle);
    inner.appendChild(title);
    inner.appendChild(sort);
    th.appendChild(inner);

    // click to sort
    th.addEventListener("click", (ev) => {
      // if they are dragging, ignore
      if (th.classList.contains("dragging")) return;
      toggleSort(key);
    });

    // drag events
    th.addEventListener("dragstart", onDragStart);
    th.addEventListener("dragover", onDragOver);
    th.addEventListener("drop", onDrop);
    th.addEventListener("dragend", onDragEnd);

    tr.appendChild(th);
  }
  thead.innerHTML = "";
  thead.appendChild(tr);

  // body
  const frag = document.createDocumentFragment();
  for (const r of filteredRows) {
    const rowEl = document.createElement("tr");
    for (const key of columnOrder) {
      const td = document.createElement("td");
      td.appendChild(renderCell(r, key));
      rowEl.appendChild(td);
    }
    frag.appendChild(rowEl);
  }
  tbody.innerHTML = "";
  tbody.appendChild(frag);
}

function renderCell(r, key) {
  // date display in MM/DD/YYYY
  if (key === "sched_start_date") return textNode(toUSDate(r.sched_start_date__d));
  if (key === "original_pm_due_date") return textNode(toUSDate(r.original_pm_due_date__d));
  if (key === "sched_end_date") return textNode(toUSDate(r.sched_end_date__d));

  const v = String(r[key] ?? "").trim();

  if (key === "status" && v) {
    const s = document.createElement("span");
    s.className = "cell-badge";
    s.textContent = v;
    return s;
  }
  if (!v) {
    const s = document.createElement("span");
    s.className = "cell-muted";
    s.textContent = "—";
    return s;
  }
  return textNode(v);
}

function textNode(s) {
  return document.createTextNode(s);
}

function toggleSort(key) {
  if (sortState.key !== key) {
    sortState = { key, dir: "asc" };
  } else {
    sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
  }
  sortRows(filteredRows, sortState.key, sortState.dir);
  renderAll();
}

// Drag-to-reorder headers
let dragKey = null;

function onDragStart(e) {
  dragKey = e.currentTarget.dataset.key;
  e.currentTarget.classList.add("dragging");
  e.dataTransfer.effectAllowed = "move";
  e.dataTransfer.setData("text/plain", dragKey);
}

function onDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = "move";
}

function onDrop(e) {
  e.preventDefault();
  const fromKey = e.dataTransfer.getData("text/plain");
  const toKey = e.currentTarget.dataset.key;
  if (!fromKey || !toKey || fromKey === toKey) return;

  const fromIdx = columnOrder.indexOf(fromKey);
  const toIdx = columnOrder.indexOf(toKey);
  if (fromIdx === -1 || toIdx === -1) return;

  columnOrder.splice(fromIdx, 1);
  columnOrder.splice(toIdx, 0, fromKey);

  saveColumnOrder();
  renderAll();
}

function onDragEnd(e) {
  e.currentTarget.classList.remove("dragging");
  dragKey = null;
}

// Summary pivot: counts per employee per day (Sched. Start Date + Assigned To)
function renderSummary() {
  const useFiltered = el.useCurrentFiltersForCounts.checked;
  const base = useFiltered ? filteredRows : allRows;

  const byDay = new Map(); // dayKey -> Map(employee -> count)
  const employees = new Set();

  for (const r of base) {
    const d = r.sched_start_date__d;
    if (!d) continue;
    const dayKey = toUSDate(d); // MM/DD/YYYY
    const emp = (r.assigned_to || "").trim() || "(Unassigned)";
    employees.add(emp);

    if (!byDay.has(dayKey)) byDay.set(dayKey, new Map());
    const m = byDay.get(dayKey);
    m.set(emp, (m.get(emp) || 0) + 1);
  }

  const days = Array.from(byDay.keys()).sort((a,b)=> {
    // sort by real date
    const da = parseDateLoose(a);
    const db = parseDateLoose(b);
    return (da?.getTime() ?? 0) - (db?.getTime() ?? 0);
  });
  const emps = Array.from(employees).sort((a,b)=>a.localeCompare(b));

  const thead = el.summaryTable.querySelector("thead");
  const tbody = el.summaryTable.querySelector("tbody");

  // Build header
  const trh = document.createElement("tr");
  trh.appendChild(thText("Date"));
  for (const emp of emps) trh.appendChild(thText(emp));
  trh.appendChild(thText("Total"));
  thead.innerHTML = "";
  thead.appendChild(trh);

  // Build rows
  const frag = document.createDocumentFragment();
  for (const day of days) {
    const tr = document.createElement("tr");
    tr.appendChild(tdText(day));

    const m = byDay.get(day);
    let total = 0;

    for (const emp of emps) {
      const c = m?.get(emp) || 0;
      total += c;
      tr.appendChild(tdText(c ? String(c) : "0"));
    }
    tr.appendChild(tdText(String(total)));
    frag.appendChild(tr);
  }

  // If no data
  if (!days.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = Math.max(2, emps.length + 2);
    td.className = "cell-muted";
    td.style.padding = "14px";
    td.textContent = "No scheduled start dates found to summarize (or filters removed everything).";
    tr.appendChild(td);
    tbody.innerHTML = "";
    tbody.appendChild(tr);
    return;
  }

  tbody.innerHTML = "";
  tbody.appendChild(frag);
}

function thText(s) {
  const th = document.createElement("th");
  th.textContent = s;
  return th;
}
function tdText(s) {
  const td = document.createElement("td");
  td.textContent = s;
  return td;
}

function exportSummaryCSV() {
  // export what summary currently shows
  const table = el.summaryTable;
  const rows = Array.from(table.querySelectorAll("tr"));
  if (!rows.length) return;

  const lines = rows.map(tr => {
    const cells = Array.from(tr.children).map(cell => {
      const v = cell.textContent ?? "";
      // CSV escape
      const needs = /[",\n]/.test(v);
      const escaped = v.replace(/"/g, '""');
      return needs ? `"${escaped}"` : escaped;
    });
    return cells.join(",");
  });

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "workorder_summary_counts.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
