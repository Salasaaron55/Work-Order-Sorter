/* Work Order Viewer (CSV/XLSX) - client-side only
   Features:
   - Upload CSV/XLSX
   - Filters + search + date range
   - Sortable table (click headers)
   - Drag to reorder columns + remember in localStorage
   - Daily counts by employee (pivot-style)
*/

const STORAGE_KEY = "wo_viewer_column_order_v1";

const DESIRED_COLUMNS = [
  "Work Order",
  "Description",
  "Status",
  "Type",
  "Department",
  "Equipment",
  "Equipment Description",
  "Sched. Start Date",
  "Original PM Due Date",
  "Sched. End Date",
  "Assigned To"
];

const DATE_FIELDS = [
  { label: "Scheduled start date", key: "Sched. Start Date" },
  { label: "Original PM due date", key: "Original PM Due Date" },
  { label: "Scheduled end date", key: "Sched. End Date" }
];

// --- DOM
const fileInput = document.getElementById("fileInput");
const fileMeta = document.getElementById("fileMeta");

const searchInput = document.getElementById("searchInput");
const statusFilter = document.getElementById("statusFilter");
const typeFilter = document.getElementById("typeFilter");
const deptFilter = document.getElementById("deptFilter");
const assignedFilter = document.getElementById("assignedFilter");

const dateFieldSelect = document.getElementById("dateFieldSelect");
const dateFrom = document.getElementById("dateFrom");
const dateTo = document.getElementById("dateTo");

const clearFiltersBtn = document.getElementById("clearFiltersBtn");
const resetColumnsBtn = document.getElementById("resetColumnsBtn");

const dataWrap = document.getElementById("dataWrap");
const countsWrap = document.getElementById("countsWrap");

const rowCountEl = document.getElementById("rowCount");
const employeeCountEl = document.getElementById("employeeCount");
const dateRangeLabelEl = document.getElementById("dateRangeLabel");

// --- State
let allRows = [];
let filteredRows = [];
let currentSort = { col: null, dir: "asc" }; // dir: asc|desc
let columnOrder = loadColumnOrder() ?? [...DESIRED_COLUMNS];

// --- Init
initDateFieldSelect();
initFilterSelects();
wireEvents();
renderEmpty();

function wireEvents() {
  fileInput.addEventListener("change", onFileSelected);

  searchInput.addEventListener("input", applyFiltersAndRender);
  statusFilter.addEventListener("change", applyFiltersAndRender);
  typeFilter.addEventListener("change", applyFiltersAndRender);
  deptFilter.addEventListener("change", applyFiltersAndRender);
  assignedFilter.addEventListener("change", applyFiltersAndRender);

  dateFieldSelect.addEventListener("change", applyFiltersAndRender);
  dateFrom.addEventListener("change", applyFiltersAndRender);
  dateTo.addEventListener("change", applyFiltersAndRender);

  clearFiltersBtn.addEventListener("click", () => {
    searchInput.value = "";
    statusFilter.value = "";
    typeFilter.value = "";
    deptFilter.value = "";
    assignedFilter.value = "";
    dateFrom.value = "";
    dateTo.value = "";
    applyFiltersAndRender();
  });

  resetColumnsBtn.addEventListener("click", () => {
    columnOrder = [...DESIRED_COLUMNS];
    saveColumnOrder(columnOrder);
    render();
  });
}

function initDateFieldSelect() {
  dateFieldSelect.innerHTML = "";
  for (const f of DATE_FIELDS) {
    const opt = document.createElement("option");
    opt.value = f.key;
    opt.textContent = f.label;
    dateFieldSelect.appendChild(opt);
  }
  dateFieldSelect.value = "Sched. Start Date";
}

function initFilterSelects() {
  for (const sel of [statusFilter, typeFilter, deptFilter, assignedFilter]) {
    sel.innerHTML = "";
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "All";
    sel.appendChild(opt);
  }
}

function renderEmpty() {
  dataWrap.innerHTML = `<div class="empty">Upload a file to view work orders.</div>`;
  countsWrap.innerHTML = `<div class="empty">Upload a file to see counts.</div>`;
  rowCountEl.textContent = "0";
  employeeCountEl.textContent = "0";
  dateRangeLabelEl.textContent = "—";
}

async function onFileSelected(e) {
  const file = e.target.files?.[0];
  if (!file) return;

  fileMeta.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;

  try {
    const ext = (file.name.split(".").pop() || "").toLowerCase();
    if (ext === "csv") {
      allRows = await parseCSV(file);
    } else if (ext === "xlsx" || ext === "xls") {
      allRows = await parseXLSX(file);
    } else {
      throw new Error("Unsupported file type. Please upload a CSV, XLSX, or XLS.");
    }

    // Normalize headers, keep only desired columns (case-insensitive mapping)
    allRows = normalizeAndPickColumns(allRows);

    // Build dropdown values from data
    rebuildFilterOptions(allRows);

    applyFiltersAndRender();
  } catch (err) {
    console.error(err);
    dataWrap.innerHTML = `<div class="empty"><strong>Upload error:</strong> ${escapeHtml(err.message)}</div>`;
    countsWrap.innerHTML = `<div class="empty">—</div>`;
  }
}

function parseCSV(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete: (results) => {
        if (results.errors?.length) {
          reject(new Error(results.errors[0].message));
          return;
        }
        resolve(results.data || []);
      }
    });
  });
}

function parseXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Could not read the file."));
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];

        // raw: false helps with date formatting in many cases
        const json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
        resolve(json);
      } catch (e) {
        reject(new Error("Failed to parse XLSX. Make sure the sheet has a header row."));
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function normalizeAndPickColumns(rows) {
  if (!rows.length) return [];

  // Build a case-insensitive map of actual headers
  const sample = rows[0];
  const actualHeaders = Object.keys(sample);
  const headerMap = new Map();
  for (const h of actualHeaders) {
    headerMap.set(normKey(h), h);
  }

  // Map desired columns to whatever exists in file
  const resolved = {};
  for (const desired of DESIRED_COLUMNS) {
    const found = headerMap.get(normKey(desired));
    if (found) resolved[desired] = found;
  }

  // If key columns are missing, still proceed, but those columns will be blank
  const out = rows.map((r) => {
    const obj = {};
    for (const desired of DESIRED_COLUMNS) {
      const actual = resolved[desired];
      obj[desired] = actual ? (r[actual] ?? "") : "";
    }
    return obj;
  });

  // Ensure columnOrder only contains allowed columns
  columnOrder = columnOrder.filter(c => DESIRED_COLUMNS.includes(c));
  if (columnOrder.length !== DESIRED_COLUMNS.length) {
    // Add missing at end (in case storage is older)
    for (const c of DESIRED_COLUMNS) if (!columnOrder.includes(c)) columnOrder.push(c);
    saveColumnOrder(columnOrder);
  }

  return out;
}

function rebuildFilterOptions(rows) {
  const unique = (arr) => [...new Set(arr.map(v => String(v ?? "").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));

  setSelectOptions(statusFilter, unique(rows.map(r => r["Status"])));
  setSelectOptions(typeFilter, unique(rows.map(r => r["Type"])));
  setSelectOptions(deptFilter, unique(rows.map(r => r["Department"])));
  setSelectOptions(assignedFilter, unique(rows.map(r => r["Assigned To"])));
}

function setSelectOptions(select, values) {
  const current = select.value;
  select.innerHTML = "";
  const allOpt = document.createElement("option");
  allOpt.value = "";
  allOpt.textContent = "All";
  select.appendChild(allOpt);

  for (const v of values) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    select.appendChild(opt);
  }

  // Keep current selection if still present
  select.value = values.includes(current) ? current : "";
}

function applyFiltersAndRender() {
  const q = (searchInput.value || "").trim().toLowerCase();
  const status = statusFilter.value || "";
  const type = typeFilter.value || "";
  const dept = deptFilter.value || "";
  const assigned = assignedFilter.value || "";
  const dateField = dateFieldSelect.value || "Sched. Start Date";

  const from = dateFrom.value ? new Date(dateFrom.value + "T00:00:00") : null;
  const to = dateTo.value ? new Date(dateTo.value + "T23:59:59") : null;

  filteredRows = allRows.filter(r => {
    if (status && String(r["Status"]).trim() !== status) return false;
    if (type && String(r["Type"]).trim() !== type) return false;
    if (dept && String(r["Department"]).trim() !== dept) return false;
    if (assigned && String(r["Assigned To"]).trim() !== assigned) return false;

    if (q) {
      const hay = [
        r["Work Order"],
        r["Description"],
        r["Equipment"],
        r["Equipment Description"],
        r["Department"],
        r["Status"],
        r["Type"],
        r["Assigned To"]
      ].join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }

    if (from || to) {
      const d = parseDate(r[dateField]);
      if (!d) return false; // if filtering by date, rows without a date are excluded
      if (from && d < from) return false;
      if (to && d > to) return false;
    }

    return true;
  });

  // Apply sort
  if (currentSort.col) {
    const { col, dir } = currentSort;
    filteredRows.sort((a,b) => compareValues(a[col], b[col], col) * (dir === "asc" ? 1 : -1));
  }

  render();
}

function render() {
  // Stats
  rowCountEl.textContent = String(filteredRows.length);

  const employees = [...new Set(filteredRows.map(r => String(r["Assigned To"] || "").trim()).filter(Boolean))];
  employeeCountEl.textContent = String(employees.length);

  // Date range label based on selected date field
  const dateField = dateFieldSelect.value || "Sched. Start Date";
  const dates = filteredRows.map(r => toISODate(parseDate(r[dateField]))).filter(Boolean).sort();
  if (dates.length) dateRangeLabelEl.textContent = `${dates[0]} → ${dates[dates.length - 1]}`;
  else dateRangeLabelEl.textContent = "—";

  // Work order table
  if (!allRows.length) {
    renderEmpty();
    return;
  }
  dataWrap.innerHTML = "";
  dataWrap.appendChild(renderDataTable(filteredRows));

  // Counts table
  countsWrap.innerHTML = "";
  countsWrap.appendChild(renderCountsTable(filteredRows));
}

function renderDataTable(rows) {
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");

  columnOrder.forEach((col, idx) => {
    const th = document.createElement("th");
    th.textContent = col;
    th.title = "Click to sort • Drag to reorder";

    // Sort on click
    th.addEventListener("click", () => {
      if (currentSort.col === col) {
        currentSort.dir = currentSort.dir === "asc" ? "desc" : "asc";
      } else {
        currentSort.col = col;
        currentSort.dir = "asc";
      }
      applyFiltersAndRender();
    });

    // Drag + drop for column reorder
    th.setAttribute("draggable", "true");
    th.dataset.col = col;
    th.dataset.index = String(idx);

    th.addEventListener("dragstart", (e) => {
      th.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", col);
    });

    th.addEventListener("dragend", () => th.classList.remove("dragging"));

    th.addEventListener("dragover", (e) => {
      e.preventDefault();
      th.classList.add("drop-target");
      e.dataTransfer.dropEffect = "move";
    });

    th.addEventListener("dragleave", () => th.classList.remove("drop-target"));

    th.addEventListener("drop", (e) => {
      e.preventDefault();
      th.classList.remove("drop-target");
      const draggedCol = e.dataTransfer.getData("text/plain");
      const targetCol = col;
      if (!draggedCol || draggedCol === targetCol) return;

      columnOrder = reorder(columnOrder, draggedCol, targetCol);
      saveColumnOrder(columnOrder);
      render(); // re-render without changing filters
    });

    trh.appendChild(th);
  });

  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (const r of rows) {
    const tr = document.createElement("tr");
    for (const col of columnOrder) {
      const td = document.createElement("td");
      const val = r[col] ?? "";

      if (col === "Status" || col === "Type" || col === "Department" || col === "Assigned To") {
        td.innerHTML = val ? `<span class="badge">${escapeHtml(String(val))}</span>` : "";
      } else {
        td.textContent = String(val ?? "");
      }

      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  return table;
}

function renderCountsTable(rows) {
  const dateField = dateFieldSelect.value || "Sched. Start Date";

  // Group by date + employee
  const map = new Map(); // date -> Map(employee -> count)
  const employeeSet = new Set();

  for (const r of rows) {
    const d = toISODate(parseDate(r[dateField]));
    if (!d) continue;

    const empRaw = String(r["Assigned To"] ?? "").trim();
    const emp = empRaw || "Unassigned";
    employeeSet.add(emp);

    if (!map.has(d)) map.set(d, new Map());
    const inner = map.get(d);
    inner.set(emp, (inner.get(emp) || 0) + 1);
  }

  const dates = [...map.keys()].sort();
  const employees = [...employeeSet].sort((a,b)=>a.localeCompare(b));

  if (!dates.length) {
    const div = document.createElement("div");
    div.className = "empty";
    div.textContent = "No dated rows to count (try a different date field, or clear date filters).";
    return div;
  }

  // Build pivot table
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");

  const th0 = document.createElement("th");
  th0.textContent = "Date";
  trh.appendChild(th0);

  for (const emp of employees) {
    const th = document.createElement("th");
    th.textContent = emp;
    trh.appendChild(th);
  }

  const thT = document.createElement("th");
  thT.textContent = "Total";
  trh.appendChild(thT);

  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  const totalsByEmp = new Map(employees.map(e => [e, 0]));
  let grandTotal = 0;

  for (const d of dates) {
    const tr = document.createElement("tr");
    const tdDate = document.createElement("td");
    tdDate.textContent = d;
    tr.appendChild(tdDate);

    let rowTotal = 0;
    const inner = map.get(d);

    for (const emp of employees) {
      const td = document.createElement("td");
      const c = inner.get(emp) || 0;
      td.textContent = c ? String(c) : "";
      tr.appendChild(td);

      rowTotal += c;
      totalsByEmp.set(emp, (totalsByEmp.get(emp) || 0) + c);
    }

    const tdTotal = document.createElement("td");
    tdTotal.textContent = String(rowTotal);
    tr.appendChild(tdTotal);

    grandTotal += rowTotal;
    tbody.appendChild(tr);
  }

  // Totals row
  const trTot = document.createElement("tr");
  const tdLabel = document.createElement("td");
  tdLabel.innerHTML = `<span class="badge">Totals</span>`;
  trTot.appendChild(tdLabel);

  for (const emp of employees) {
    const td = document.createElement("td");
    td.textContent = String(totalsByEmp.get(emp) || 0);
    trTot.appendChild(td);
  }

  const tdGrand = document.createElement("td");
  tdGrand.textContent = String(grandTotal);
  trTot.appendChild(tdGrand);

  tbody.appendChild(trTot);
  table.appendChild(tbody);

  return table;
}

// -------- Helpers

function reorder(arr, draggedItem, targetItem) {
  const a = [...arr];
  const from = a.indexOf(draggedItem);
  const to = a.indexOf(targetItem);
  if (from < 0 || to < 0) return a;
  a.splice(from, 1);
  a.splice(to, 0, draggedItem);
  return a;
}

function loadColumnOrder() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return null;
    return parsed;
  } catch {
    return null;
  }
}

function saveColumnOrder(order) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(order));
  } catch {
    // ignore storage errors
  }
}

function normKey(s) {
  return String(s ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[.\u00A0]/g, "."); // keep periods somewhat consistent
}

// Tries to parse common date formats; returns Date or null
function parseDate(value) {
  if (!value) return null;

  // If already Date
  if (value instanceof Date && !isNaN(value)) return value;

  const s = String(value).trim();
  if (!s) return null;

  // Try ISO-ish
  const iso = new Date(s);
  if (!isNaN(iso)) return iso;

  // Try MM/DD/YYYY or M/D/YYYY
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s.*)?$/);
  if (m) {
    const mm = parseInt(m[1], 10);
    const dd = parseInt(m[2], 10);
    let yy = parseInt(m[3], 10);
    if (yy < 100) yy += 2000;
    const d = new Date(yy, mm - 1, dd);
    return isNaN(d) ? null : d;
  }

  return null;
}

function toISODate(d) {
  if (!d || isNaN(d)) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function compareValues(a, b, colName) {
  const sa = String(a ?? "").trim();
  const sb = String(b ?? "").trim();

  // Try date compare for date fields
  if (DESIRED_COLUMNS.includes(colName) && colName.toLowerCase().includes("date")) {
    const da = parseDate(sa);
    const db = parseDate(sb);
    if (da && db) return da - db;
    if (da && !db) return 1;
    if (!da && db) return -1;
    return sa.localeCompare(sb);
  }

  // Work Order often numeric-ish
  if (colName === "Work Order") {
    const na = Number(sa);
    const nb = Number(sb);
    const aIsNum = !Number.isNaN(na) && sa !== "";
    const bIsNum = !Number.isNaN(nb) && sb !== "";
    if (aIsNum && bIsNum) return na - nb;
    if (aIsNum && !bIsNum) return 1;
    if (!aIsNum && bIsNum) return -1;
  }

  return sa.localeCompare(sb, undefined, { numeric: true, sensitivity: "base" });
}

function escapeHtml(str) {
  return String(str ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
