/* Work Order Viewer (CSV + XLSX)
   - Drag headers to reorder columns (saved in localStorage)
   - Filters + counts pivot by Assigned To x chosen date field
   - MM/DD/YYYY formatting
*/

const CANON = {
  WORK_ORDER: "Work Order",
  DESCRIPTION: "Description",
  STATUS: "Status",
  TYPE: "Type",
  DEPARTMENT: "Department",
  EQUIPMENT: "Equipment",
  EQUIPMENT_DESC: "Equipment Description",
  SCHED_START: "Sched. Start Date",
  PM_DUE: "Original PM Due Date",
  SCHED_END: "Sched. End Date",
  ASSIGNED_TO: "Assigned To",
};

const DISPLAY_NAMES = {
  [CANON.WORK_ORDER]: "Work Order",
  [CANON.DESCRIPTION]: "Description",
  [CANON.STATUS]: "Status",
  [CANON.TYPE]: "Type",
  [CANON.DEPARTMENT]: "Department",
  [CANON.EQUIPMENT]: "Equipment",
  [CANON.EQUIPMENT_DESC]: "Equipment Description",
  [CANON.SCHED_START]: "Scheduled start date",
  [CANON.PM_DUE]: "Original PM due date",
  [CANON.SCHED_END]: "Scheduled end date",
  [CANON.ASSIGNED_TO]: "Assigned to",
};

const REQUIRED_CANON_ORDER = [
  CANON.WORK_ORDER,
  CANON.DESCRIPTION,
  CANON.STATUS,
  CANON.TYPE,
  CANON.DEPARTMENT,
  CANON.EQUIPMENT,
  CANON.EQUIPMENT_DESC,
  CANON.SCHED_START,
  CANON.PM_DUE,
  CANON.SCHED_END,
  CANON.ASSIGNED_TO,
];

const STORAGE_KEY = "wo_viewer_column_order_v1";

const el = (id) => document.getElementById(id);

const state = {
  rawRows: [],
  filteredRows: [],
  columnOrder: [...REQUIRED_CANON_ORDER],
  dragCol: null,
  lastFileSig: null,
};

function normalizeHeader(s) {
  return String(s ?? "")
    .trim()
    .toLowerCase()
    .replace(/[\s\.\-_/\\()]+/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function guessCanonical(header) {
  const h = normalizeHeader(header);

  // Work Order
  if (h === "workorder" || h === "wo" || h.includes("workorder")) return CANON.WORK_ORDER;

  // Description
  if (h === "description" || h.includes("desc")) return CANON.DESCRIPTION;

  // Status / Type / Department
  if (h === "status") return CANON.STATUS;
  if (h === "type") return CANON.TYPE;
  if (h === "department" || h === "dept") return CANON.DEPARTMENT;

  // Equipment & Equipment Description
  if (h === "equipment") return CANON.EQUIPMENT;
  if (h === "equipmentdescription" || (h.includes("equipment") && h.includes("description")))
    return CANON.EQUIPMENT_DESC;

  // Dates (match common variants)
  if (h.includes("sched") && h.includes("start")) return CANON.SCHED_START;
  if (h.includes("scheduled") && h.includes("start")) return CANON.SCHED_START;
  if (h.includes("original") && h.includes("pm") && h.includes("due")) return CANON.PM_DUE;
  if (h.includes("pm") && h.includes("due")) return CANON.PM_DUE;
  if (h.includes("sched") && h.includes("end")) return CANON.SCHED_END;
  if (h.includes("scheduled") && h.includes("end")) return CANON.SCHED_END;

  // Assigned To
  if (h.includes("assigned") && h.includes("to")) return CANON.ASSIGNED_TO;

  return null;
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function formatMMDDYYYY(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj)) return "";
  return `${pad2(dateObj.getMonth() + 1)}/${pad2(dateObj.getDate())}/${dateObj.getFullYear()}`;
}

function toISODate(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj)) return "";
  return `${dateObj.getFullYear()}-${pad2(dateObj.getMonth() + 1)}-${pad2(dateObj.getDate())}`;
}

// Excel serial date to JS Date (handles typical modern Excel)
function excelSerialToDate(serial) {
  // Excel incorrectly treats 1900 as leap year; base at 1899-12-30 works for most cases
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400; // seconds
  return new Date(utcValue * 1000);
}

function parseMaybeDate(v) {
  if (v == null || v === "") return null;

  // If already a Date
  if (v instanceof Date && !isNaN(v)) return v;

  // If an Excel serial number
  if (typeof v === "number" && v > 20000 && v < 60000) {
    const d = excelSerialToDate(v);
    return isNaN(d) ? null : d;
  }

  // If string, try Date parse (works if CSV is like 12/27/2025 or 2025-12-27)
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;

    // Prefer MM/DD/YYYY parsing when it looks like it
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
      const mm = Number(m[1]);
      const dd = Number(m[2]);
      const yy = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
      const d = new Date(yy, mm - 1, dd);
      return isNaN(d) ? null : d;
    }

    const d = new Date(s);
    return isNaN(d) ? null : d;
  }

  return null;
}

function loadSavedColumnOrder() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
    if (Array.isArray(saved) && saved.length) {
      // only keep columns we support, preserve order, add missing
      const filtered = saved.filter((c) => REQUIRED_CANON_ORDER.includes(c));
      const missing = REQUIRED_CANON_ORDER.filter((c) => !filtered.includes(c));
      state.columnOrder = [...filtered, ...missing];
    } else {
      state.columnOrder = [...REQUIRED_CANON_ORDER];
    }
  } catch {
    state.columnOrder = [...REQUIRED_CANON_ORDER];
  }
}

function saveColumnOrder() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.columnOrder));
}

function fileSignature(headersCanon) {
  return REQUIRED_CANON_ORDER.map((c) => (headersCanon.includes(c) ? "1" : "0")).join("");
}

function setStats(rows) {
  el("statRows").textContent = rows.length.toLocaleString();

  const people = new Set(rows.map((r) => (r[CANON.ASSIGNED_TO] || "").trim()).filter(Boolean));
  el("statPeople").textContent = people.size.toLocaleString();

  // date span based on chosen count date column
  const dateField = el("countDateColumn").value;
  const dates = rows.map((r) => r.__dates?.[dateField]).filter(Boolean);
  if (!dates.length) {
    el("statSpan").textContent = "—";
  } else {
    dates.sort((a, b) => a - b);
    el("statSpan").textContent = `${formatMMDDYYYY(dates[0])} → ${formatMMDDYYYY(dates[dates.length - 1])}`;
  }
}

function rebuildFilterOptions(rows) {
  // Unique values from RAW data
  const uniq = (arr) => Array.from(new Set(arr.filter(Boolean))).sort((a, b) => a.localeCompare(b));

  const statuses = uniq(rows.map((r) => r[CANON.STATUS]));
  const types = uniq(rows.map((r) => r[CANON.TYPE]));
  const depts = uniq(rows.map((r) => r[CANON.DEPARTMENT]));
  const people = uniq(rows.map((r) => r[CANON.ASSIGNED_TO]));

  fillSelect(el("status"), statuses, true);
  fillSelect(el("type"), types, true);
  fillSelect(el("department"), depts, true);
  fillMultiSelect(el("assignedTo"), people);
}

function fillSelect(selectEl, values, keepAllOption) {
  const current = selectEl.value;
  selectEl.innerHTML = "";
  if (keepAllOption) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "All";
    selectEl.appendChild(opt);
  }
  values.forEach((v) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    selectEl.appendChild(opt);
  });
  // restore if exists
  if ([...selectEl.options].some((o) => o.value === current)) selectEl.value = current;
}

function fillMultiSelect(selectEl, values) {
  const prev = new Set([...selectEl.selectedOptions].map((o) => o.value));
  selectEl.innerHTML = "";
  values.forEach((v) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    if (prev.has(v)) opt.selected = true;
    selectEl.appendChild(opt);
  });
}

function getSelectedMulti(selectEl) {
  return new Set([...selectEl.selectedOptions].map((o) => o.value));
}

function textHaystack(row) {
  const parts = [
    row[CANON.WORK_ORDER],
    row[CANON.DESCRIPTION],
    row[CANON.STATUS],
    row[CANON.TYPE],
    row[CANON.DEPARTMENT],
    row[CANON.EQUIPMENT],
    row[CANON.EQUIPMENT_DESC],
    row[CANON.ASSIGNED_TO],
    row[CANON.SCHED_START],
    row[CANON.PM_DUE],
    row[CANON.SCHED_END],
  ];
  return parts.map((p) => String(p ?? "")).join(" ").toLowerCase();
}

function applyFilters() {
  const kw = el("keyword").value.trim().toLowerCase();
  const status = el("status").value;
  const type = el("type").value;
  const dept = el("department").value;
  const peopleSet = getSelectedMulti(el("assignedTo"));

  const dateField = el("countDateColumn").value;
  const fromISO = el("dateFrom").value; // yyyy-mm-dd
  const toISO = el("dateTo").value;

  const from = fromISO ? new Date(fromISO + "T00:00:00") : null;
  const to = toISO ? new Date(toISO + "T23:59:59") : null;

  const rows = state.rawRows.filter((r) => {
    if (kw && !textHaystack(r).includes(kw)) return false;
    if (status && (r[CANON.STATUS] || "") !== status) return false;
    if (type && (r[CANON.TYPE] || "") !== type) return false;
    if (dept && (r[CANON.DEPARTMENT] || "") !== dept) return false;

    if (peopleSet.size) {
      const who = (r[CANON.ASSIGNED_TO] || "").trim();
      if (!peopleSet.has(who)) return false;
    }

    // Date range applies to chosen dateField
    const d = r.__dates?.[dateField] || null;
    if (from && (!d || d < from)) return false;
    if (to && (!d || d > to)) return false;

    return true;
  });

  state.filteredRows = rows;
  setStats(rows);
  renderDataTable(rows);
  renderCounts(rows);
}

function renderDataTable(rows) {
  const table = el("dataTable");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  // Header
  const tr = document.createElement("tr");
  state.columnOrder.forEach((canonKey, idx) => {
    const th = document.createElement("th");
    th.textContent = DISPLAY_NAMES[canonKey] || canonKey;
    th.draggable = true;
    th.dataset.col = canonKey;
    th.dataset.idx = String(idx);

    th.addEventListener("dragstart", (e) => {
      state.dragCol = canonKey;
      th.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
    });
    th.addEventListener("dragend", () => {
      th.classList.remove("dragging");
      state.dragCol = null;
    });
    th.addEventListener("dragover", (e) => e.preventDefault());
    th.addEventListener("drop", (e) => {
      e.preventDefault();
      const targetCol = th.dataset.col;
      if (!state.dragCol || state.dragCol === targetCol) return;

      const order = [...state.columnOrder];
      const from = order.indexOf(state.dragCol);
      const to = order.indexOf(targetCol);
      order.splice(from, 1);
      order.splice(to, 0, state.dragCol);

      state.columnOrder = order;
      saveColumnOrder();
      renderDataTable(state.filteredRows);
    });

    tr.appendChild(th);
  });
  thead.appendChild(tr);

  // Body
  if (!rows.length) {
    const empty = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = state.columnOrder.length;
    td.style.color = "var(--muted)";
    td.textContent = "No rows match your filters (or no data loaded yet).";
    empty.appendChild(td);
    tbody.appendChild(empty);
    return;
  }

  for (const r of rows) {
    const rowEl = document.createElement("tr");
    for (const canonKey of state.columnOrder) {
      const td = document.createElement("td");
      td.textContent = String(r[canonKey] ?? "");
      rowEl.appendChild(td);
    }
    tbody.appendChild(rowEl);
  }
}

function renderCounts(rows) {
  const dateField = el("countDateColumn").value;
  const whoField = CANON.ASSIGNED_TO;

  const people = Array.from(new Set(rows.map((r) => (r[whoField] || "").trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b));

  // Dates based on parsed dates for dateField
  const dates = Array.from(
    new Set(
      rows
        .map((r) => r.__dates?.[dateField])
        .filter(Boolean)
        .map((d) => toISODate(d))
    )
  ).sort();

  const pivot = new Map(); // key: person -> Map(dateISO -> count)
  for (const p of people) pivot.set(p, new Map());

  for (const r of rows) {
    const p = (r[whoField] || "").trim();
    const dObj = r.__dates?.[dateField] || null;
    if (!p || !dObj) continue;
    const dISO = toISODate(dObj);
    const m = pivot.get(p) || new Map();
    m.set(dISO, (m.get(dISO) || 0) + 1);
    pivot.set(p, m);
  }

  const table = el("countsTable");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  // If no data/dates
  if (!rows.length || !dates.length || !people.length) {
    thead.innerHTML = `<tr><th>Assigned To</th><th>Count</th></tr>`;
    tbody.innerHTML = `<tr><td style="color:var(--muted)" colspan="2">No counts to show yet.</td></tr>`;
    return;
  }

  // Header row: Assigned To + dates + Total
  const hr = document.createElement("tr");
  const h0 = document.createElement("th");
  h0.textContent = "Assigned To";
  h0.draggable = false;
  hr.appendChild(h0);

  for (const dISO of dates) {
    const th = document.createElement("th");
    const d = new Date(dISO + "T00:00:00");
    th.textContent = formatMMDDYYYY(d);
    th.draggable = false;
    hr.appendChild(th);
  }

  const hT = document.createElement("th");
  hT.textContent = "Total";
  hT.draggable = false;
  hr.appendChild(hT);
  thead.appendChild(hr);

  // Body rows per person
  const colTotals = new Map(dates.map((d) => [d, 0]));
  let grandTotal = 0;

  for (const p of people) {
    const tr = document.createElement("tr");
    const tdName = document.createElement("td");
    tdName.textContent = p;
    tr.appendChild(tdName);

    let rowTotal = 0;
    const m = pivot.get(p) || new Map();

    for (const dISO of dates) {
      const td = document.createElement("td");
      const c = m.get(dISO) || 0;
      td.textContent = String(c);
      rowTotal += c;
      colTotals.set(dISO, (colTotals.get(dISO) || 0) + c);
      tr.appendChild(td);
    }

    const tdTot = document.createElement("td");
    tdTot.textContent = String(rowTotal);
    tr.appendChild(tdTot);

    grandTotal += rowTotal;
    tbody.appendChild(tr);
  }

  // Totals row
  const trTot = document.createElement("tr");
  const tdLabel = document.createElement("td");
  tdLabel.textContent = "TOTAL";
  tdLabel.style.fontWeight = "800";
  trTot.appendChild(tdLabel);

  for (const dISO of dates) {
    const td = document.createElement("td");
    td.textContent = String(colTotals.get(dISO) || 0);
    td.style.fontWeight = "800";
    trTot.appendChild(td);
  }

  const tdGrand = document.createElement("td");
  tdGrand.textContent = String(grandTotal);
  tdGrand.style.fontWeight = "900";
  trTot.appendChild(tdGrand);
  tbody.appendChild(trTot);
}

function standardizeRows(rawObjects) {
  // Build a mapping from source headers -> canonical
  const headers = Object.keys(rawObjects[0] || {});
  const headerToCanon = new Map();
  for (const h of headers) {
    const canon = guessCanonical(h);
    if (canon) headerToCanon.set(h, canon);
  }

  // Build standardized rows with canonical keys
  const out = rawObjects.map((obj) => {
    const r = {};
    for (const canonKey of REQUIRED_CANON_ORDER) r[canonKey] = "";

    // dates parsed stored separately
    r.__dates = {
      [CANON.SCHED_START]: null,
      [CANON.PM_DUE]: null,
      [CANON.SCHED_END]: null,
    };

    for (const [srcHeader, canonKey] of headerToCanon.entries()) {
      const v = obj[srcHeader];

      // If this is one of our date columns, parse+format to MM/DD/YYYY
      if (canonKey === CANON.SCHED_START || canonKey === CANON.PM_DUE || canonKey === CANON.SCHED_END) {
        const d = parseMaybeDate(v);
        r.__dates[canonKey] = d;
        r[canonKey] = d ? formatMMDDYYYY(d) : (v == null ? "" : String(v).trim());
      } else {
        r[canonKey] = v == null ? "" : String(v).trim();
      }
    }

    return r;
  });

  // Keep a signature so saved column order applies safely across files
  const presentCanon = Array.from(new Set(Array.from(headerToCanon.values())));
  state.lastFileSig = fileSignature(presentCanon);

  return out;
}

function readCSV(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete: (results) => resolve(results.data || []),
      error: (err) => reject(err),
    });
  });
}

function readXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const wsName = wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

async function handleUpload(file) {
  if (!file) return;

  const name = file.name.toLowerCase();
  let rawObjects = [];

  try {
    if (name.endsWith(".csv")) {
      rawObjects = await readCSV(file);
    } else if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      rawObjects = await readXLSX(file);
    } else {
      alert("Please upload a .csv or .xlsx/.xls file.");
      return;
    }
  } catch (err) {
    console.error(err);
    alert("Failed to read the file. Check the console for details.");
    return;
  }

  if (!rawObjects.length) {
    alert("No rows found in that file.");
    return;
  }

  // Standardize
  const standardized = standardizeRows(rawObjects);

  // Validate at least Work Order exists
  const hasWO = standardized.some((r) => String(r[CANON.WORK_ORDER] || "").trim() !== "");
  if (!hasWO) {
    alert(
      "I loaded the file, but I couldn't find a 'Work Order' column (or it was empty). " +
      "If your headers are different, tell me what they are and I’ll map them."
    );
  }

  state.rawRows = standardized;

  // Build filter dropdowns from raw data
  rebuildFilterOptions(state.rawRows);

  // Apply filters + render
  applyFilters();
}

function resetAll() {
  // Clear UI filters
  el("keyword").value = "";
  el("status").value = "";
  el("type").value = "";
  el("department").value = "";
  el("dateFrom").value = "";
  el("dateTo").value = "";
  el("countDateColumn").value = CANON.SCHED_START;

  // Clear people selection
  [...el("assignedTo").options].forEach((o) => (o.selected = false));

  // Reset column order to saved, else default
  loadSavedColumnOrder();

  applyFilters();
}

function wireUI() {
  el("fileInput").addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    handleUpload(file);
  });

  el("resetBtn").addEventListener("click", () => resetAll());

  // Filter events
  ["keyword", "status", "type", "department", "dateFrom", "dateTo", "assignedTo", "countDateColumn"].forEach((id) => {
    el(id).addEventListener("input", applyFilters);
    el(id).addEventListener("change", applyFilters);
  });

  el("selectAllPeople").addEventListener("click", () => {
    [...el("assignedTo").options].forEach((o) => (o.selected = true));
    applyFilters();
  });

  el("clearPeople").addEventListener("click", () => {
    [...el("assignedTo").options].forEach((o) => (o.selected = false));
    applyFilters();
  });
}

// Init
loadSavedColumnOrder();
wireUI();

// Initial empty render
renderDataTable([]);
renderCounts([]);
setStats([]);
