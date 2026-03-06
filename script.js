const HEADERS = [
  "ID",
  "Floor No.",
  "Room No.",
  "Full Name",
  "Head of Family",
  "Gender",
  "Age",
  "Displaced From",
  "Marital Status",
  "Medical Condition",
  "Required Medication",
  "Blood Type",
  "Phone Number",
  "Nationality",
  "Notes"
];

const SEARCH_HEADERS = ["ID", "Full Name", "Phone Number", "Room No.", "Head of Family"];
const EDITABLE_HEADERS = HEADERS.filter((h) => h !== "ID");

const ARABIC_TO_ENGLISH_HEADERS = {
  "عدد": "ID",
  "رقم الطابق": "Floor No.",
  "رقم الغرفة": "Room No.",
  "الاسم الثلاثي": "Full Name",
  "ارباب العائلات": "Head of Family",
  "أرباب العائلات": "Head of Family",
  "رب العائلة": "Head of Family",
  "ارباب العائلة": "Head of Family",
  "الجنس": "Gender",
  "العمر": "Age",
  "نازح من منطقة": "Displaced From",
  "نازح من المنطقه": "Displaced From",
  "وضع العائلي": "Marital Status",
  "الوضع العائلي": "Marital Status",
  "وضع عائلي": "Marital Status",
  "الحالة المرضية ان وجدت": "Medical Condition",
  "الحالة المرضية إن وجدت": "Medical Condition",
  "الحالة المرضية": "Medical Condition",
  "الدواء المطلوب": "Required Medication",
  "الدواء": "Required Medication",
  "فئة الدم": "Blood Type",
  "رقم الهاتف": "Phone Number",
  "رقم تليفون": "Phone Number",
  "رقم الموبايل": "Phone Number",
  "الجنسية": "Nationality",
  "الملاحظة": "Notes",
  "الملاحظات": "Notes",
  "ملاحظة": "Notes"
};

let dataRows = [];
let selectedIds = new Set();
let fileName = "shelter_data.xlsx";
let editingIndex = -1;
let roomCapacity = 6;
let pendingSwitchToRecordsAfterAdd = false;
let roomMetadata = {}; // { "floor|room": "description" }
let fileHandle = null;
let originalWorkbook = null;
let currentRecordsSheetName = "Records";

const uiState = {
  view: "dashboard",
  page: 1,
  pageSize: 20
};

// IndexedDB for FileHandle persistence
const DB_NAME = "ShelterDataDB";
const STORE_NAME = "FileHandles";

function getDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE_NAME);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveFileHandle(handle) {
  const db = await getDB();
  const tx = db.transaction(STORE_NAME, "readwrite");
  tx.objectStore(STORE_NAME).put(handle, "recentHandle");
  return tx.complete;
}

async function getSavedFileHandle() {
  const db = await getDB();
  return new Promise((resolve) => {
    const request = db.transaction(STORE_NAME).objectStore(STORE_NAME).get("recentHandle");
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => resolve(null);
  });
}

async function autoOpenRecentFile() {
  const handle = await getSavedFileHandle();
  if (handle) {
    const recentDiv = document.getElementById("recentFileContainer");
    if (recentDiv) {
      recentDiv.classList.remove("hidden");
      const recentName = document.getElementById("recentFileName");
      recentName.textContent = handle.name;

      const openRecentBtn = document.getElementById("openRecentBtn");
      openRecentBtn.onclick = async () => {
        try {
          const permission = await handle.queryPermission({ mode: "readwrite" });
          if (permission === "granted" || (await handle.requestPermission({ mode: "readwrite" })) === "granted") {
            fileHandle = handle;
            const file = await handle.getFile();
            processExcelFile(file);
          }
        } catch (err) {
          console.error("Manual re-open failed:", err);
        }
      };

      try {
        const permission = await handle.queryPermission({ mode: "readwrite" });
        if (permission === "granted") {
          fileHandle = handle;
          const file = await handle.getFile();
          processExcelFile(file);
        }
      } catch (err) {
        console.error("Auto-open permission check failed:", err);
      }
    }
  }
}

const el = {
  fileInput: document.getElementById("fileInput"),
  saveExcel: document.getElementById("saveExcel"),
  mainContent: document.getElementById("mainContent"),
  tableHeader: document.getElementById("tableHeader"),
  tableBody: document.getElementById("tableBody"),
  searchInput: document.getElementById("searchInput"),
  filterGender: document.getElementById("filterGender"),
  filterFloor: document.getElementById("filterFloor"),
  filterRoom: document.getElementById("filterRoom"),
  filterMedical: document.getElementById("filterMedical"),
  filterAgeGroup: document.getElementById("filterAgeGroup"),
  filterNationality: document.getElementById("filterNationality"),
  filterBloodType: document.getElementById("filterBloodType"),
  clearFilters: document.getElementById("clearFilters"),
  registrationForm: document.getElementById("registrationForm"),
  btnAddAndGoRecords: document.getElementById("btnAddAndGoRecords"),
  editModal: document.getElementById("editModal"),
  editForm: document.getElementById("editForm"),
  cancelEdit: document.getElementById("cancelEdit"),
  floorsContainer: document.getElementById("floorsContainer"),
  capacityInput: document.getElementById("capacityInput"),
  pageSizeSelect: document.getElementById("pageSizeSelect"),
  bulkActions: document.getElementById("bulkActions"),
  selectedCount: document.getElementById("selectedCount"),
  bulkDeleteBtn: document.getElementById("bulkDeleteBtn"),
  bulkMoveBtn: document.getElementById("bulkMoveBtn"),
  prevPageBtn: document.getElementById("prevPageBtn"),
  nextPageBtn: document.getElementById("nextPageBtn"),
  pageInfo: document.getElementById("pageInfo"),
  resultCount: document.getElementById("resultCount"),
  libraryWarning: document.getElementById("libraryWarning"),
  headerWarning: document.getElementById("headerWarning"),
  roomModal: document.getElementById("roomModal"),
  roomModalTitle: document.getElementById("roomModalTitle"),
  roomDescriptionInput: document.getElementById("roomDescriptionInput"),
  saveRoomDescription: document.getElementById("saveRoomDescription"),
  roomOccupantsList: document.getElementById("roomOccupantsList"),
  closeRoomModal: document.getElementById("closeRoomModal")
};

const statIds = {
  totalPeople: document.getElementById("statTotalPeople"),
  totalRooms: document.getElementById("statTotalRooms"),
  totalFloors: document.getElementById("statTotalFloors"),
  medicalCases: document.getElementById("statMedicalCases"),
  male: document.getElementById("statMale"),
  female: document.getElementById("statFemale"),
  children: document.getElementById("statChildren"),
  adults: document.getElementById("statAdults"),
  avgPerRoom: document.getElementById("statAvgPerRoom"),
  busyFloor: document.getElementById("statBusyFloor")
};

bootstrap();

function bootstrap() {
  if (typeof XLSX === "undefined") {
    el.libraryWarning.classList.remove("hidden");
    return;
  }

  buildRegistrationForm();
  buildTableHeader();
  uiState.pageSize = parseInt(el.pageSizeSelect.value, 10) || 20;

  // Tab handling
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  el.fileInput.addEventListener("change", onLoadExcelFile);
  const openPickerBtn = document.getElementById("openFilePicker");
  if (openPickerBtn) {
    openPickerBtn.addEventListener("click", onOpenFilePicker);
  }

  autoOpenRecentFile();

  el.registrationForm.addEventListener("submit", onAddRecord);
  el.searchInput.addEventListener("input", onFiltersChanged);
  el.filterGender.addEventListener("change", onFiltersChanged);
  el.filterFloor.addEventListener("input", onFiltersChanged);
  el.filterRoom.addEventListener("input", onFiltersChanged);
  el.filterMedical.addEventListener("change", onFiltersChanged);
  el.filterAgeGroup.addEventListener("change", onFiltersChanged);
  el.filterNationality.addEventListener("input", onFiltersChanged);
  el.filterBloodType.addEventListener("change", onFiltersChanged);
  el.clearFilters.addEventListener("click", clearAllFilters);
  document.getElementById("findDuplicates")?.addEventListener("click", findDuplicates);
  el.saveExcel.addEventListener("click", saveExcelFile);
  el.cancelEdit.addEventListener("click", closeEditModal);
  el.editForm.addEventListener("submit", onSaveEdit);
  el.btnAddAndGoRecords.addEventListener("click", () => {
    pendingSwitchToRecordsAfterAdd = true;
    el.registrationForm.requestSubmit();
  });
  el.pageSizeSelect.addEventListener("change", () => {
    uiState.pageSize = parseInt(el.pageSizeSelect.value, 10) || 20;
    uiState.page = 1;
    renderTable();
  });
  el.prevPageBtn.addEventListener("click", () => changePage(-1));
  el.nextPageBtn.addEventListener("click", () => changePage(1));

  el.bulkDeleteBtn.addEventListener("click", onBulkDelete);
  el.bulkMoveBtn.addEventListener("click", onBulkMove);
  el.capacityInput.addEventListener("input", () => {
    roomCapacity = Math.max(1, parseInt(el.capacityInput.value, 10) || 6);
    renderRoomVisuals();
    renderTable();
  });

  el.saveRoomDescription.addEventListener("click", onSaveRoomDescription);
  el.closeRoomModal.addEventListener("click", () => el.roomModal.classList.add("hidden"));

  document.getElementById("exportStatsPDF").addEventListener("click", exportStatsPDF);
  document.getElementById("exportAllRoomsPDF").addEventListener("click", exportAllRoomsPDF);
  document.getElementById("exportRoomPDF").addEventListener("click", exportSingleRoomPDF);
  document.getElementById("statChildrenBox").addEventListener("click", exportChildrenPDF);

  startClock();
}

function startClock() {
  const timeEl = document.getElementById("currentTime");
  const dateEl = document.getElementById("currentDate");

  function update() {
    const now = new Date();
    timeEl.textContent = now.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true });
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    dateEl.textContent = now.toLocaleDateString('en-US', options);
  }

  update();
  setInterval(update, 1000);
}

function normalizeHeader(rawKey) {
  const trimmedKey = String(rawKey || "").trim();

  // Check direct English match
  if (HEADERS.includes(trimmedKey)) return trimmedKey;

  // Check Arabic mapping
  if (ARABIC_TO_ENGLISH_HEADERS[trimmedKey]) return ARABIC_TO_ENGLISH_HEADERS[trimmedKey];

  // Robust Arabic matching with canonicalization
  const canonicalKey = canonicalizeArabic(trimmedKey);
  for (const [arabic, english] of Object.entries(ARABIC_TO_ENGLISH_HEADERS)) {
    if (canonicalizeArabic(arabic) === canonicalKey) return english;
  }

  return trimmedKey;
}

function normalizeRow(rawRow) {
  const row = {};
  HEADERS.forEach((h) => {
    row[h] = "";
  });

  Object.keys(rawRow || {}).forEach((rawKey) => {
    const normalizedKey = normalizeHeader(rawKey);
    if (HEADERS.includes(normalizedKey)) {
      let val = String(rawRow[rawKey] ?? "").trim();
      if (normalizedKey === "Gender") {
        const lower = val.toLowerCase();
        if (lower === "m" || lower === "male" || val === "ذكر") val = "Male";
        else if (lower === "f" || lower === "female" || val === "أنثى") val = "Female";
      }
      row[normalizedKey] = val;
    }
  });

  return row;
}

function validateHeaders(excelHeaders) {
  const normalized = excelHeaders.map((h) => normalizeHeader(h));
  return HEADERS.filter((required) => !normalized.includes(required));
}

function canonicalizeArabic(value) {
  return String(value || "")
    .trim()
    .replace(/[\u064B-\u065F\u0670]/g, "")
    .replace(/[إأآٱ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/ة/g, "ه")
    .replace(/[^\w\u0600-\u06FF+]/g, "")
    .toLowerCase();
}

function detectHeaderRow(matrix) {
  let bestIndex = -1;
  let bestScore = -1;

  const scanRows = Math.min(matrix.length, 25);
  for (let i = 0; i < scanRows; i += 1) {
    const row = matrix[i] || [];
    const normalized = row.map((cell) => normalizeHeader(cell));
    const score = HEADERS.filter((h) => normalized.includes(h)).length;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return { bestIndex, bestScore };
}

function rowsFromMatrix(matrix, headerRowIndex) {
  const rawHeaders = matrix[headerRowIndex] || [];
  const mappedHeaders = rawHeaders.map((h) => normalizeHeader(h));
  const rows = [];

  for (let i = headerRowIndex + 1; i < matrix.length; i += 1) {
    const raw = matrix[i] || [];
    const hasValue = raw.some((cell) => String(cell || "").trim() !== "");
    if (!hasValue) continue;

    const row = {};
    mappedHeaders.forEach((header, idx) => {
      if (HEADERS.includes(header)) row[header] = String(raw[idx] ?? "").trim();
    });
    const normalized = normalizeRow(row);
    if (isMeaningfulRow(normalized)) rows.push(normalized);
  }

  return rows;
}

function isMeaningfulRow(row) {
  const primaryFields = [
    "Full Name",
    "Phone Number",
    "Floor No.",
    "Room No.",
    "Head of Family",
    "Age",
    "Gender"
  ];
  return primaryFields.some((h) => String(row[h] || "").trim() !== "");
}

async function onOpenFilePicker() {
  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{ description: "Excel Files", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }],
      multiple: false
    });
    fileHandle = handle;
    await saveFileHandle(handle);
    const file = await handle.getFile();
    processExcelFile(file);
  } catch (err) {
    console.error("Picker cancelled or failed:", err);
  }
}

function onLoadExcelFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  fileHandle = null; // Manual upload clears handle
  processExcelFile(file);
}

function processExcelFile(file) {
  fileName = file.name;
  const reader = new FileReader();

  reader.onload = (e) => {
    const bytes = new Uint8Array(e.target.result);
    const workbook = XLSX.read(bytes, { type: "array" });
    originalWorkbook = workbook;

    // Look for records sheet
    // We try to find the one with the most matching headers if there are multiple sheets
    let bestSheetName = workbook.SheetNames[0];
    let bestMatrix = [];
    let bestHeaderInfo = { bestIndex: -1, bestScore: -1 };

    workbook.SheetNames.forEach(name => {
      const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "" });
      const info = detectHeaderRow(matrix);
      if (info.bestScore > bestHeaderInfo.bestScore) {
        bestHeaderInfo = info;
        bestSheetName = name;
        bestMatrix = matrix;
      }
    });

    if (bestHeaderInfo.bestIndex < 0 || bestHeaderInfo.bestScore < 5) {
      el.headerWarning.classList.remove("hidden");
      el.headerWarning.textContent = "Could not recognize the column headers. Please ensure the file contains the required columns in English or Arabic.";
      return;
    }

    const excelHeaders = (bestMatrix[bestHeaderInfo.bestIndex] || []).map((x) => String(x).trim());
    const missingHeaders = validateHeaders(excelHeaders);

    if (missingHeaders.length) {
      el.headerWarning.classList.remove("hidden");
      el.headerWarning.textContent = `The following columns are missing: ${missingHeaders.join(" - ")}`;
      return;
    }

    el.headerWarning.classList.add("hidden");
    dataRows = rowsFromMatrix(bestMatrix, bestHeaderInfo.bestIndex);
    currentRecordsSheetName = bestSheetName;

    // Load Room Metadata if exists
    const metaSheetName = workbook.SheetNames.find(n => n === "RoomMetadata" || n === "وصف الغرف");
    if (metaSheetName) {
      const metaSheet = workbook.Sheets[metaSheetName];
      const metaRows = XLSX.utils.sheet_to_json(metaSheet);
      roomMetadata = {};
      metaRows.forEach(row => {
        const floorKey = row["Floor No."] || row["رقم الطابق"];
        const roomKey = row["Room No."] || row["رقم الغرفة"];
        const desc = row["Description"] || row["الوصف"] || "";
        if (floorKey && roomKey) {
          roomMetadata[`${floorKey}|${roomKey}`] = desc;
        }
      });
    }

    ensureUniqueNumericIds();

    el.mainContent.classList.remove("hidden");
    document.getElementById("loadPanel").classList.add("hidden");
    el.saveExcel.disabled = false;
    switchTab("dashboardTab");
    renderAll();
  };

  reader.readAsArrayBuffer(file);
}

function buildRegistrationForm() {
  el.registrationForm.innerHTML = "";

  EDITABLE_HEADERS.forEach((header) => {
    const group = document.createElement("div");
    group.className = "form-group";

    const label = document.createElement("label");
    label.textContent = header;

    let input;
    if (header === "Gender") {
      input = document.createElement("select");
      ["", "Male", "Female"].forEach((v) => {
        const option = document.createElement("option");
        option.value = v;
        option.textContent = v || "Select";
        input.appendChild(option);
      });
    } else if (header === "Blood Type") {
      input = document.createElement("select");
      ["", "A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"].forEach((v) => {
        const option = document.createElement("option");
        option.value = v;
        option.textContent = v || "Select";
        input.appendChild(option);
      });
    } else if (header === "Notes" || header === "Medical Condition") {
      input = document.createElement("textarea");
    } else {
      input = document.createElement("input");
      input.type = ["Age", "Floor No.", "Room No."].includes(header) ? "number" : "text";
    }

    input.name = header;
    input.required = ["Floor No.", "Room No.", "Full Name", "Gender"].includes(header);

    group.appendChild(label);
    group.appendChild(input);
    el.registrationForm.appendChild(group);
  });
}

function buildTableHeader() {
  el.tableHeader.innerHTML = "";

  const selectTh = document.createElement("th");
  const selectAll = document.createElement("input");
  selectAll.type = "checkbox";
  selectAll.id = "selectAllRecords";
  selectAll.onclick = (e) => toggleSelectAll(e.target.checked);
  selectTh.appendChild(selectAll);
  el.tableHeader.appendChild(selectTh);

  HEADERS.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    el.tableHeader.appendChild(th);
  });
  const actionTh = document.createElement("th");
  actionTh.textContent = "Actions";
  el.tableHeader.appendChild(actionTh);
}

function getNextId() {
  const maxId = dataRows.reduce((max, row) => {
    const v = parseInt(row["ID"], 10);
    return Number.isFinite(v) ? Math.max(max, v) : max;
  }, 0);
  return String(maxId + 1);
}

function ensureUniqueNumericIds() {
  const seen = new Set();
  dataRows.forEach((row) => {
    let id = String(row["ID"] || "").trim();
    const numeric = parseInt(id, 10);
    if (!id || Number.isNaN(numeric) || seen.has(String(numeric))) {
      id = getNextId();
    } else {
      id = String(numeric);
    }
    row["ID"] = id;
    seen.add(id);
  });
}

function onAddRecord(e) {
  e.preventDefault();
  const formData = new FormData(el.registrationForm);
  const row = {};

  HEADERS.forEach((header) => {
    if (header === "ID") {
      row[header] = getNextId();
    } else {
      row[header] = String(formData.get(header) || "").trim();
    }
  });

  dataRows.push(row);
  el.registrationForm.reset();
  if (pendingSwitchToRecordsAfterAdd) {
    switchTab("recordsTab");
    pendingSwitchToRecordsAfterAdd = false;
  }
  renderAll();
}

function getFilteredRows() {
  const q = String(el.searchInput.value || "").trim().toLowerCase();
  const gender = String(el.filterGender.value || "").trim();
  const floor = String(el.filterFloor.value || "").trim().toLowerCase();
  const room = String(el.filterRoom.value || "").trim().toLowerCase();
  const medical = String(el.filterMedical.value || "").trim();
  const ageGroup = String(el.filterAgeGroup.value || "").trim();
  const nationality = String(el.filterNationality.value || "").trim().toLowerCase();
  const bloodType = String(el.filterBloodType.value || "").trim();

  return dataRows.filter((row) => {
    const matchesSearch = !q || SEARCH_HEADERS.some((h) => String(row[h] || "").toLowerCase().includes(q));
    const matchesGender = !gender || String(row["Gender"] || "").trim() === gender;
    const matchesFloor = !floor || String(row["Floor No."] || "").toLowerCase().includes(floor);
    const matchesRoom = !room || String(row["Room No."] || "").toLowerCase().includes(room);
    const hasMedical = String(row["Medical Condition"] || "").trim() !== "";
    const matchesMedical = !medical || (medical === "yes" ? hasMedical : !hasMedical);
    const age = parseInt(String(row["Age"] || "").trim(), 10);
    const matchesAge = !ageGroup
      || (ageGroup === "children" && Number.isFinite(age) && age < 18)
      || (ageGroup === "adults" && Number.isFinite(age) && age >= 18);
    const matchesNationality = !nationality || String(row["Nationality"] || "").toLowerCase().includes(nationality);
    const matchesBlood = !bloodType || String(row["Blood Type"] || "").trim() === bloodType;

    return matchesSearch && matchesGender && matchesFloor && matchesRoom && matchesMedical && matchesAge && matchesNationality && matchesBlood;
  });
}

function getRoomCounts(rows = dataRows) {
  const map = new Map();
  rows.forEach((row) => {
    const floor = String(row["Floor No."] || "").trim();
    const room = String(row["Room No."] || "").trim();
    if (!floor && !room) return;
    const key = `${floor}|${room}`;
    map.set(key, (map.get(key) || 0) + 1);
  });
  return map;
}

function renderTable(rowsToRender = null) {
  const rows = rowsToRender || getFilteredRows();
  updateSelectAllCheckbox(rows);

  const roomCounts = getRoomCounts();
  const totalPages = Math.max(1, Math.ceil(rows.length / uiState.pageSize));
  if (uiState.page > totalPages) uiState.page = totalPages;
  const start = (uiState.page - 1) * uiState.pageSize;
  const end = start + uiState.pageSize;
  const pagedRows = rows.slice(start, end);
  el.tableBody.innerHTML = "";

  pagedRows.forEach((row) => {
    const tr = document.createElement("tr");
    const rowId = String(row["ID"]);

    if (String(row["Medical Condition"] || "").trim()) {
      tr.classList.add("medical-alert");
    }

    const selectTd = document.createElement("td");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = selectedIds.has(rowId);
    cb.onclick = (e) => toggleSelection(rowId, e.target.checked);
    selectTd.appendChild(cb);
    tr.appendChild(selectTd);

    HEADERS.forEach((h) => {
      const td = document.createElement("td");
      td.textContent = row[h] || "";
      tr.appendChild(td);
    });

    const actions = document.createElement("td");
    actions.className = "actions-cell";

    const editBtn = document.createElement("button");
    editBtn.className = "btn btn-primary";
    editBtn.textContent = "Edit";
    editBtn.type = "button";
    editBtn.addEventListener("click", () => openEditModalById(row["ID"]));

    const duplicateBtn = document.createElement("button");
    duplicateBtn.className = "btn btn-success";
    duplicateBtn.textContent = "Duplicate";
    duplicateBtn.type = "button";
    duplicateBtn.addEventListener("click", () => duplicateById(row["ID"]));

    const deleteBtn = document.createElement("button");
    deleteBtn.className = "btn btn-danger";
    deleteBtn.textContent = "Delete";
    deleteBtn.type = "button";
    deleteBtn.addEventListener("click", () => deleteById(row["ID"]));

    actions.appendChild(editBtn);
    actions.appendChild(duplicateBtn);
    actions.appendChild(deleteBtn);
    tr.appendChild(actions);

    const floor = String(row["Floor No."] || "").trim();
    const room = String(row["Room No."] || "").trim();
    const occupancy = roomCounts.get(`${floor}|${room}`) || 0;
    if (occupancy > roomCapacity) {
      tr.title = `Warning: Room ${room} on Floor ${floor} has ${occupancy} people`;
    }

    el.tableBody.appendChild(tr);
  });

  el.resultCount.textContent = `${rows.length} Results`;
  updateBulkActionsUI();
  el.pageInfo.textContent = `Page ${uiState.page} of ${totalPages}`;
  el.prevPageBtn.disabled = uiState.page <= 1;
  el.nextPageBtn.disabled = uiState.page >= totalPages;
}

function renderDashboard() {
  statIds.totalPeople.textContent = String(dataRows.length);

  const floors = new Set();
  const roomsMap = new Map(); // "floor|room" -> count
  const floorCounts = new Map();
  let medical = 0;
  let male = 0;
  let female = 0;
  let children = 0;
  let adults = 0;

  dataRows.forEach((row) => {
    const floor = String(row["Floor No."] || "").trim();
    const room = String(row["Room No."] || "").trim();
    const gender = String(row["Gender"] || "").trim();
    const age = parseInt(String(row["Age"] || "").trim(), 10);

    if (floor) {
      floors.add(floor);
      floorCounts.set(floor, (floorCounts.get(floor) || 0) + 1);
    }

    const roomKey = `${floor}|${room}`;
    roomsMap.set(roomKey, (roomsMap.get(roomKey) || 0) + 1);

    if (String(row["Medical Condition"] || "").trim()) medical += 1;
    if (gender === "Male") male += 1;
    if (gender === "Female") female += 1;

    if (Number.isFinite(age)) {
      if (age < 18) children += 1;
      else adults += 1;
    }
  });

  let fullRooms = 0;
  let overRooms = 0;
  roomsMap.forEach((count) => {
    if (count === roomCapacity) fullRooms++;
    else if (count > roomCapacity) overRooms++;
  });

  statIds.totalRooms.textContent = String(roomsMap.size);
  statIds.totalFloors.textContent = String(floors.size);
  statIds.medicalCases.textContent = String(medical);
  statIds.male.textContent = String(male);
  statIds.female.textContent = String(female);
  statIds.children.textContent = String(children);
  statIds.adults.textContent = String(adults);

  const avg = roomsMap.size > 0 ? (dataRows.length / roomsMap.size).toFixed(1) : "0";
  if (statIds.avgPerRoom) statIds.avgPerRoom.textContent = avg;

  let busyFloor = "-";
  let maxCount = -1;
  floorCounts.forEach((count, floor) => {
    if (count > maxCount) {
      maxCount = count;
      busyFloor = floor;
    }
  });
  if (statIds.busyFloor) statIds.busyFloor.textContent = busyFloor;

  updateCircularCharts(dataRows.length, medical);
}

function updateCircularCharts(total, medical) {
  const circleTotal = document.getElementById('circleTotalPeople');
  const circleMedical = document.getElementById('circleMedicalCases');

  if (circleTotal) {
    // Total people doesn't have a % limit, let's assume 1000 is 100% for visual
    const percentage = Math.min(100, (total / 1000) * 100);
    circleTotal.setAttribute('stroke-dasharray', `${percentage}, 100`);
  }

  if (circleMedical) {
    const percentage = total > 0 ? (medical / total) * 100 : 0;
    circleMedical.setAttribute('stroke-dasharray', `${percentage}, 100`);
  }
}

function renderRoomVisuals() {
  el.floorsContainer.innerHTML = "";
  const floorMap = new Map(); // floor -> [room]

  dataRows.forEach(row => {
    const f = String(row["Floor No."] || "").trim() || "No Floor";
    const r = String(row["Room No."] || "").trim() || "No Room";
    if (!floorMap.has(f)) floorMap.set(f, new Set());
    floorMap.get(f).add(r);
  });

  if (floorMap.size === 0) {
    el.floorsContainer.textContent = "No room data currently available.";
    return;
  }

  const sortedFloors = Array.from(floorMap.keys()).sort();
  sortedFloors.forEach(floorName => {
    const floorBlock = document.createElement("div");
    floorBlock.className = "floor-block";

    const title = document.createElement("h3");
    title.className = "floor-title";
    title.textContent = `Floor: ${floorName}`;
    floorBlock.appendChild(title);

    const grid = document.createElement("div");
    grid.className = "rooms-grid";

    const sortedRooms = Array.from(floorMap.get(floorName)).sort();
    sortedRooms.forEach(roomName => {
      const occupants = dataRows.filter(r =>
        (String(r["Floor No."] || "").trim() || "No Floor") === floorName &&
        (String(r["Room No."] || "").trim() || "No Room") === roomName
      );

      const card = document.createElement("div");
      card.className = "room-card";

      const count = occupants.length;
      if (count === 0) card.classList.add("empty");
      else if (count <= roomCapacity * 0.5) card.classList.add("low");
      else if (count < roomCapacity) card.classList.add("medium");
      else if (count === roomCapacity) card.classList.add("full");
      else card.classList.add("over");

      const roomKey = `${floorName}|${roomName}`;
      const desc = roomMetadata[roomKey] || "";

      card.innerHTML = `
        <div class="room-number">Room ${roomName}</div>
        <div class="room-occupancy">${count} / ${roomCapacity} People</div>
        <div class="room-desc-preview">${desc}</div>
      `;

      card.addEventListener("click", () => openRoomModal(floorName, roomName, occupants));
      grid.appendChild(card);
    });

    floorBlock.appendChild(grid);
    el.floorsContainer.appendChild(floorBlock);
  });
}

let currentActiveRoomKey = null;

function openRoomModal(floor, room, occupants) {
  currentActiveRoomKey = `${floor}|${room}`;
  el.roomModalTitle.textContent = `Floor ${floor} | Room ${room} Details`;
  el.roomDescriptionInput.value = roomMetadata[currentActiveRoomKey] || "";

  el.roomOccupantsList.innerHTML = "";
  occupants.forEach(person => {
    const div = document.createElement("div");
    div.className = "occupant-item";
    div.innerHTML = `
      <div class="occupant-info">
        <span class="occupant-name">${person["Full Name"]}</span>
        <span class="occupant-meta">ID: ${person["ID"]} | Phone: ${person["Phone Number"] || "---"}</span>
      </div>
      <div class="occupant-actions">
        <button class="btn btn-secondary btn-sm" onclick="movePerson('${person["ID"]}')">Move</button>
      </div>
    `;
    el.roomOccupantsList.appendChild(div);
  });

  el.roomModal.classList.remove("hidden");
}

function onSaveRoomDescription() {
  if (currentActiveRoomKey) {
    roomMetadata[currentActiveRoomKey] = el.roomDescriptionInput.value;
    renderRoomVisuals();
    el.roomModal.classList.add("hidden");
  }
}

window.movePerson = function(personId) {
  const person = dataRows.find(r => String(r["ID"]) === String(personId));
  if (!person) return;

  const newFloor = prompt("Enter new Floor No.:", person["Floor No."]);
  if (newFloor === null) return;
  const newRoom = prompt("Enter new Room No.:", person["Room No."]);
  if (newRoom === null) return;

  person["Floor No."] = newFloor.trim();
  person["Room No."] = newRoom.trim();

  el.roomModal.classList.add("hidden");
  renderAll();
};

function openEditModalById(id) {
  editingIndex = dataRows.findIndex((r) => String(r["ID"]) === String(id));
  if (editingIndex < 0) return;

  const row = dataRows[editingIndex];
  el.editForm.innerHTML = "";

  HEADERS.forEach((header) => {
    const group = document.createElement("div");
    group.className = "form-group";

    const label = document.createElement("label");
    label.textContent = header;

    const input = document.createElement("input");
    input.name = header;
    input.value = row[header] || "";

    if (header === "ID") {
      input.readOnly = true;
      input.disabled = true;
    }

    group.appendChild(label);
    group.appendChild(input);
    el.editForm.appendChild(group);
  });

  el.editModal.classList.remove("hidden");
}

function closeEditModal() {
  el.editModal.classList.add("hidden");
  editingIndex = -1;
}

function onSaveEdit(e) {
  e.preventDefault();
  if (editingIndex < 0) return;

  const formData = new FormData(el.editForm);
  EDITABLE_HEADERS.forEach((header) => {
    dataRows[editingIndex][header] = String(formData.get(header) || "").trim();
  });

  closeEditModal();
  renderAll();
}

function deleteById(id) {
  const index = dataRows.findIndex((r) => String(r["ID"]) === String(id));
  if (index < 0) return;

  const ok = window.confirm("Are you sure you want to delete this record?");
  if (!ok) return;

  dataRows.splice(index, 1);
  renderAll();
}

function duplicateById(id) {
  const index = dataRows.findIndex((r) => String(r["ID"]) === String(id));
  if (index < 0) return;

  const original = dataRows[index];
  const copy = { ...original };
  copy["ID"] = getNextId();
  copy["Full Name"] = original["Full Name"] + " (Copy)";

  dataRows.push(copy);
  renderAll();
}

async function saveExcelFile() {
  const rowsForExport = dataRows.map((row) => {
    const out = {};
    HEADERS.forEach((h) => {
      out[h] = row[h] || "";
    });
    return out;
  });

  const ws = XLSX.utils.json_to_sheet(rowsForExport, { header: HEADERS });

  // Use original workbook to preserve other sheets
  const wb = originalWorkbook || XLSX.utils.book_new();

  // Update or Add Records sheet
  const recordsSheetName = currentRecordsSheetName || "Records";
  if (wb.SheetNames.includes(recordsSheetName)) {
    wb.Sheets[recordsSheetName] = ws;
  } else {
    XLSX.utils.book_append_sheet(wb, ws, recordsSheetName);
  }

  // Save Room Metadata
  const metaForExport = Object.entries(roomMetadata).map(([key, desc]) => {
    const [floor, room] = key.split("|");
    return {
      "Floor No.": floor,
      "Room No.": room,
      "Description": desc
    };
  });

  const metaSheetName = "RoomMetadata";
  if (metaForExport.length > 0) {
    const wsMeta = XLSX.utils.json_to_sheet(metaForExport);
    if (wb.SheetNames.includes(metaSheetName)) {
      wb.Sheets[metaSheetName] = wsMeta;
    } else {
      XLSX.utils.book_append_sheet(wb, wsMeta, metaSheetName);
    }
  }

  if (fileHandle) {
    try {
      const writable = await fileHandle.createWritable();
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      await writable.write(wbout);
      await writable.close();
      alert("Saved successfully to the original file.");
    } catch (err) {
      console.error("Failed to save via File System API:", err);
      XLSX.writeFile(wb, fileName || "shelter_data.xlsx");
    }
  } else {
    XLSX.writeFile(wb, fileName || "shelter_data.xlsx");
  }
}

function renderAll() {
  selectedIds.clear();
  renderDashboard();
  renderRoomVisuals();
  renderTable();
  updateCharts();
}

function onFiltersChanged() {
  uiState.page = 1;
  renderTable();
}

function toggleSelection(id, isSelected) {
  if (isSelected) {
    selectedIds.add(id);
  } else {
    selectedIds.delete(id);
  }
  updateBulkActionsUI();
}

function toggleSelectAll(isSelected) {
  const filtered = getFilteredRows();
  filtered.forEach(row => {
    const id = String(row["ID"]);
    if (isSelected) selectedIds.add(id);
    else selectedIds.delete(id);
  });
  renderTable();
}

function updateSelectAllCheckbox(filteredRows) {
  const selectAll = document.getElementById("selectAllRecords");
  if (!selectAll) return;
  if (filteredRows.length === 0) {
    selectAll.checked = false;
    return;
  }
  const allFilteredSelected = filteredRows.every(row => selectedIds.has(String(row["ID"])));
  selectAll.checked = allFilteredSelected;
}

function updateBulkActionsUI() {
  if (selectedIds.size > 0) {
    el.bulkActions.classList.remove("hidden");
    el.selectedCount.textContent = selectedIds.size;
  } else {
    el.bulkActions.classList.add("hidden");
  }
}

function onBulkDelete() {
  if (selectedIds.size === 0) return;
  const ok = window.confirm(`Delete ${selectedIds.size} selected records?`);
  if (!ok) return;

  dataRows = dataRows.filter(row => !selectedIds.has(String(row["ID"])));
  renderAll();
}

function onBulkMove() {
  if (selectedIds.size === 0) return;
  const floor = prompt("Move selected to which Floor No.?");
  if (floor === null) return;
  const room = prompt("Move selected to which Room No.?");
  if (room === null) return;

  dataRows.forEach(row => {
    if (selectedIds.has(String(row["ID"]))) {
      row["Floor No."] = floor.trim();
      row["Room No."] = room.trim();
    }
  });
  renderAll();
}

function clearAllFilters() {
  el.searchInput.value = "";
  el.filterGender.value = "";
  el.filterFloor.value = "";
  el.filterRoom.value = "";
  el.filterMedical.value = "";
  el.filterAgeGroup.value = "";
  el.filterNationality.value = "";
  el.filterBloodType.value = "";
  onFiltersChanged();
}

function changePage(delta) {
  const filteredRows = getFilteredRows();
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / uiState.pageSize));
  uiState.page = Math.max(1, Math.min(totalPages, uiState.page + delta));
  renderTable();
}

function exportToPDF(element, filename) {
  const opt = {
    margin: 10,
    filename: filename,
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
  };
  html2pdf().set(opt).from(element).save();
}

function exportStatsPDF() {
  const container = document.getElementById("advancedStatsContainer");
  exportToPDF(container, "advanced_stats.pdf");
}

function exportAllRoomsPDF() {
  const container = document.getElementById("floorsContainer");
  exportToPDF(container, "all_rooms_analysis.pdf");
}

function exportSingleRoomPDF() {
  const modalContent = document.querySelector("#roomModal .modal-card").cloneNode(true);
  // Remove buttons from the clone before exporting
  modalContent.querySelectorAll(".actions-row").forEach(el => el.remove());
  modalContent.querySelector("textarea").replaceWith(document.createElement("p")).textContent = el.roomDescriptionInput.value;

  exportToPDF(modalContent, `Room_Details_${currentActiveRoomKey}.pdf`);
}

function exportChildrenPDF() {
  const children = dataRows.filter(r => {
    const age = parseInt(r["Age"], 10);
    return Number.isFinite(age) && age < 18;
  });

  if (children.length === 0) {
    alert("No children found to export.");
    return;
  }

  const wrapper = document.createElement("div");
  wrapper.style.padding = "20px";
  wrapper.innerHTML = `<h1>Children Records (${children.length})</h1>`;

  const table = document.createElement("table");
  table.style.width = "100%";
  table.style.borderCollapse = "collapse";

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  ["ID", "Full Name", "Age", "Gender", "Floor No.", "Room No."].forEach(h => {
    const th = document.createElement("th");
    th.textContent = h;
    th.style.border = "1px solid #ddd";
    th.style.padding = "8px";
    th.style.textAlign = "left";
    th.style.backgroundColor = "#f2f2f2";
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  children.forEach(c => {
    const tr = document.createElement("tr");
    ["ID", "Full Name", "Age", "Gender", "Floor No.", "Room No."].forEach(h => {
      const td = document.createElement("td");
      td.textContent = c[h] || "";
      td.style.border = "1px solid #ddd";
      td.style.padding = "8px";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  wrapper.appendChild(table);

  exportToPDF(wrapper, "children_records.pdf");
}

function switchTab(tabId) {
  document.querySelectorAll('.tab-pane').forEach(pane => pane.classList.add('hidden'));
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.remove('active');
    btn.setAttribute('aria-selected', 'false');
  });

  document.getElementById(tabId).classList.remove('hidden');
  const activeBtn = document.querySelector(`[data-tab="${tabId}"]`);
  activeBtn.classList.add('active');
  activeBtn.setAttribute('aria-selected', 'true');

  if (tabId === 'statsTab') {
    updateCharts();
  }
}

let genderChart = null;
let ageChart = null;

function findDuplicates() {
  const seen = new Map(); // Name -> Array of IDs
  const duplicates = new Set();

  dataRows.forEach(row => {
    const name = String(row["Full Name"] || "").trim().toLowerCase();
    if (name) {
      if (seen.has(name)) {
        seen.get(name).push(row["ID"]);
        duplicates.add(name);
      } else {
        seen.set(name, [row["ID"]]);
      }
    }
  });

  if (duplicates.size === 0) {
    alert("No potential duplicates found by Name.");
    return;
  }

  const dupIds = new Set();
  duplicates.forEach(name => {
    seen.get(name).forEach(id => dupIds.add(id));
  });

  uiState.page = 1;
  const filteredRows = dataRows.filter(row => dupIds.has(String(row["ID"])));
  renderTable(filteredRows);
  alert(`Found ${duplicates.size} groups of potential duplicates.`);
}

function updateCharts() {
  if (typeof Chart === 'undefined') return;

  const male = dataRows.filter(r => r["Gender"] === "Male").length;
  const female = dataRows.filter(r => r["Gender"] === "Female").length;

  const ageBrackets = {
    '0-2': 0,
    '3-5': 0,
    '6-14': 0,
    '15-17': 0,
    '18-59': 0,
    '60+': 0,
    'Unknown': 0
  };

  dataRows.forEach(r => {
    const age = parseInt(r["Age"], 10);
    if (!Number.isFinite(age)) {
      ageBrackets['Unknown']++;
    } else if (age <= 2) ageBrackets['0-2']++;
    else if (age <= 5) ageBrackets['3-5']++;
    else if (age <= 14) ageBrackets['6-14']++;
    else if (age <= 17) ageBrackets['15-17']++;
    else if (age <= 59) ageBrackets['18-59']++;
    else ageBrackets['60+']++;
  });

  const ctxGender = document.getElementById('genderChart')?.getContext('2d');
  if (ctxGender) {
    if (genderChart) genderChart.destroy();
    genderChart = new Chart(ctxGender, {
      type: 'doughnut',
      data: {
        labels: ['Males', 'Females'],
        datasets: [{
          data: [male, female],
          backgroundColor: ['#2563eb', '#db2777']
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false
      }
    });
  }

  const ctxAge = document.getElementById('ageChart')?.getContext('2d');
  if (ctxAge) {
    if (ageChart) ageChart.destroy();
    ageChart = new Chart(ctxAge, {
      type: 'bar',
      data: {
        labels: Object.keys(ageBrackets),
        datasets: [{
          label: 'Number of People',
          data: Object.values(ageBrackets),
          backgroundColor: [
            '#60a5fa', '#34d399', '#fbbf24', '#f87171', '#818cf8', '#a78bfa', '#9ca3af'
          ]
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false }
        },
        scales: {
          y: { beginAtZero: true, ticks: { stepSize: 1 } }
        }
      }
    });
  }
}
