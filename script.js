const HEADERS = [
  "عدد",
  "رقم الطابق",
  "رقم الغرفة",
  "الاسم الثلاثي",
  "ارباب العائلات",
  "الجنس",
  "العمر",
  "نازح من منطقة",
  "وضع العائلي",
  "الحالة المرضية ان وجدت",
  "الدواء المطلوب",
  "فئة الدم",
  "رقم الهاتف",
  "الجنسية",
  "الملاحظة"
];

const SEARCH_HEADERS = ["عدد", "الاسم الثلاثي", "رقم الهاتف", "رقم الغرفة", "ارباب العائلات"];
const EDITABLE_HEADERS = HEADERS.filter((h) => h !== "عدد");
const HEADER_ALIASES = {
  "أرباب العائلات": "ارباب العائلات",
  "رب العائلة": "ارباب العائلات",
  "ارباب العائلة": "ارباب العائلات",
  "الوضع العائلي": "وضع العائلي",
  "وضع عائلي": "وضع العائلي",
  "نازح من المنطقه": "نازح من منطقة",
  "الحالة المرضية": "الحالة المرضية ان وجدت",
  "الحالة المرضية إن وجدت": "الحالة المرضية ان وجدت",
  "الدواء": "الدواء المطلوب",
  "رقم تليفون": "رقم الهاتف",
  "رقم الموبايل": "رقم الهاتف",
  "الملاحظات": "الملاحظة",
  "ملاحظة": "الملاحظة"
};

let dataRows = [];
let fileName = "shelter_data.xlsx";
let editingIndex = -1;
let roomCapacity = 6;
let pendingSwitchToRecordsAfterAdd = false;
let roomMetadata = {}; // { "floor|room": "description" }

const uiState = {
  view: "dashboard",
  page: 1,
  pageSize: 20
};

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
  clearFilters: document.getElementById("clearFilters"),
  registrationForm: document.getElementById("registrationForm"),
  btnAddAndGoRecords: document.getElementById("btnAddAndGoRecords"),
  editModal: document.getElementById("editModal"),
  editForm: document.getElementById("editForm"),
  cancelEdit: document.getElementById("cancelEdit"),
  floorsContainer: document.getElementById("floorsContainer"),
  capacityInput: document.getElementById("capacityInput"),
  pageSizeSelect: document.getElementById("pageSizeSelect"),
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
  el.registrationForm.addEventListener("submit", onAddRecord);
  el.searchInput.addEventListener("input", onFiltersChanged);
  el.filterGender.addEventListener("change", onFiltersChanged);
  el.filterFloor.addEventListener("input", onFiltersChanged);
  el.filterRoom.addEventListener("input", onFiltersChanged);
  el.filterMedical.addEventListener("change", onFiltersChanged);
  el.filterAgeGroup.addEventListener("change", onFiltersChanged);
  el.clearFilters.addEventListener("click", clearAllFilters);
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
  el.capacityInput.addEventListener("input", () => {
    roomCapacity = Math.max(1, parseInt(el.capacityInput.value, 10) || 6);
    renderRoomVisuals();
    renderTable();
  });

  el.saveRoomDescription.addEventListener("click", onSaveRoomDescription);
  el.closeRoomModal.addEventListener("click", () => el.roomModal.classList.add("hidden"));

  startClock();
}

function startClock() {
  const timeEl = document.getElementById("currentTime");
  const dateEl = document.getElementById("currentDate");

  function update() {
    const now = new Date();
    timeEl.textContent = now.toLocaleTimeString('en-GB');
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    dateEl.textContent = now.toLocaleDateString('ar-LB', options);
  }

  update();
  setInterval(update, 1000);
}

function normalizeHeader(rawKey) {
  const key = canonicalizeArabic(rawKey);
  const direct = HEADERS.find((h) => canonicalizeArabic(h) === key);
  if (direct) return direct;

  const aliased = Object.keys(HEADER_ALIASES).find((h) => canonicalizeArabic(h) === key);
  if (aliased) return HEADER_ALIASES[aliased];

  return String(rawKey || "").trim();
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
      if (normalizedKey === "الجنس") {
        const lower = val.toLowerCase();
        if (lower === "m" || lower === "male") val = "ذكر";
        else if (lower === "f" || lower === "female") val = "أنثى";
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
    "الاسم الثلاثي",
    "رقم الهاتف",
    "رقم الطابق",
    "رقم الغرفة",
    "ارباب العائلات",
    "العمر",
    "الجنس"
  ];
  return primaryFields.some((h) => String(row[h] || "").trim() !== "");
}

function onLoadExcelFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  fileName = file.name;
  const reader = new FileReader();

  reader.onload = (e) => {
    const bytes = new Uint8Array(e.target.result);
    const workbook = XLSX.read(bytes, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const { bestIndex, bestScore } = detectHeaderRow(matrix);

    if (bestIndex < 0 || bestScore < 5) {
      el.headerWarning.classList.remove("hidden");
      el.headerWarning.textContent = "لم أتعرف على صف الأعمدة. تأكد أن الملف يحتوي أعمدة السجل بالعربية.";
      return;
    }

    const excelHeaders = (matrix[bestIndex] || []).map((x) => String(x).trim());
    const missingHeaders = validateHeaders(excelHeaders);

    if (missingHeaders.length) {
      el.headerWarning.classList.remove("hidden");
      el.headerWarning.textContent = `الأعمدة التالية مفقودة من الملف: ${missingHeaders.join(" - ")}`;
      return;
    }

    el.headerWarning.classList.add("hidden");
    dataRows = rowsFromMatrix(matrix, bestIndex);

    // Load Room Metadata if exists
    const metaSheetName = workbook.SheetNames.find(n => n === "RoomMetadata" || n === "وصف الغرف");
    if (metaSheetName) {
      const metaSheet = workbook.Sheets[metaSheetName];
      const metaRows = XLSX.utils.sheet_to_json(metaSheet);
      roomMetadata = {};
      metaRows.forEach(row => {
        const key = `${row["رقم الطابق"]}|${row["رقم الغرفة"]}`;
        roomMetadata[key] = row["الوصف"] || "";
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
    if (header === "الجنس") {
      input = document.createElement("select");
      ["", "ذكر", "أنثى"].forEach((v) => {
        const option = document.createElement("option");
        option.value = v;
        option.textContent = v || "اختر";
        input.appendChild(option);
      });
    } else if (header === "فئة الدم") {
      input = document.createElement("select");
      ["", "A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"].forEach((v) => {
        const option = document.createElement("option");
        option.value = v;
        option.textContent = v || "اختر";
        input.appendChild(option);
      });
    } else if (header === "الملاحظة" || header === "الحالة المرضية ان وجدت") {
      input = document.createElement("textarea");
    } else {
      input = document.createElement("input");
      input.type = ["العمر", "رقم الطابق", "رقم الغرفة"].includes(header) ? "number" : "text";
    }

    input.name = header;
    input.required = ["رقم الطابق", "رقم الغرفة", "الاسم الثلاثي", "الجنس"].includes(header);

    group.appendChild(label);
    group.appendChild(input);
    el.registrationForm.appendChild(group);
  });
}

function buildTableHeader() {
  el.tableHeader.innerHTML = "";
  HEADERS.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    el.tableHeader.appendChild(th);
  });
  const actionTh = document.createElement("th");
  actionTh.textContent = "إجراءات";
  el.tableHeader.appendChild(actionTh);
}

function getNextId() {
  const maxId = dataRows.reduce((max, row) => {
    const v = parseInt(row["عدد"], 10);
    return Number.isFinite(v) ? Math.max(max, v) : max;
  }, 0);
  return String(maxId + 1);
}

function ensureUniqueNumericIds() {
  const seen = new Set();
  dataRows.forEach((row) => {
    let id = String(row["عدد"] || "").trim();
    const numeric = parseInt(id, 10);
    if (!id || Number.isNaN(numeric) || seen.has(String(numeric))) {
      id = getNextId();
    } else {
      id = String(numeric);
    }
    row["عدد"] = id;
    seen.add(id);
  });
}

function onAddRecord(e) {
  e.preventDefault();
  const formData = new FormData(el.registrationForm);
  const row = {};

  HEADERS.forEach((header) => {
    if (header === "عدد") {
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

  return dataRows.filter((row) => {
    const matchesSearch = !q || SEARCH_HEADERS.some((h) => String(row[h] || "").toLowerCase().includes(q));
    const matchesGender = !gender || String(row["الجنس"] || "").trim() === gender;
    const matchesFloor = !floor || String(row["رقم الطابق"] || "").toLowerCase().includes(floor);
    const matchesRoom = !room || String(row["رقم الغرفة"] || "").toLowerCase().includes(room);
    const hasMedical = String(row["الحالة المرضية ان وجدت"] || "").trim() !== "";
    const matchesMedical = !medical || (medical === "yes" ? hasMedical : !hasMedical);
    const age = parseInt(String(row["العمر"] || "").trim(), 10);
    const matchesAge = !ageGroup
      || (ageGroup === "children" && Number.isFinite(age) && age < 18)
      || (ageGroup === "adults" && Number.isFinite(age) && age >= 18);

    return matchesSearch && matchesGender && matchesFloor && matchesRoom && matchesMedical && matchesAge;
  });
}

function getRoomCounts(rows = dataRows) {
  const map = new Map();
  rows.forEach((row) => {
    const floor = String(row["رقم الطابق"] || "").trim();
    const room = String(row["رقم الغرفة"] || "").trim();
    if (!floor && !room) return;
    const key = `${floor}|${room}`;
    map.set(key, (map.get(key) || 0) + 1);
  });
  return map;
}

function renderTable() {
  const filteredRows = getFilteredRows();
  const roomCounts = getRoomCounts();
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / uiState.pageSize));
  if (uiState.page > totalPages) uiState.page = totalPages;
  const start = (uiState.page - 1) * uiState.pageSize;
  const end = start + uiState.pageSize;
  const pagedRows = filteredRows.slice(start, end);
  el.tableBody.innerHTML = "";

  pagedRows.forEach((row) => {
    const tr = document.createElement("tr");

    if (String(row["الحالة المرضية ان وجدت"] || "").trim()) {
      tr.classList.add("medical-alert");
    }

    HEADERS.forEach((h) => {
      const td = document.createElement("td");
      td.textContent = row[h] || "";
      tr.appendChild(td);
    });

    const actions = document.createElement("td");
    actions.className = "actions-cell";

    const editBtn = document.createElement("button");
    editBtn.className = "btn btn-primary";
    editBtn.textContent = "تعديل";
    editBtn.type = "button";
    editBtn.addEventListener("click", () => openEditModalById(row["عدد"]));

    const deleteBtn = document.createElement("button");
    deleteBtn.className = "btn btn-danger";
    deleteBtn.textContent = "حذف";
    deleteBtn.type = "button";
    deleteBtn.addEventListener("click", () => deleteById(row["عدد"]));

    actions.appendChild(editBtn);
    actions.appendChild(deleteBtn);
    tr.appendChild(actions);

    const floor = String(row["رقم الطابق"] || "").trim();
    const room = String(row["رقم الغرفة"] || "").trim();
    const occupancy = roomCounts.get(`${floor}|${room}`) || 0;
    if (occupancy > roomCapacity) {
      tr.title = `تحذير: الغرفة ${room} في الطابق ${floor} فيها ${occupancy} أشخاص`;
    }

    el.tableBody.appendChild(tr);
  });

  el.resultCount.textContent = `${filteredRows.length} نتيجة`;
  el.pageInfo.textContent = `صفحة ${uiState.page} من ${totalPages}`;
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
    const floor = String(row["رقم الطابق"] || "").trim();
    const room = String(row["رقم الغرفة"] || "").trim();
    const gender = String(row["الجنس"] || "").trim();
    const age = parseInt(String(row["العمر"] || "").trim(), 10);

    if (floor) {
      floors.add(floor);
      floorCounts.set(floor, (floorCounts.get(floor) || 0) + 1);
    }

    const roomKey = `${floor}|${room}`;
    roomsMap.set(roomKey, (roomsMap.get(roomKey) || 0) + 1);

    if (String(row["الحالة المرضية ان وجدت"] || "").trim()) medical += 1;
    if (gender === "ذكر") male += 1;
    if (gender === "أنثى") female += 1;

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
    const f = String(row["رقم الطابق"] || "").trim() || "بدون طابق";
    const r = String(row["رقم الغرفة"] || "").trim() || "بدون رقم";
    if (!floorMap.has(f)) floorMap.set(f, new Set());
    floorMap.get(f).add(r);
  });

  if (floorMap.size === 0) {
    el.floorsContainer.textContent = "لا توجد بيانات غرف حالياً.";
    return;
  }

  const sortedFloors = Array.from(floorMap.keys()).sort();
  sortedFloors.forEach(floorName => {
    const floorBlock = document.createElement("div");
    floorBlock.className = "floor-block";

    const title = document.createElement("h3");
    title.className = "floor-title";
    title.textContent = `طابق: ${floorName}`;
    floorBlock.appendChild(title);

    const grid = document.createElement("div");
    grid.className = "rooms-grid";

    const sortedRooms = Array.from(floorMap.get(floorName)).sort();
    sortedRooms.forEach(roomName => {
      const occupants = dataRows.filter(r =>
        (String(r["رقم الطابق"] || "").trim() || "بدون طابق") === floorName &&
        (String(r["رقم الغرفة"] || "").trim() || "بدون رقم") === roomName
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
        <div class="room-number">غرفة ${roomName}</div>
        <div class="room-occupancy">${count} / ${roomCapacity} شخص</div>
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
  el.roomModalTitle.textContent = `تفاصيل طابق ${floor} | غرفة ${room}`;
  el.roomDescriptionInput.value = roomMetadata[currentActiveRoomKey] || "";

  el.roomOccupantsList.innerHTML = "";
  occupants.forEach(person => {
    const div = document.createElement("div");
    div.className = "occupant-item";
    div.innerHTML = `
      <div class="occupant-info">
        <span class="occupant-name">${person["الاسم الثلاثي"]}</span>
        <span class="occupant-meta">رقم: ${person["عدد"]} | هاتف: ${person["رقم الهاتف"] || "---"}</span>
      </div>
      <div class="occupant-actions">
        <button class="btn btn-secondary btn-sm" onclick="movePerson('${person["عدد"]}')">نقل</button>
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
  const person = dataRows.find(r => String(r["عدد"]) === String(personId));
  if (!person) return;

  const newFloor = prompt("أدخل رقم الطابق الجديد:", person["رقم الطابق"]);
  if (newFloor === null) return;
  const newRoom = prompt("أدخل رقم الغرفة الجديد:", person["رقم الغرفة"]);
  if (newRoom === null) return;

  person["رقم الطابق"] = newFloor.trim();
  person["رقم الغرفة"] = newRoom.trim();

  el.roomModal.classList.add("hidden");
  renderAll();
};

function openEditModalById(id) {
  editingIndex = dataRows.findIndex((r) => String(r["عدد"]) === String(id));
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

    if (header === "عدد") {
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
  const index = dataRows.findIndex((r) => String(r["عدد"]) === String(id));
  if (index < 0) return;

  const ok = window.confirm("هل تريد حذف هذا السجل؟");
  if (!ok) return;

  dataRows.splice(index, 1);
  renderAll();
}

function saveExcelFile() {
  const rowsForExport = dataRows.map((row) => {
    const out = {};
    HEADERS.forEach((h) => {
      out[h] = row[h] || "";
    });
    return out;
  });

  const ws = XLSX.utils.json_to_sheet(rowsForExport, { header: HEADERS });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Records");

  // Save Room Metadata
  const metaForExport = Object.entries(roomMetadata).map(([key, desc]) => {
    const [floor, room] = key.split("|");
    return {
      "رقم الطابق": floor,
      "رقم الغرفة": room,
      "الوصف": desc
    };
  });
  if (metaForExport.length > 0) {
    const wsMeta = XLSX.utils.json_to_sheet(metaForExport);
    XLSX.utils.book_append_sheet(wb, wsMeta, "RoomMetadata");
  }

  XLSX.writeFile(wb, fileName || "shelter_data.xlsx");
}

function renderAll() {
  renderDashboard();
  renderRoomVisuals();
  renderTable();
  updateCharts();
}

function onFiltersChanged() {
  uiState.page = 1;
  renderTable();
}

function clearAllFilters() {
  el.searchInput.value = "";
  el.filterGender.value = "";
  el.filterFloor.value = "";
  el.filterRoom.value = "";
  el.filterMedical.value = "";
  el.filterAgeGroup.value = "";
  onFiltersChanged();
}

function changePage(delta) {
  const filteredRows = getFilteredRows();
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / uiState.pageSize));
  uiState.page = Math.max(1, Math.min(totalPages, uiState.page + delta));
  renderTable();
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

function updateCharts() {
  if (typeof Chart === 'undefined') return;

  const male = dataRows.filter(r => r["الجنس"] === "ذكر").length;
  const female = dataRows.filter(r => r["الجنس"] === "أنثى").length;

  const children = dataRows.filter(r => {
    const age = parseInt(r["العمر"], 10);
    return Number.isFinite(age) && age < 18;
  }).length;
  const adults = dataRows.filter(r => {
    const age = parseInt(r["العمر"], 10);
    return Number.isFinite(age) && age >= 18;
  }).length;

  const ctxGender = document.getElementById('genderChart')?.getContext('2d');
  if (ctxGender) {
    if (genderChart) genderChart.destroy();
    genderChart = new Chart(ctxGender, {
      type: 'doughnut',
      data: {
        labels: ['ذكور', 'إناث'],
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
      type: 'pie',
      data: {
        labels: ['أطفال', 'بالغون'],
        datasets: [{
          data: [children, adults],
          backgroundColor: ['#10b981', '#f59e0b']
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false
      }
    });
  }
}
