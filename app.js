const HOURS_PER_WORKDAY = 7.57;
const SHIFT_HOURS = 12;
const EMPLOYEES_STORAGE_KEY = 'grafik_employees_v1';
const REQUESTS_STORAGE_KEY = 'grafik_requests_v1';
const SCHEDULE_STORAGE_KEY = 'grafik_saved_schedule_v1';

const state = {
  month: new Date().getMonth() + 1,
  year: new Date().getFullYear(),
  employees: [],
  nextEmployeeId: 1,
  requestsByMonth: {},
  lastResult: null
};

function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function countBusinessDays(year, month) {
  const days = daysInMonth(year, month);
  let count = 0;
  for (let day = 1; day <= days; day++) {
    const weekDay = new Date(year, month - 1, day).getDay();
    if (weekDay !== 0 && weekDay !== 6) count += 1;
  }
  return count;
}

function isoWeekKey(year, month, day) {
  const date = new Date(Date.UTC(year, month - 1, day));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const y = date.getUTCFullYear();
  const yearStart = new Date(Date.UTC(y, 0, 1));
  const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return `${y}-W${weekNo}`;
}

function weekdayLabel(year, month, day) {
  const labels = ['ND', 'PN', 'WT', 'ŚR', 'CZ', 'PT', 'SO'];
  return labels[new Date(year, month - 1, day).getDay()];
}

function shuffle(array) {
  const out = array.slice();
  for (let i = out.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [out[i], out[j]] = [out[j], out[i]];
  }
  return out;
}

function ensureTopScrollbar(wrapper) {
  let top = wrapper.previousElementSibling;
  if (!top || !top.classList.contains('top-scroll')) {
    top = document.createElement('div');
    top.className = 'top-scroll';
    top.innerHTML = '<div class="top-scroll-inner"></div>';
    wrapper.parentNode.insertBefore(top, wrapper);
  }

  const inner = top.firstElementChild;
  const table = wrapper.querySelector('table');
  if (!inner || !table) return;
  inner.style.width = `${table.scrollWidth}px`;

  if (wrapper.dataset.topScrollBound === '1') return;
  wrapper.dataset.topScrollBound = '1';

  let syncing = false;
  top.addEventListener('scroll', () => {
    if (syncing) return;
    syncing = true;
    wrapper.scrollLeft = top.scrollLeft;
    syncing = false;
  });

  wrapper.addEventListener('scroll', () => {
    if (syncing) return;
    syncing = true;
    top.scrollLeft = wrapper.scrollLeft;
    syncing = false;
  });
}

function syncEmployeeScheduleStickyRows() {
  const table = document.getElementById('employeeScheduleTable');
  if (!table || table.offsetParent === null) return;

  const rows = Array.from(table.querySelectorAll('thead tr'));
  if (!rows.length) return;

  let top = 0;
  rows.forEach((row) => {
    const rowHeight = row.getBoundingClientRect().height;
    row.querySelectorAll('th').forEach((th) => {
      th.style.top = `${top}px`;
    });
    top += rowHeight;
  });
}

function refreshTopScrollbars() {
  document.querySelectorAll('.table-wrap').forEach(ensureTopScrollbar);
  syncEmployeeScheduleStickyRows();
}

function setActiveTab(tab) {
  const employeesTab = document.getElementById('employeesTab');
  const planTab = document.getElementById('planTab');
  const resultTab = document.getElementById('resultTab');

  const employeesBtn = document.getElementById('tabEmployeesBtn');
  const planBtn = document.getElementById('tabPlanBtn');
  const resultBtn = document.getElementById('tabResultBtn');

  employeesTab.classList.toggle('hidden', tab !== 'employees');
  planTab.classList.toggle('hidden', tab !== 'plan');
  resultTab.classList.toggle('hidden', tab !== 'result');

  employeesBtn.classList.toggle('active', tab === 'employees');
  planBtn.classList.toggle('active', tab === 'plan');
  resultBtn.classList.toggle('active', tab === 'result');

  requestAnimationFrame(refreshTopScrollbars);
}

function populateMonthSelect() {
  const monthInput = document.getElementById('monthInput');
  monthInput.innerHTML = '';
  const names = [
    'Styczeń', 'Luty', 'Marzec', 'Kwiecień', 'Maj', 'Czerwiec',
    'Lipiec', 'Sierpień', 'Wrzesień', 'Październik', 'Listopad', 'Grudzień'
  ];
  names.forEach((name, idx) => {
    const opt = document.createElement('option');
    opt.value = String(idx + 1);
    opt.textContent = name;
    if (idx + 1 === state.month) opt.selected = true;
    monthInput.appendChild(opt);
  });
}

function readCurrentInputs() {
  state.month = Number(document.getElementById('monthInput').value);
  state.year = Number(document.getElementById('yearInput').value);
}

function employeeContractLabel(contract) {
  if (contract === 'half') return 'pół etatu';
  if (contract === 'outside') return 'poza grafikiem';
  return 'pełny etat';
}

function createDefaultEmployees(count) {
  state.employees = Array.from({ length: count }, (_, i) => ({
    id: i + 1,
    name: '',
    contract: 'full'
  }));
  state.nextEmployeeId = count + 1;
}

function saveEmployeesToStorage() {
  syncEmployeesFromTable();
  const payload = {
    nextEmployeeId: state.nextEmployeeId,
    employees: state.employees
  };
  localStorage.setItem(EMPLOYEES_STORAGE_KEY, JSON.stringify(payload));
}

function loadEmployeesFromStorage() {
  const raw = localStorage.getItem(EMPLOYEES_STORAGE_KEY);
  if (!raw) return false;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.employees)) return false;
    state.employees = parsed.employees.map((e) => ({
      id: Number(e.id),
      name: typeof e.name === 'string' ? e.name : '',
      contract: e.contract === 'half' || e.contract === 'outside' ? e.contract : 'full'
    })).filter((e) => Number.isFinite(e.id));
    if (!state.employees.length) return false;
    const maxId = state.employees.reduce((m, e) => Math.max(m, e.id), 0);
    state.nextEmployeeId = Math.max(Number(parsed.nextEmployeeId) || 1, maxId + 1);
    return true;
  } catch {
    return false;
  }
}

function setSaveStatus(text) {
  const el = document.getElementById('employeesSaveStatus');
  if (!el) return;
  el.textContent = text;
}

function setDemandStatus(text) {
  const el = document.getElementById('demandSaveStatus');
  if (!el) return;
  el.textContent = text;
}

function monthKey(year, month) {
  return `${year}-${String(month).padStart(2, '0')}`;
}

function loadRequestsFromStorage() {
  const raw = localStorage.getItem(REQUESTS_STORAGE_KEY);
  if (!raw) {
    state.requestsByMonth = {};
    return;
  }
  try {
    const parsed = JSON.parse(raw);
    state.requestsByMonth = parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    state.requestsByMonth = {};
  }
}

function saveRequestsToStorage() {
  localStorage.setItem(REQUESTS_STORAGE_KEY, JSON.stringify(state.requestsByMonth));
}

function loadScheduleFromStorage() {
  const raw = localStorage.getItem(SCHEDULE_STORAGE_KEY);
  if (!raw) return false;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return false;
    if (!Array.isArray(parsed.assignments) || !Array.isArray(parsed.employeeStats) || !Array.isArray(parsed.daySummary)) return false;
    state.lastResult = parsed;
    return true;
  } catch {
    return false;
  }
}

function saveScheduleToStorage() {
  syncLastResultFromResultTable();
  if (!state.lastResult) return false;
  localStorage.setItem(SCHEDULE_STORAGE_KEY, JSON.stringify(state.lastResult));
  return true;
}

function collectCurrentDemandMap() {
  const out = {};
  document.querySelectorAll('#requestTable .cell-select').forEach((sel) => {
    const code = sel.value || '';
    if (!code) return;
    out[`${sel.dataset.eid}-${sel.dataset.d}`] = code;
  });
  return out;
}

function saveCurrentDemandForMonth() {
  readCurrentInputs();
  const key = monthKey(state.year, state.month);
  state.requestsByMonth[key] = collectCurrentDemandMap();
  saveRequestsToStorage();
}

function clearCurrentDemandForMonth() {
  readCurrentInputs();
  const key = monthKey(state.year, state.month);
  delete state.requestsByMonth[key];
  saveRequestsToStorage();
}

function renderEmployeesTable() {
  const tbody = document.querySelector('#employeesTable tbody');
  tbody.innerHTML = '';

  state.employees.forEach((emp, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td><input class="employee-name" data-id="${emp.id}" type="text" value="${emp.name}" placeholder="Nazwa pracownika" /></td>
      <td>
        <select class="employee-contract" data-id="${emp.id}">
          <option value="full" ${emp.contract === 'full' ? 'selected' : ''}>pełny etat</option>
          <option value="half" ${emp.contract === 'half' ? 'selected' : ''}>pół etatu</option>
          <option value="outside" ${emp.contract === 'outside' ? 'selected' : ''}>poza grafikiem</option>
        </select>
      </td>
      <td><button class="danger" type="button" data-remove-id="${emp.id}">Usuń</button></td>
    `;
    tbody.appendChild(tr);
  });
  refreshTopScrollbars();
}

function syncEmployeesFromTable() {
  const rows = Array.from(document.querySelectorAll('#employeesTable tbody tr'));
  if (!rows.length) return;

  const byId = new Map(state.employees.map((e) => [e.id, e]));
  state.employees = rows.map((row) => {
    const nameInput = row.querySelector('input.employee-name');
    const contractSelect = row.querySelector('select.employee-contract');
    const id = Number((nameInput || contractSelect).dataset.id);
    const current = byId.get(id) || { id, name: '', contract: 'full' };
    return {
      id,
      name: nameInput ? nameInput.value.trim() : current.name,
      contract: contractSelect ? contractSelect.value : current.contract
    };
  });
}

function snapshotGridValues() {
  const map = new Map();
  document.querySelectorAll('#requestTable .cell-select').forEach((sel) => {
    map.set(`${sel.dataset.eid}-${sel.dataset.d}`, sel.value);
  });
  return map;
}

function applyRequestCellClass(selectEl) {
  const td = selectEl.closest('td');
  if (!td) return;
  td.classList.remove('req-blue', 'req-u');
  if (selectEl.value === 'U') {
    td.classList.add('req-u');
    return;
  }
  if (['D', 'N', 'W', 'bN', 'bD'].includes(selectEl.value)) {
    td.classList.add('req-blue');
  }
}

function applyResultCellClass(selectEl) {
  const td = selectEl.closest('td');
  if (!td) return;
  td.classList.remove('code-blue', 'code-u', 'req-demand');
  if (selectEl.dataset.req === '1') {
    td.classList.add('req-demand');
    return;
  }
  if (selectEl.value === 'U' || selectEl.value === 'W') {
    td.classList.add('code-u');
    return;
  }
  if (['D', 'N', 'K', 'bN', 'bD'].includes(selectEl.value)) {
    td.classList.add('code-blue');
  }
}

function updateResultTotals(dayCount) {
  const sumD = Array(dayCount).fill(0);
  const sumN = Array(dayCount).fill(0);
  const sumK = Array(dayCount).fill(0);
  const sumU = Array(dayCount).fill(0);

  document.querySelectorAll('#employeeScheduleTable .result-cell-select').forEach((sel) => {
    const dayIdx = Number(sel.dataset.d) - 1;
    if (dayIdx < 0 || dayIdx >= dayCount) return;
    if (sel.value === 'D') sumD[dayIdx] += 1;
    if (sel.value === 'N') sumN[dayIdx] += 1;
    if (sel.value === 'K') sumK[dayIdx] += 1;
    if (sel.value === 'U') sumU[dayIdx] += 1;
  });

  for (let d = 1; d <= dayCount; d++) {
      const dCell = document.querySelector(`#sumD-${d}`);
      const nCell = document.querySelector(`#sumN-${d}`);
      const kCell = document.querySelector(`#sumK-${d}`);
      const uCell = document.querySelector(`#sumU-${d}`);
      if (dCell) dCell.textContent = String(sumD[d - 1]);
      if (nCell) nCell.textContent = String(sumN[d - 1]);
      if (kCell) kCell.textContent = String(sumK[d - 1]);
      if (uCell) uCell.textContent = String(sumU[d - 1]);
  }
}

function updateResultHours(dayCount) {
  const rowCount = Number(document.getElementById('employeeScheduleTable').dataset.rowCount || 0);
  for (let p = 1; p <= rowCount; p++) {
    let hours = 0;
    const shortHours = Number(document.getElementById('employeeScheduleTable').dataset[`short${p}`] || 0);
    for (let d = 1; d <= dayCount; d++) {
      const sel = document.querySelector(`#employeeScheduleTable .result-cell-select[data-p="${p}"][data-d="${d}"]`);
      if (!sel) continue;
      if (sel.value === 'D' || sel.value === 'N' || sel.value === 'U') hours += SHIFT_HOURS;
      if (sel.value === 'K') hours += shortHours;
    }
    const cell = document.getElementById(`hours-${p}`);
    if (cell) cell.textContent = hours.toFixed(2);
  }
}

function updateResultSummaryCards() {
  const table = document.getElementById('employeeScheduleTable');
  const rowCount = Number(table.dataset.rowCount || 0);
  const hasShortByPerson = Array(rowCount).fill(false);

  document.querySelectorAll('#employeeScheduleTable .result-cell-select').forEach((sel) => {
    if (sel.value !== 'K') return;
    const p = Number(sel.dataset.p);
    if (p < 1 || p > rowCount) return;
    hasShortByPerson[p - 1] = true;
  });

  let missingHoursCount = 0;
  let shortMissingCount = 0;

  for (let p = 1; p <= rowCount; p++) {
    const targetHours = Number(table.dataset[`target${p}`] || 0);
    const requiresShortShift = table.dataset[`requireShort${p}`] === '1';
    const hoursCell = document.getElementById(`hours-${p}`);
    const currentHours = Number(hoursCell ? hoursCell.textContent : 0);

    if (currentHours + 0.01 < targetHours) missingHoursCount += 1;
    if (requiresShortShift && !hasShortByPerson[p - 1]) shortMissingCount += 1;
  }

  const missingEl = document.getElementById('summaryMissingHours');
  if (missingEl) {
    missingEl.textContent = `${missingHoursCount} osób`;
    missingEl.classList.toggle('bad', missingHoursCount > 0);
  }

  const shortMissingEl = document.getElementById('summaryShortMissing');
  if (shortMissingEl) {
    shortMissingEl.textContent = `${shortMissingCount} osób`;
    shortMissingEl.classList.toggle('bad', shortMissingCount > 0);
  }
}

function updateResultAggregates(dayCount) {
  updateResultTotals(dayCount);
  updateResultHours(dayCount);
  updateResultSummaryCards();
}

function rebuildRequestGrid() {
  readCurrentInputs();
  syncEmployeesFromTable();

  const previous = snapshotGridValues();
  const dayCount = daysInMonth(state.year, state.month);
  const key = monthKey(state.year, state.month);
  const table = document.getElementById('requestTable');
  const reusePrevious = table.dataset.monthKey === key;
  const stored = state.requestsByMonth[key] || {};
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  const header = document.createElement('tr');
  let headerHtmlTop = '<th rowspan="2">Pracownik / etat</th>';
  for (let d = 1; d <= dayCount; d++) headerHtmlTop += `<th>${weekdayLabel(state.year, state.month, d)}</th>`;
  header.innerHTML = headerHtmlTop;
  thead.appendChild(header);

  const header2 = document.createElement('tr');
  let headerHtmlBottom = '';
  for (let d = 1; d <= dayCount; d++) headerHtmlBottom += `<th>${d}</th>`;
  header2.innerHTML = headerHtmlBottom;
  thead.appendChild(header2);

  state.employees.forEach((emp, rowIndex) => {
    const tr = document.createElement('tr');
    const label = `${rowIndex + 1}. ${emp.name || 'brak nazwy'} (${employeeContractLabel(emp.contract)})`;
    let row = `<td class="person-fixed">${label}</td>`;

    for (let d = 1; d <= dayCount; d++) {
      const key = `${emp.id}-${d}`;
      const val = (reusePrevious ? previous.get(key) : undefined) ?? stored[key] ?? '';
      row += `
        <td class="request-cell">
          <select class="cell-select" data-r="${rowIndex + 1}" data-eid="${emp.id}" data-d="${d}">
            <option value="" ${val === '' ? 'selected' : ''}>-</option>
            <option value="D" ${val === 'D' ? 'selected' : ''}>D</option>
            <option value="N" ${val === 'N' ? 'selected' : ''}>N</option>
            <option value="U" ${val === 'U' ? 'selected' : ''}>U</option>
            <option value="W" ${val === 'W' ? 'selected' : ''}>W</option>
            <option value="bN" ${val === 'bN' ? 'selected' : ''}>bN</option>
            <option value="bD" ${val === 'bD' ? 'selected' : ''}>bD</option>
          </select>
        </td>
      `;
    }

    tr.innerHTML = row;
    tbody.appendChild(tr);
  });

  document.querySelectorAll('#requestTable .cell-select').forEach(applyRequestCellClass);
  table.dataset.monthKey = key;

  const businessDays = countBusinessDays(state.year, state.month);
  const monthHours = businessDays * HOURS_PER_WORKDAY;
  const fullShifts = Math.floor(monthHours / SHIFT_HOURS);
  const shortHours = Number((monthHours - fullShifts * SHIFT_HOURS).toFixed(2));

  document.getElementById('metaBox').textContent =
    `Liczba pracowników: ${state.employees.length}. Dni miesiąca: ${dayCount}. Dni robocze: ${businessDays}. Pełny etat: ${monthHours.toFixed(2)}h = ${fullShifts}x12h + ${shortHours.toFixed(2)}h.`;
  refreshTopScrollbars();
}

function focusRequestCell(row, day) {
  const next = document.querySelector(`#requestTable .cell-select[data-r="${row}"][data-d="${day}"]`);
  if (next) next.focus();
}

function handleRequestGridArrows(event) {
  const target = event.target;
  if (!(target instanceof HTMLSelectElement)) return;
  if (!target.classList.contains('cell-select')) return;

  const row = Number(target.dataset.r);
  const day = Number(target.dataset.d);
  if (!Number.isFinite(row) || !Number.isFinite(day)) return;

  let nextRow = row;
  let nextDay = day;

  if (event.key === 'ArrowLeft') nextDay -= 1;
  if (event.key === 'ArrowRight') nextDay += 1;
  if (event.key === 'ArrowUp') nextRow -= 1;
  if (event.key === 'ArrowDown') nextRow += 1;

  if (nextRow === row && nextDay === day) return;
  event.preventDefault();
  focusRequestCell(nextRow, nextDay);
}

function focusResultCell(person, day) {
  const next = document.querySelector(`#employeeScheduleTable .result-cell-select[data-p="${person}"][data-d="${day}"]`);
  if (next) next.focus();
}

function handleResultGridArrows(event) {
  const target = event.target;
  if (!(target instanceof HTMLSelectElement)) return;
  if (!target.classList.contains('result-cell-select')) return;

  const person = Number(target.dataset.p);
  const day = Number(target.dataset.d);
  if (!Number.isFinite(person) || !Number.isFinite(day)) return;

  let nextPerson = person;
  let nextDay = day;

  if (event.key === 'ArrowLeft') nextDay -= 1;
  if (event.key === 'ArrowRight') nextDay += 1;
  if (event.key === 'ArrowUp') nextPerson -= 1;
  if (event.key === 'ArrowDown') nextPerson += 1;

  if (nextPerson === person && nextDay === day) return;
  event.preventDefault();
  focusResultCell(nextPerson, nextDay);
}

function collectRequests() {
  const dayCount = daysInMonth(state.year, state.month);
  const requests = Array.from({ length: state.employees.length }, () => Array(dayCount).fill(''));
  const indexById = new Map(state.employees.map((e, idx) => [e.id, idx]));

  document.querySelectorAll('#requestTable .cell-select').forEach((sel) => {
    const idx = indexById.get(Number(sel.dataset.eid));
    if (idx === undefined) return;
    const d = Number(sel.dataset.d) - 1;
    requests[idx][d] = sel.value;
  });

  return requests;
}

function isDayShift(type) {
  return type === 'D' || type === 'K';
}

function isWorkingAssignment(code) {
  return code === 'D' || code === 'N' || code === 'K';
}

function blocksShiftByRequest(requestCode, type) {
  if (requestCode === 'U' || requestCode === 'W') return true;
  if (requestCode === 'bN' && type === 'N') return true;
  if (requestCode === 'bD' && (type === 'D' || type === 'K')) return true;
  return false;
}

function creditedHoursNow(personIdx, stats, personTargetShort) {
  return stats.fullShiftCount[personIdx] * SHIFT_HOURS
    + stats.vacationCreditHours[personIdx]
    + stats.shortCount[personIdx] * personTargetShort[personIdx];
}

function canAddHours(personIdx, addHours, stats, personTargetHours, personTargetShort) {
  return creditedHoursNow(personIdx, stats, personTargetShort) + addHours <= personTargetHours[personIdx] + 0.01;
}

function canPlaceShift(assignments, requests, personIdx, dayIdx, type, stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads, options = {}) {
  if (dayIdx < 0 || dayIdx >= assignments[personIdx].length) return false;
  if (assignments[personIdx][dayIdx] !== '') return false;
  if (blocksShiftByRequest(requests[personIdx][dayIdx], type)) return false;

  if (dayLoads) {
    if (type === 'K') {
      if (dayLoads.K[dayIdx] >= 1) return false;
      if (dayLoads.D[dayIdx] + dayLoads.K[dayIdx] >= 7) return false;
    }
    if (type === 'N') {
      if (dayLoads.N[dayIdx] >= 7) return false;
    }
    if (type === 'D') {
      if (dayLoads.D[dayIdx] + dayLoads.K[dayIdx] >= 7) return false;
    }
  }

  let leftWorking = 0;
  for (let d = dayIdx - 1; d >= 0; d--) {
    if (!isWorkingAssignment(assignments[personIdx][d])) break;
    leftWorking += 1;
  }
  let rightWorking = 0;
  for (let d = dayIdx + 1; d < assignments[personIdx].length; d++) {
    if (!isWorkingAssignment(assignments[personIdx][d])) break;
    rightWorking += 1;
  }
  if (leftWorking + 1 + rightWorking > 3) return false;

  const prev1 = assignments[personIdx][dayIdx - 1];
  const prev2 = assignments[personIdx][dayIdx - 2];
  if (prev1 === 'N') return false;
  if (isDayShift(type) && prev2 === 'N') return false;

  if (type === 'N') {
    const next1 = assignments[personIdx][dayIdx + 1];
    const next2 = assignments[personIdx][dayIdx + 2];
    if (next1 && next1 !== '') return false;
    if (next2 === 'D' || next2 === 'K') return false;
  }

  if (type === 'D' || type === 'N') {
    const weekKey = weekKeyByDay[dayIdx];
    const weekHours = stats.weeklyDNHours[personIdx].get(weekKey) || 0;
    if (weekHours + SHIFT_HOURS > 48) return false;
    if (!canAddHours(personIdx, SHIFT_HOURS, stats, personTargetHours, personTargetShort)) return false;
  }

  if (type === 'K') {
    const shortHours = personTargetShort[personIdx];
    if (shortHours <= 0.01 && !options.forceShortShift) return false;
    if (shortHours > 0.01 && !canAddHours(personIdx, shortHours, stats, personTargetHours, personTargetShort)) return false;
  }

  return true;
}

function assignFull(assignments, dayLoads, personIdx, dayIdx, type, stats, weekKeyByDay) {
  assignments[personIdx][dayIdx] = type;
  dayLoads[type][dayIdx] += 1;
  stats.fullShiftCount[personIdx] += 1;
  const weekKey = weekKeyByDay[dayIdx];
  const current = stats.weeklyDNHours[personIdx].get(weekKey) || 0;
  stats.weeklyDNHours[personIdx].set(weekKey, current + SHIFT_HOURS);
}

function removeFull(assignments, dayLoads, personIdx, dayIdx, stats, weekKeyByDay) {
  const type = assignments[personIdx][dayIdx];
  if (type !== 'D' && type !== 'N') return;
  dayLoads[type][dayIdx] -= 1;
  stats.fullShiftCount[personIdx] -= 1;
  const weekKey = weekKeyByDay[dayIdx];
  const current = stats.weeklyDNHours[personIdx].get(weekKey) || 0;
  stats.weeklyDNHours[personIdx].set(weekKey, Math.max(0, current - SHIFT_HOURS));
  assignments[personIdx][dayIdx] = '';
}

function pickSlotForFull(assignments, requests, dayLoads, personIdx, dayCount, stats, weekKeyByDay, personTargetHours, personTargetShort) {
  const open = [];
  for (let d = 0; d < dayCount; d++) {
    if (!canPlaceShift(assignments, requests, personIdx, d, 'D', stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads) && !canPlaceShift(assignments, requests, personIdx, d, 'N', stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads)) continue;

    const dLoad = dayLoads.D[d];
    const nLoad = dayLoads.N[d];
    const bestType = dLoad <= nLoad ? 'D' : 'N';
    if (!canPlaceShift(assignments, requests, personIdx, d, bestType, stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads)) continue;
    open.push({ day: d, type: bestType, load: Math.min(dLoad, nLoad) });
  }

  if (!open.length) return null;
  open.sort((a, b) => a.load - b.load);
  const top = open.slice(0, Math.min(8, open.length));
  return top[Math.floor(Math.random() * top.length)];
}

function tryReassignPerson(assignments, requests, dayLoads, personIdx, dayCount, maxPerShift, stats, weekKeyByDay, personTargetHours, personTargetShort) {
  const candidates = [];
  for (let d = 0; d < dayCount; d++) {
    if (assignments[personIdx][d] !== '') continue;

    if (dayLoads.D[d] < maxPerShift && canPlaceShift(assignments, requests, personIdx, d, 'D', stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads)) {
      candidates.push({ day: d, type: 'D', load: dayLoads.D[d] + dayLoads.N[d] });
    }
    if (dayLoads.N[d] < maxPerShift && canPlaceShift(assignments, requests, personIdx, d, 'N', stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads)) {
      candidates.push({ day: d, type: 'N', load: dayLoads.D[d] + dayLoads.N[d] });
    }
  }
  if (!candidates.length) return null;
  candidates.sort((a, b) => a.load - b.load);
  return candidates[Math.floor(Math.random() * Math.min(10, candidates.length))];
}

function enforceMonthlyHourCaps(assignments, requests, dayLoads, stats, weekKeyByDay, personTargetHours, personTargetShort) {
  const peopleCount = assignments.length;
  const dayCount = assignments[0]?.length || 0;

  for (let p = 0; p < peopleCount; p++) {
    while (creditedHoursNow(p, stats, personTargetShort) > personTargetHours[p] + 0.01) {
      let changed = false;

      for (let d = dayCount - 1; d >= 0; d--) {
        const code = assignments[p][d];
        if (code !== 'D' && code !== 'N') continue;
        const req = requests[p][d];
        if (req === code) continue;
        removeFull(assignments, dayLoads, p, d, stats, weekKeyByDay);
        changed = true;
        break;
      }

      if (changed) continue;

      for (let d = dayCount - 1; d >= 0; d--) {
        const code = assignments[p][d];
        if (code !== 'D' && code !== 'N') continue;
        removeFull(assignments, dayLoads, p, d, stats, weekKeyByDay);
        changed = true;
        break;
      }

      if (changed) continue;

      if (stats.shortCount[p] > 0) {
        for (let d = dayCount - 1; d >= 0; d--) {
          if (assignments[p][d] === 'K') {
            assignments[p][d] = '';
            dayLoads.K[d] = Math.max(0, dayLoads.K[d] - 1);
            stats.shortCount[p] = 0;
            changed = true;
            break;
          }
        }
      }

      if (!changed) break;
    }
  }
}

function backfillVacationFromDaysOff(assignments, stats, personTargetHours, personTargetShort) {
  const peopleCount = assignments.length;
  const dayCount = assignments[0]?.length || 0;

  for (let p = 0; p < peopleCount; p++) {
    while (creditedHoursNow(p, stats, personTargetShort) + 0.01 < personTargetHours[p]) {
      if (!canAddHours(p, SHIFT_HOURS, stats, personTargetHours, personTargetShort)) break;

      let changed = false;
      for (let d = 0; d < dayCount; d++) {
        if (assignments[p][d] !== 'W') continue;
        assignments[p][d] = 'U';
        stats.vacationCreditHours[p] += SHIFT_HOURS;
        changed = true;
        break;
      }

      if (!changed) break;
    }
  }
}

function buildPhaseOrders(requests, activePeople) {
  const phaseNoUAndW = [];
  const phaseUWithoutW = [];
  const phaseWithW = [];

  for (const p of activePeople) {
    const row = requests[p] || [];
    const hasVacation = row.includes('U');
    const hasDayOff = row.includes('W');

    if (!hasVacation && !hasDayOff) {
      phaseNoUAndW.push(p);
      continue;
    }
    if (hasVacation && !hasDayOff) {
      phaseUWithoutW.push(p);
      continue;
    }
    let dayOffCount = 0;
    for (const code of row) {
      if (code === 'W') dayOffCount += 1;
    }
    phaseWithW.push({ person: p, dayOffCount });
  }

  return [
    shuffle(phaseNoUAndW),
    shuffle(phaseUWithoutW),
    phaseWithW
      .sort((a, b) => {
        if (a.dayOffCount !== b.dayOffCount) return a.dayOffCount - b.dayOffCount;
        return a.person - b.person;
      })
      .map((entry) => entry.person)
  ];
}

function assignShortShiftsByPhase(phaseOrders, assignments, requests, dayLoads, dayCount, stats, weekKeyByDay, personTargetHours, personTargetShort, forceShortForEveryone) {
  for (const phase of phaseOrders) {
    for (const p of phase) {
      const shouldTryShortShift = forceShortForEveryone
        ? state.employees[p].contract !== 'outside'
        : personTargetShort[p] > 0.01;
      if (!shouldTryShortShift) continue;
      if (stats.shortCount[p] > 0) continue;

      for (let d = 0; d < dayCount; d++) {
        if (!canPlaceShift(assignments, requests, p, d, 'K', stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads, { forceShortShift: forceShortForEveryone })) continue;
        assignments[p][d] = 'K';
        dayLoads.K[d] += 1;
        stats.shortCount[p] = 1;
        break;
      }
    }
  }
}

function enforceShiftMinimum(assignments, requests, dayLoads, stats, weekKeyByDay, personTargetHours, personTargetShort, type, minPerDay, oppositeType, oppositeMinPerDay) {
  const peopleCount = assignments.length;
  const dayCount = assignments[0]?.length || 0;

  for (let d = 0; d < dayCount; d++) {
    while (dayLoads[type][d] < minPerDay) {
      const emptyCandidates = [];
      for (let p = 0; p < peopleCount; p++) {
        if (!canPlaceShift(assignments, requests, p, d, type, stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads)) continue;
        emptyCandidates.push({ p, load: stats.fullShiftCount[p] });
      }

      if (emptyCandidates.length) {
        emptyCandidates.sort((a, b) => a.load - b.load);
        assignFull(assignments, dayLoads, emptyCandidates[0].p, d, type, stats, weekKeyByDay);
        continue;
      }

      const switchCandidates = [];
      for (let p = 0; p < peopleCount; p++) {
        if (assignments[p][d] !== oppositeType) continue;
        if (dayLoads[oppositeType][d] - 1 < oppositeMinPerDay) continue;
        const req = requests[p][d];
        removeFull(assignments, dayLoads, p, d, stats, weekKeyByDay);
        const canSwitch = canPlaceShift(assignments, requests, p, d, type, stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads);
        assignFull(assignments, dayLoads, p, d, oppositeType, stats, weekKeyByDay);
        if (!canSwitch) continue;
        switchCandidates.push({ p, preferred: req !== oppositeType ? 1 : 0, load: stats.fullShiftCount[p] });
      }

      if (!switchCandidates.length) break;

      switchCandidates.sort((a, b) => {
        if (a.preferred !== b.preferred) return b.preferred - a.preferred;
        return a.load - b.load;
      });
      const chosen = switchCandidates[0].p;
      removeFull(assignments, dayLoads, chosen, d, stats, weekKeyByDay);
      assignFull(assignments, dayLoads, chosen, d, type, stats, weekKeyByDay);
    }
  }
}

function generateSchedule(requests, year, month) {
  const peopleCount = state.employees.length;
  const dayCount = daysInMonth(year, month);
  const businessDays = countBusinessDays(year, month);
  const monthHours = businessDays * HOURS_PER_WORKDAY;
  const monthRequiresShortShift = Math.abs(monthHours - (Math.floor(monthHours / SHIFT_HOURS) * SHIFT_HOURS)) > 0.01;

  const personTargetHours = state.employees.map((emp) => {
    if (emp.contract === 'half') return Number((monthHours * 0.5).toFixed(2));
    if (emp.contract === 'outside') return 0;
    return Number(monthHours.toFixed(2));
  });

  const vacationRequestedCount = requests.map((row, idx) => {
    if (state.employees[idx].contract === 'outside') return 0;
    return row.filter((code) => code === 'U').length;
  });
  const vacationCreditHours = personTargetHours.map((h, i) => Math.min(h, vacationRequestedCount[i] * SHIFT_HOURS));
  const personShiftHours = personTargetHours.map((h, i) => Math.max(0, Number((h - vacationCreditHours[i]).toFixed(2))));
  const personTargetFull = personShiftHours.map((h) => Math.floor(h / SHIFT_HOURS));
  const personTargetShort = personShiftHours.map((h, i) => Number((h - personTargetFull[i] * SHIFT_HOURS).toFixed(2)));
  const weekKeyByDay = Array.from({ length: dayCount }, (_, i) => isoWeekKey(year, month, i + 1));

  const assignments = Array.from({ length: peopleCount }, () => Array(dayCount).fill(''));
  const dayLoads = { D: Array(dayCount).fill(0), N: Array(dayCount).fill(0), K: Array(dayCount).fill(0) };
  const stats = {
    fullShiftCount: Array(peopleCount).fill(0),
    vacationCreditHours: vacationCreditHours.slice(),
    shortCount: Array(peopleCount).fill(0),
    weeklyDNHours: Array.from({ length: peopleCount }, () => new Map())
  };

  const activePeople = Array.from({ length: peopleCount }, (_, i) => i)
    .filter((i) => state.employees[i].contract !== 'outside');
  const phaseOrders = buildPhaseOrders(requests, activePeople);
  const requiresShortShiftByPerson = state.employees.map((emp, idx) => {
    if (emp.contract === 'outside') return false;
    if (monthRequiresShortShift) return true;
    return personTargetShort[idx] > 0.01;
  });

  if (monthRequiresShortShift) {
    assignShortShiftsByPhase(phaseOrders, assignments, requests, dayLoads, dayCount, stats, weekKeyByDay, personTargetHours, personTargetShort, true);
  }

  for (const phase of phaseOrders) {
    for (const p of phase) {
      for (let d = 0; d < dayCount; d++) {
        const req = requests[p][d];
        if ((req === 'D' || req === 'N') && stats.fullShiftCount[p] < personTargetFull[p] && canPlaceShift(assignments, requests, p, d, req, stats, weekKeyByDay, personTargetHours, personTargetShort, dayLoads)) {
          assignFull(assignments, dayLoads, p, d, req, stats, weekKeyByDay);
        }
      }
    }
  }

  for (const phase of phaseOrders) {
    for (const p of phase) {
      while (stats.fullShiftCount[p] < personTargetFull[p]) {
        const slot = pickSlotForFull(assignments, requests, dayLoads, p, dayCount, stats, weekKeyByDay, personTargetHours, personTargetShort);
        if (!slot) break;
        assignFull(assignments, dayLoads, p, slot.day, slot.type, stats, weekKeyByDay);
      }
    }
  }

  for (let p = 0; p < peopleCount; p++) {
    for (let d = 0; d < dayCount; d++) {
      if (requests[p][d] !== 'U') continue;
      if (assignments[p][d] === 'D' || assignments[p][d] === 'N') {
        removeFull(assignments, dayLoads, p, d, stats, weekKeyByDay);
        assignments[p][d] = 'U';
      } else if (assignments[p][d] === '') {
        assignments[p][d] = 'U';
      }
    }
  }

  for (let p = 0; p < peopleCount; p++) {
    for (let d = 0; d < dayCount; d++) {
      if (requests[p][d] !== 'W') continue;
      if (assignments[p][d] === '') {
        assignments[p][d] = 'W';
      }
    }
  }

  const maxPerShift = 7;
  for (let d = 0; d < dayCount; d++) {
    for (const type of ['D', 'N']) {
      while (dayLoads[type][d] > maxPerShift) {
        const overloadedPeople = [];
        for (let p = 0; p < peopleCount; p++) {
          if (assignments[p][d] === type) overloadedPeople.push(p);
        }
        if (!overloadedPeople.length) break;

        let chosen = overloadedPeople.find((p) => requests[p][d] !== type);
        if (chosen === undefined) chosen = overloadedPeople[Math.floor(Math.random() * overloadedPeople.length)];

        removeFull(assignments, dayLoads, chosen, d, stats, weekKeyByDay);
        const newSlot = tryReassignPerson(assignments, requests, dayLoads, chosen, dayCount, maxPerShift, stats, weekKeyByDay, personTargetHours, personTargetShort);
        if (newSlot) assignFull(assignments, dayLoads, chosen, newSlot.day, newSlot.type, stats, weekKeyByDay);
      }
    }
  }

  if (!monthRequiresShortShift) {
    assignShortShiftsByPhase(phaseOrders, assignments, requests, dayLoads, dayCount, stats, weekKeyByDay, personTargetHours, personTargetShort, false);
  }

  enforceMonthlyHourCaps(assignments, requests, dayLoads, stats, weekKeyByDay, personTargetHours, personTargetShort);
  enforceShiftMinimum(assignments, requests, dayLoads, stats, weekKeyByDay, personTargetHours, personTargetShort, 'N', 5, 'D', 5);
  enforceShiftMinimum(assignments, requests, dayLoads, stats, weekKeyByDay, personTargetHours, personTargetShort, 'D', 5, 'N', 5);
  backfillVacationFromDaysOff(assignments, stats, personTargetHours, personTargetShort);

  const daySummary = Array.from({ length: dayCount }, (_, i) => ({ day: i + 1, D: [], N: [], K: [], U: [] }));
  for (let p = 0; p < peopleCount; p++) {
    for (let d = 0; d < dayCount; d++) {
      const code = assignments[p][d];
      if (!code) continue;
      if (daySummary[d][code]) {
        daySummary[d][code].push(p + 1);
      }
    }
  }

  const employeeStats = state.employees.map((emp, p) => {
    const creditedHours = stats.fullShiftCount[p] * SHIFT_HOURS
      + stats.vacationCreditHours[p]
      + stats.shortCount[p] * personTargetShort[p];
    return {
      person: p + 1,
      name: emp.name || '',
      contract: employeeContractLabel(emp.contract),
      requiresShortShift: requiresShortShiftByPerson[p],
      targetHours: personTargetHours[p],
      shiftTargetHours: personShiftHours[p],
      targetFullShifts: personTargetFull[p],
      targetShortHours: personTargetShort[p],
      fullShiftCount: stats.fullShiftCount[p],
      vacationCreditHours: stats.vacationCreditHours[p],
      shortCount: stats.shortCount[p],
      creditedHours
    };
  });
  const requestedFlags = requests.map((row) => row.map((code) => code !== ''));

  return {
    year,
    month,
    dayCount,
    businessDays,
    monthHours,
    shortNeededCount: requiresShortShiftByPerson.filter((v) => v).length,
    shortAssignedCount: stats.shortCount.filter((v) => v > 0).length,
    assignments,
    requestedFlags,
    daySummary,
    employeeStats
  };
}

function renderResult(result) {
  const summary = document.getElementById('summary');
  const missing = result.employeeStats.filter((e) => e.creditedHours + 0.01 < e.targetHours);
  const overMax = result.daySummary.filter((d) => d.D.length > 7 || d.N.length > 7 || (d.D.length + d.K.length > 7));
  const shortMissing = result.shortNeededCount - result.shortAssignedCount;
  const halfCount = result.employeeStats.filter((e) => e.contract === 'pół etatu').length;

  summary.innerHTML = `
    <article class="card"><h3>Dni robocze</h3><p>${result.businessDays}</p></article>
    <article class="card"><h3>Cel pełnego etatu</h3><p>${result.monthHours.toFixed(2)}h</p></article>
    <article class="card"><h3>Liczba pół etatów</h3><p>${halfCount}</p></article>
    <article class="card"><h3>Braki godzin</h3><p id="summaryMissingHours" class="${missing.length ? 'bad' : ''}">${missing.length} osób</p></article>
    <article class="card"><h3>Przekroczenie 7 os./godz.</h3><p class="${overMax.length ? 'bad' : ''}">${overMax.length} dni</p></article>
    <article class="card"><h3>Brak krótkiej zmiany</h3><p id="summaryShortMissing" class="${shortMissing > 0 ? 'bad' : ''}">${Math.max(0, shortMissing)} osób</p></article>
  `;

  const employeeTable = document.getElementById('employeeScheduleTable');
  const etHead = employeeTable.querySelector('thead');
  const etBody = employeeTable.querySelector('tbody');
  etHead.innerHTML = '';
  etBody.innerHTML = '';

  const eh = document.createElement('tr');
  let ehHtmlTop = '<th rowspan="2">Osoba</th>';
  for (let d = 1; d <= result.dayCount; d++) ehHtmlTop += `<th>${weekdayLabel(result.year, result.month, d)}</th>`;
  ehHtmlTop += '<th rowspan="2">Etat</th><th rowspan="2">Godz. zaliczone</th>';
  eh.innerHTML = ehHtmlTop;
  etHead.appendChild(eh);

  const eh2 = document.createElement('tr');
  let ehHtmlBottom = '';
  for (let d = 1; d <= result.dayCount; d++) ehHtmlBottom += `<th>${d}</th>`;
  eh2.innerHTML = ehHtmlBottom;
  etHead.appendChild(eh2);

  const appendSummaryHeaderRow = (label, prefix, cls) => {
    const tr = document.createElement('tr');
    tr.className = 'schedule-summary-row';
    let row = `<th>${label}</th>`;
    for (let d = 0; d < result.dayCount; d++) row += `<th id="${prefix}-${d + 1}" class="${cls}">0</th>`;
    row += '<th>-</th><th>-</th>';
    tr.innerHTML = row;
    etHead.appendChild(tr);
  };

  appendSummaryHeaderRow('Suma D', 'sumD', 'code-blue');
  appendSummaryHeaderRow('Suma N', 'sumN', 'code-blue');
  appendSummaryHeaderRow('Suma K', 'sumK', 'code-blue');
  appendSummaryHeaderRow('Suma U', 'sumU', 'code-u');

  employeeTable.dataset.rowCount = String(result.assignments.length);
  employeeTable.dataset.dayCount = String(result.dayCount);

  for (let p = 0; p < result.assignments.length; p++) {
    const tr = document.createElement('tr');
    const label = result.employeeStats[p].name ? `${p + 1} - ${result.employeeStats[p].name}` : `${p + 1}`;
    let row = `<td>${label}</td>`;
    for (let d = 0; d < result.dayCount; d++) {
      const code = result.assignments[p][d] || '-';
      const hadRequest = Boolean(result.requestedFlags && result.requestedFlags[p] && result.requestedFlags[p][d]);
      row += `
        <td>
          <select class="cell-select result-cell-select" data-p="${p + 1}" data-d="${d + 1}" data-req="${hadRequest ? '1' : '0'}">
            <option value="" ${code === '-' ? 'selected' : ''}>-</option>
            <option value="D" ${code === 'D' ? 'selected' : ''}>D</option>
            <option value="N" ${code === 'N' ? 'selected' : ''}>N</option>
            <option value="K" ${code === 'K' ? 'selected' : ''}>K</option>
            <option value="U" ${code === 'U' ? 'selected' : ''}>U</option>
            <option value="W" ${code === 'W' ? 'selected' : ''}>W</option>
            <option value="bN" ${code === 'bN' ? 'selected' : ''}>bN</option>
            <option value="bD" ${code === 'bD' ? 'selected' : ''}>bD</option>
          </select>
        </td>
      `;
    }
    employeeTable.dataset[`short${p + 1}`] = String(result.employeeStats[p].targetShortHours || 0);
    employeeTable.dataset[`target${p + 1}`] = String(result.employeeStats[p].targetHours || 0);
    employeeTable.dataset[`targetShort${p + 1}`] = String(result.employeeStats[p].targetShortHours || 0);
    employeeTable.dataset[`requireShort${p + 1}`] = result.employeeStats[p].requiresShortShift ? '1' : '0';
    row += `<td>${result.employeeStats[p].contract}</td><td id="hours-${p + 1}">${result.employeeStats[p].creditedHours.toFixed(2)}</td>`;
    tr.innerHTML = row;
    etBody.appendChild(tr);
  }

  document.querySelectorAll('#employeeScheduleTable .result-cell-select').forEach(applyResultCellClass);
  updateResultAggregates(result.dayCount);

  const dayBody = document.querySelector('#dayScheduleTable tbody');
  dayBody.innerHTML = '';
  for (const day of result.daySummary) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${day.day}</td>
      <td>${day.D.join(', ') || '-'}</td>
      <td>${day.N.join(', ') || '-'}</td>
      <td>${day.K.join(', ') || '-'}</td>
      <td>${day.U.join(', ') || '-'}</td>
    `;
    dayBody.appendChild(tr);
  }
  refreshTopScrollbars();
}

function syncLastResultFromResultTable() {
  if (!state.lastResult || !Array.isArray(state.lastResult.employeeStats)) return;

  const employeeTable = document.getElementById('employeeScheduleTable');
  const rowCount = Number(employeeTable.dataset.rowCount || 0);
  const dayCount = Number(employeeTable.dataset.dayCount || 0);
  if (!rowCount || !dayCount) return;

  const assignments = Array.from({ length: rowCount }, () => Array(dayCount).fill(''));
  document.querySelectorAll('#employeeScheduleTable .result-cell-select').forEach((sel) => {
    const p = Number(sel.dataset.p) - 1;
    const d = Number(sel.dataset.d) - 1;
    if (p < 0 || p >= rowCount || d < 0 || d >= dayCount) return;
    assignments[p][d] = sel.value || '';
  });

  const daySummary = Array.from({ length: dayCount }, (_, i) => ({ day: i + 1, D: [], N: [], K: [], U: [] }));
  for (let p = 0; p < rowCount; p++) {
    for (let d = 0; d < dayCount; d++) {
      const code = assignments[p][d];
      if (!code) continue;
      if (daySummary[d][code]) daySummary[d][code].push(p + 1);
    }
  }

  const employeeStats = state.lastResult.employeeStats.map((stat, idx) => {
    const hoursCell = document.getElementById(`hours-${idx + 1}`);
    const creditedHours = Number(hoursCell ? hoursCell.textContent : stat.creditedHours || 0);
    const hasShort = assignments[idx] ? assignments[idx].some((code) => code === 'K') : false;
    return {
      ...stat,
      creditedHours,
      shortCount: hasShort ? 1 : 0
    };
  });

  state.lastResult = {
    ...state.lastResult,
    assignments,
    daySummary,
    employeeStats,
    shortAssignedCount: employeeStats.filter((e) => e.shortCount > 0).length
  };
}

function exportScheduleAsXls() {
  const employeeTable = document.getElementById('employeeScheduleTable');
  const dayTable = document.getElementById('dayScheduleTable');
  if (!employeeTable || !dayTable) return false;
  if (!employeeTable.querySelector('tbody tr')) return false;

  const cloneTableForExport = (tableEl) => {
    const clone = tableEl.cloneNode(true);
    clone.querySelectorAll('select').forEach((sel) => {
      const td = sel.closest('td');
      if (!td) return;
      td.textContent = sel.value || '-';
    });
    return clone.outerHTML;
  };

  const y = state.lastResult?.year || state.year;
  const m = state.lastResult?.month || state.month;
  const ym = `${y}-${String(m).padStart(2, '0')}`;

  const html = `
    <html>
      <head>
        <meta charset="utf-8" />
      </head>
      <body>
        <h1>Grafik ${ym}</h1>
        <h2>Widok pracownika</h2>
        ${cloneTableForExport(employeeTable)}
        <h2>Widok dnia</h2>
        ${cloneTableForExport(dayTable)}
      </body>
    </html>
  `;

  const blob = new Blob([`\ufeff${html}`], { type: 'application/vnd.ms-excel;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `grafik-${ym}.xls`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  return true;
}

function generateFromGrid() {
  readCurrentInputs();
  syncEmployeesFromTable();
  const requests = collectRequests();
  const result = generateSchedule(requests, state.year, state.month);
  state.lastResult = result;
  renderResult(result);
  setActiveTab('result');
  requestAnimationFrame(refreshTopScrollbars);
}

function init() {
  if (!loadEmployeesFromStorage()) {
    createDefaultEmployees(31);
  }
  loadRequestsFromStorage();
  populateMonthSelect();
  document.getElementById('yearInput').value = String(state.year);

  renderEmployeesTable();
  rebuildRequestGrid();
  const hasSavedSchedule = loadScheduleFromStorage();
  if (hasSavedSchedule) {
    renderResult(state.lastResult);
  }

  document.getElementById('tabEmployeesBtn').addEventListener('click', () => setActiveTab('employees'));
  document.getElementById('tabPlanBtn').addEventListener('click', () => setActiveTab('plan'));
  document.getElementById('tabResultBtn').addEventListener('click', () => setActiveTab('result'));

  document.getElementById('monthInput').addEventListener('change', rebuildRequestGrid);
  document.getElementById('yearInput').addEventListener('change', rebuildRequestGrid);
  document.getElementById('refreshPeopleBtn').addEventListener('click', rebuildRequestGrid);
  document.getElementById('saveDemandBtn').addEventListener('click', () => {
    saveCurrentDemandForMonth();
    setDemandStatus('Zapisano zapotrzebowanie dla miesiąca.');
  });
  document.getElementById('saveScheduleBtn').addEventListener('click', () => {
    const saved = saveScheduleToStorage();
    setDemandStatus(saved ? 'Zapisano grafik.' : 'Brak grafiku do zapisu.');
  });
  document.getElementById('rebuildGridBtn').addEventListener('click', () => {
    clearCurrentDemandForMonth();
    rebuildRequestGrid();
    setDemandStatus('Wyczyszczono zapotrzebowanie dla miesiąca.');
  });

  document.getElementById('requestTable').addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) return;
    if (!target.classList.contains('cell-select')) return;
    applyRequestCellClass(target);
    setDemandStatus('Zmiany niezapisane.');
  });
  document.getElementById('requestTable').addEventListener('keydown', handleRequestGridArrows);
  document.getElementById('employeeScheduleTable').addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) return;
    if (!target.classList.contains('result-cell-select')) return;
    applyResultCellClass(target);
    const dayCount = Number(document.getElementById('employeeScheduleTable').dataset.dayCount || 0);
    updateResultAggregates(dayCount);
    syncLastResultFromResultTable();
  });
  document.getElementById('employeeScheduleTable').addEventListener('keydown', handleResultGridArrows);

  document.getElementById('employeesTable').addEventListener('input', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (!target.classList.contains('employee-name')) return;
    const id = Number(target.dataset.id);
    const emp = state.employees.find((e) => e.id === id);
    if (emp) emp.name = target.value;
    setSaveStatus('Zmiany niezapisane.');
  });

  document.getElementById('employeesTable').addEventListener('change', (event) => {
    const target = event.target;
    if (target instanceof HTMLSelectElement && target.classList.contains('employee-contract')) {
      const id = Number(target.dataset.id);
      const emp = state.employees.find((e) => e.id === id);
      if (emp) emp.contract = target.value;
      setSaveStatus('Zmiany niezapisane.');
      return;
    }
  });

  document.getElementById('employeesTable').addEventListener('click', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLButtonElement)) return;
    const removeId = Number(target.dataset.removeId);
    if (!Number.isFinite(removeId)) return;
    if (state.employees.length <= 1) return;
    state.employees = state.employees.filter((e) => e.id !== removeId);
    renderEmployeesTable();
    setSaveStatus('Zmiany niezapisane.');
  });

  document.getElementById('addEmployeeBtn').addEventListener('click', () => {
    state.employees.push({ id: state.nextEmployeeId++, name: '', contract: 'full' });
    renderEmployeesTable();
    setSaveStatus('Zmiany niezapisane.');
  });

  document.getElementById('saveEmployeesBtn').addEventListener('click', () => {
    saveEmployeesToStorage();
    setSaveStatus('Zapisano listę pracowników.');
  });

  document.getElementById('generateBtn').addEventListener('click', generateFromGrid);
  document.getElementById('regenerateBtn').addEventListener('click', generateFromGrid);
  document.getElementById('downloadXlsBtn').addEventListener('click', () => {
    exportScheduleAsXls();
  });

  setSaveStatus('Lista pracowników wczytana.');
  setDemandStatus(hasSavedSchedule ? 'Wczytano zapotrzebowanie i zapisany grafik.' : 'Zapotrzebowanie wczytane.');
  refreshTopScrollbars();
  window.addEventListener('resize', refreshTopScrollbars);
  setActiveTab(hasSavedSchedule ? 'result' : 'employees');
}

init();





