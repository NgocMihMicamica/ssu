/* global XLSX */

const state = {
  c1Rows: [],
  ddvRows: [],
  rankRows: [],
  images: [],
  rawFiles: {
    c1: null,
    ddv: null,
    ddvPlan: null,
  },
};

const dom = {
  c1File: document.getElementById('c1File'),
  ddvFile: document.getElementById('ddvFile'),
  ddvPlanFile: document.getElementById('ddvPlanFile'),
  attendanceImages: document.getElementById('attendanceImages'),
  parseButton: document.getElementById('parseButton'),
  sampleButton: document.getElementById('sampleButton'),
  exportButton: document.getElementById('exportButton'),
  statusBox: document.getElementById('statusBox'),
  logBox: document.getElementById('logBox'),
  imageList: document.getElementById('imageList'),
  c1TableBody: document.getElementById('c1TableBody'),
  ddvTableBody: document.getElementById('ddvTableBody'),
  rankTableBody: document.getElementById('rankTableBody'),
  c1ShiftFilter: document.getElementById('c1ShiftFilter'),
  c1Search: document.getElementById('c1Search'),
  ddvShiftFilter: document.getElementById('ddvShiftFilter'),
  ddvSearch: document.getElementById('ddvSearch'),
};

const defaultSheetNameHints = {
  c1: ['Lịch ctv đăng kí', 'Lich ctv dang ki', 'Lich ctv dang ky'],
  ddv: ['DDV'],
  ddvPlan: ['Phân công DDV', 'Phan cong DDV', 'Lịch đăng kí', 'Lich dang ki'],
};

function log(message) {
  const current = dom.logBox.textContent.trim();
  dom.logBox.textContent = current && current !== 'Chưa có log.' ? `${current}\n${message}` : message;
}

function setStatus(message, tone = 'info') {
  const colors = {
    info: '#6c4d5a',
    success: '#2f7752',
    warning: '#9a6a1b',
    danger: '#9b3949',
  };
  dom.statusBox.innerHTML = `<span style="color:${colors[tone] || colors.info}">${message}</span>`;
}

function normalizeText(value) {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeName(value) {
  return normalizeText(value).replace(/\u00a0/g, ' ');
}

function normalizeShift(value) {
  const text = normalizeText(value).toLowerCase();
  if (['sáng', 'sang', 'morning', 'am'].includes(text)) return 'Sang';
  if (['chiều', 'chieu', 'afternoon', 'pm'].includes(text)) return 'Chieu';
  if (['tối', 'toi', 'evening', 'night'].includes(text)) return 'Toi';
  return normalizeText(value) || 'Khac';
}

function isAllCaps(value) {
  const text = normalizeName(value);
  return Boolean(text) && text === text.toUpperCase() && /[A-ZÀ-Ỵ]/.test(text);
}

function detectLikelyPriority(row) {
  const name = row.name || '';
  return Boolean(row.priority) || isAllCaps(name) || /\bVIP\b/i.test(name);
}

function sheetNameFromWorkbook(workbook, hints) {
  const normalizedHints = hints.map((hint) => hint.toLowerCase());
  return workbook.SheetNames.find((name) => normalizedHints.includes(name.toLowerCase())) || workbook.SheetNames[0];
}

async function readUploadedFile(file) {
  if (!file) return null;
  const data = await file.arrayBuffer();
  const lower = file.name.toLowerCase();
  if (lower.endsWith('.csv')) {
    const text = new TextDecoder('utf-8').decode(data);
    return XLSX.read(text, { type: 'string' });
  }
  return XLSX.read(data, { type: 'array' });
}

function workbookToRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false, blankrows: false });
}

function findColumn(row, candidates) {
  const keys = Object.keys(row || {});
  const lowerCandidates = candidates.map((candidate) => candidate.toLowerCase());
  return keys.find((key) => lowerCandidates.includes(key.toLowerCase()));
}

function getCell(row, candidates, fallback = '') {
  const matchedKey = findColumn(row, candidates);
  if (matchedKey) return row[matchedKey];
  return fallback;
}

function parseC1Rows(rows) {
  return rows.map((row, index) => {
    const name = normalizeName(
      getCell(row, ['Họ tên', 'Ho ten', 'Tên', 'Ten', 'Name', 'CTV', 'Cộng tác viên'])
    );
    const shift = normalizeShift(getCell(row, ['Ca', 'Shift', 'Buổi', 'Session']));
    const attended = Boolean(getCell(row, ['Đi trực', 'Di truc', 'Đã đi', 'Da di', 'Present', 'Attended']));
    const scoreCell = Number(getCell(row, ['Điểm', 'Diem', 'Score'], 0));
    const priority = detectLikelyPriority({ name, priority: getCell(row, ['Ưu tiên', 'Uu tien', 'Priority']) });

    return {
      id: `c1-${index}`,
      group: 'C1',
      name,
      shift,
      priority,
      attended,
      score: Number.isFinite(scoreCell) ? scoreCell : 0,
      assignmentStatus: attended ? 'Đã đi' : 'Chưa đi',
      note: priority ? 'Tên in hoa / ưu tiên cao' : 'Bình thường',
      origin: row,
    };
  }).filter((row) => row.name);
}

function parseDdvRows(rows) {
  return rows.map((row, index) => {
    const name = normalizeName(getCell(row, ['Họ tên', 'Ho ten', 'Tên', 'Ten', 'Name', 'CTV']));
    const shift = normalizeShift(getCell(row, ['Ca', 'Shift', 'Buổi', 'Session']));
    const existingCount = Number(getCell(row, ['Đã đi', 'Da di', 'Số lần đi', 'So lan di', 'Count'], 0));
    const fixed = Boolean(getCell(row, ['Cố định', 'Co dinh', 'Fixed']));
    const preferred = Boolean(getCell(row, ['Ưu tiên', 'Uu tien', 'Priority']));
    return {
      id: `ddv-${index}`,
      group: 'DDV',
      name,
      shift,
      existingCount: Number.isFinite(existingCount) ? existingCount : 0,
      fixed,
      preferred,
      assignedTo: '',
      attended: false,
      note: fixed ? 'Cố định' : preferred ? 'Ưu tiên' : 'Theo lịch',
      origin: row,
    };
  }).filter((row) => row.name);
}

function parsePlanRows(rows) {
  return rows.map((row, index) => {
    const name = normalizeName(getCell(row, ['Họ tên', 'Ho ten', 'Tên', 'Ten', 'Name', 'CTV']));
    const shift = normalizeShift(getCell(row, ['Ca', 'Shift', 'Buổi', 'Session']));
    const historyCount = Number(getCell(row, ['Đã đi', 'Da di', 'Số lần đi', 'So lan di', 'Count'], 0));
    const note = normalizeText(getCell(row, ['Ghi chú', 'Ghi chu', 'Note']));
    return {
      id: `plan-${index}`,
      group: 'Plan',
      name,
      shift,
      historyCount: Number.isFinite(historyCount) ? historyCount : 0,
      note,
      assignedTo: '',
      attended: false,
      priority: false,
      origin: row,
    };
  }).filter((row) => row.name);
}

function rankScoreForPerson(person) {
  const attendanceBoost = person.attended ? 1 : -1;
  const priorityBoost = person.priority ? 2 : 0;
  const fixedBoost = person.fixed ? -2 : 0;
  const historyPenalty = -(person.existingCount ?? person.historyCount ?? 0) * 0.5;
  return Number((attendanceBoost + priorityBoost + fixedBoost + historyPenalty).toFixed(2));
}

function sortByPriorityAndHistory(list, countKey = 'existingCount') {
  return [...list].sort((a, b) => {
    const priorityDelta = Number(Boolean(b.priority)) - Number(Boolean(a.priority));
    if (priorityDelta !== 0) return priorityDelta;

    const fixedDelta = Number(Boolean(a.fixed)) - Number(Boolean(b.fixed));
    if (fixedDelta !== 0) return fixedDelta;

    const countDelta = Number(a[countKey] ?? 0) - Number(b[countKey] ?? 0);
    if (countDelta !== 0) return countDelta;

    return a.name.localeCompare(b.name, 'vi');
  });
}

function assignDdvSchedule(ddvRows, planRows) {
  const groupedByShift = {
    Sang: [],
    Chieu: [],
    Khac: [],
  };

  ddvRows.forEach((row) => {
    const bucket = groupedByShift[row.shift] ? row.shift : 'Khac';
    groupedByShift[bucket].push(row);
  });

  const planAssignments = planRows.map((plan) => {
    const pool = sortByPriorityAndHistory(
      (groupedByShift[plan.shift] || []).filter((candidate) => !candidate.assignedTo),
      'existingCount'
    );

    const chosen = pool[0] || null;
    if (chosen) {
      chosen.assignedTo = plan.name;
      chosen.shift = plan.shift;
      chosen.attended = Boolean(plan.attended);
      chosen.note = plan.note || chosen.note;
    }

    return {
      ...plan,
      assignedTo: chosen ? chosen.name : 'Chưa có người phù hợp',
      selectedPerson: chosen,
    };
  });

  return planAssignments;
}

function mergeRankRows(c1Rows, ddvRows, planRows) {
  const c1Rank = c1Rows.map((row) => ({
    name: row.name,
    group: 'C1',
    shift: row.shift,
    times: Number(row.score || 0),
    rank: rankScoreForPerson(row),
    note: row.assignmentStatus,
  }));

  const ddvRank = ddvRows.map((row) => ({
    name: row.name,
    group: 'DDV',
    shift: row.shift,
    times: row.existingCount,
    rank: rankScoreForPerson(row),
    note: row.note,
  }));

  const planRank = planRows
    .filter((row) => row.name)
    .map((row) => ({
      name: row.name,
      group: 'Plan',
      shift: row.shift,
      times: row.historyCount,
      rank: rankScoreForPerson(row),
      note: row.note || row.assignedTo || 'Lịch đăng kí',
    }));

  return [...c1Rank, ...ddvRank, ...planRank].sort((a, b) => b.rank - a.rank || a.name.localeCompare(b.name, 'vi'));
}

function renderC1Rows() {
  const query = normalizeText(dom.c1Search.value).toLowerCase();
  const shiftFilter = dom.c1ShiftFilter.value;
  const rows = state.c1Rows.filter((row) => {
    const shiftMatch = shiftFilter === 'all' || row.shift === shiftFilter;
    const queryMatch = !query || row.name.toLowerCase().includes(query);
    return shiftMatch && queryMatch;
  });

  dom.c1TableBody.innerHTML = rows
    .map((row, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.name)}</td>
        <td><span class="badge badge--soft">${escapeHtml(row.shift)}</span></td>
        <td>${row.priority ? '<span class="badge badge--warn">Ưu tiên</span>' : 'Bình thường'}</td>
        <td>${row.assignmentStatus === 'Đã đi' ? '<span class="badge badge--success">Đã đi</span>' : '<span class="badge badge--danger">Chưa đi</span>'}</td>
        <td>${row.score}</td>
        <td>
          <label class="checkbox-cell">
            <input type="checkbox" data-type="c1-attendance" data-id="${row.id}" ${row.attended ? 'checked' : ''} />
            Xác nhận
          </label>
        </td>
      </tr>
    `)
    .join('');

  dom.c1TableBody.querySelectorAll('input[data-type="c1-attendance"]').forEach((input) => {
    input.addEventListener('change', (event) => {
      const row = state.c1Rows.find((item) => item.id === event.target.dataset.id);
      if (!row) return;
      row.attended = event.target.checked;
      row.assignmentStatus = row.attended ? 'Đã đi' : 'Chưa đi';
      row.score += row.attended ? 1 : -1;
      updateRankTable();
      renderC1Rows();
    });
  });
}

function renderDdvRows() {
  const query = normalizeText(dom.ddvSearch.value).toLowerCase();
  const shiftFilter = dom.ddvShiftFilter.value;
  const rows = state.ddvRows.filter((row) => {
    const shiftMatch = shiftFilter === 'all' || row.shift === shiftFilter;
    const queryMatch = !query || row.name.toLowerCase().includes(query);
    return shiftMatch && queryMatch;
  });

  dom.ddvTableBody.innerHTML = rows
    .map((row, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.name)}</td>
        <td><span class="badge badge--soft">${escapeHtml(row.shift)}</span></td>
        <td>${row.existingCount}</td>
        <td>${row.fixed ? '<span class="badge badge--warn">Cố định</span>' : row.preferred ? '<span class="badge badge--soft">Ưu tiên</span>' : '<span class="badge">Mới / ít đi</span>'}</td>
        <td>${escapeHtml(row.assignedTo || 'Tự động')}</td>
        <td>
          <label class="checkbox-cell">
            <input type="checkbox" data-type="ddv-attendance" data-id="${row.id}" ${row.attended ? 'checked' : ''} />
            Đã đi
          </label>
        </td>
      </tr>
    `)
    .join('');

  dom.ddvTableBody.querySelectorAll('input[data-type="ddv-attendance"]').forEach((input) => {
    input.addEventListener('change', (event) => {
      const row = state.ddvRows.find((item) => item.id === event.target.dataset.id);
      if (!row) return;
      row.attended = event.target.checked;
      row.note = row.attended ? 'Đã tick thủ công' : row.note.replace('Đã tick thủ công', '').trim();
      updateRankTable();
      renderDdvRows();
    });
  });
}

function updateRankTable() {
  state.rankRows = mergeRankRows(state.c1Rows, state.ddvRows, state.ddvPlanRows || []);
  dom.rankTableBody.innerHTML = state.rankRows
    .map((row, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(row.name)}</td>
        <td><span class="badge badge--soft">${escapeHtml(row.group)}</span></td>
        <td>${escapeHtml(row.shift)}</td>
        <td>${row.times}</td>
        <td><strong>${row.rank.toFixed(2)}</strong></td>
        <td>${escapeHtml(row.note || '')}</td>
      </tr>
    `)
    .join('');
}

function renderImages() {
  dom.imageList.innerHTML = state.images.length
    ? state.images
        .map(
          (image) => `
          <div class="image-item">
            <div>
              <strong>${escapeHtml(image.name)}</strong>
              <br />
              <small>${Math.round(image.size / 1024)} KB</small>
            </div>
            <small>${escapeHtml(image.type || 'image')}</small>
          </div>
        `
        )
        .join('')
    : '<div class="image-item"><div><strong>Chưa có ảnh</strong><br /><small>Upload ảnh chấm công để ghi nhận bằng chứng.</small></div></div>';
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function seedSampleData() {
  state.c1Rows = [
    { id: 'c1-1', group: 'C1', name: 'NGUYEN AN', shift: 'Sang', priority: true, attended: true, score: 2, assignmentStatus: 'Đã đi', note: 'Tên in hoa / ưu tiên cao' },
    { id: 'c1-2', group: 'C1', name: 'Tran Binh', shift: 'Chieu', priority: false, attended: false, score: 0, assignmentStatus: 'Chưa đi', note: 'Bình thường' },
    { id: 'c1-3', group: 'C1', name: 'LE THAO', shift: 'Sang', priority: true, attended: true, score: 1, assignmentStatus: 'Đã đi', note: 'Tên in hoa / ưu tiên cao' },
  ];

  state.ddvRows = [
    { id: 'ddv-1', group: 'DDV', name: 'Pham Minh', shift: 'Sang', existingCount: 0, fixed: false, preferred: false, assignedTo: 'Lịch 1', attended: false, note: 'Theo lịch' },
    { id: 'ddv-2', group: 'DDV', name: 'Vo Linh', shift: 'Chieu', existingCount: 1, fixed: false, preferred: true, assignedTo: 'Lịch 2', attended: true, note: 'Ưu tiên' },
    { id: 'ddv-3', group: 'DDV', name: 'Do My', shift: 'Sang', existingCount: 3, fixed: true, preferred: false, assignedTo: 'Lịch 3', attended: false, note: 'Cố định' },
  ];

  state.ddvPlanRows = [
    { id: 'plan-1', group: 'Plan', name: 'Lịch 01', shift: 'Sang', historyCount: 0, note: 'Mẫu', assignedTo: '', attended: false, priority: false },
    { id: 'plan-2', group: 'Plan', name: 'Lịch 02', shift: 'Chieu', historyCount: 1, note: 'Mẫu', assignedTo: '', attended: false, priority: false },
  ];

  state.images = [{ name: 'sample-proof.png', size: 145000, type: 'image/png' }];
  assignDdvSchedule(state.ddvRows, state.ddvPlanRows);
  updateRankTable();
  renderC1Rows();
  renderDdvRows();
  renderImages();
  setStatus('Đã nạp dữ liệu mẫu. Bạn có thể dùng ngay hoặc upload file thật để thay thế.', 'success');
  log('Sample data loaded successfully.');
}

async function parseAllFiles() {
  try {
    const c1Workbook = await readUploadedFile(dom.c1File.files[0]);
    const ddvWorkbook = await readUploadedFile(dom.ddvFile.files[0]);
    const ddvPlanWorkbook = await readUploadedFile(dom.ddvPlanFile.files[0]);

    if (!c1Workbook && !ddvWorkbook && !ddvPlanWorkbook) {
      seedSampleData();
      return;
    }

    const c1Sheet = c1Workbook ? sheetNameFromWorkbook(c1Workbook, defaultSheetNameHints.c1) : null;
    const ddvSheet = ddvWorkbook ? sheetNameFromWorkbook(ddvWorkbook, defaultSheetNameHints.ddv) : null;
    const ddvPlanSheet = ddvPlanWorkbook ? sheetNameFromWorkbook(ddvPlanWorkbook, defaultSheetNameHints.ddvPlan) : null;

    state.c1Rows = c1Workbook ? parseC1Rows(workbookToRows(c1Workbook, c1Sheet)) : [];
    state.ddvRows = ddvWorkbook ? parseDdvRows(workbookToRows(ddvWorkbook, ddvSheet)) : [];
    state.ddvPlanRows = ddvPlanWorkbook ? parsePlanRows(workbookToRows(ddvPlanWorkbook, ddvPlanSheet)) : [];

    const imgFiles = Array.from(dom.attendanceImages.files || []);
    state.images = imgFiles.map((file) => ({ name: file.name, size: file.size, type: file.type }));

    assignDdvSchedule(state.ddvRows, state.ddvPlanRows);
    updateRankTable();
    renderC1Rows();
    renderDdvRows();
    renderImages();

    const totalC1 = state.c1Rows.length;
    const totalDdv = state.ddvRows.length;
    const totalPlan = state.ddvPlanRows.length;
    setStatus(`Đã xử lý xong dữ liệu: ${totalC1} C1, ${totalDdv} DDV, ${totalPlan} lịch phân công.`, 'success');
    log(`Parsed files successfully. C1=${totalC1}, DDV=${totalDdv}, Plan=${totalPlan}.`);
    log('Assignment rules applied: shift matching, priority sorting, and attendance scoring.');
  } catch (error) {
    console.error(error);
    setStatus('Có lỗi khi đọc file. Hãy kiểm tra lại định dạng Excel/CSV hoặc tên sheet.', 'danger');
    log(`Error: ${error.message}`);
  }
}

function exportWorkbook() {
  const workbook = XLSX.utils.book_new();

  const c1Sheet = XLSX.utils.json_to_sheet(
    state.c1Rows.map((row) => ({
      Ho_ten: row.name,
      Ca: row.shift,
      Uu_tien: row.priority ? 'Co' : 'Khong',
      Trang_thai: row.assignmentStatus,
      Diem: row.score,
      Ghi_chu: row.note,
    }))
  );

  const ddvSheet = XLSX.utils.json_to_sheet(
    state.ddvRows.map((row) => ({
      Ho_ten: row.name,
      Ca: row.shift,
      So_lan_di: row.existingCount,
      Co_dinh: row.fixed ? 'Co' : 'Khong',
      Uu_tien: row.preferred ? 'Co' : 'Khong',
      Da_gan_lich: row.assignedTo || '',
      Da_di: row.attended ? 'Co' : 'Khong',
      Ghi_chu: row.note,
    }))
  );

  const rankSheet = XLSX.utils.json_to_sheet(
    state.rankRows.map((row) => ({
      Ho_ten: row.name,
      Nhom: row.group,
      Ca: row.shift,
      So_lan_di: row.times,
      Diem_Rank: row.rank,
      Ghi_chu: row.note,
    }))
  );

  XLSX.utils.book_append_sheet(workbook, c1Sheet, 'C1');
  XLSX.utils.book_append_sheet(workbook, ddvSheet, 'DDV');
  XLSX.utils.book_append_sheet(workbook, rankSheet, 'RankTong');
  XLSX.writeFile(workbook, `ssu-ctv-phan-cong-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function wireEvents() {
  dom.parseButton.addEventListener('click', parseAllFiles);
  dom.sampleButton.addEventListener('click', seedSampleData);
  dom.exportButton.addEventListener('click', exportWorkbook);

  [dom.c1ShiftFilter, dom.c1Search].forEach((input) => input.addEventListener('input', renderC1Rows));
  [dom.ddvShiftFilter, dom.ddvSearch].forEach((input) => input.addEventListener('input', renderDdvRows));

  dom.attendanceImages.addEventListener('change', () => {
    state.images = Array.from(dom.attendanceImages.files || []).map((file) => ({
      name: file.name,
      size: file.size,
      type: file.type,
    }));
    renderImages();
  });
}

function initializeApp() {
  dom.logBox.textContent = 'Chưa có log.';
  renderImages();
  wireEvents();
  setStatus('Sẵn sàng. Hãy upload file Excel/CSV hoặc bấm tải dữ liệu mẫu để xem thử.', 'info');
}

initializeApp();
