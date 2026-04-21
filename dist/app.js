const CONFIG = {
  spreadsheetId: '1QZDUhkwln01ymUqJlPf6RUmMyEvpMHvKpZLCfrRB8ek',
  mainSheetName: 'Список',
  mainSheetGid: null,
  feedbackForm: {
    baseUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSfXFCl2TKFHFKdXwYQmIsVpKpSvzUJE4YkaDHkHHzyVOb8yAw/viewform',
    fioEntryId: '1750988741',

  },
  mainColumns: {
    fio: 'Физическое лицо',
    date: 'Дата выгрузки',
    department: 'Подразделение',
    position: 'Должность',
    state: 'Состояние',
    reason: 'Причина',
    status: 'Статус',
    lastPassDate: 'Дата последнего прохода'
  }
};

const tabButtons = Array.from(document.querySelectorAll('.tab-btn'));
const tabPanels = {
  main: document.getElementById('tabMain'),
  history: document.getElementById('tabHistory'),
  source: document.getElementById('tabSource')
};

const ui = {
  main: {
    statusEl: document.getElementById('statusMain'),
    resultSectionEl: document.getElementById('resultSectionMain'),
    resultBodyEl: document.getElementById('resultBodyMain'),
    departmentFilterEl: document.getElementById('departmentFilterMain')
  },
  history: {
    statusEl: document.getElementById('statusHistory'),
    resultSectionEl: document.getElementById('resultSectionHistory'),
    resultBodyEl: document.getElementById('resultBodyHistory'),
    departmentFilterEl: document.getElementById('departmentFilterHistory')
  },
  source: {
    statusEl: document.getElementById('statusSource'),
    resultSectionEl: document.getElementById('resultSectionSource'),
    resultBodyEl: document.getElementById('resultBodySource'),
    dateFilterEl: document.getElementById('dateFilterSource'),
    departmentFilterEl: document.getElementById('departmentFilterSource')
  }
};

function normalizeKey(value) {
  return String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function parseGoogleVisualization(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1) {
    throw new Error('Не удалось прочитать ответ Google Sheets.');
  }
  return JSON.parse(text.slice(start, end + 1));
}

async function fetchWorksheetTitles() {
  const url = `https://spreadsheets.google.com/feeds/worksheets/${CONFIG.spreadsheetId}/public/basic?alt=json`;
  const response = await fetch(url);
  if (!response.ok) return [];

  const json = await response.json();
  const entries = json.feed?.entry ?? [];
  return entries
    .map((entry) => String(entry.title?.$t ?? '').trim())
    .filter(Boolean);
}

async function resolveActualSheetName(sheetName) {
  const worksheetTitles = await fetchWorksheetTitles();
  if (!worksheetTitles.length) {
    return { resolvedName: sheetName, availableSheets: [] };
  }

  const normalizedRequested = normalizeKey(sheetName);
  const exact = worksheetTitles.find((title) => title === sheetName);
  if (exact) {
    return { resolvedName: exact, availableSheets: worksheetTitles };
  }

  const normalizedMatch = worksheetTitles.find((title) => normalizeKey(title) === normalizedRequested);
  return {
    resolvedName: normalizedMatch || sheetName,
    availableSheets: worksheetTitles
  };
}

function buildGvizUrl({ sheetName, sheetGid }) {
  const params = new URLSearchParams({ tqx: 'out:json' });
  if (sheetGid !== null && sheetGid !== undefined && String(sheetGid).trim() !== '') {
    params.set('gid', String(sheetGid).trim());
  } else {
    params.set('sheet', sheetName);
  }
  return `https://docs.google.com/spreadsheets/d/${CONFIG.spreadsheetId}/gviz/tq?${params.toString()}`;
}


function resolveFeedbackFormEntryId() {
  const params = new URLSearchParams(window.location.search);
  const fromUrl = params.get('formEntryId');
  const fromGlobal = window.FEEDBACK_FORM_ENTRY_ID;

  return String(fromUrl || fromGlobal || CONFIG.feedbackForm.fioEntryId).trim();
}

function buildFeedbackFormUrl(fio) {
  const entryId = resolveFeedbackFormEntryId();
  const params = new URLSearchParams({
    usp: 'pp_url',
    [`entry.${entryId}`]: fio
  });

  Object.entries(CONFIG.feedbackForm.staticEntries || {}).forEach(([id, value]) => {
    params.set(`entry.${id}`, value);
  });

  return `${CONFIG.feedbackForm.baseUrl}?${params.toString()}`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

async function fetchSheet(sheetConfig) {
  const config = typeof sheetConfig === 'string' ? { name: sheetConfig } : sheetConfig;
  const sheetName = config.name;
  const sheetGid = config.gid;

  const { resolvedName, availableSheets } = sheetGid !== null && sheetGid !== undefined && String(sheetGid).trim() !== ''
    ? { resolvedName: sheetName, availableSheets: [] }
    : await resolveActualSheetName(sheetName);

  const url = buildGvizUrl({ sheetName: resolvedName, sheetGid });
  const response = await fetch(url);
  if (!response.ok) {
    const titlesHint = availableSheets.length ? `. Доступные листы: ${availableSheets.join(', ')}` : '';
    throw new Error(`Ошибка загрузки листа «${sheetName}»: ${response.status}${titlesHint}`);
  }

  const text = await response.text();
  const json = parseGoogleVisualization(text);

  if (json.status === 'error') {
    const details = json.errors?.[0]?.detailed_message || json.errors?.[0]?.message || 'неизвестная ошибка';
    throw new Error(`Google Sheets вернул ошибку для листа «${sheetName}»: ${details}`);
  }

  if (!json.table?.cols || !json.table?.rows) {
    throw new Error('Лист не содержит читаемой таблицы. Проверьте доступ ("Anyone with the link") и корректность вкладки.');
  }

  const cols = json.table.cols.map((c, idx) => ({
    key: c.label || c.id || `col_${idx}`,
    index: idx
  }));

  return json.table.rows.map((row) => {
    const output = {};
    cols.forEach((col) => {
      const cell = row.c?.[col.index];
      output[col.key] = cell?.f ?? cell?.v ?? '';
    });
    return output;
  });
}

function findColumnName(sampleRow, aliases) {
  const keys = Object.keys(sampleRow || {});
  for (const key of keys) {
    const normalized = normalizeKey(key);
    if (aliases.some((alias) => normalized.includes(alias))) {
      return key;
    }
  }
  return null;
}

function findExactColumnName(sampleRow, expectedName) {
  const keys = Object.keys(sampleRow || {});
  const normalizedExpected = normalizeKey(expectedName);
  return keys.find((key) => normalizeKey(key) === normalizedExpected) || null;
}

function resolveColumnName(sampleRow, expectedName, aliases, required = true) {
  const resolved = findExactColumnName(sampleRow, expectedName) || findColumnName(sampleRow, aliases);
  if (!required) return resolved;
  return resolved;
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;

  const str = String(value).trim();

  const gvizDate = str.match(/^Date\((\d+),(\d+),(\d+)/);
  if (gvizDate) {
    return new Date(Number(gvizDate[1]), Number(gvizDate[2]), Number(gvizDate[3]));
  }

  const ru = str.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (ru) {
    const year = ru[3].length === 2 ? `20${ru[3]}` : ru[3];
    return new Date(Number(year), Number(ru[2]) - 1, Number(ru[1]));
  }

  const parsed = new Date(str);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate();
}

function isCommentEmpty(value) {
  return String(value ?? '').trim() === '';
}

function getDaysSince(date) {
  if (!date) return null;

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const source = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const diffMs = today.getTime() - source.getTime();
  return Math.floor(diffMs / (1000 * 60 * 60 * 24));
}

function toDisplay(value) {
  const text = String(value ?? '').trim();
  return text || '—';
}

function toRuDateOrDash(date) {
  return date ? date.toLocaleDateString('ru-RU') : '—';
}

function dateSortDesc(a, b) {
  return b.getTime() - a.getTime();
}

function getDepartmentValue(value) {
  const text = String(value ?? '').trim();
  return text || '—';
}

function dateSortDesc(a, b) {
  return b.getTime() - a.getTime();
}

function getDepartmentValue(value) {
  const text = String(value ?? '').trim();
  return text || '—';
}

function setStatus(tab, message, isError = false) {
  const { statusEl } = ui[tab];
  statusEl.textContent = message;
  statusEl.classList.toggle('error', isError);
}

function renderRows(tab, rows, columns) {
  const { resultBodyEl, resultSectionEl } = ui[tab];
  resultBodyEl.innerHTML = '';

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    tr.innerHTML = columns.map((column) => {
      const value = row[column];
      if (value && typeof value === 'object' && value.rawHtml) {
        return `<td>${value.rawHtml}</td>`;
      }

      return `<td>${escapeHtml(value)}</td>`;
    }).join('');

    resultBodyEl.appendChild(tr);
  });

  resultSectionEl.hidden = false;
}

function fillSelect(selectEl, options, allLabel) {
  if (!selectEl) return;
  selectEl.innerHTML = '';

  const allOption = document.createElement('option');
  allOption.value = '';
  allOption.textContent = allLabel;
  selectEl.appendChild(allOption);

  options.forEach((option) => {
    const el = document.createElement('option');
    el.value = option;
    el.textContent = option;
    selectEl.appendChild(el);
  });
}

function applyDepartmentFilter(rows, department) {
  if (!department) return rows;
  return rows.filter((row) => row.department === department);
}

function setupTabs() {
  tabButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const activeTab = button.dataset.tab;
      tabButtons.forEach((item) => {
        const isActive = item === button;
        item.classList.toggle('active', isActive);
        item.setAttribute('aria-selected', String(isActive));
      });

      Object.entries(tabPanels).forEach(([tabKey, panel]) => {
        panel.hidden = tabKey !== activeTab;
      });
    });
  });
}

async function init() {
  setupTabs();

  try {
    const mainRows = await fetchSheet({ name: CONFIG.mainSheetName, gid: CONFIG.mainSheetGid });

    if (!mainRows.length) {
      setStatus('main', 'Основной лист пуст.', true);
      setStatus('history', 'Основной лист пуст.', true);
      setStatus('source', 'Основной лист пуст.', true);
      return;
    }

    const sample = mainRows[0];
    const fioKey = resolveColumnName(sample, CONFIG.mainColumns.fio, ['фио', 'сотруд', 'физическ', 'person']);
    const dateKey = resolveColumnName(sample, CONFIG.mainColumns.date, ['дата', 'выгрузк', 'upload']);
    const departmentKey = resolveColumnName(sample, CONFIG.mainColumns.department, ['подраздел', 'отдел', 'департамент', 'локац', 'department']);
    const positionKey = resolveColumnName(sample, CONFIG.mainColumns.position, ['должност', 'position', 'role']);
    const stateKey = resolveColumnName(sample, CONFIG.mainColumns.state, ['состояни', 'state'], false);
    const reasonKey = resolveColumnName(sample, CONFIG.mainColumns.reason, ['причин', 'reason']);
    const statusKey = resolveColumnName(sample, CONFIG.mainColumns.status, ['статус', 'status']);
    const lastPassDateKey = resolveColumnName(sample, CONFIG.mainColumns.lastPassDate, ['последнего прохода', 'последн', 'проход', 'last pass']);

    if (!fioKey || !dateKey || !departmentKey || !positionKey || !reasonKey || !statusKey || !lastPassDateKey) {
      const columns = Object.keys(sample).join(', ');
      setStatus('main', `Не удалось определить нужные столбцы. Найдены колонки: ${columns}.`, true);
      setStatus('history', `Не удалось определить нужные столбцы. Найдены колонки: ${columns}.`, true);
      setStatus('source', `Не удалось определить нужные столбцы. Найдены колонки: ${columns}.`, true);
      return;
    }

    const rowsWithDate = mainRows
      .map((row) => ({ row, parsedDate: parseDate(row[dateKey]) }))
      .filter((item) => item.parsedDate);

    if (!rowsWithDate.length) {
      setStatus('main', 'В основном листе нет корректных дат.', true);
      setStatus('history', 'В основном листе нет корректных дат.', true);
      setStatus('source', 'В основном листе нет корректных дат.', true);
      return;
    }

    const maxDate = rowsWithDate.reduce((max, item) => (item.parsedDate > max ? item.parsedDate : max), rowsWithDate[0].parsedDate);
    const maxDateRows = rowsWithDate.filter(({ parsedDate }) => sameDay(parsedDate, maxDate));

    const noFeedbackRowsBase = maxDateRows
      .filter(({ row }) => isCommentEmpty(row[reasonKey]) && isCommentEmpty(row[statusKey]))
      .map(({ row }) => {
        const lastPassDate = parseDate(row[lastPassDateKey]);
        const daysSinceLastPass = getDaysSince(lastPassDate);
        const fio = toDisplay(row[fioKey]);

        return {
          fio: {
            rawHtml: `<a class="person-link" href="${buildFeedbackFormUrl(fio)}" target="_blank" rel="noopener noreferrer">${escapeHtml(fio)}</a>`
          },
          department: getDepartmentValue(row[departmentKey]),
          position: toDisplay(row[positionKey]),
          lastPassDateText: toRuDateOrDash(lastPassDate),
          daysSinceLastPassText: Number.isInteger(daysSinceLastPass) ? String(daysSinceLastPass) : '—'
        };
      });

    const mainDepartments = Array.from(new Set(noFeedbackRowsBase.map((row) => row.department))).sort((a, b) => a.localeCompare(b, 'ru'));
    fillSelect(ui.main.departmentFilterEl, mainDepartments, 'Все подразделения');

    const renderMain = () => {
      const filtered = applyDepartmentFilter(noFeedbackRowsBase, ui.main.departmentFilterEl.value);
      if (!filtered.length) {
        setStatus('main', `Нет записей по выбранному подразделению. Дата: ${maxDate.toLocaleDateString('ru-RU')}.`);
        renderRows('main', [], ['fio', 'department', 'position', 'lastPassDateText', 'daysSinceLastPassText']);
        return;
      }

      setStatus('main', `Найдено записей: ${filtered.length}. Дата: ${maxDate.toLocaleDateString('ru-RU')}.`);
      renderRows('main', filtered, ['fio', 'department', 'position', 'lastPassDateText', 'daysSinceLastPassText']);
    };

    ui.main.departmentFilterEl.addEventListener('change', renderMain);
    renderMain();

    const departmentsOnEarlierDates = new Set(
      rowsWithDate
        .filter(({ parsedDate }) => parsedDate < maxDate)
        .map(({ row }) => normalizeKey(row[departmentKey]))
        .filter(Boolean)
    );

    const historyRowsBase = maxDateRows
      .filter(({ row }) => peopleOnEarlierDates.has(normalizeKey(row[fioKey])))
      .map(({ row, parsedDate }) => ({
        exportDateText: toRuDateOrDash(parsedDate),
        fio: toDisplay(row[fioKey]),
        firstSeenDateText: '—',
        entryCountText: '—',
        department: getDepartmentValue(row[departmentKey]),
        position: toDisplay(row[positionKey]),
        lastPassDateText: toRuDateOrDash(parseDate(row[lastPassDateKey])),
        state: stateKey ? toDisplay(row[stateKey]) : '—',
        reason: toDisplay(row[reasonKey]),
        status: toDisplay(row[statusKey])
      }));

    const historyDepartments = Array.from(new Set(historyRowsBase.map((row) => row.department))).sort((a, b) => a.localeCompare(b, 'ru'));
    fillSelect(ui.history.departmentFilterEl, historyDepartments, 'Все подразделения');

    const renderHistory = () => {
      const filtered = applyDepartmentFilter(historyRowsBase, ui.history.departmentFilterEl.value);
      if (!filtered.length) {
        setStatus('history', `Нет записей по выбранному подразделению. Дата: ${maxDate.toLocaleDateString('ru-RU')}.`);
        renderRows('history', [], ['exportDateText', 'fio', 'firstSeenDateText', 'entryCountText', 'department', 'position', 'lastPassDateText', 'state', 'reason', 'status']);
        return;
      }

      setStatus('history', `Найдено записей: ${filtered.length}. Дата: ${maxDate.toLocaleDateString('ru-RU')}.`);
      renderRows('history', filtered, ['exportDateText', 'fio', 'firstSeenDateText', 'entryCountText', 'department', 'position', 'lastPassDateText', 'state', 'reason', 'status']);
    };

    ui.history.departmentFilterEl.addEventListener('change', renderHistory);
    renderHistory();

    const sourceRowsBase = rowsWithDate.map(({ row, parsedDate }) => ({
      parsedDate,
      dateKey: parsedDate.toISOString().slice(0, 10),
      exportDateText: toRuDateOrDash(parsedDate),
      fio: toDisplay(row[fioKey]),
      department: getDepartmentValue(row[departmentKey]),
      position: toDisplay(row[positionKey]),
      lastPassDateText: toRuDateOrDash(parseDate(row[lastPassDateKey])),
      state: stateKey ? toDisplay(row[stateKey]) : '—',
      reason: toDisplay(row[reasonKey]),
      status: toDisplay(row[statusKey])
    }));

    const sourceDates = Array.from(
      new Map(sourceRowsBase.map((row) => [row.dateKey, row.parsedDate])).entries()
    )
      .sort((a, b) => dateSortDesc(a[1], b[1]))
      .map(([key, date]) => ({ key, label: toRuDateOrDash(date) }));

    ui.source.dateFilterEl.innerHTML = '';
    const sourceAllDateOption = document.createElement('option');
    sourceAllDateOption.value = '';
    sourceAllDateOption.textContent = 'Все даты';
    ui.source.dateFilterEl.appendChild(sourceAllDateOption);
    sourceDates.forEach((item) => {
      const optionEl = document.createElement('option');
      optionEl.value = item.key;
      optionEl.textContent = item.label;
      ui.source.dateFilterEl.appendChild(optionEl);
    });

    const sourceDepartments = Array.from(new Set(sourceRowsBase.map((row) => row.department))).sort((a, b) => a.localeCompare(b, 'ru'));
    fillSelect(ui.source.departmentFilterEl, sourceDepartments, 'Все подразделения');

    const renderSource = () => {
      const dateValue = ui.source.dateFilterEl.value;
      const departmentValue = ui.source.departmentFilterEl.value;

      const filteredByDate = dateValue
        ? sourceRowsBase.filter((row) => row.dateKey === dateValue)
        : sourceRowsBase;
      const filtered = applyDepartmentFilter(filteredByDate, departmentValue)
        .slice()
        .sort((a, b) => dateSortDesc(a.parsedDate, b.parsedDate));

      if (!filtered.length) {
        setStatus('source', 'Нет записей по выбранным фильтрам.');
        renderRows('source', [], ['exportDateText', 'fio', 'department', 'position', 'lastPassDateText', 'state', 'reason', 'status']);
        return;
      }

      setStatus('source', `Найдено записей: ${filtered.length}.`);
      renderRows('source', filtered, ['exportDateText', 'fio', 'department', 'position', 'lastPassDateText', 'state', 'reason', 'status']);
    };

    ui.source.dateFilterEl.addEventListener('change', renderSource);
    ui.source.departmentFilterEl.addEventListener('change', renderSource);
    renderSource();
  } catch (error) {
    setStatus('main', `Ошибка: ${error.message}`, true);
    setStatus('history', `Ошибка: ${error.message}`, true);
    setStatus('source', `Ошибка: ${error.message}`, true);
  }
}

init();
