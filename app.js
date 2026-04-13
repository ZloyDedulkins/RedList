const CONFIG = {
  spreadsheetId: '1QZDUhkwln01ymUqJlPf6RUmMyEvpMHvKpZLCfrRB8ek',
  mainSheetName: 'Список',
  mainSheetGid: null,
  mainColumns: {
    fio: 'Физическое лицо',
    date: 'Дата выгрузки',
    department: 'Подразделение',
    position: 'Должность',
    reason: 'Причина',
    status: 'Статус',
    lastPassDate: 'Дата последнего прохода'
  }
};

const statusEl = document.getElementById('status');
const resultSectionEl = document.getElementById('resultSection');
const resultBodyEl = document.getElementById('resultBody');

function normalizeKey(value) {
  return String(value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function createDepartmentLookupKeys(value) {
  const base = normalizeKey(value);
  if (!base) return [];

  const compact = base
    .replace(/\u00a0/g, ' ')
    .replace(/ё/g, 'е')
    .replace(/[«»"'`]/g, '')
    .replace(/[–—−-]+/g, ' ')
    .replace(/[(){}[\],.;:/\\|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const noSpaces = compact.replace(/\s+/g, '');
  return [...new Set([base, compact, noSpaces].filter(Boolean))];
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
  if (!response.ok) {
    return [];
  }

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
    const titlesHint = availableSheets.length
      ? `. Доступные листы: ${availableSheets.join(', ')}`
      : '';
    throw new Error(`Ошибка загрузки листа «${sheetName}»: ${response.status}${titlesHint}`);
  }
  const text = await response.text();
  const json = parseGoogleVisualization(text);
  if (json.status === 'error') {
    throw new Error(`Google Sheets вернул ошибку для листа «${sheetName}»: ${json.errors?.[0]?.detailed_message || json.errors?.[0]?.message || 'неизвестная ошибка'}`);
  }
  if (!json.table?.cols || !json.table?.rows) {
    throw new Error(`Лист «${sheetName}» не содержит читаемой таблицы. Проверьте доступ (\"Anyone with the link\") и корректность вкладки.`);
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

function resolveColumnName(sampleRow, expectedName, aliases) {
  return findExactColumnName(sampleRow, expectedName) || findColumnName(sampleRow, aliases);
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

function isCommentEmpty(comment) {
  return String(comment ?? '').trim() === '';
}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle('error', isError);
}

function renderRows(rows) {
  resultBodyEl.innerHTML = '';

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.fio}</td>
      <td>${row.department}</td>
      <td>${row.position}</td>
      <td>${row.lastPassDateText}</td>
      <td>${row.daysSinceLastPassText}</td>
    `;
    resultBodyEl.appendChild(tr);
  });

  resultSectionEl.hidden = false;
}

function getDaysSince(date) {
  if (!date) return null;

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const source = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const diffMs = today.getTime() - source.getTime();
  return Math.floor(diffMs / (1000 * 60 * 60 * 24));
}

async function init() {
  try {
    const mainRows = await fetchSheet({ name: CONFIG.mainSheetName, gid: CONFIG.mainSheetGid });

    if (!mainRows.length) {
      setStatus('Основной лист пуст.', true);
      return;
    }

    const sample = mainRows[0];
    const fioKey = resolveColumnName(sample, CONFIG.mainColumns.fio, ['фио', 'сотруд', 'физическ', 'person']);
    const dateKey = resolveColumnName(sample, CONFIG.mainColumns.date, ['дата', 'выгрузк', 'upload']);
    const departmentKey = resolveColumnName(sample, CONFIG.mainColumns.department, ['подраздел', 'отдел', 'департамент', 'локац', 'department']);
    const positionKey = resolveColumnName(sample, CONFIG.mainColumns.position, ['должност', 'position', 'role']);
    const reasonKey = resolveColumnName(sample, CONFIG.mainColumns.reason, ['причин', 'reason']);
    const statusKey = resolveColumnName(sample, CONFIG.mainColumns.status, ['статус', 'status']);
    const lastPassDateKey = resolveColumnName(sample, CONFIG.mainColumns.lastPassDate, ['последнего прохода', 'последн', 'проход', 'last pass']);

    if (!fioKey || !dateKey || !departmentKey || !positionKey || !reasonKey || !statusKey || !lastPassDateKey) {
      setStatus(`Не удалось определить нужные столбцы. Найдены колонки: ${Object.keys(sample).join(', ')}.`, true);
      return;
    }

    const rowsWithDate = mainRows
      .map((row) => ({ row, parsedDate: parseDate(row[dateKey]) }))
      .filter((item) => item.parsedDate);

    if (!rowsWithDate.length) {
      setStatus('В основном листе нет корректных дат.', true);
      return;
    }

    const maxDate = rowsWithDate.reduce((max, item) =>
      item.parsedDate > max ? item.parsedDate : max,
      rowsWithDate[0].parsedDate
    );

    const resultRows = rowsWithDate
      .filter(({ row, parsedDate }) =>
        sameDay(parsedDate, maxDate) &&
        isCommentEmpty(row[reasonKey]) &&
        isCommentEmpty(row[statusKey])
      )
      .map(({ row }) => {
        const department = String(row[departmentKey] ?? '').trim() || '—';
        const fio = String(row[fioKey] ?? '').trim() || '—';
        const position = String(row[positionKey] ?? '').trim() || '—';
        const lastPassDate = parseDate(row[lastPassDateKey]);
        const daysSinceLastPass = getDaysSince(lastPassDate);
        const lastPassDateText = lastPassDate ? lastPassDate.toLocaleDateString('ru-RU') : '—';
        const daysSinceLastPassText = Number.isInteger(daysSinceLastPass) ? String(daysSinceLastPass) : '—';

        return {
          fio,
          department,
          position,
          lastPassDateText,
          daysSinceLastPassText
        };
      });

    if (!resultRows.length) {
      setStatus('На максимальную дату нет записей с пустыми полями «Причина» и «Статус».');
      return;
    }

    setStatus(`Найдено записей: ${resultRows.length}. Дата: ${maxDate.toLocaleDateString('ru-RU')}.`);
    renderRows(resultRows);
  } catch (error) {
    setStatus(`Ошибка: ${error.message}`, true);
  }
}

init();
