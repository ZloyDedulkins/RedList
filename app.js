const CONFIG = {
  spreadsheetId: '1QZDUhkwln01ymUqJlPf6RUmMyEvpMHvKpZLCfrRB8ek',
  mainSheetName: 'Список',
  bridgeSheetName: 'Мост (Имена подразделений)'
};

const statusEl = document.getElementById('status');
const resultSectionEl = document.getElementById('resultSection');
const resultBodyEl = document.getElementById('resultBody');

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

async function fetchSheetByName(sheetName) {
  const { resolvedName, availableSheets } = await resolveActualSheetName(sheetName);
  const url = `https://docs.google.com/spreadsheets/d/${CONFIG.spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(resolvedName)}`;
  const response = await fetch(url);
  if (!response.ok) {
    const titlesHint = availableSheets.length
      ? `. Доступные листы: ${availableSheets.join(', ')}`
      : '';
    throw new Error(`Ошибка загрузки листа «${sheetName}»: ${response.status}${titlesHint}`);
  }
  const text = await response.text();
  const json = parseGoogleVisualization(text);
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

function buildBridgeMap(bridgeRows) {
  if (!bridgeRows.length) return new Map();

  const sample = bridgeRows[0];
  const sourceKey = findColumnName(sample, ['исход', 'raw', 'как в']);
  const targetKey = findColumnName(sample, ['подраздел', 'норм', 'целев', 'итог']);

  const keys = Object.keys(sample);
  const fallbackSource = keys[0];
  const fallbackTarget = keys[1] || keys[0];

  const map = new Map();
  bridgeRows.forEach((row) => {
    const src = normalizeKey(row[sourceKey || fallbackSource]);
    const dst = String(row[targetKey || fallbackTarget] ?? '').trim();
    if (src && dst) {
      map.set(src, dst);
    }
  });

  return map;
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
      <td>${row.formatText}</td>
    `;
    resultBodyEl.appendChild(tr);
  });

  resultSectionEl.hidden = false;
}

async function init() {
  try {
    const [mainRows, bridgeRows] = await Promise.all([
      fetchSheetByName(CONFIG.mainSheetName),
      fetchSheetByName(CONFIG.bridgeSheetName)
    ]);

    if (!mainRows.length) {
      setStatus('Основной лист пуст.', true);
      return;
    }

    const sample = mainRows[0];
    const fioKey = findColumnName(sample, ['фио', 'сотруд']);
    const dateKey = findColumnName(sample, ['дата']);
    const commentKey = findColumnName(sample, ['коммент']);
    const departmentKey = findColumnName(sample, ['подраздел', 'отдел', 'департамент']);

    if (!fioKey || !dateKey || !commentKey || !departmentKey) {
      setStatus('Не удалось автоматически определить нужные столбцы. Проверьте названия колонок.', true);
      return;
    }

    const bridgeMap = buildBridgeMap(bridgeRows);

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
      .filter(({ row, parsedDate }) => sameDay(parsedDate, maxDate) && isCommentEmpty(row[commentKey]))
      .map(({ row }) => {
        const rawDepartment = String(row[departmentKey] ?? '').trim();
        const department = bridgeMap.get(normalizeKey(rawDepartment)) || rawDepartment || '—';
        const fio = String(row[fioKey] ?? '').trim() || '—';

        return {
          fio,
          department,
          formatText: `${fio} - ${department}`
        };
      });

    if (!resultRows.length) {
      setStatus('На максимальную дату нет записей с пустым комментарием.');
      return;
    }

    setStatus(`Найдено записей: ${resultRows.length}. Дата: ${maxDate.toLocaleDateString('ru-RU')}.`);
    renderRows(resultRows);
  } catch (error) {
    setStatus(`Ошибка: ${error.message}`, true);
  }
}

init();
