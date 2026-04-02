const Q_INDEX = 16;
const STORAGE_KEY = "pcpmastergrid_saved_v1";
const STORAGE_ROW_LIMIT = 8000;
const STORAGE_CELL_LIMIT = 180000;
const FILTER_OPTION_LIMIT = 1500;
const CHART_ITEM_LIMIT = 40;
const PARSE_YIELD_EVERY = 250;

const state = {
  rows: [],
  headers: [],
  headerLetters: [],
  formulaIgnored: 0,
  qFallbackUsed: false,
  sourceFiles: [],
  selectedFiles: [],
  persistenceEnabled: true,
  lastStatusMessage: "",
  charts: {
    count: null,
    share: null,
    metric: null,
    compare: null
  }
};

const els = {
  files: document.getElementById("files"),
  selectedFilesWrap: document.getElementById("selectedFilesWrap"),
  selectedFilesList: document.getElementById("selectedFilesList"),
  clearSelection: document.getElementById("clearSelection"),
  process: document.getElementById("process"),
  clearData: document.getElementById("clearData"),
  metricColumn: document.getElementById("metricColumn"),
  qFilter: document.getElementById("qFilter"),
  sheetFilter: document.getElementById("sheetFilter"),
  azColumn: document.getElementById("azColumn"),
  azValue: document.getElementById("azValue"),
  compareBy: document.getElementById("compareBy"),
  maxRows: document.getElementById("maxRows"),
  download: document.getElementById("download"),
  kpiFiles: document.getElementById("kpiFiles"),
  kpiRows: document.getElementById("kpiRows"),
  kpiQ: document.getElementById("kpiQ"),
  kpiGroups: document.getElementById("kpiGroups"),
  kpiFormula: document.getElementById("kpiFormula"),
  tableMeta: document.getElementById("tableMeta"),
  dataTable: document.getElementById("dataTable"),
  compareTable: document.getElementById("compareTable"),
  statusBar: document.getElementById("statusBar"),
  statusTitle: document.getElementById("statusTitle"),
  statusMeta: document.getElementById("statusMeta"),
  statusFill: document.getElementById("statusFill")
};

function toColName(idx) {
  let n = idx + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function updateStatus(title, meta = "", progress = null, keepVisible = true) {
  state.lastStatusMessage = title;
  els.statusTitle.textContent = title;
  els.statusMeta.textContent = meta;
  els.statusBar.hidden = !keepVisible;
  els.statusFill.style.width = progress === null ? "0%" : `${Math.max(0, Math.min(100, progress))}%`;
}

function hideStatus() {
  els.statusBar.hidden = true;
}

function yieldToUi() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => resolve());
  });
}

function estimateCellCount(rows, headers) {
  return rows.length * Math.max(1, headers.length);
}

function ensureUniqueHeaders(headers) {
  const map = new Map();
  return headers.map((raw, index) => {
    const base = String(raw || "").trim() || toColName(index);
    const count = map.get(base) || 0;
    map.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

function getSheetCell(sheet, r, c) {
  if (Array.isArray(sheet)) {
    const row = sheet[r];
    return Array.isArray(row) ? row[c] : undefined;
  }

  const addr = XLSX.utils.encode_cell({ r, c });
  return sheet[addr];
}

function getCellValue(sheet, r, c) {
  const cell = getSheetCell(sheet, r, c);
  if (!cell) return "";

  if (cell.f) {
    state.formulaIgnored += 1;
  }

  if (!("v" in cell)) return "";

  const value = cell.v;
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "string") return value.trim();
  return value;
}

function parseNumeric(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string") return null;

  const cleaned = value.trim().replace(/\s/g, "");
  if (!cleaned) return null;

  const hasComma = cleaned.includes(",");
  const hasDot = cleaned.includes(".");

  let normalized = cleaned;
  if (hasComma && hasDot) {
    normalized = cleaned.lastIndexOf(",") > cleaned.lastIndexOf(".")
      ? cleaned.replace(/\./g, "").replace(",", ".")
      : cleaned.replace(/,/g, "");
  } else if (hasComma) {
    normalized = cleaned.replace(/\./g, "").replace(",", ".");
  }

  const n = Number(normalized);
  return Number.isFinite(n) ? n : null;
}

function pickQHeader(headers, startCol) {
  const byIndex = headers[Q_INDEX - startCol];
  if (byIndex) {
    const t = String(byIndex).trim();
    if (t) return { header: byIndex, usedFallback: false };
  }

  const byName = headers.find((h) => String(h || "").trim().toUpperCase() === "Q");
  if (byName) return { header: byName, usedFallback: false };

  const fallback = headers.find((h) => String(h || "").trim() !== "");
  if (fallback) return { header: fallback, usedFallback: true };

  return { header: "", usedFallback: false };
}

async function parseSheet(fileName, sheetName, sheet, progressInfo) {
  if (!sheet["!ref"]) return { headers: [], rows: [], usedQFallback: false };

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headerRow = range.s.r;

  const rawHeaders = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    rawHeaders.push(getCellValue(sheet, headerRow, c));
  }
  const headers = ensureUniqueHeaders(rawHeaders);
  const qPick = pickQHeader(headers, range.s.c);
  if (!qPick.header) return { headers, rows: [], usedQFallback: false };

  const rows = [];
  const totalRows = Math.max(0, range.e.r - headerRow);

  for (let r = headerRow + 1; r <= range.e.r; r += 1) {
    const row = {};
    let hasAny = false;

    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const value = getCellValue(sheet, r, c);
      if (value !== "") hasAny = true;
      row[headers[c - range.s.c]] = value;
    }

    if (!hasAny) continue;

    const qValue = row[qPick.header];
    if (qValue === "" || qValue === null || qValue === undefined) continue;

    row.__Q = String(qValue).trim();
    row.__arquivo = fileName;
    row.__aba = sheetName;
    rows.push(row);

    if ((r - headerRow) % PARSE_YIELD_EVERY === 0) {
      const localProgress = totalRows ? ((r - headerRow) / totalRows) * 100 : 100;
      const overallProgress = progressInfo.baseProgress + (localProgress * progressInfo.weight);
      updateStatus(
        "Processando planilhas pesadas",
        `${fileName} - ${sheetName} (${rows.length.toLocaleString("pt-BR")} linhas validas)`,
        overallProgress
      );
      await yieldToUi();
    }
  }

  return { headers, rows, usedQFallback: qPick.usedFallback };
}

function fileKey(file) {
  return `${file.name}::${file.size}::${file.lastModified}`;
}

function getFileExtension(fileName) {
  const idx = String(fileName || "").lastIndexOf(".");
  if (idx < 0) return "";
  return String(fileName).slice(idx + 1).toLowerCase();
}

function arrayBufferToBinaryString(buffer) {
  const bytes = new Uint8Array(buffer);
  let out = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    const slice = bytes.subarray(i, i + chunk);
    out += String.fromCharCode(...slice);
  }
  return out;
}

function readWorkbookWithFallback(fileName, data) {
  const ext = getFileExtension(fileName);
  const baseOpts = { cellDates: true, cellFormula: true, dense: true, sheetStubs: false };

  try {
    return XLSX.read(data, { ...baseOpts, type: "array" });
  } catch (firstErr) {
    try {
      const binary = arrayBufferToBinaryString(data);
      return XLSX.read(binary, { ...baseOpts, type: "binary" });
    } catch (secondErr) {
      if (ext === "csv") {
        const text = new TextDecoder("utf-8").decode(new Uint8Array(data));
        return XLSX.read(text, { ...baseOpts, type: "string" });
      }
      throw secondErr || firstErr;
    }
  }
}

function resetFileSelection() {
  state.selectedFiles = [];
  els.files.value = "";
  renderSelectedFiles();
}

function renderSelectedFiles() {
  if (!state.selectedFiles.length) {
    els.selectedFilesWrap.hidden = true;
    els.selectedFilesList.innerHTML = "";
    return;
  }

  els.selectedFilesWrap.hidden = false;
  els.selectedFilesList.innerHTML = state.selectedFiles.map((file) => {
    const key = fileKey(file);
    const kb = Math.max(1, Math.round(file.size / 1024));
    return `<li><span title="${escapeHtml(file.name)}">${escapeHtml(file.name)} (${kb} KB)</span><button type="button" class="btn-link" data-file-key="${escapeHtml(key)}">Retirar</button></li>`;
  }).join("");
}

async function readFiles(files) {
  state.rows = [];
  state.headers = [];
  state.headerLetters = [];
  state.formulaIgnored = 0;
  state.qFallbackUsed = false;
  state.sourceFiles = [];
  state.persistenceEnabled = true;

  const globalHeaders = new Set();
  const processedFiles = [];
  const failedFiles = [];
  const totalFiles = files.length || 1;

  for (let fileIndex = 0; fileIndex < files.length; fileIndex += 1) {
    const file = files[fileIndex];
    const baseProgress = (fileIndex / totalFiles) * 100;
    const fileWeight = 100 / totalFiles;

    try {
      updateStatus(
        "Abrindo arquivo",
        `${file.name} (${fileIndex + 1}/${totalFiles})`,
        baseProgress
      );

      const data = await file.arrayBuffer();
      const workbook = readWorkbookWithFallback(file.name, data);

      const totalSheets = workbook.SheetNames.length || 1;
      for (let sheetIndex = 0; sheetIndex < workbook.SheetNames.length; sheetIndex += 1) {
        const sheetName = workbook.SheetNames[sheetIndex];
        const parsed = await parseSheet(file.name, sheetName, workbook.Sheets[sheetName], {
          baseProgress: baseProgress + ((sheetIndex / totalSheets) * fileWeight),
          weight: fileWeight / totalSheets
        });
        parsed.headers.forEach((h) => globalHeaders.add(h));
        state.rows.push(...parsed.rows);
        if (parsed.usedQFallback) state.qFallbackUsed = true;
        await yieldToUi();
      }

      state.sourceFiles.push(file.name);
      processedFiles.push(file.name);
    } catch (err) {
      console.error(`Falha ao ler planilha: ${file.name}`, err);
      failedFiles.push(file.name);
    }
  }

  state.headers = Array.from(globalHeaders);
  state.headerLetters = state.headers.slice(0, 26).map((header, index) => {
    const letter = toColName(index);
    return {
      letter,
      header,
      label: `${letter} - ${header || letter}`
    };
  });

  const estimatedCells = estimateCellCount(state.rows, state.headers);
  if (state.rows.length > STORAGE_ROW_LIMIT || estimatedCells > STORAGE_CELL_LIMIT) {
    state.persistenceEnabled = false;
    clearPersistedState();
  }

  updateStatus(
    "Leitura concluida",
    `${state.rows.length.toLocaleString("pt-BR")} linhas prontas para analise`,
    100
  );

  return { processedFiles, failedFiles };
}

function aggregateByQ(rows, metric) {
  const map = new Map();

  rows.forEach((row) => {
    const q = row.__Q;
    if (!map.has(q)) {
      map.set(q, { q, count: 0, metricSum: 0 });
    }
    const item = map.get(q);
    item.count += 1;

    if (metric) {
      const n = parseNumeric(row[metric]);
      if (n !== null) item.metricSum += n;
    }
  });

  return Array.from(map.values()).sort((a, b) => b.count - a.count);
}

function aggregateByOrigin(rows, metric, mode) {
  const map = new Map();

  rows.forEach((row) => {
    let key = row.__arquivo;
    if (mode === "aba") key = row.__aba;
    if (mode === "arquivo_aba") key = `${row.__arquivo} | ${row.__aba}`;

    if (!map.has(key)) {
      map.set(key, { origin: key, count: 0, metricSum: 0 });
    }

    const item = map.get(key);
    item.count += 1;
    if (metric) {
      const n = parseNumeric(row[metric]);
      if (n !== null) item.metricSum += n;
    }
  });

  return Array.from(map.values()).sort((a, b) => b.count - a.count);
}

function pickNumericColumns(rows, headers) {
  return headers.filter((header) => {
    let numericHits = 0;
    let sample = 0;
    for (const row of rows) {
      if (!(header in row)) continue;
      const n = parseNumeric(row[header]);
      sample += 1;
      if (n !== null) numericHits += 1;
      if (sample >= 60) break;
    }
    return numericHits > 0;
  });
}

function palette(size) {
  const base = [
    "#3f7d38", "#5a9c58", "#7bb36f", "#98c987", "#b8dba6", "#2f6b8a", "#5f90b8",
    "#f0aa33", "#d86a34", "#9a4f9f", "#c2698f", "#3b8d78", "#d15f5f"
  ];
  return Array.from({ length: size }, (_, i) => base[i % base.length]);
}

function compressAggForChart(items, labelKey) {
  if (items.length <= CHART_ITEM_LIMIT) return items;

  const visible = items.slice(0, CHART_ITEM_LIMIT - 1);
  const tail = items.slice(CHART_ITEM_LIMIT - 1);
  const other = tail.reduce((acc, item) => {
    acc.count += item.count || 0;
    acc.metricSum += item.metricSum || 0;
    return acc;
  }, { count: 0, metricSum: 0 });

  visible.push({
    [labelKey]: `Outros (${tail.length})`,
    count: other.count,
    metricSum: other.metricSum
  });

  return visible;
}

function destroyCharts() {
  Object.keys(state.charts).forEach((key) => {
    if (state.charts[key]) {
      state.charts[key].destroy();
      state.charts[key] = null;
    }
  });
}

function renderCharts(agg, metric) {
  destroyCharts();

  if (!agg.length) {
    return;
  }

  const chartAgg = compressAggForChart(agg, "q");
  const labels = chartAgg.map((x) => x.q);
  const counts = chartAgg.map((x) => x.count);
  const sums = chartAgg.map((x) => Number(x.metricSum.toFixed(2)));
  const colors = palette(labels.length);

  state.charts.count = new Chart(document.getElementById("chartCount"), {
    type: "bar",
    data: {
      labels,
      datasets: [{ label: "Linhas por Q", data: counts, backgroundColor: colors }]
    },
    options: { responsive: true, maintainAspectRatio: false }
  });

  state.charts.share = new Chart(document.getElementById("chartShare"), {
    type: "doughnut",
    data: {
      labels,
      datasets: [{ label: "Participacao", data: counts, backgroundColor: colors }]
    },
    options: { responsive: true, maintainAspectRatio: false }
  });

  state.charts.metric = new Chart(document.getElementById("chartMetric"), {
    type: "line",
    data: {
      labels,
      datasets: [{
        label: metric ? `Soma de ${metric} por Q` : "Selecione coluna numerica",
        data: sums,
        borderColor: "#184f28",
        backgroundColor: "rgba(24,79,40,0.12)",
        fill: true,
        tension: 0.25
      }]
    },
    options: { responsive: true, maintainAspectRatio: false }
  });
}

function renderCompare(compareAgg, metric) {
  if (state.charts.compare) {
    state.charts.compare.destroy();
    state.charts.compare = null;
  }

  if (!compareAgg.length) {
    els.compareTable.innerHTML = "<thead><tr><th>Sem comparacao</th></tr></thead><tbody></tbody>";
    return;
  }

  const chartAgg = compressAggForChart(compareAgg, "origin");
  const labels = chartAgg.map((x) => x.origin);
  const counts = chartAgg.map((x) => x.count);
  const sums = chartAgg.map((x) => Number(x.metricSum.toFixed(2)));
  const colors = palette(labels.length);

  state.charts.compare = new Chart(document.getElementById("chartCompare"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Quantidade",
          data: counts,
          backgroundColor: colors
        },
        {
          label: metric ? `Soma (${metric})` : "Soma da metrica",
          data: sums,
          backgroundColor: "rgba(24,79,40,0.24)",
          borderColor: "#184f28",
          borderWidth: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });

  const rowsHtml = compareAgg.slice(0, 500).map((item) => {
    return `<tr><td>${escapeHtml(item.origin)}</td><td>${item.count}</td><td>${item.metricSum.toFixed(2)}</td></tr>`;
  }).join("");

  const truncatedInfo = compareAgg.length > 500
    ? '<tr><td colspan="3">Tabela limitada aos primeiros 500 grupos para manter a tela leve.</td></tr>'
    : "";

  els.compareTable.innerHTML = `
    <thead>
      <tr>
        <th>Origem</th>
        <th>Quantidade</th>
        <th>${metric ? `Soma (${escapeHtml(metric)})` : "Soma da metrica"}</th>
      </tr>
    </thead>
    <tbody>${rowsHtml}${truncatedInfo}</tbody>
  `;
}

function renderTable(rows) {
  const maxRowsValue = String(els.maxRows.value || "100");
  const maxRows = maxRowsValue === "all" ? Infinity : Number(maxRowsValue);
  const shown = Number.isFinite(maxRows) ? rows.slice(0, maxRows) : rows;

  if (!shown.length) {
    els.dataTable.innerHTML = "<thead><tr><th>Sem dados</th></tr></thead><tbody></tbody>";
    const qText = state.qFallbackUsed ? "coluna Q (com fallback automatico)" : "coluna Q";
    els.tableMeta.textContent = `Nenhuma linha valida encontrada na ${qText}.`;
    return;
  }

  const headers = ["__arquivo", "__aba", "__Q", ...state.headers.filter((h) => h !== "")];

  const thead = `<thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${shown.map((row) => {
    const tds = headers.map((h) => `<td>${escapeHtml(row[h] ?? "")}</td>`).join("");
    return `<tr>${tds}</tr>`;
  }).join("")}</tbody>`;

  els.dataTable.innerHTML = `${thead}${tbody}`;
  const qText = state.qFallbackUsed ? "coluna Q (com fallback automatico)" : "coluna Q";
  const storageNote = state.persistenceEnabled ? "" : " Persistencia automatica desativada para evitar travamento.";
  els.tableMeta.textContent = `Mostrando ${shown.length} de ${rows.length} linhas validas pela ${qText}.${storageNote}`;
}

function ensureAllRowsOption() {
  const hasAll = Array.from(els.maxRows.options).some((opt) => opt.value === "all");
  if (hasAll) return;

  const opt = document.createElement("option");
  opt.value = "all";
  opt.textContent = "Todas";
  els.maxRows.insertBefore(opt, els.maxRows.firstChild);
}

function renderKpis(rows, agg) {
  els.kpiFiles.textContent = String(state.sourceFiles.length);
  els.kpiRows.textContent = String(rows.length);
  els.kpiQ.textContent = String(rows.length);
  els.kpiGroups.textContent = String(agg.length);
  els.kpiFormula.textContent = String(state.formulaIgnored);
}

function fillMetricSelect(rows) {
  const previous = els.metricColumn.value;
  const numericCols = pickNumericColumns(rows, state.headers).filter((h) => h !== "__Q");

  els.metricColumn.innerHTML = "";
  const first = document.createElement("option");
  first.value = "";
  first.textContent = numericCols.length ? "Sem metrica (zera soma)" : "Sem colunas numericas";
  els.metricColumn.appendChild(first);

  numericCols.forEach((col) => {
    const opt = document.createElement("option");
    opt.value = col;
    opt.textContent = col;
    els.metricColumn.appendChild(opt);
  });

  if (numericCols.includes(previous)) {
    els.metricColumn.value = previous;
  }
}

function getUniqueValuesLimited(values) {
  const unique = Array.from(new Set(values));
  unique.sort((a, b) => String(a).localeCompare(String(b), "pt-BR", { numeric: true, sensitivity: "base" }));
  return {
    values: unique.slice(0, FILTER_OPTION_LIMIT),
    truncated: unique.length > FILTER_OPTION_LIMIT,
    total: unique.length
  };
}

function fillQFilterSelect(rows) {
  const previous = els.qFilter.value;
  const result = getUniqueValuesLimited(rows.map((r) => r.__Q));

  els.qFilter.innerHTML = "";
  const all = document.createElement("option");
  all.value = "";
  all.textContent = result.truncated
    ? `Todos os valores de Q (primeiros ${FILTER_OPTION_LIMIT} de ${result.total})`
    : "Todos os valores de Q";
  els.qFilter.appendChild(all);

  result.values.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    els.qFilter.appendChild(opt);
  });

  if (result.values.includes(previous)) {
    els.qFilter.value = previous;
  }
}

function fillSheetFilterSelect(rows) {
  const previous = els.sheetFilter.value;
  const result = getUniqueValuesLimited(
    rows
      .map((r) => String(r.__aba || "").trim())
      .filter((value) => value)
  );

  els.sheetFilter.innerHTML = "";
  const all = document.createElement("option");
  all.value = "";
  all.textContent = "Todas as abas";
  els.sheetFilter.appendChild(all);

  result.values.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    els.sheetFilter.appendChild(opt);
  });

  if (result.values.includes(previous)) {
    els.sheetFilter.value = previous;
  }
}

function fillAzColumnSelect() {
  const previous = els.azColumn.value;
  els.azColumn.innerHTML = "";

  const none = document.createElement("option");
  none.value = "";
  none.textContent = "Sem filtro A-Z";
  els.azColumn.appendChild(none);

  state.headerLetters.forEach((item) => {
    const opt = document.createElement("option");
    opt.value = item.letter;
    opt.textContent = item.label;
    els.azColumn.appendChild(opt);
  });

  if (state.headerLetters.some((item) => item.letter === previous)) {
    els.azColumn.value = previous;
  }
}

function getSelectedAzHeader() {
  const selected = els.azColumn.value;
  if (!selected) return "";
  const found = state.headerLetters.find((item) => item.letter === selected);
  return found ? found.header : "";
}

function fillAzValueSelect(rows) {
  const previous = els.azValue.value;
  const selectedHeader = getSelectedAzHeader();
  els.azValue.innerHTML = "";

  const all = document.createElement("option");
  all.value = "";
  all.textContent = "Todos os valores";
  els.azValue.appendChild(all);

  if (!selectedHeader) return;

  const result = getUniqueValuesLimited(
    rows
      .map((row) => row[selectedHeader])
      .filter((value) => value !== "" && value !== null && value !== undefined)
      .map((value) => String(value))
  );

  if (result.truncated) {
    all.textContent = `Todos os valores (primeiros ${FILTER_OPTION_LIMIT} de ${result.total})`;
  }

  result.values.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    els.azValue.appendChild(opt);
  });

  if (result.values.includes(previous)) {
    els.azValue.value = previous;
  }
}

function getRowsAfterQFilter(rows) {
  const sheetRows = getRowsAfterSheetFilter(rows);
  const selectedQ = els.qFilter.value;
  if (!selectedQ) return sheetRows;
  return sheetRows.filter((row) => row.__Q === selectedQ);
}

function getRowsAfterSheetFilter(rows) {
  const selectedSheet = els.sheetFilter.value;
  if (!selectedSheet) return rows;
  return rows.filter((row) => row.__aba === selectedSheet);
}

function getFilteredRows(rows) {
  const qRows = getRowsAfterQFilter(rows);
  const selectedHeader = getSelectedAzHeader();
  const selectedValue = els.azValue.value;

  if (!selectedHeader || !selectedValue) return qRows;
  return qRows.filter((row) => String(row[selectedHeader] ?? "") === selectedValue);
}

function savePersistedState() {
  if (!state.rows.length || !state.persistenceEnabled) {
    if (!state.persistenceEnabled) clearPersistedState();
    return;
  }

  const payload = {
    rows: state.rows,
    headers: state.headers,
    headerLetters: state.headerLetters,
    formulaIgnored: state.formulaIgnored,
    sourceFiles: state.sourceFiles,
    ui: {
      metricColumn: els.metricColumn.value,
      sheetFilter: els.sheetFilter.value,
      qFilter: els.qFilter.value,
      azColumn: els.azColumn.value,
      azValue: els.azValue.value,
      compareBy: els.compareBy.value,
      maxRows: els.maxRows.value
    }
  };

  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  } catch (err) {
    state.persistenceEnabled = false;
    clearPersistedState();
    console.warn("Falha ao salvar estado no localStorage.", err);
  }
}

function clearPersistedState() {
  localStorage.removeItem(STORAGE_KEY);
}

function applyUiFromSaved(savedUi) {
  if (!savedUi) return;

  if (savedUi.maxRows && Array.from(els.maxRows.options).some((o) => o.value === String(savedUi.maxRows))) {
    els.maxRows.value = String(savedUi.maxRows);
  }

  if (savedUi.compareBy && Array.from(els.compareBy.options).some((o) => o.value === String(savedUi.compareBy))) {
    els.compareBy.value = String(savedUi.compareBy);
  }

  if (savedUi.metricColumn && Array.from(els.metricColumn.options).some((o) => o.value === String(savedUi.metricColumn))) {
    els.metricColumn.value = String(savedUi.metricColumn);
  }

  if (savedUi.sheetFilter && Array.from(els.sheetFilter.options).some((o) => o.value === String(savedUi.sheetFilter))) {
    els.sheetFilter.value = String(savedUi.sheetFilter);
  }

  fillQFilterSelect(getRowsAfterSheetFilter(state.rows));

  if (savedUi.qFilter && Array.from(els.qFilter.options).some((o) => o.value === String(savedUi.qFilter))) {
    els.qFilter.value = String(savedUi.qFilter);
  }

  if (savedUi.azColumn && Array.from(els.azColumn.options).some((o) => o.value === String(savedUi.azColumn))) {
    els.azColumn.value = String(savedUi.azColumn);
  }

  fillAzValueSelect(getRowsAfterQFilter(state.rows));

  if (savedUi.azValue && Array.from(els.azValue.options).some((o) => o.value === String(savedUi.azValue))) {
    els.azValue.value = String(savedUi.azValue);
  }
}

function restorePersistedState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed.rows) || !Array.isArray(parsed.headers)) return;

    state.rows = parsed.rows;
    state.headers = parsed.headers;
    state.headerLetters = Array.isArray(parsed.headerLetters) ? parsed.headerLetters : [];
    state.formulaIgnored = Number(parsed.formulaIgnored || 0);
    state.sourceFiles = Array.isArray(parsed.sourceFiles) ? parsed.sourceFiles : [];
    state.persistenceEnabled = true;

    fillMetricSelect(state.rows);
    fillSheetFilterSelect(state.rows);
    fillQFilterSelect(getRowsAfterSheetFilter(state.rows));
    fillAzColumnSelect();
    fillAzValueSelect(getRowsAfterQFilter(state.rows));
    applyUiFromSaved(parsed.ui || {});
    runPipeline();
  } catch (err) {
    console.warn("Falha ao restaurar estado salvo.", err);
  }
}

function runPipeline() {
  const metric = els.metricColumn.value;
  const compareMode = els.compareBy.value;
  const filteredRows = getFilteredRows(state.rows);
  const agg = aggregateByQ(filteredRows, metric);
  const compareAgg = aggregateByOrigin(filteredRows, metric, compareMode);

  renderKpis(filteredRows, agg);
  renderCharts(agg, metric);
  renderCompare(compareAgg, metric);
  renderTable(filteredRows);
  savePersistedState();

  const extras = [];
  if (!state.persistenceEnabled) extras.push("persistencia automatica desligada");
  if (agg.length > CHART_ITEM_LIMIT) extras.push(`graficos resumidos em top ${CHART_ITEM_LIMIT}`);
  const meta = extras.length ? extras.join(" | ") : "analise pronta";
  updateStatus("Painel atualizado", meta, 100, true);
}

function resetUiAfterClear() {
  destroyCharts();
  els.compareTable.innerHTML = "";
  els.dataTable.innerHTML = "";
  els.tableMeta.textContent = "Nenhum dado carregado.";

  els.metricColumn.innerHTML = '<option value="">Selecione apos carregar</option>';
  els.sheetFilter.innerHTML = '<option value="">Todas as abas</option>';
  els.qFilter.innerHTML = '<option value="">Todos os valores de Q</option>';
  els.azColumn.innerHTML = '<option value="">Sem filtro A-Z</option>';
  els.azValue.innerHTML = '<option value="">Todos os valores</option>';
  els.compareBy.value = "arquivo";
  els.maxRows.value = "100";

  els.kpiFiles.textContent = "0";
  els.kpiRows.textContent = "0";
  els.kpiQ.textContent = "0";
  els.kpiGroups.textContent = "0";
  els.kpiFormula.textContent = "0";
  hideStatus();
}

function clearAllData() {
  state.rows = [];
  state.headers = [];
  state.headerLetters = [];
  state.formulaIgnored = 0;
  state.sourceFiles = [];
  state.persistenceEnabled = true;
  resetFileSelection();
  clearPersistedState();
  resetUiAfterClear();
}

function exportSummary() {
  const metric = els.metricColumn.value;
  const filteredRows = getFilteredRows(state.rows);
  const agg = aggregateByQ(filteredRows, metric);
  if (!agg.length) {
    alert("Nao ha dados para exportar.");
    return;
  }

  const lines = [
    ["Q", "quantidade", metric ? `soma_${metric}` : "soma_metrica"].join(";")
  ];

  agg.forEach((x) => {
    lines.push([x.q, x.count, x.metricSum.toFixed(2)].join(";"));
  });

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "resumo_por_coluna_Q.csv";
  a.click();
  URL.revokeObjectURL(a.href);
}

els.files.addEventListener("change", () => {
  const incoming = Array.from(els.files.files || []);
  if (!incoming.length) return;

  const known = new Set(state.selectedFiles.map((file) => fileKey(file)));
  incoming.forEach((file) => {
    const key = fileKey(file);
    if (!known.has(key)) {
      state.selectedFiles.push(file);
      known.add(key);
    }
  });

  els.files.value = "";
  renderSelectedFiles();
});

els.selectedFilesList.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const key = target.getAttribute("data-file-key");
  if (!key) return;

  state.selectedFiles = state.selectedFiles.filter((file) => fileKey(file) !== key);
  renderSelectedFiles();
});

els.clearSelection.addEventListener("click", () => {
  resetFileSelection();
});

els.process.addEventListener("click", async () => {
  const files = state.selectedFiles.slice();
  if (!files.length) {
    alert("Selecione ao menos uma planilha.");
    return;
  }

  els.process.disabled = true;
  els.process.textContent = "Processando...";
  updateStatus("Preparando leitura", `${files.length} arquivo(s) na fila`, 0);

  try {
    const { processedFiles, failedFiles } = await readFiles(files);
    if (!processedFiles.length) {
      alert("Nao foi possivel ler nenhuma planilha selecionada.");
      return;
    }

    fillMetricSelect(state.rows);
    fillSheetFilterSelect(state.rows);
    fillQFilterSelect(getRowsAfterSheetFilter(state.rows));
    fillAzColumnSelect();
    fillAzValueSelect(getRowsAfterQFilter(state.rows));
    runPipeline();

      if (failedFiles.length) {
        alert(`Algumas planilhas nao puderam ser lidas: ${failedFiles.join(", ")}`);
      }
    } catch (err) {
      console.error(err);
      updateStatus("Falha ao processar", "Verifique o formato das planilhas e tente novamente.", 0);
      alert("Erro ao processar planilhas. Verifique o formato dos arquivos.");
    } finally {
      els.process.disabled = false;
    els.process.textContent = "Gerar graficos";
  }
});

els.metricColumn.addEventListener("change", () => {
  if (!state.rows.length) return;
  runPipeline();
});

els.sheetFilter.addEventListener("change", () => {
  if (!state.rows.length) return;
  fillQFilterSelect(getRowsAfterSheetFilter(state.rows));
  fillAzValueSelect(getRowsAfterQFilter(state.rows));
  runPipeline();
});

els.qFilter.addEventListener("change", () => {
  if (!state.rows.length) return;
  fillAzValueSelect(getRowsAfterQFilter(state.rows));
  runPipeline();
});

els.azColumn.addEventListener("change", () => {
  if (!state.rows.length) return;
  fillAzValueSelect(getRowsAfterQFilter(state.rows));
  runPipeline();
});

[els.azValue, els.compareBy].forEach((node) => {
  node.addEventListener("change", () => {
    if (!state.rows.length) return;
    runPipeline();
  });
});

els.maxRows.addEventListener("change", () => {
  if (!state.rows.length) return;
  renderTable(getFilteredRows(state.rows));
  savePersistedState();
});

els.download.addEventListener("click", exportSummary);
els.clearData.addEventListener("click", clearAllData);

renderSelectedFiles();
ensureAllRowsOption();
restorePersistedState();
