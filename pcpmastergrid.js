const Q_INDEX = 16;
const STORAGE_KEY = "pcpmastergrid_saved_v1";

const state = {
  rows: [],
  headers: [],
  headerLetters: [],
  formulaIgnored: 0,
  sourceFiles: [],
  selectedFiles: [],
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
  compareTable: document.getElementById("compareTable")
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

function ensureUniqueHeaders(headers) {
  const map = new Map();
  return headers.map((raw, index) => {
    const base = String(raw || "").trim() || toColName(index);
    const count = map.get(base) || 0;
    map.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

function getCellValue(sheet, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = sheet[addr];
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

function parseSheet(fileName, sheetName, sheet) {
  if (!sheet["!ref"]) return { headers: [], rows: [] };

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headerRow = range.s.r;

  const rawHeaders = [];
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    rawHeaders.push(getCellValue(sheet, headerRow, c));
  }
  const headers = ensureUniqueHeaders(rawHeaders);

  const rows = [];
  for (let r = headerRow + 1; r <= range.e.r; r += 1) {
    const row = {};
    let hasAny = false;

    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const value = getCellValue(sheet, r, c);
      if (value !== "") hasAny = true;
      row[headers[c - range.s.c]] = value;
    }

    if (!hasAny) continue;

    const qHeader = headers[Q_INDEX - range.s.c];
    const qValue = qHeader ? row[qHeader] : "";
    if (qValue === "" || qValue === null || qValue === undefined) continue;

    row.__Q = String(qValue).trim();
    row.__arquivo = fileName;
    row.__aba = sheetName;
    rows.push(row);
  }

  return { headers, rows };
}

function fileKey(file) {
  return `${file.name}::${file.size}::${file.lastModified}`;
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
    return `<li><span title="${file.name}">${file.name} (${kb} KB)</span><button type="button" class="btn-link" data-file-key="${key}">Retirar</button></li>`;
  }).join("");
}

async function readFiles(files) {
  state.rows = [];
  state.headers = [];
  state.headerLetters = [];
  state.formulaIgnored = 0;
  state.sourceFiles = [];

  const globalHeaders = new Set();

  for (const file of files) {
    state.sourceFiles.push(file.name);
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true, cellFormula: true });

    for (const sheetName of workbook.SheetNames) {
      const parsed = parseSheet(file.name, sheetName, workbook.Sheets[sheetName]);
      parsed.headers.forEach((h) => globalHeaders.add(h));
      state.rows.push(...parsed.rows);
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

  const labels = agg.map((x) => x.q);
  const counts = agg.map((x) => x.count);
  const sums = agg.map((x) => Number(x.metricSum.toFixed(2)));
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
  const labels = compareAgg.map((x) => x.origin);
  const counts = compareAgg.map((x) => x.count);
  const sums = compareAgg.map((x) => Number(x.metricSum.toFixed(2)));
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

  if (!compareAgg.length) {
    els.compareTable.innerHTML = "<thead><tr><th>Sem comparacao</th></tr></thead><tbody></tbody>";
    return;
  }

  const rowsHtml = compareAgg.map((item) => {
    return `<tr><td>${item.origin}</td><td>${item.count}</td><td>${item.metricSum.toFixed(2)}</td></tr>`;
  }).join("");

  els.compareTable.innerHTML = `
    <thead>
      <tr>
        <th>Origem</th>
        <th>Quantidade</th>
        <th>${metric ? `Soma (${metric})` : "Soma da metrica"}</th>
      </tr>
    </thead>
    <tbody>${rowsHtml}</tbody>
  `;
}

function renderTable(rows) {
  const maxRows = Number(els.maxRows.value);
  const shown = rows.slice(0, maxRows);

  if (!shown.length) {
    els.dataTable.innerHTML = "<thead><tr><th>Sem dados</th></tr></thead><tbody></tbody>";
    els.tableMeta.textContent = "Nenhuma linha valida encontrada na coluna Q.";
    return;
  }

  const headers = ["__arquivo", "__aba", "__Q", ...state.headers.filter((h) => h !== "")];

  const thead = `<thead><tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${shown.map((row) => {
    const tds = headers.map((h) => `<td>${row[h] ?? ""}</td>`).join("");
    return `<tr>${tds}</tr>`;
  }).join("")}</tbody>`;

  els.dataTable.innerHTML = `${thead}${tbody}`;
  els.tableMeta.textContent = `Mostrando ${shown.length} de ${rows.length} linhas validas pela coluna Q.`;
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

function fillQFilterSelect(rows) {
  const previous = els.qFilter.value;
  const values = Array.from(new Set(rows.map((r) => r.__Q))).sort((a, b) =>
    String(a).localeCompare(String(b), "pt-BR", { numeric: true, sensitivity: "base" })
  );

  els.qFilter.innerHTML = "";
  const all = document.createElement("option");
  all.value = "";
  all.textContent = "Todos os valores de Q";
  els.qFilter.appendChild(all);

  values.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    els.qFilter.appendChild(opt);
  });

  if (values.includes(previous)) {
    els.qFilter.value = previous;
  }
}

function fillSheetFilterSelect(rows) {
  const previous = els.sheetFilter.value;
  const values = Array.from(
    new Set(
      rows
        .map((r) => String(r.__aba || "").trim())
        .filter((value) => value)
    )
  ).sort((a, b) => a.localeCompare(b, "pt-BR", { numeric: true, sensitivity: "base" }));

  els.sheetFilter.innerHTML = "";
  const all = document.createElement("option");
  all.value = "";
  all.textContent = "Todas as abas";
  els.sheetFilter.appendChild(all);

  values.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    els.sheetFilter.appendChild(opt);
  });

  if (values.includes(previous)) {
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

  const values = Array.from(
    new Set(
      rows
        .map((row) => row[selectedHeader])
        .filter((value) => value !== "" && value !== null && value !== undefined)
        .map((value) => String(value))
    )
  ).sort((a, b) => a.localeCompare(b, "pt-BR", { numeric: true, sensitivity: "base" }));

  values.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    els.azValue.appendChild(opt);
  });

  if (values.includes(previous)) {
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
  if (!state.rows.length) return;

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
}

function clearAllData() {
  state.rows = [];
  state.headers = [];
  state.headerLetters = [];
  state.formulaIgnored = 0;
  state.sourceFiles = [];
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

  try {
    await readFiles(files);
    fillMetricSelect(state.rows);
    fillSheetFilterSelect(state.rows);
    fillQFilterSelect(getRowsAfterSheetFilter(state.rows));
    fillAzColumnSelect();
    fillAzValueSelect(getRowsAfterQFilter(state.rows));
    runPipeline();
  } catch (err) {
    console.error(err);
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
restorePersistedState();
