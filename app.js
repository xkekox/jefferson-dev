const state = {
  stockRows: [],
  salesRows: [],
  stockLoaded: false,
  salesLoaded: false,
  processed: [],
  months: [],
  monitoredTerms: [],
  productConfigs: {}
};

const refs = {
  stockFile: document.getElementById("stockFile"),
  salesFile: document.getElementById("salesFile"),
  stockText: document.getElementById("stockText"),
  salesText: document.getElementById("salesText"),
  stockStatus: document.getElementById("stockStatus"),
  salesStatus: document.getElementById("salesStatus"),
  stockExampleButton: document.getElementById("stockExampleButton"),
  salesExampleButton: document.getElementById("salesExampleButton"),
  processStockButton: document.getElementById("processStockButton"),
  processSalesButton: document.getElementById("processSalesButton"),
  processButton: document.getElementById("processButton"),
  clearButton: document.getElementById("clearButton"),
  messageBox: document.getElementById("messageBox"),
  monitoredInput: document.getElementById("monitoredInput"),
  tableHead: document.getElementById("tableHead"),
  tableBody: document.getElementById("tableBody"),
  productCount: document.getElementById("productCount"),
  totalStock: document.getElementById("totalStock"),
  currentMonthSales: document.getElementById("currentMonthSales"),
  currentMonthProjection: document.getElementById("currentMonthProjection"),
  monthsLabel: document.getElementById("monthsLabel"),
  searchInput: document.getElementById("searchInput"),
  riskFilter: document.getElementById("riskFilter"),
  sortSelect: document.getElementById("sortSelect")
};

const STOCK_ALIASES = {
  code: ["codigo", "cod", "codigo produto", "produto codigo", "sku", "item"],
  name: ["nome", "descricao", "descrição", "produto", "nome produto", "descricao produto"],
  stock: ["estoque", "estoque atual", "saldo", "saldo estoque", "qtd estoque", "quantidade estoque"]
};

const SALES_ALIASES = {
  code: ["codigo", "cod", "codigo produto", "produto codigo", "sku", "item"],
  name: ["nome", "descricao", "descrição", "produto", "nome produto", "descricao produto"],
  date: ["data", "data venda", "emissao", "data emissao", "data faturamento"],
  quantity: ["quantidade", "qtd", "qtde", "vendido", "qtd vendida", "quantidade vendida"]
};

function normalizeKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .trim()
    .toLowerCase();
}

function detectDelimiter(text) {
  const firstLine = text.split(/\r?\n/).find((line) => line.trim());
  if (!firstLine) {
    return ",";
  }

  const candidates = [",", ";", "\t", "|"];
  let best = ",";
  let bestCount = -1;

  for (const candidate of candidates) {
    const count = firstLine.split(candidate).length;
    if (count > bestCount) {
      best = candidate;
      bestCount = count;
    }
  }

  return best;
}

function parseDelimitedText(text) {
  const trimmed = String(text || "").replace(/^\uFEFF/, "").trim();
  if (!trimmed) {
    return [];
  }

  const delimiter = detectDelimiter(trimmed);
  const lines = trimmed
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length < 2) {
    return [];
  }

  const headers = lines[0].split(delimiter).map((value) => value.trim());

  return lines.slice(1).map((line) => {
    const values = line.split(delimiter).map((value) => value.trim());
    return headers.reduce((accumulator, header, index) => {
      accumulator[header] = values[index] ?? "";
      return accumulator;
    }, {});
  });
}

function findColumn(row, aliases) {
  const entries = Object.keys(row);
  const normalizedAliases = aliases.map((alias) => normalizeKey(alias));

  for (const key of entries) {
    const normalized = normalizeKey(key);
    if (normalizedAliases.includes(normalized)) {
      return key;
    }
    if (normalizedAliases.some((alias) => normalized.includes(alias) || alias.includes(normalized))) {
      return key;
    }
  }

  return null;
}

function parseNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  const raw = String(value || "").trim();
  if (!raw) {
    return 0;
  }

  let normalized = raw.replace(/[^\d,.-]/g, "");

  if (normalized.includes(".") && normalized.includes(",")) {
    const lastDot = normalized.lastIndexOf(".");
    const lastComma = normalized.lastIndexOf(",");
    if (lastComma > lastDot) {
      normalized = normalized.replace(/\./g, "").replace(",", ".");
    } else {
      normalized = normalized.replace(/,/g, "");
    }
  } else if (normalized.includes(",")) {
    const commaMatches = normalized.match(/,/g) || [];
    normalized =
      commaMatches.length > 1 || /,\d{3}$/.test(normalized)
        ? normalized.replace(/,/g, "")
        : normalized.replace(",", ".");
  } else if (normalized.includes(".")) {
    const dotMatches = normalized.match(/\./g) || [];
    if (dotMatches.length > 1 || /\.\d{3}$/.test(normalized)) {
      normalized = normalized.replace(/\./g, "");
    }
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseDate(value) {
  if (!value) {
    return null;
  }

  if (typeof value === "number" && typeof XLSX !== "undefined") {
    const excelDate = XLSX.SSF.parse_date_code(value);
    if (excelDate) {
      return new Date(excelDate.y, excelDate.m - 1, excelDate.d, 12, 0, 0);
    }
  }

  const raw = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return new Date(`${raw}T12:00:00`);
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) {
    const [day, month, year] = raw.split("/");
    return new Date(`${year}-${month}-${day}T12:00:00`);
  }

  const fallback = new Date(raw);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function monthKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function monthLabel(key) {
  const [year, month] = key.split("-").map(Number);
  return new Date(year, month - 1, 1).toLocaleDateString("pt-BR", {
    month: "long",
    year: "numeric"
  });
}

function getLastThreeMonthKeys(referenceDate) {
  return [2, 1, 0].map((offset) => {
    const date = new Date(referenceDate.getFullYear(), referenceDate.getMonth() - offset, 1);
    return monthKey(date);
  });
}

function countBusinessDaysInMonth(date) {
  const year = date.getFullYear();
  const month = date.getMonth();
  const lastDay = new Date(year, month + 1, 0).getDate();
  let total = 0;

  for (let day = 1; day <= lastDay; day += 1) {
    const current = new Date(year, month, day);
    const weekDay = current.getDay();
    if (weekDay !== 0 && weekDay !== 6) {
      total += 1;
    }
  }

  return total;
}

function countBusinessDaysElapsed(date) {
  const year = date.getFullYear();
  const month = date.getMonth();
  const currentDay = date.getDate();
  let total = 0;

  for (let day = 1; day <= currentDay; day += 1) {
    const current = new Date(year, month, day);
    const weekDay = current.getDay();
    if (weekDay !== 0 && weekDay !== 6) {
      total += 1;
    }
  }

  return Math.max(total, 1);
}

function formatInteger(value) {
  return new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 0 }).format(value);
}

function formatDecimal(value) {
  return new Intl.NumberFormat("pt-BR", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1
  }).format(value);
}

function getMonitoredTerms() {
  return refs.monitoredInput.value
    .split(/\r?\n/)
    .map((value) => value.trim())
    .filter(Boolean)
    .map((value) => value.toLowerCase());
}

function matchesMonitoredTerms(product, terms) {
  if (!terms.length) {
    return true;
  }

  const haystack = `${product.code} ${product.name}`.toLowerCase();
  return terms.some((term) => haystack.includes(term));
}

function getProductConfig(code) {
  return (
    state.productConfigs[code] || {
      leadDays: "",
      ignoreAlert: false,
      ignoreReason: ""
    }
  );
}

function updateProductConfig(code, patch) {
  state.productConfigs[code] = {
    ...getProductConfig(code),
    ...patch
  };
}

async function readTextFile(file) {
  return file.text();
}

async function readSpreadsheetFile(file) {
  if (typeof XLSX === "undefined") {
    throw new Error("Leitor de Excel indisponivel no momento. Recarregue a pagina e tente novamente.");
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];

  return XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: true
  });
}

async function readUploadedFile(file) {
  const extension = (file.name.split(".").pop() || "").toLowerCase();
  if (["xlsx", "xls"].includes(extension)) {
    return readSpreadsheetFile(file);
  }
  const text = await readTextFile(file);
  return parseDelimitedText(text);
}

async function loadDataset({ fileInput, textArea, aliases, kind }) {
  const file = fileInput.files?.[0];
  const pasted = textArea.value.trim();

  if (!file && !pasted) {
    return null;
  }

  const rows = file ? await readUploadedFile(file) : parseDelimitedText(pasted);

  if (!rows.length) {
    throw new Error(
      `Nao foi possivel ler o relatorio de ${kind}. Verifique se o arquivo tem cabecalho e ao menos uma linha de dados.`
    );
  }

  const firstRow = rows[0];
  const headers = Object.keys(firstRow);
  const mappedColumns = Object.fromEntries(
    Object.entries(aliases).map(([field, fieldAliases], index) => {
      const found = findColumn(firstRow, fieldAliases);
      return [field, found || headers[index] || null];
    })
  );

  const missing = Object.entries(mappedColumns)
    .filter(([, column]) => !column)
    .map(([field]) => field);

  if (missing.length) {
    throw new Error(`Colunas obrigatorias ausentes no relatorio de ${kind}: ${missing.join(", ")}.`);
  }

  return rows.map((row) => {
    const normalized = {};
    for (const [field, column] of Object.entries(mappedColumns)) {
      normalized[field] = row[column];
    }
    return normalized;
  });
}

function setMessage(type, text) {
  refs.messageBox.className = `message-box ${type}`;
  refs.messageBox.textContent = text;
}

function setStatus(element, label, loaded) {
  element.textContent = label;
  element.classList.toggle("is-ready", loaded);
}

async function loadExample(kind) {
  const fileMap = {
    estoque: "./examples/relatorio-estoque-exemplo.csv",
    vendas: "./examples/relatorio-vendas-exemplo.csv"
  };

  const target = fileMap[kind];
  const response = await fetch(target);
  if (!response.ok) {
    throw new Error(`Nao foi possivel carregar o exemplo de ${kind}.`);
  }

  return response.text();
}

function processData() {
  const referenceDate = new Date();
  const monthKeys = getLastThreeMonthKeys(referenceDate);
  const currentMonthKey = monthKeys[2];
  const businessDaysElapsed = countBusinessDaysElapsed(referenceDate);
  const businessDaysTotal = countBusinessDaysInMonth(referenceDate);
  const monitoredTerms = getMonitoredTerms();
  const products = new Map();

  for (const item of state.stockRows) {
    const code = String(item.code || "").trim();
    if (!code) {
      continue;
    }

    const existing = products.get(code) || {
      code,
      name: String(item.name || "").trim(),
      stock: 0,
      monthlySales: Object.fromEntries(monthKeys.map((key) => [key, 0]))
    };

    existing.name = existing.name || String(item.name || "").trim();
    existing.stock = parseNumber(item.stock);
    products.set(code, existing);
  }

  for (const sale of state.salesRows) {
    const code = String(sale.code || "").trim();
    const date = parseDate(sale.date);

    if (!code || !date) {
      continue;
    }

    const key = monthKey(date);
    if (!monthKeys.includes(key)) {
      continue;
    }

    const existing = products.get(code) || {
      code,
      name: String(sale.name || "").trim(),
      stock: 0,
      monthlySales: Object.fromEntries(monthKeys.map((month) => [month, 0]))
    };

    existing.name = existing.name || String(sale.name || "").trim();
    existing.monthlySales[key] += parseNumber(sale.quantity);
    products.set(code, existing);
  }

  state.months = monthKeys;
  state.monitoredTerms = monitoredTerms;
  state.processed = Array.from(products.values())
    .filter((product) => matchesMonitoredTerms(product, monitoredTerms))
    .map((product) => {
      const config = getProductConfig(product.code);
      const currentSales = product.monthlySales[currentMonthKey] || 0;
      const projection = (currentSales / businessDaysElapsed) * businessDaysTotal;
      const coverageDays =
        projection > 0 ? (product.stock / projection) * businessDaysTotal : Number.POSITIVE_INFINITY;
      const riskLevel = product.stock === 0 || product.stock < projection ? "risk" : "healthy";
      const leadDays = parseNumber(config.leadDays);
      const reorderAlert =
        leadDays > 0 &&
        coverageDays !== Number.POSITIVE_INFINITY &&
        coverageDays < leadDays &&
        !config.ignoreAlert;

      return {
        ...product,
        currentSales,
        projection,
        coverageDays,
        riskLevel,
        leadDays,
        ignoreAlert: Boolean(config.ignoreAlert),
        ignoreReason: config.ignoreReason || "",
        reorderAlert
      };
    });
}

function updateSummary() {
  const totalProducts = state.processed.length;
  const totalStock = state.processed.reduce((sum, item) => sum + item.stock, 0);
  const totalCurrentSales = state.processed.reduce((sum, item) => sum + item.currentSales, 0);
  const totalProjection = state.processed.reduce((sum, item) => sum + item.projection, 0);

  refs.productCount.textContent = formatInteger(totalProducts);
  refs.totalStock.textContent = formatInteger(totalStock);
  refs.currentMonthSales.textContent = formatInteger(totalCurrentSales);
  refs.currentMonthProjection.textContent = formatInteger(totalProjection);
  refs.monthsLabel.textContent = state.months.map((key) => monthLabel(key)).join(" | ");
}

function buildTableHead() {
  refs.tableHead.innerHTML = `
    <tr>
      <th>Codigo</th>
      <th>Produto</th>
      <th>Estoque atual</th>
      <th>${monthLabel(state.months[0])}</th>
      <th>${monthLabel(state.months[1])}</th>
      <th>${monthLabel(state.months[2])}</th>
      <th>Projecao ${monthLabel(state.months[2])}</th>
      <th>Cobertura em dias</th>
      <th>Prazo reposicao</th>
      <th>Risco</th>
      <th>Alerta compra</th>
      <th>Ignorar alerta</th>
      <th>Justificativa</th>
    </tr>
  `;
}

function buildRow(item) {
  const coverage =
    item.coverageDays === Number.POSITIVE_INFINITY ? "Sem consumo" : formatDecimal(item.coverageDays);
  const riskText = item.riskLevel === "risk" ? "Reposicao urgente" : "Cobertura suficiente";
  const riskClass = item.riskLevel === "risk" ? "risk-cell" : "ok-cell";
  const reorderText = item.reorderAlert
    ? "Comprar agora"
    : item.ignoreAlert
      ? "Ignorado"
      : "Dentro do prazo";
  const reorderClass = item.reorderAlert ? "risk-cell" : "ok-cell";

  return `
    <tr>
      <td>${item.code}</td>
      <td>${item.name || "-"}</td>
      <td>${formatInteger(item.stock)}</td>
      <td>${formatInteger(item.monthlySales[state.months[0]] || 0)}</td>
      <td>${formatInteger(item.monthlySales[state.months[1]] || 0)}</td>
      <td>${formatInteger(item.monthlySales[state.months[2]] || 0)}</td>
      <td>${formatInteger(item.projection)}</td>
      <td>${coverage}</td>
      <td>
        <input
          class="table-input"
          type="number"
          min="0"
          step="1"
          data-config-field="leadDays"
          data-product-code="${item.code}"
          value="${item.leadDays || ""}"
          placeholder="Dias"
        />
      </td>
      <td class="${riskClass}">${riskText}</td>
      <td class="${reorderClass}">${reorderText}</td>
      <td>
        <label class="inline-toggle">
          <input
            type="checkbox"
            data-config-field="ignoreAlert"
            data-product-code="${item.code}"
            ${item.ignoreAlert ? "checked" : ""}
          />
          <span>Ignorar</span>
        </label>
      </td>
      <td>
        <input
          class="table-input"
          type="text"
          data-config-field="ignoreReason"
          data-product-code="${item.code}"
          value="${item.ignoreReason || ""}"
          placeholder="Ex.: saindo de linha"
        />
      </td>
    </tr>
  `;
}

function filteredRows() {
  const search = refs.searchInput.value.trim().toLowerCase();
  const risk = refs.riskFilter.value;
  const sort = refs.sortSelect.value;

  let rows = [...state.processed];

  if (search) {
    rows = rows.filter((item) => {
      const haystack = `${item.code} ${item.name}`.toLowerCase();
      return haystack.includes(search);
    });
  }

  if (risk !== "all") {
    rows = rows.filter((item) => item.riskLevel === risk || (risk === "risk" && item.reorderAlert));
  }

  rows.sort((left, right) => {
    if (sort === "name") {
      return left.name.localeCompare(right.name, "pt-BR");
    }
    if (sort === "stock") {
      return left.stock - right.stock;
    }
    if (sort === "coverage") {
      return left.coverageDays - right.coverageDays;
    }
    return right.projection - left.projection;
  });

  return rows;
}

function renderTable() {
  if (!state.processed.length) {
    refs.tableHead.innerHTML = "";
    refs.tableBody.innerHTML = `
      <tr>
        <td colspan="13" class="empty-state">Nenhum produto consolidado para exibir.</td>
      </tr>
    `;
    return;
  }

  buildTableHead();
  const rows = filteredRows();

  refs.tableBody.innerHTML = rows.length
    ? rows.map(buildRow).join("")
    : `
      <tr>
        <td colspan="13" class="empty-state">Nenhum item encontrado com os filtros atuais.</td>
      </tr>
    `;
}

function resetState() {
  state.stockRows = [];
  state.salesRows = [];
  state.stockLoaded = false;
  state.salesLoaded = false;
  state.processed = [];
  state.months = [];
  state.monitoredTerms = [];
  state.productConfigs = {};

  refs.stockFile.value = "";
  refs.salesFile.value = "";
  refs.stockText.value = "";
  refs.salesText.value = "";
  refs.monitoredInput.value = "";
  refs.searchInput.value = "";
  refs.riskFilter.value = "all";
  refs.sortSelect.value = "projection";

  setStatus(refs.stockStatus, "Aguardando", false);
  setStatus(refs.salesStatus, "Aguardando", false);
  setMessage("info", "Carregue um ou ambos os relatorios para montar a visao operacional de Zain na Pichau.");

  refs.productCount.textContent = "0";
  refs.totalStock.textContent = "0";
  refs.currentMonthSales.textContent = "0";
  refs.currentMonthProjection.textContent = "0";
  refs.monthsLabel.textContent = "Ultimos 3 meses ainda nao carregados.";
  renderTable();
}

async function handleProcess(mode) {
  try {
    setMessage("info", "Lendo relatorios do ERP do cliente Zain para a operacao de gabinetes...");

    const stockRows =
      mode === "sales"
        ? null
        : await loadDataset({
            fileInput: refs.stockFile,
            textArea: refs.stockText,
            aliases: STOCK_ALIASES,
            kind: "estoque"
          });

    const salesRows =
      mode === "stock"
        ? null
        : await loadDataset({
            fileInput: refs.salesFile,
            textArea: refs.salesText,
            aliases: SALES_ALIASES,
            kind: "vendas"
          });

    if (!stockRows && !salesRows) {
      throw new Error(
        mode === "stock"
          ? "Carregue um relatorio de estoque para processar."
          : mode === "sales"
            ? "Carregue um relatorio de vendas para processar."
            : "Carregue pelo menos um relatorio para processar."
      );
    }

    state.stockRows = stockRows || [];
    state.salesRows = salesRows || [];
    state.stockLoaded = Boolean(stockRows);
    state.salesLoaded = Boolean(salesRows);

    setStatus(refs.stockStatus, stockRows ? "Carregado" : "Nao carregado", Boolean(stockRows));
    setStatus(refs.salesStatus, salesRows ? "Carregado" : "Nao carregado", Boolean(salesRows));

    processData();
    updateSummary();
    renderTable();

    if (stockRows && salesRows) {
      setMessage("success", "Relatorios de estoque e vendas processados com sucesso.");
    } else if (stockRows) {
      setMessage("success", "Relatorio de estoque processado com sucesso.");
    } else {
      setMessage("success", "Relatorio de vendas processado com sucesso.");
    }
  } catch (error) {
    setMessage("error", error.message || "Falha ao processar os relatorios.");
  }
}

function handleProcessStock() {
  return handleProcess("stock");
}

function handleProcessSales() {
  return handleProcess("sales");
}

function handleProcessBoth() {
  return handleProcess("both");
}

async function handleExample(kind) {
  try {
    const content = await loadExample(kind);

    if (kind === "estoque") {
      refs.stockText.value = content;
      refs.stockFile.value = "";
      setStatus(refs.stockStatus, "Exemplo carregado", true);
    } else {
      refs.salesText.value = content;
      refs.salesFile.value = "";
      setStatus(refs.salesStatus, "Exemplo carregado", true);
    }

    setMessage("info", `Exemplo de ${kind} carregado. Agora processe os relatorios.`);
  } catch (error) {
    setMessage("error", error.message || `Falha ao carregar exemplo de ${kind}.`);
  }
}

function reprocessIfLoaded() {
  if (state.stockRows.length || state.salesRows.length) {
    processData();
    updateSummary();
    renderTable();
  }
}

function handleConfigChange(event) {
  const target = event.target;
  const code = target.dataset.productCode;
  const field = target.dataset.configField;

  if (!code || !field) {
    return;
  }

  const value = target.type === "checkbox" ? target.checked : target.value;
  updateProductConfig(code, { [field]: value });
  reprocessIfLoaded();
}

refs.processButton.addEventListener("click", handleProcessBoth);
refs.processStockButton.addEventListener("click", handleProcessStock);
refs.processSalesButton.addEventListener("click", handleProcessSales);
refs.clearButton.addEventListener("click", resetState);
refs.searchInput.addEventListener("input", renderTable);
refs.riskFilter.addEventListener("change", renderTable);
refs.sortSelect.addEventListener("change", renderTable);
refs.monitoredInput.addEventListener("input", reprocessIfLoaded);
refs.tableBody.addEventListener("change", handleConfigChange);
refs.tableBody.addEventListener("input", (event) => {
  if (event.target.dataset.configField === "ignoreReason") {
    handleConfigChange(event);
  }
});
refs.stockExampleButton.addEventListener("click", () => handleExample("estoque"));
refs.salesExampleButton.addEventListener("click", () => handleExample("vendas"));
refs.stockFile.addEventListener("change", () => {
  if (refs.stockFile.files?.[0]) {
    setStatus(refs.stockStatus, refs.stockFile.files[0].name, true);
    refs.stockText.value = "";
  }
});
refs.salesFile.addEventListener("change", () => {
  if (refs.salesFile.files?.[0]) {
    setStatus(refs.salesStatus, refs.salesFile.files[0].name, true);
    refs.salesText.value = "";
  }
});

resetState();
