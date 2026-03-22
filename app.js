const state = {
  stockRows: [],
  salesRows: [],
  stockLoaded: false,
  salesLoaded: false,
  processed: [],
  months: [],
  monitoredTerms: [],
  persistedProducts: [],
  productConfigs: {},
  savedOrders: [],
  importHistory: [],
  orderDraft: null
};

const refs = {
  tabButtons: Array.from(document.querySelectorAll("[data-tab-target]")),
  tabPanels: Array.from(document.querySelectorAll("[data-tab-panel]")),
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
  monitoredList: document.getElementById("monitoredList"),
  monitorEmpty: document.getElementById("monitorEmpty"),
  monitoredCount: document.getElementById("monitoredCount"),
  activeAlertCount: document.getElementById("activeAlertCount"),
  ignoredAlertCount: document.getElementById("ignoredAlertCount"),
  importHistoryBody: document.getElementById("importHistoryBody"),
  tableHead: document.getElementById("tableHead"),
  tableBody: document.getElementById("tableBody"),
  productCount: document.getElementById("productCount"),
  totalStock: document.getElementById("totalStock"),
  currentMonthSales: document.getElementById("currentMonthSales"),
  currentMonthProjection: document.getElementById("currentMonthProjection"),
  dashboardStockStatus: document.getElementById("dashboardStockStatus"),
  dashboardSalesStatus: document.getElementById("dashboardSalesStatus"),
  dashboardProductBaseCount: document.getElementById("dashboardProductBaseCount"),
  supabaseUrlInput: document.getElementById("supabaseUrlInput"),
  supabaseAnonKeyInput: document.getElementById("supabaseAnonKeyInput"),
  saveSupabaseConfigButton: document.getElementById("saveSupabaseConfigButton"),
  clearSupabaseConfigButton: document.getElementById("clearSupabaseConfigButton"),
  monthsLabel: document.getElementById("monthsLabel"),
  searchInput: document.getElementById("searchInput"),
  riskFilter: document.getElementById("riskFilter"),
  sortSelect: document.getElementById("sortSelect"),
  orderPi: document.getElementById("orderPi"),
  orderOc: document.getElementById("orderOc"),
  orderPch: document.getElementById("orderPch"),
  orderSplitCount: document.getElementById("orderSplitCount"),
  orderEntryPercent: document.getElementById("orderEntryPercent"),
  orderPaymentType: document.getElementById("orderPaymentType"),
  orderPaid: document.getElementById("orderPaid"),
  orderPaymentDate: document.getElementById("orderPaymentDate"),
  orderSplitBadge: document.getElementById("orderSplitBadge"),
  orderSplitPreview: document.getElementById("orderSplitPreview"),
  orderSplitHint: document.getElementById("orderSplitHint"),
  orderTotalValue: document.getElementById("orderTotalValue"),
  orderEntryValue: document.getElementById("orderEntryValue"),
  orderBalanceValue: document.getElementById("orderBalanceValue"),
  orderItemsBody: document.getElementById("orderItemsBody"),
  addOrderItemButton: document.getElementById("addOrderItemButton"),
  saveOrderButton: document.getElementById("saveOrderButton"),
  clearOrderButton: document.getElementById("clearOrderButton"),
  ordersTableHead: document.getElementById("ordersTableHead"),
  ordersTableBody: document.getElementById("ordersTableBody")
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

STOCK_ALIASES.code.push("cod produto", "codigo sku", "sku produto");
STOCK_ALIASES.name.push("descricao item", "produto descricao");
STOCK_ALIASES.stock.push("saldo atual", "qtd disponivel", "quantidade disponivel", "estoque disponivel");

SALES_ALIASES.code.push("cod produto", "codigo sku", "sku produto");
SALES_ALIASES.name.push("descricao item", "produto descricao");
SALES_ALIASES.date.push("dt faturamento", "data movimento");
SALES_ALIASES.quantity.push("quantidade faturada", "qtd faturada");

const STORAGE_KEYS = {
  configs: "jefferson-dev-product-configs",
  products: "jefferson-dev-products",
  orders: "jefferson-dev-orders",
  imports: "jefferson-dev-imports",
  processedSnapshot: "jefferson-dev-processed-snapshot",
  orderDraft: "jefferson-dev-order-draft",
  activeTab: "jefferson-dev-active-tab",
  supabaseConfig: "jefferson-dev-supabase-config"
};

const PERSISTENCE = {
  client: null,
  mode: "local",
  tables: {
    products: "zain_products",
    productConfigs: "zain_product_configs",
    orders: "zain_orders",
    imports: "zain_import_batches"
  },
  scope: "zain-pichau-console"
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

function scoreHeaderMatch(header, aliases) {
  const normalizedHeader = normalizeKey(header);
  if (!normalizedHeader) {
    return 0;
  }

  let bestScore = 0;
  for (const alias of aliases) {
    const normalizedAlias = normalizeKey(alias);
    if (!normalizedAlias) {
      continue;
    }

    if (normalizedHeader === normalizedAlias) {
      bestScore = Math.max(bestScore, 100);
      continue;
    }

    if (normalizedHeader.startsWith(`${normalizedAlias} `) || normalizedHeader.endsWith(` ${normalizedAlias}`)) {
      bestScore = Math.max(bestScore, 80);
      continue;
    }

    if (normalizedHeader.includes(normalizedAlias)) {
      bestScore = Math.max(bestScore, 65);
      continue;
    }

    const aliasTokens = normalizedAlias.split(" ");
    const headerTokens = normalizedHeader.split(" ");
    const tokenMatches = aliasTokens.filter((token) => headerTokens.includes(token)).length;
    if (tokenMatches) {
      bestScore = Math.max(bestScore, tokenMatches * 20);
    }
  }

  return bestScore;
}

function scoreValueByField(field, value) {
  if (value == null || value === "") {
    return 0;
  }

  if (field === "stock" || field === "quantity") {
    return parseNumber(value) !== 0 || String(value).trim() === "0" ? 25 : -10;
  }

  if (field === "date") {
    return parseDate(value) ? 25 : -15;
  }

  if (field === "code") {
    const raw = String(value).trim();
    if (!raw) {
      return 0;
    }
    return /[a-zA-Z]/.test(raw) || /\d/.test(raw) ? 15 : -10;
  }

  if (field === "name") {
    const raw = String(value).trim();
    if (!raw) {
      return 0;
    }
    return /[a-zA-Z]/.test(raw) ? 15 : -10;
  }

  return 0;
}

function findColumnByScore(rows, aliasesByField) {
  const sampleRows = rows.slice(0, 20);
  const headers = Object.keys(rows[0] || {});
  const usedHeaders = new Set();
  const mapped = {};

  for (const [field, aliases] of Object.entries(aliasesByField)) {
    let bestHeader = null;
    let bestScore = -Infinity;

    for (const header of headers) {
      if (usedHeaders.has(header)) {
        continue;
      }

      let score = scoreHeaderMatch(header, aliases);
      for (const row of sampleRows) {
        score += scoreValueByField(field, row[header]);
      }

      if (score > bestScore) {
        bestScore = score;
        bestHeader = header;
      }
    }

    mapped[field] = bestScore > 0 ? bestHeader : null;
    if (mapped[field]) {
      usedHeaders.add(mapped[field]);
    }
  }

  return mapped;
}

function scoreHeaderRow(values, aliasesByField) {
  return Object.values(aliasesByField).reduce((sum, aliases) => {
    const rowBest = values.reduce((best, value) => Math.max(best, scoreHeaderMatch(value, aliases)), 0);
    return sum + rowBest;
  }, 0);
}

function rowsFromHeaderMatrix(matrix, headerRowIndex) {
  const headerRow = matrix[headerRowIndex].map((value) => String(value || "").trim());
  const rows = [];

  for (let index = headerRowIndex + 1; index < matrix.length; index += 1) {
    const values = matrix[index];
    if (!values.some((value) => String(value || "").trim())) {
      continue;
    }

    const row = {};
    headerRow.forEach((header, headerIndex) => {
      if (!header) {
        return;
      }
      row[header] = values[headerIndex] ?? "";
    });
    rows.push(row);
  }

  return rows;
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

function readStorage(key, fallback) {
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function writeStorage(key, value) {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // ignore storage failures
  }
}

function updateDatabaseStatus(label) {
  return label;
}

function setDashboardSourceStatus(kind, label) {
  if (kind === "estoque" && refs.dashboardStockStatus) {
    refs.dashboardStockStatus.textContent = label;
  }
  if (kind === "vendas" && refs.dashboardSalesStatus) {
    refs.dashboardSalesStatus.textContent = label;
  }
}

function buildSourceStatusLabel(sourceName, rowCount) {
  const safeSource = sourceName || "Manual";
  const safeRows = Number(rowCount || 0);
  return `${safeSource} • ${formatInteger(safeRows)} linhas`;
}

function getSupabaseConfig() {
  const saved = readStorage(STORAGE_KEYS.supabaseConfig, {});
  const config = { ...(window.__SUPABASE_CONFIG__ || {}), ...saved };
  return {
    url: String(config.url || "").trim(),
    anonKey: String(config.anonKey || "").trim(),
    projectScope: String(config.projectScope || "").trim() || PERSISTENCE.scope
  };
}

function saveSupabaseConfig(config) {
  writeStorage(STORAGE_KEYS.supabaseConfig, config);
}

function renderSupabaseConfigForm() {
  const config = getSupabaseConfig();
  if (refs.supabaseUrlInput) {
    refs.supabaseUrlInput.value = config.url || "";
  }
  if (refs.supabaseAnonKeyInput) {
    refs.supabaseAnonKeyInput.value = config.anonKey || "";
  }
}

async function initializePersistence() {
  const config = getSupabaseConfig();
  PERSISTENCE.scope = config.projectScope;

  if (!config.url || !config.anonKey || !window.supabase?.createClient) {
    PERSISTENCE.mode = "local";
    updateDatabaseStatus("Local");
    return;
  }

  try {
    PERSISTENCE.client = window.supabase.createClient(config.url, config.anonKey);
    const { error } = await PERSISTENCE.client
      .from(PERSISTENCE.tables.orders)
      .select("id", { head: true, count: "exact" })
      .eq("project_scope", PERSISTENCE.scope);

    if (error) {
      throw error;
    }

    PERSISTENCE.mode = "supabase";
    updateDatabaseStatus("Supabase");
    renderSupabaseConfigForm();
  } catch (error) {
    console.warn("Supabase indisponivel, mantendo modo local.", error);
    PERSISTENCE.client = null;
    PERSISTENCE.mode = "local";
    updateDatabaseStatus("Local");
    renderSupabaseConfigForm();
  }
}

async function reloadPersistentData() {
  state.persistedProducts = await loadProductsFromSupabase();
  state.productConfigs = await loadProductConfigsFromSupabase();
  state.savedOrders = await loadOrdersFromSupabase();
  state.importHistory = await loadImportsFromSupabase();
  const snapshot = loadProcessedSnapshot();
  const hydrated =
    hydrateFromImportHistory() ||
    hydrateFromPersistedProducts() ||
    (!isSimulatedSnapshot(snapshot) && restoreProcessedSnapshot(snapshot));

  if (hydrated) {
    saveProcessedSnapshot();
  } else if (isSimulatedSnapshot(snapshot)) {
    clearProcessedSnapshot();
  }
  renderPersistedProductCount();
  renderMonitorList();
  renderTable();
  renderOrderForm();
  renderOrderItems();
  renderOrderTotals();
  renderSavedOrders();
  renderImportHistory();
}

async function applySupabaseConfig() {
  const nextConfig = {
    url: String(refs.supabaseUrlInput?.value || "").trim(),
    anonKey: String(refs.supabaseAnonKeyInput?.value || "").trim(),
    projectScope: PERSISTENCE.scope
  };

  saveSupabaseConfig(nextConfig);
  await initializePersistence();
  await reloadPersistentData();
  setMessage(
    PERSISTENCE.mode === "supabase" ? "success" : "info",
    PERSISTENCE.mode === "supabase"
      ? "Supabase conectado com sucesso."
      : "Configuracao salva, mas a conexao com o Supabase nao foi validada. Mantido modo local."
  );
}

async function clearSupabaseConfig() {
  saveSupabaseConfig({ url: "", anonKey: "", projectScope: PERSISTENCE.scope });
  await initializePersistence();
  await reloadPersistentData();
  setMessage("info", "Configuracao do Supabase removida. O app voltou para modo local.");
}

async function loadProductConfigsFromSupabase() {
  if (PERSISTENCE.mode !== "supabase" || !PERSISTENCE.client) {
    return readStorage(STORAGE_KEYS.configs, {});
  }

  const { data, error } = await PERSISTENCE.client
    .from(PERSISTENCE.tables.productConfigs)
    .select("product_code, lead_days, ignore_alert, ignore_reason")
    .eq("project_scope", PERSISTENCE.scope);

  if (error) {
    console.warn("Falha ao carregar configuracoes do Supabase.", error);
    return readStorage(STORAGE_KEYS.configs, {});
  }

  const mapped = Object.fromEntries(
    (data || []).map((row) => [
      row.product_code,
      {
        leadDays: row.lead_days ?? "",
        ignoreAlert: Boolean(row.ignore_alert),
        ignoreReason: row.ignore_reason || ""
      }
    ])
  );

  writeStorage(STORAGE_KEYS.configs, mapped);
  return mapped;
}

async function loadProductsFromSupabase() {
  if (PERSISTENCE.mode !== "supabase" || !PERSISTENCE.client) {
    return readStorage(STORAGE_KEYS.products, []);
  }

  const { data, error } = await PERSISTENCE.client
    .from(PERSISTENCE.tables.products)
    .select("product_code, sku, product_name, latest_stock, current_sales, projected_sales, coverage_days, updated_at")
    .eq("project_scope", PERSISTENCE.scope)
    .order("updated_at", { ascending: false });

  if (error) {
    console.warn("Falha ao carregar produtos do Supabase.", error);
    return readStorage(STORAGE_KEYS.products, []);
  }

  const mapped = (data || []).map((row) => ({
    code: row.product_code,
    sku: row.sku || "",
    name: row.product_name || "",
    stock: Number(row.latest_stock || 0),
    currentSales: Number(row.current_sales || 0),
    projection: Number(row.projected_sales || 0),
    coverageDays: row.coverage_days == null ? null : Number(row.coverage_days),
    updatedAt: row.updated_at || ""
  }));

  writeStorage(STORAGE_KEYS.products, mapped);
  return mapped;
}

async function loadOrdersFromSupabase() {
  if (PERSISTENCE.mode !== "supabase" || !PERSISTENCE.client) {
    return readStorage(STORAGE_KEYS.orders, []);
  }

  const { data, error } = await PERSISTENCE.client
    .from(PERSISTENCE.tables.orders)
    .select("*")
    .eq("project_scope", PERSISTENCE.scope)
    .order("created_at", { ascending: false });

  if (error) {
    console.warn("Falha ao carregar pedidos do Supabase.", error);
    return readStorage(STORAGE_KEYS.orders, []);
  }

  const mapped = (data || []).map((row) => ({
    id: row.id,
    pi: row.pi || "",
    oc: row.oc || "",
    pch: row.pch || "",
    splitCount: Number(row.split_count || 1),
    splitLabels: Array.isArray(row.split_labels) ? row.split_labels : [],
    entryPercent: Number(row.entry_percent || 0),
    paymentType: row.payment_type || "antes-carregamento",
    paid: Boolean(row.paid),
    paymentDate: row.payment_date || "",
    total: Number(row.total || 0),
    entry: Number(row.entry || 0),
    balance: Number(row.balance || 0),
    items: Array.isArray(row.items) ? row.items : []
  }));

  writeStorage(STORAGE_KEYS.orders, mapped);
  return mapped;
}

async function loadImportsFromSupabase() {
  if (PERSISTENCE.mode !== "supabase" || !PERSISTENCE.client) {
    return readStorage(STORAGE_KEYS.imports, []);
  }

  const { data, error } = await PERSISTENCE.client
    .from(PERSISTENCE.tables.imports)
    .select("id, import_type, source_name, row_count, created_at, payload")
    .eq("project_scope", PERSISTENCE.scope)
    .order("created_at", { ascending: false })
    .limit(12);

  if (error) {
    console.warn("Falha ao carregar importacoes do Supabase.", error);
    return readStorage(STORAGE_KEYS.imports, []);
  }

  const mapped = (data || []).map((row) => ({
    id: row.id,
    importType: row.import_type,
    sourceName: row.source_name || "Manual",
    rowCount: Number(row.row_count || 0),
    createdAt: row.created_at || "",
    payload: Array.isArray(row.payload) ? row.payload : []
  }));

  writeStorage(STORAGE_KEYS.imports, mapped);
  return mapped;
}

async function syncProductConfigToSupabase(code) {
  if (PERSISTENCE.mode !== "supabase" || !PERSISTENCE.client || !code) {
    return;
  }

  const config = getProductConfig(code);
  const payload = {
    project_scope: PERSISTENCE.scope,
    product_code: code,
    lead_days: config.leadDays === "" ? null : Number(config.leadDays) || 0,
    ignore_alert: Boolean(config.ignoreAlert),
    ignore_reason: config.ignoreReason || ""
  };

  const { error } = await PERSISTENCE.client.from(PERSISTENCE.tables.productConfigs).upsert(payload, {
    onConflict: "project_scope,product_code"
  });

  if (error) {
    console.warn("Falha ao salvar configuracao no Supabase.", error);
  }
}

async function insertOrderToSupabase(order) {
  if (PERSISTENCE.mode !== "supabase" || !PERSISTENCE.client) {
    return;
  }

  const payload = {
    id: order.id,
    project_scope: PERSISTENCE.scope,
    pi: order.pi,
    oc: order.oc,
    pch: order.pch,
    split_count: order.splitCount,
    split_labels: order.splitLabels,
    entry_percent: order.entryPercent,
    payment_type: order.paymentType,
    paid: order.paid,
    payment_date: order.paymentDate || null,
    total: order.total,
    entry: order.entry,
    balance: order.balance,
    items: order.items
  };

  const { error } = await PERSISTENCE.client.from(PERSISTENCE.tables.orders).upsert(payload);
  if (error) {
    console.warn("Falha ao salvar pedido no Supabase.", error);
  }
}

async function deleteOrderFromSupabase(orderId) {
  if (PERSISTENCE.mode !== "supabase" || !PERSISTENCE.client || !orderId) {
    return;
  }

  const { error } = await PERSISTENCE.client
    .from(PERSISTENCE.tables.orders)
    .delete()
    .eq("project_scope", PERSISTENCE.scope)
    .eq("id", orderId);

  if (error) {
    console.warn("Falha ao remover pedido no Supabase.", error);
  }
}

async function insertImportBatchToSupabase(batch) {
  if (PERSISTENCE.mode !== "supabase" || !PERSISTENCE.client) {
    return;
  }

  const payload = {
    id: batch.id,
    project_scope: PERSISTENCE.scope,
    import_type: batch.importType,
    source_name: batch.sourceName,
    row_count: batch.rowCount,
    payload: batch.payload
  };

  const { error } = await PERSISTENCE.client.from(PERSISTENCE.tables.imports).upsert(payload);
  if (error) {
    console.warn("Falha ao salvar importacao no Supabase.", error);
  }
}

async function upsertProcessedProductsToSupabase(products) {
  if (PERSISTENCE.mode !== "supabase" || !PERSISTENCE.client || !products.length) {
    return;
  }

  const payload = products.map((product) => ({
    project_scope: PERSISTENCE.scope,
    product_code: product.code,
    sku: product.sku || "",
    product_name: product.name || "",
    latest_stock: Number(product.stock || 0),
    current_sales: Number(product.currentSales || 0),
    projected_sales: Number(product.projection || 0),
    coverage_days: Number.isFinite(product.coverageDays) ? product.coverageDays : null
  }));

  const { error } = await PERSISTENCE.client.from(PERSISTENCE.tables.products).upsert(payload, {
    onConflict: "project_scope,product_code"
  });

  if (error) {
    console.warn("Falha ao salvar produtos no Supabase.", error);
  }
}

function setActiveTab(tabName) {
  refs.tabButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.tabTarget === tabName);
  });
  refs.tabPanels.forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.tabPanel === tabName);
  });
  writeStorage(STORAGE_KEYS.activeTab, tabName);
}

function moneyUSD(value) {
  return `US$ ${Number(value || 0).toFixed(2)}`;
}

function formatPaymentTypeLabel(value) {
  const labels = {
    "antes-carregamento": "Antes do carregamento",
    "30-apos-embarque": "30 dias apos embarque",
    "60-apos-embarque": "60 dias apos embarque",
    "saldo-documentos": "Saldo contra documentos"
  };

  return labels[value] || value || "-";
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function newId(prefix = "item") {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
}

function emptyOrderItem() {
  return {
    id: newId("line"),
    type: "novo",
    code: "",
    name: "",
    quantity: 1,
    unitPriceUSD: 0
  };
}

function emptyOrderDraft() {
  return {
    editingOrderId: null,
    pi: "",
    oc: "",
    pch: "",
    splitCount: "1",
    entryPercent: "30",
    paymentType: "antes-carregamento",
    paid: false,
    paymentDate: "",
    items: [emptyOrderItem()]
  };
}

function orderDraftFromStorage() {
  const draft = readStorage(STORAGE_KEYS.orderDraft, null);
  if (!draft || typeof draft !== "object") {
    return emptyOrderDraft();
  }

  return {
    ...emptyOrderDraft(),
    ...draft,
    items: Array.isArray(draft.items) && draft.items.length
      ? draft.items.map((item) => ({ ...emptyOrderItem(), ...item }))
      : [emptyOrderItem()]
  };
}

function saveOrderDraft() {
  writeStorage(STORAGE_KEYS.orderDraft, state.orderDraft);
}

function saveProductConfigs() {
  writeStorage(STORAGE_KEYS.configs, state.productConfigs);
}

function saveOrders() {
  writeStorage(STORAGE_KEYS.orders, state.savedOrders);
}

function saveImports() {
  writeStorage(STORAGE_KEYS.imports, state.importHistory);
}

function saveProducts() {
  writeStorage(STORAGE_KEYS.products, state.persistedProducts);
}

function saveProcessedSnapshot() {
  writeStorage(STORAGE_KEYS.processedSnapshot, {
    processed: state.processed,
    months: state.months,
    monitoredInput: refs.monitoredInput?.value || "",
    stockStatus: refs.dashboardStockStatus?.textContent || "Sem leitura",
    salesStatus: refs.dashboardSalesStatus?.textContent || "Sem leitura"
  });
}

function loadProcessedSnapshot() {
  return readStorage(STORAGE_KEYS.processedSnapshot, null);
}

function isSimulatedStatus(label) {
  const normalized = normalizeKey(label);
  return normalized.includes("exemplo") || normalized.includes("simulado") || normalized.includes("demo");
}

function isSimulatedSnapshot(snapshot) {
  if (!snapshot) {
    return false;
  }

  return isSimulatedStatus(snapshot.stockStatus) || isSimulatedStatus(snapshot.salesStatus);
}

function getLatestImportBatch(kind) {
  return state.importHistory.find((batch) => batch.importType === kind && Array.isArray(batch.payload) && batch.payload.length);
}

function hydrateFromImportHistory() {
  const latestStockBatch = getLatestImportBatch("estoque");
  const latestSalesBatch = getLatestImportBatch("vendas");

  if (!latestStockBatch && !latestSalesBatch) {
    return false;
  }

  state.stockRows = latestStockBatch?.payload || [];
  state.salesRows = latestSalesBatch?.payload || [];
  state.stockLoaded = Boolean(state.stockRows.length);
  state.salesLoaded = Boolean(state.salesRows.length);

  setStatus(
    refs.stockStatus,
    latestStockBatch ? `Base persistida: ${latestStockBatch.sourceName || "Manual"}` : "Nao carregado",
    Boolean(latestStockBatch)
  );
  setStatus(
    refs.salesStatus,
    latestSalesBatch ? `Base persistida: ${latestSalesBatch.sourceName || "Manual"}` : "Nao carregado",
    Boolean(latestSalesBatch)
  );

  if (latestStockBatch) {
    setDashboardSourceStatus("estoque", buildSourceStatusLabel(latestStockBatch.sourceName, latestStockBatch.rowCount));
  }
  if (latestSalesBatch) {
    setDashboardSourceStatus("vendas", buildSourceStatusLabel(latestSalesBatch.sourceName, latestSalesBatch.rowCount));
  }

  processData();
  updateSummary();
  renderTable();
  renderMonitorList();
  return true;
}

function hydrateFromPersistedProducts() {
  if (!state.persistedProducts.length) {
    return false;
  }

  const monthKeys = getLastThreeMonthKeys(new Date());
  const currentMonthKey = monthKeys[2];

  state.months = monthKeys;
  state.monitoredTerms = getMonitoredTerms();
  state.stockRows = [];
  state.salesRows = [];
  state.stockLoaded = false;
  state.salesLoaded = false;
  state.processed = state.persistedProducts
    .filter((product) => matchesMonitoredTerms(product, state.monitoredTerms))
    .map((product) => {
      const config = getProductConfig(product.code);
      const projection = Number(product.projection || 0);
      const stock = Number(product.stock || 0);
      const currentSales = Number(product.currentSales || 0);
      const coverageDays =
        Number.isFinite(Number(product.coverageDays)) && product.coverageDays !== null
          ? Number(product.coverageDays)
          : projection > 0
            ? (stock / projection) * countBusinessDaysInMonth(new Date())
            : Number.POSITIVE_INFINITY;
      const riskLevel = stock === 0 || stock < projection ? "risk" : "healthy";
      const leadDays = parseNumber(config.leadDays);
      const reorderAlert =
        leadDays > 0 &&
        coverageDays !== Number.POSITIVE_INFINITY &&
        coverageDays < leadDays &&
        !config.ignoreAlert;

      return {
        code: product.code,
        sku: product.sku || product.code || "",
        name: product.name || "",
        stock,
        monthlySales: Object.fromEntries(monthKeys.map((key) => [key, key === currentMonthKey ? currentSales : 0])),
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

  refs.dashboardStockStatus.textContent = "Base persistida";
  refs.dashboardSalesStatus.textContent = "Base persistida";
  updateSummary();
  renderTable();
  renderMonitorList();
  return true;
}

function clearProcessedSnapshot() {
  try {
    window.localStorage.removeItem(STORAGE_KEYS.processedSnapshot);
  } catch {
    // ignore storage failures
  }
}

function formatOrderValue(value) {
  return moneyUSD(Number(value || 0));
}

function formatDateTime(value) {
  if (!value) {
    return "-";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return date.toLocaleString("pt-BR");
}

function getOrderItems() {
  return state.orderDraft?.items || [];
}

function computeOrderTotals() {
  const items = getOrderItems();
  const total = items.reduce((sum, item) => sum + (Number(item.quantity) || 0) * (Number(item.unitPriceUSD) || 0), 0);
  const entryPercent = Number(state.orderDraft?.entryPercent || 0);
  const entry = total * (entryPercent / 100);
  const balance = total - entry;
  return { total, entry, balance };
}

function getSplitLabels() {
  const base = String(state.orderDraft?.pch || "").trim();
  const splitCount = Math.max(1, Number(state.orderDraft?.splitCount || 1));
  if (!base) {
    return [];
  }
  return Array.from({ length: splitCount }, (_, index) => `${base}-${index + 1}`);
}

function syncOrderDraftToStorage() {
  saveOrderDraft();
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

  const haystack = `${product.code} ${product.sku || ""} ${product.name}`.toLowerCase();
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
  saveProductConfigs();
  void syncProductConfigToSupabase(code);
}

function renderMonitorList() {
  const items = state.processed;
  refs.monitoredCount.textContent = String(items.length);
  refs.activeAlertCount.textContent = String(items.filter((item) => item.reorderAlert).length);
  refs.ignoredAlertCount.textContent = String(items.filter((item) => item.ignoreAlert).length);
  refs.monitorEmpty.hidden = items.length > 0;
  refs.monitoredList.innerHTML = items.length
    ? items
        .map((item) => {
          const alertClass = item.reorderAlert ? "is-alert" : item.ignoreAlert ? "is-muted" : "is-safe";
          const alertLabel = item.reorderAlert ? "Comprar agora" : item.ignoreAlert ? "Ignorado" : "Cobertura ok";
          return `
            <article class="monitor-card">
              <div class="monitor-card-head">
                <div>
                  <p class="monitor-card-code">${item.code}</p>
                  <h4>${item.name || "-"}</h4>
                  <p class="monitor-meta-line">
                    SKU: ${item.sku || item.code || "-"} | Estoque: ${formatInteger(item.stock)} | Vendas: ${formatInteger(item.currentSales)}
                  </p>
                </div>
                <span class="monitor-badge ${alertClass}">${alertLabel}</span>
              </div>
              <div class="monitor-grid">
                <label>
                  Prazo de reposicao (dias)
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
                </label>
                <label class="inline-toggle">
                  <input
                    type="checkbox"
                    data-config-field="ignoreAlert"
                    data-product-code="${item.code}"
                    ${item.ignoreAlert ? "checked" : ""}
                  />
                  <span>Ignorar alerta</span>
                </label>
                <label class="monitor-note">
                  Justificativa
                  <input
                    class="table-input"
                    type="text"
                    data-config-field="ignoreReason"
                    data-product-code="${item.code}"
                    value="${item.ignoreReason || ""}"
                    placeholder="Ex.: pedido em caminho"
                  />
                </label>
              </div>
            </article>
          `;
        })
        .join("")
    : "";
}

function updateOrderHeader() {
  const splitLabels = getSplitLabels();
  refs.orderSplitBadge.textContent = splitLabels.length > 1 ? `${splitLabels.length} divisões` : "Base unica";
  refs.orderSplitPreview.textContent = splitLabels.length ? splitLabels.join(" / ") : "PCH-123-1 / PCH-123-2";
  refs.orderSplitHint.textContent = splitLabels.length
    ? "O sistema gera sufixos por pedido separado."
    : "Informe a PCH para visualizar a divisão."
}

function renderOrderItems() {
  refs.orderItemsBody.innerHTML = getOrderItems()
    .map((item) => {
      const total = (Number(item.quantity) || 0) * (Number(item.unitPriceUSD) || 0);
      return `
        <tr data-item-id="${item.id}">
          <td>
            <select data-order-field="type">
              <option value="novo" ${item.type === "novo" ? "selected" : ""}>Novo</option>
              <option value="recompra" ${item.type === "recompra" ? "selected" : ""}>Recompra</option>
            </select>
          </td>
          <td><input class="table-input" data-order-field="code" type="text" value="${item.code || ""}" placeholder="Codigo" /></td>
          <td><input class="table-input" data-order-field="name" type="text" value="${item.name || ""}" placeholder="Nome" /></td>
          <td><input class="table-input" data-order-field="quantity" type="number" min="1" step="1" value="${item.quantity || 1}" /></td>
          <td><input class="table-input" data-order-field="unitPriceUSD" type="number" min="0" step="0.01" value="${item.unitPriceUSD || 0}" /></td>
          <td>${formatOrderValue(total)}</td>
          <td><button class="button-inline" type="button" data-remove-item="${item.id}">Remover</button></td>
        </tr>
      `;
    })
    .join("");
  if (!getOrderItems().length) {
    refs.orderItemsBody.innerHTML = `
      <tr>
        <td colspan="7" class="empty-state">Adicione itens para montar o pedido.</td>
      </tr>
    `;
  }
}

function renderOrderTotals() {
  const totals = computeOrderTotals();
  refs.orderTotalValue.textContent = formatOrderValue(totals.total);
  refs.orderEntryValue.textContent = formatOrderValue(totals.entry);
  refs.orderBalanceValue.textContent = formatOrderValue(totals.balance);
  updateOrderHeader();
}

function renderSavedOrders() {
  refs.ordersTableHead.innerHTML = `
    <tr>
      <th>Pedido</th>
      <th>Itens</th>
      <th>Total</th>
      <th>Entrada</th>
      <th>Balance</th>
      <th>Entrada %</th>
      <th>Pagamento</th>
      <th>Status</th>
      <th>Divisoes</th>
      <th>Acoes</th>
    </tr>
  `;
  if (!state.savedOrders.length) {
    refs.ordersTableBody.innerHTML = `
      <tr>
        <td colspan="10" class="empty-state">Nenhum pedido registrado ainda.</td>
      </tr>
    `;
    return;
  }
  refs.ordersTableBody.innerHTML = state.savedOrders
    .map((order) => `
      <tr data-saved-order-id="${order.id}">
        <td>
          <strong>${[order.pi, order.oc, order.pch].filter(Boolean).join(" | ") || "-"}</strong>
        </td>
        <td>${order.items.map((item) => `${item.code || item.name || "Item"} x${item.quantity}`).join("<br />")}</td>
        <td>${formatOrderValue(order.total)}</td>
        <td>${formatOrderValue(order.entry)}</td>
        <td>${formatOrderValue(order.balance)}</td>
        <td>${order.entryPercent}%</td>
        <td>${formatPaymentTypeLabel(order.paymentType)}</td>
        <td>${order.paid ? `Pago${order.paymentDate ? ` em ${order.paymentDate}` : ""}` : "Pendente"}</td>
        <td>${(order.splitLabels || []).join("<br />") || "-"}</td>
        <td>
          <div class="inline-actions">
            <button class="button-inline" type="button" data-edit-order="${order.id}">Editar</button>
            <button class="button-inline" type="button" data-delete-order="${order.id}">Excluir</button>
          </div>
        </td>
      </tr>
    `)
    .join("");
}

function renderImportHistory() {
  if (!refs.importHistoryBody) {
    return;
  }

  if (!state.importHistory.length) {
    refs.importHistoryBody.innerHTML = `
      <tr>
        <td colspan="4" class="empty-state">Nenhuma importacao registrada ainda.</td>
      </tr>
    `;
    return;
  }

  refs.importHistoryBody.innerHTML = state.importHistory
    .map((batch) => `
      <tr>
        <td>${batch.importType === "estoque" ? "Estoque" : "Vendas"}</td>
        <td>${batch.sourceName || "Manual"}</td>
        <td>${formatInteger(batch.rowCount || 0)}</td>
        <td>${formatDateTime(batch.createdAt)}</td>
      </tr>
    `)
    .join("");
}

function renderPersistedProductCount() {
  if (refs.dashboardProductBaseCount) {
    refs.dashboardProductBaseCount.textContent = formatInteger(state.persistedProducts.length);
  }
}

function restoreProcessedSnapshot(snapshot) {
  if (!snapshot || !Array.isArray(snapshot.processed) || !snapshot.processed.length) {
    return false;
  }

  state.processed = snapshot.processed;
  state.months = Array.isArray(snapshot.months) ? snapshot.months : [];
  refs.monitoredInput.value = snapshot.monitoredInput || "";
  refs.dashboardStockStatus.textContent = snapshot.stockStatus || "Sem leitura";
  refs.dashboardSalesStatus.textContent = snapshot.salesStatus || "Sem leitura";
  updateSummary();
  renderTable();
  renderMonitorList();
  return true;
}

function setOrderDraft(nextDraft) {
  state.orderDraft = nextDraft;
  saveOrderDraft();
  renderOrderForm();
  renderOrderItems();
  renderOrderTotals();
}

function updateOrderDraftField(field, value) {
  setOrderDraft({
    ...state.orderDraft,
    [field]: value
  });
}

function updateOrderItem(itemId, field, value) {
  const items = getOrderItems().map((item) => {
    if (item.id !== itemId) {
      return item;
    }

    return {
      ...item,
      [field]: field === "quantity" || field === "unitPriceUSD" ? Number(value) || 0 : value
    };
  });

  setOrderDraft({
    ...state.orderDraft,
    items
  });
}

function addOrderItem() {
  setOrderDraft({
    ...state.orderDraft,
    items: [...getOrderItems(), emptyOrderItem()]
  });
}

function removeOrderItem(itemId) {
  const items = getOrderItems().filter((item) => item.id !== itemId);
  setOrderDraft({
    ...state.orderDraft,
    items: items.length ? items : [emptyOrderItem()]
  });
}

function renderOrderForm() {
  refs.orderPi.value = state.orderDraft.pi || "";
  refs.orderOc.value = state.orderDraft.oc || "";
  refs.orderPch.value = state.orderDraft.pch || "";
  refs.orderSplitCount.value = state.orderDraft.splitCount || "1";
  refs.orderEntryPercent.value = state.orderDraft.entryPercent || "30";
  refs.orderPaymentType.value = state.orderDraft.paymentType || "antes-carregamento";
  refs.orderPaid.checked = Boolean(state.orderDraft.paid);
  refs.orderPaymentDate.value = state.orderDraft.paymentDate || "";
  refs.saveOrderButton.textContent = state.orderDraft.editingOrderId ? "Atualizar pedido" : "Salvar pedido";
  refs.clearOrderButton.textContent = state.orderDraft.editingOrderId ? "Cancelar edicao" : "Limpar pedido";
}

function saveCurrentOrder() {
  const items = getOrderItems()
    .map((item) => ({
      ...item,
      quantity: Math.max(1, Number(item.quantity) || 1),
      unitPriceUSD: Number(item.unitPriceUSD) || 0
    }))
    .filter((item) => item.code || item.name);

  if (!items.length) {
    setMessage("error", "Adicione ao menos um item no pedido.");
    return;
  }

  const totals = computeOrderTotals();
  const splitLabels = getSplitLabels();
  const orderId = state.orderDraft.editingOrderId || newId("order");
  const order = {
    id: orderId,
    pi: String(state.orderDraft.pi || "").trim(),
    oc: String(state.orderDraft.oc || "").trim(),
    pch: String(state.orderDraft.pch || "").trim(),
    splitCount: Number(state.orderDraft.splitCount || 1),
    splitLabels,
    entryPercent: Number(state.orderDraft.entryPercent || 0),
    paymentType: state.orderDraft.paymentType,
    paid: Boolean(state.orderDraft.paid),
    paymentDate: String(state.orderDraft.paymentDate || "").trim() || (state.orderDraft.paid ? todayISO() : ""),
    total: totals.total,
    entry: totals.entry,
    balance: totals.balance,
    items
  };

  if (state.orderDraft.editingOrderId) {
    state.savedOrders = state.savedOrders.map((item) => (item.id === order.id ? order : item));
  } else {
    state.savedOrders = [order, ...state.savedOrders];
  }
  saveOrders();
  void insertOrderToSupabase(order);
  renderSavedOrders();
  setMessage(
    "success",
    `Pedido ${order.pch || order.pi || order.id} ${state.orderDraft.editingOrderId ? "atualizado" : "salvo"} com sucesso.`
  );
  clearOrderDraft();
}

function clearOrderDraft() {
  setOrderDraft(emptyOrderDraft());
}

function editSavedOrder(orderId) {
  const order = state.savedOrders.find((item) => item.id === orderId);
  if (!order) {
    return;
  }

  setOrderDraft({
    editingOrderId: order.id,
    pi: order.pi || "",
    oc: order.oc || "",
    pch: order.pch || "",
    splitCount: String(order.splitCount || 1),
    entryPercent: String(order.entryPercent || 0),
    paymentType: order.paymentType || "antes-carregamento",
    paid: Boolean(order.paid),
    paymentDate: order.paymentDate || "",
    items: Array.isArray(order.items) && order.items.length
      ? order.items.map((item) => ({ ...emptyOrderItem(), ...item, id: item.id || newId("line") }))
      : [emptyOrderItem()]
  });
  setMessage("info", `Pedido ${order.pch || order.pi || order.id} carregado para edicao.`);
}

function handleOrderFieldChange(event) {
  const target = event.target;
  const itemRow = target.closest("tr[data-item-id]");
  const orderField = target.dataset.orderField;

  if (itemRow && orderField) {
    updateOrderItem(itemRow.dataset.itemId, orderField, target.value);
    return;
  }

  if (target === refs.orderPi) updateOrderDraftField("pi", target.value);
  if (target === refs.orderOc) updateOrderDraftField("oc", target.value);
  if (target === refs.orderPch) updateOrderDraftField("pch", target.value);
  if (target === refs.orderSplitCount) updateOrderDraftField("splitCount", target.value);
  if (target === refs.orderEntryPercent) updateOrderDraftField("entryPercent", target.value);
  if (target === refs.orderPaymentType) updateOrderDraftField("paymentType", target.value);
  if (target === refs.orderPaid) updateOrderDraftField("paid", target.checked);
  if (target === refs.orderPaymentDate) updateOrderDraftField("paymentDate", target.value);
}

function handleOrderTableClick(event) {
  const button = event.target.closest("[data-remove-item],[data-delete-order],[data-edit-order]");
  if (!button) {
    return;
  }

  if (button.dataset.removeItem) {
    removeOrderItem(button.dataset.removeItem);
  }

  if (button.dataset.editOrder) {
    editSavedOrder(button.dataset.editOrder);
  }

  if (button.dataset.deleteOrder) {
    state.savedOrders = state.savedOrders.filter((order) => order.id !== button.dataset.deleteOrder);
    saveOrders();
    void deleteOrderFromSupabase(button.dataset.deleteOrder);
    renderSavedOrders();
  }
}

async function readTextFile(file) {
  return file.text();
}

function parseStockSpreadsheetByFixedColumns(matrix) {
  return matrix
    .map((row) => ({
      code: String(row[3] ?? "").trim(),
      sku: String(row[4] ?? "").trim(),
      name: String(row[5] ?? "").trim(),
      stock: row[10] ?? ""
    }))
    .filter((row) => {
      const normalizedCode = normalizeKey(row.code);
      const normalizedSku = normalizeKey(row.sku);
      const normalizedName = normalizeKey(row.name);
      const normalizedStock = normalizeKey(row.stock);

      const headerLike =
        ["codigo", "cod"].includes(normalizedCode) ||
        normalizedSku === "sku" ||
        normalizedName.includes("descricao") ||
        normalizedStock.includes("estoque");

      return !headerLike && Boolean(row.code) && (Boolean(row.name) || Boolean(row.sku));
    });
}

async function readSpreadsheetFile(file, aliasesByField, kind) {
  if (typeof XLSX === "undefined") {
    throw new Error("Leitor de Excel indisponivel no momento. Recarregue a pagina e tente novamente.");
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: true
  });

  if (kind === "estoque") {
    const fixedRows = parseStockSpreadsheetByFixedColumns(matrix);
    if (fixedRows.length) {
      return fixedRows;
    }
  }

  const candidateRows = matrix
    .map((row, index) => ({ index, row: row.map((value) => String(value || "").trim()) }))
    .filter(({ row }) => row.some(Boolean))
    .slice(0, 10);

  if (!candidateRows.length) {
    return [];
  }

  let bestHeaderRowIndex = candidateRows[0].index;
  let bestScore = -Infinity;

  for (const candidate of candidateRows) {
    const score = scoreHeaderRow(candidate.row, aliasesByField);
    if (score > bestScore) {
      bestScore = score;
      bestHeaderRowIndex = candidate.index;
    }
  }

  return rowsFromHeaderMatrix(matrix, bestHeaderRowIndex);
}

async function readUploadedFile(file, aliasesByField, kind) {
  const extension = (file.name.split(".").pop() || "").toLowerCase();
  if (["xlsx", "xls"].includes(extension)) {
    return readSpreadsheetFile(file, aliasesByField, kind);
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

  const rows = file ? await readUploadedFile(file, aliases, kind) : parseDelimitedText(pasted);

  if (!rows.length) {
    throw new Error(
      `Nao foi possivel ler o relatorio de ${kind}. Verifique se o arquivo tem cabecalho e ao menos uma linha de dados.`
    );
  }

  if (
    kind === "estoque" &&
    Object.prototype.hasOwnProperty.call(rows[0], "code") &&
    Object.prototype.hasOwnProperty.call(rows[0], "sku") &&
    Object.prototype.hasOwnProperty.call(rows[0], "name") &&
    Object.prototype.hasOwnProperty.call(rows[0], "stock")
  ) {
    return rows;
  }

  const mappedColumns = findColumnByScore(rows, aliases);

  const missing = Object.entries(mappedColumns)
    .filter(([, column]) => !column)
    .map(([field]) => field);

  if (missing.length) {
    throw new Error(`Colunas obrigatorias ausentes no relatorio de ${kind}: ${missing.join(", ")}.`);
  }

  const normalizedRows = rows.map((row) => {
    const normalized = {};
    for (const [field, column] of Object.entries(mappedColumns)) {
      normalized[field] = row[column];
    }
    return normalized;
  });

  if (!normalizedRows.some((row) => String(row.code || "").trim())) {
    throw new Error(
      `O relatorio de ${kind} foi lido, mas nenhuma linha trouxe codigo de produto reconhecivel. Revise os cabecalhos do arquivo.`
    );
  }

  return normalizedRows;
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
      sku: String(item.sku || item.code || "").trim(),
      name: String(item.name || "").trim(),
      stock: 0,
      monthlySales: Object.fromEntries(monthKeys.map((key) => [key, 0]))
    };

    existing.name = existing.name || String(item.name || "").trim();
    existing.sku = existing.sku || String(item.sku || item.code || "").trim();
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
      sku: String(sale.sku || sale.code || "").trim(),
      name: String(sale.name || "").trim(),
      stock: 0,
      monthlySales: Object.fromEntries(monthKeys.map((month) => [month, 0]))
    };

    existing.name = existing.name || String(sale.name || "").trim();
    existing.sku = existing.sku || String(sale.sku || sale.code || "").trim();
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

  if (!state.processed.length) {
    throw new Error(
      "Nenhum produto consolidado foi encontrado. Verifique se o relatorio possui codigo do produto e valores nas colunas esperadas."
    );
  }

  renderMonitorList();
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
  renderMonitorList();
}

function buildTableHead() {
  refs.tableHead.innerHTML = `
    <tr>
      <th>Codigo</th>
      <th>SKU</th>
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
      <td>${item.sku || item.code || "-"}</td>
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
      const haystack = `${item.code} ${item.sku || ""} ${item.name}`.toLowerCase();
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
        <td colspan="14" class="empty-state">Nenhum produto consolidado para exibir.</td>
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
        <td colspan="14" class="empty-state">Nenhum item encontrado com os filtros atuais.</td>
      </tr>
    `;
}

function resetState(options = {}) {
  const { clearSnapshot = false } = options;
  state.stockRows = [];
  state.salesRows = [];
  state.stockLoaded = false;
  state.salesLoaded = false;
  state.processed = [];
  state.months = [];
  state.monitoredTerms = [];
  state.persistedProducts = [];
  state.productConfigs = {};
  state.importHistory = [];
  state.orderDraft = emptyOrderDraft();
  if (clearSnapshot) {
    clearProcessedSnapshot();
  }

  refs.stockFile.value = "";
  refs.salesFile.value = "";
  refs.stockText.value = "";
  refs.salesText.value = "";
  refs.monitoredInput.value = "";
  refs.searchInput.value = "";
  refs.riskFilter.value = "all";
  refs.sortSelect.value = "projection";
  refs.orderPi.value = "";
  refs.orderOc.value = "";
  refs.orderPch.value = "";
  refs.orderSplitCount.value = "1";
  refs.orderEntryPercent.value = "30";
  refs.orderPaymentType.value = "antes-carregamento";
  refs.orderPaid.checked = false;
  refs.orderPaymentDate.value = "";

  setStatus(refs.stockStatus, "Aguardando", false);
  setStatus(refs.salesStatus, "Aguardando", false);
  refs.dashboardStockStatus.textContent = "Sem leitura";
  refs.dashboardSalesStatus.textContent = "Sem leitura";
  refs.dashboardProductBaseCount.textContent = "0";
  setMessage("info", "Carregue um ou ambos os relatorios para montar a visao operacional de Zain na Pichau.");

  refs.productCount.textContent = "0";
  refs.totalStock.textContent = "0";
  refs.currentMonthSales.textContent = "0";
  refs.currentMonthProjection.textContent = "0";
  refs.monthsLabel.textContent = "Ultimos 3 meses ainda nao carregados.";
  renderTable();
  renderMonitorList();
  renderOrderForm();
  renderOrderItems();
  renderOrderTotals();
  renderSavedOrders();
  renderImportHistory();
  renderPersistedProductCount();
}

function createImportBatch(kind, rows, sourceName) {
  return {
    id: newId("import"),
    importType: kind,
    sourceName: sourceName || "Manual",
    rowCount: rows.length,
    createdAt: new Date().toISOString(),
    payload: rows
  };
}

function getSourceName(fileInput, fallbackLabel) {
  return fileInput.files?.[0]?.name || fallbackLabel;
}

async function registerImport(kind, rows, sourceName) {
  const batch = createImportBatch(kind, rows, sourceName);
  state.importHistory = [batch, ...state.importHistory].slice(0, 12);
  saveImports();
  renderImportHistory();
  setDashboardSourceStatus(kind, buildSourceStatusLabel(batch.sourceName, batch.rowCount));
  await insertImportBatchToSupabase(batch);
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
    state.persistedProducts = state.processed.map((product) => ({
      code: product.code,
      sku: product.sku || "",
      name: product.name || "",
      stock: product.stock,
      currentSales: product.currentSales,
      projection: product.projection,
      coverageDays: product.coverageDays,
      updatedAt: new Date().toISOString()
    }));
    saveProducts();
    renderPersistedProductCount();
    await upsertProcessedProductsToSupabase(state.persistedProducts);
    updateSummary();
    renderTable();
    saveProcessedSnapshot();

    if (stockRows) {
      await registerImport("estoque", stockRows, getSourceName(refs.stockFile, "Texto colado"));
    }

    if (salesRows) {
      await registerImport("vendas", salesRows, getSourceName(refs.salesFile, "Texto colado"));
    }

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
      setDashboardSourceStatus("estoque", "Exemplo carregado");
    } else {
      refs.salesText.value = content;
      refs.salesFile.value = "";
      setStatus(refs.salesStatus, "Exemplo carregado", true);
      setDashboardSourceStatus("vendas", "Exemplo carregado");
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
    saveProcessedSnapshot();
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
refs.clearButton.addEventListener("click", () => resetState({ clearSnapshot: true }));
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
refs.orderItemsBody.addEventListener("change", handleOrderFieldChange);
refs.orderItemsBody.addEventListener("input", handleOrderFieldChange);
refs.orderItemsBody.addEventListener("click", handleOrderTableClick);
refs.orderPi.addEventListener("input", handleOrderFieldChange);
refs.orderOc.addEventListener("input", handleOrderFieldChange);
refs.orderPch.addEventListener("input", handleOrderFieldChange);
refs.orderSplitCount.addEventListener("change", handleOrderFieldChange);
refs.orderEntryPercent.addEventListener("change", handleOrderFieldChange);
refs.orderPaymentType.addEventListener("change", handleOrderFieldChange);
refs.orderPaid.addEventListener("change", handleOrderFieldChange);
refs.orderPaymentDate.addEventListener("change", handleOrderFieldChange);
refs.addOrderItemButton.addEventListener("click", addOrderItem);
refs.saveOrderButton.addEventListener("click", saveCurrentOrder);
refs.clearOrderButton.addEventListener("click", clearOrderDraft);
refs.stockExampleButton.addEventListener("click", () => handleExample("estoque"));
refs.salesExampleButton.addEventListener("click", () => handleExample("vendas"));
refs.stockFile.addEventListener("change", () => {
  if (refs.stockFile.files?.[0]) {
    setStatus(refs.stockStatus, refs.stockFile.files[0].name, true);
    refs.stockText.value = "";
    setDashboardSourceStatus("estoque", `${refs.stockFile.files[0].name} - pronto para processar`);
  }
});
refs.salesFile.addEventListener("change", () => {
  if (refs.salesFile.files?.[0]) {
    setStatus(refs.salesStatus, refs.salesFile.files[0].name, true);
    refs.salesText.value = "";
    setDashboardSourceStatus("vendas", `${refs.salesFile.files[0].name} - pronto para processar`);
  }
});
refs.tabButtons.forEach((button) => {
  button.addEventListener("click", () => setActiveTab(button.dataset.tabTarget));
});
refs.saveSupabaseConfigButton?.addEventListener("click", () => {
  void applySupabaseConfig();
});
refs.clearSupabaseConfigButton?.addEventListener("click", () => {
  void clearSupabaseConfig();
});

async function initializeApp() {
  resetState();
  await initializePersistence();
  state.persistedProducts = await loadProductsFromSupabase();
  state.productConfigs = await loadProductConfigsFromSupabase();
  state.savedOrders = await loadOrdersFromSupabase();
  state.importHistory = await loadImportsFromSupabase();
  state.orderDraft = orderDraftFromStorage();
  setActiveTab(readStorage(STORAGE_KEYS.activeTab, "dashboard"));
  renderSupabaseConfigForm();
  renderMonitorList();
  renderOrderForm();
  renderOrderItems();
  renderOrderTotals();
  renderSavedOrders();
  renderImportHistory();
  renderPersistedProductCount();

  if (hydrateFromImportHistory()) {
    saveProcessedSnapshot();
    return;
  }

  const snapshot = loadProcessedSnapshot();
  if (!isSimulatedSnapshot(snapshot) && restoreProcessedSnapshot(snapshot)) {
    return;
  }

  if (isSimulatedSnapshot(snapshot)) {
    clearProcessedSnapshot();
  }

  if (hydrateFromPersistedProducts()) {
    return;
  }
}

void initializeApp();
