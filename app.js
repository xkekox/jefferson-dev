const state = {
  stockRows: [],
  salesRows: [],
  products: [],
  selectedSector: 'Todos',
  selectedBrand: 'Todas',
  selectedSubgroup: 'Todos',
  searchTerm: ''
};

const elements = {
  stockFile: document.getElementById('stockFile'),
  salesFile: document.getElementById('salesFile'),
  processButton: document.getElementById('processButton'),
  clearButton: document.getElementById('clearButton'),
  searchButton: document.getElementById('searchButton'),
  clearSearchButton: document.getElementById('clearSearchButton'),
  messageBox: document.getElementById('messageBox'),
  tableBody: document.getElementById('tableBody'),
  monthHeader0: document.getElementById('monthHeader0'),
  monthHeader1: document.getElementById('monthHeader1'),
  monthHeader2: document.getElementById('monthHeader2'),
  monthsLabel: document.getElementById('monthsLabel'),
  kpiProducts: document.getElementById('kpiProducts'),
  kpiStock: document.getElementById('kpiStock'),
  kpiCurrentSales: document.getElementById('kpiCurrentSales'),
  kpiProjection: document.getElementById('kpiProjection'),
  kpiNoStock: document.getElementById('kpiNoStock'),
  kpiLowStock: document.getElementById('kpiLowStock'),
  sectorFilter: document.getElementById('sectorFilter'),
  sectorChips: document.getElementById('sectorChips'),
  productSearch: document.getElementById('productSearch'),
  brandFilter: document.getElementById('brandFilter'),
  subgroupFilter: document.getElementById('subgroupFilter'),
  brandSummaryBody: document.getElementById('brandSummaryBody'),
  scopeTitle: document.getElementById('scopeTitle'),
  scopeSubtitle: document.getElementById('scopeSubtitle'),
  scopeStock: document.getElementById('scopeStock'),
  scopeCurrentSales: document.getElementById('scopeCurrentSales'),
  scopeProjection: document.getElementById('scopeProjection'),
  scopeBrandCount: document.getElementById('scopeBrandCount')
};

const SECTOR_RULES = [
  { name: 'Gabinete', keywords: ['gabinete', 'case', 'chassis'] },
  { name: 'Placa-mae', keywords: ['placamae', 'placa mae', 'motherboard', 'mb '] },
  { name: 'Memoria', keywords: ['memoria', 'ram', 'ddr4', 'ddr5'] },
  { name: 'SSD', keywords: ['ssd', 'nvme', 'm2', 'sata ssd'] },
  { name: 'Placa de video', keywords: ['placa de video', 'gpu', 'rtx', 'gtx', 'radeon', 'geforce'] },
  { name: 'Perifericos', keywords: ['mouse', 'teclado', 'headset', 'microfone', 'webcam', 'monitor', 'cadeira', 'controle', 'mousepad'] }
];

const BRAND_PATTERNS = [
  'legion',
  'pichau',
  'mancer',
  'acer',
  'asus',
  'aorus',
  'gigabyte',
  'msi',
  'galax',
  'corsair',
  'kingston',
  'hyperx',
  'crucial',
  'adata',
  'xpg',
  'wd',
  'western digital',
  'seagate',
  'sandisk',
  'logitech',
  'redragon',
  'razer',
  'lenovo',
  'intel',
  'amd',
  'nvidia'
];

const MONTH_FORMATTER = new Intl.DateTimeFormat('pt-BR', {
  month: 'short',
  year: '2-digit'
});

const NUMBER_FORMATTER = new Intl.NumberFormat('pt-BR', {
  maximumFractionDigits: 2,
  minimumFractionDigits: 0
});

function normalizeKey(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function toNumber(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  const normalized = String(value || '')
    .trim()
    .replace(/\s/g, '')
    .replace(/\.(?=\d{3}(?:\D|$))/g, '')
    .replace(',', '.')
    .replace(/[^\d.-]/g, '');

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === 'number' && window.XLSX) {
    const converted = window.XLSX.SSF.parse_date_code(value);
    if (converted) {
      return new Date(converted.y, converted.m - 1, converted.d);
    }
  }

  const raw = String(value || '').trim();
  if (!raw) {
    return null;
  }

  const brMatch = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (brMatch) {
    const year = brMatch[3].length === 2 ? `20${brMatch[3]}` : brMatch[3];
    return new Date(Number(year), Number(brMatch[2]) - 1, Number(brMatch[1]));
  }

  const iso = new Date(raw);
  return Number.isNaN(iso.getTime()) ? null : iso;
}

function getWorkdaysInMonth(referenceDate) {
  const year = referenceDate.getFullYear();
  const month = referenceDate.getMonth();
  const today = new Date();
  let elapsed = 0;
  let total = 0;

  const lastDay = new Date(year, month + 1, 0).getDate();
  for (let day = 1; day <= lastDay; day += 1) {
    const current = new Date(year, month, day);
    const weekDay = current.getDay();
    const isWorkday = weekDay !== 0 && weekDay !== 6;

    if (isWorkday) {
      total += 1;
      if (current <= today) {
        elapsed += 1;
      }
    }
  }

  return {
    elapsed: Math.max(elapsed, 1),
    total: Math.max(total, 1)
  };
}

function getMonthBuckets() {
  const now = new Date();
  const current = new Date(now.getFullYear(), now.getMonth(), 1);
  const previous = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const previousTwo = new Date(now.getFullYear(), now.getMonth() - 2, 1);
  return [previousTwo, previous, current];
}

function formatMonthLabel(date) {
  return MONTH_FORMATTER.format(date).replace('.', '');
}

function getHeaderMap(row) {
  const map = new Map();
  Object.keys(row || {}).forEach((key) => {
    map.set(normalizeKey(key), key);
  });
  return map;
}

function findValue(row, aliases) {
  const headerMap = getHeaderMap(row);
  for (const alias of aliases) {
    const resolved = headerMap.get(normalizeKey(alias));
    if (resolved) {
      return row[resolved];
    }
  }
  return '';
}

function sanitizeBrand(value) {
  const brand = String(value || '').trim();
  if (!brand) {
    return '';
  }

  const normalized = brand
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();

  if (!normalized) {
    return '';
  }

  if (/^\d+$/.test(normalized)) {
    return '';
  }

  if (/^[\d\s./-]+$/.test(normalized)) {
    return '';
  }

  return normalized;
}

function inferBrandFromText(...values) {
  const source = values
    .join(' ')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();

  const matched = BRAND_PATTERNS.find((pattern) => source.includes(pattern));
  if (!matched) {
    return '';
  }

  if (matched === 'western digital') {
    return 'Western Digital';
  }

  if (matched === 'wd') {
    return 'WD';
  }

  return matched
    .split(' ')
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(' ');
}

function resolveBrand(row, contextValues = []) {
  const directBrand = sanitizeBrand(
    findValue(row, ['marca', 'submarca', 'fabricante', 'linha', 'colecao', 'vendor', 'brand'])
  );

  if (directBrand) {
    return directBrand;
  }

  return inferBrandFromText(...contextValues);
}

async function readFileRows(file) {
  const extension = file.name.split('.').pop().toLowerCase();

  if (extension === 'csv') {
    const text = await file.text();
    return parseCsv(text);
  }

  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, { type: 'array' });
  const firstSheet = workbook.SheetNames[0];
  return window.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });
}

function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  if (!lines.length) {
    return [];
  }

  const delimiter = detectDelimiter(lines[0]);
  const headers = splitCsvLine(lines[0], delimiter);

  return lines.slice(1).map((line) => {
    const values = splitCsvLine(line, delimiter);
    const row = {};
    headers.forEach((header, index) => {
      row[header] = values[index] || '';
    });
    return row;
  });
}

function detectDelimiter(line) {
  if (line.includes(';')) {
    return ';';
  }
  if (line.includes('\t')) {
    return '\t';
  }
  return ',';
}

function splitCsvLine(line, delimiter) {
  if (delimiter === '\t') {
    return line.split('\t').map((part) => part.trim());
  }

  const values = [];
  let current = '';
  let quoted = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    if (char === '"') {
      if (quoted && line[index + 1] === '"') {
        current += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
      continue;
    }

    if (char === delimiter && !quoted) {
      values.push(current.trim());
      current = '';
      continue;
    }

    current += char;
  }

  values.push(current.trim());
  return values;
}

function parseStockRows(rows) {
  return rows
    .map((row) => {
      const code = String(findValue(row, ['codigo do produto', 'codigo', 'codigodoproduto', 'codproduto', 'id'])).trim();
      if (!code) {
        return null;
      }

      return {
        code,
        sku: String(findValue(row, ['sku', 'referencia'])).trim(),
        name: String(findValue(row, ['nome', 'nome do produto', 'descricao', 'produto'])).trim(),
        brand: '',
        group: String(findValue(row, ['grupo'])).trim(),
        subgroup: String(findValue(row, ['subgrupo'])).trim(),
        currentStock: toNumber(findValue(row, ['estoque atual', 'estoque fisico', 'estoque', 'saldo'])),
        reservedStock: toNumber(findValue(row, ['reserva', 'estoque reservado', 'reservado'])),
        minimumStock: toNumber(findValue(row, ['estoque minimo', 'minimo'])),
        cost: toNumber(findValue(row, ['valor de custo', 'custo unitario', 'custo'])),
        price: toNumber(findValue(row, ['preco unitario', 'valor de venda', 'preco', 'valor'])),
        averageSalesHint: toNumber(findValue(row, ['media de venda 3 meses', 'media de venda', 'media 3 meses']))
      };
    })
    .filter(Boolean)
    .map((row) => ({
      ...row,
      brand: resolveBrand(
        {
          marca: row.brand,
          grupo: row.group,
          subgrupo: row.subgroup,
          nome: row.name,
          sku: row.sku
        },
        [row.name, row.group, row.subgroup, row.sku]
      )
    }))
    .filter(Boolean);
}

function parseSalesRows(rows) {
  return rows
    .map((row) => {
      const code = String(findValue(row, ['codigo do produto', 'codigo', 'codproduto', 'produto codigo'])).trim();
      const date = parseDate(findValue(row, ['data', 'data da venda', 'emissao']));
      if (!code || !date) {
        return null;
      }

      return {
        code,
        name: String(findValue(row, ['descricao', 'produto', 'nome'])).trim(),
        brand: resolveBrand(row, [
          findValue(row, ['descricao', 'produto', 'nome']),
          findValue(row, ['grupo']),
          findValue(row, ['subgrupo'])
        ]),
        group: String(findValue(row, ['grupo'])).trim(),
        subgroup: String(findValue(row, ['subgrupo'])).trim(),
        quantity: toNumber(findValue(row, ['quantidade', 'quant.', 'qtd', 'qtdvendida'])),
        value: toNumber(findValue(row, ['valor', 'total', 'valor vendido'])),
        date
      };
    })
    .filter(Boolean);
}

function consolidateProducts(stockRows, salesRows) {
  const monthBuckets = getMonthBuckets();
  const monthKeys = monthBuckets.map((date) => `${date.getFullYear()}-${date.getMonth()}`);
  const currentBucket = monthBuckets[2];
  const workdays = getWorkdaysInMonth(currentBucket);
  const productMap = new Map();

  stockRows.forEach((row) => {
    productMap.set(row.code, {
      code: row.code,
      sku: row.sku,
      name: row.name,
      brand: row.brand,
      group: row.group,
      subgroup: row.subgroup,
      currentStock: row.currentStock,
      reservedStock: row.reservedStock,
      availableStock: row.currentStock - row.reservedStock,
      minimumStock: row.minimumStock,
      cost: row.cost,
      price: row.price,
      monthlySales: [0, 0, 0],
      monthlyRevenue: [0, 0, 0],
      averageSalesHint: row.averageSalesHint
    });
  });

  salesRows.forEach((row) => {
    const bucketKey = `${row.date.getFullYear()}-${row.date.getMonth()}`;
    const monthIndex = monthKeys.indexOf(bucketKey);
    if (monthIndex === -1) {
      return;
    }

    const existing = productMap.get(row.code) || {
      code: row.code,
      sku: '',
      name: row.name,
      brand: row.brand,
      group: row.group,
      subgroup: row.subgroup,
      currentStock: 0,
      reservedStock: 0,
      availableStock: 0,
      minimumStock: 0,
      cost: 0,
      price: 0,
      monthlySales: [0, 0, 0],
      monthlyRevenue: [0, 0, 0],
      averageSalesHint: 0
    };

    existing.monthlySales[monthIndex] += row.quantity;
    existing.monthlyRevenue[monthIndex] += row.value;
    existing.name = existing.name || row.name;
    existing.brand = existing.brand || row.brand;
    existing.group = existing.group || row.group;
    existing.subgroup = existing.subgroup || row.subgroup;

    productMap.set(row.code, existing);
  });

  return Array.from(productMap.values())
    .map((product) => {
      const averageSales = product.averageSalesHint || average(product.monthlySales);
      const currentMonthSales = product.monthlySales[2];
      const projection = (currentMonthSales / workdays.elapsed) * workdays.total;
      const coverage = averageSales > 0 ? product.availableStock / averageSales : 0;
      const need = Math.max(0, projection + product.minimumStock - product.availableStock);
      const status = classifyStock(product.availableStock, product.minimumStock, coverage, projection);

      return {
        ...product,
        sector: inferSector(product),
        projection,
        averageSales,
        coverage,
        need,
        status
      };
    })
    .sort((left, right) => right.projection - left.projection);
}

function average(values) {
  const total = values.reduce((sum, current) => sum + current, 0);
  return values.length ? total / values.length : 0;
}

function classifyStock(availableStock, minimumStock, coverage, projection) {
  if (availableStock <= 0) {
    return 'Sem estoque';
  }
  if (availableStock < minimumStock) {
    return 'Estoque baixo';
  }
  if (projection > 0 && coverage > 0 && coverage < 0.5) {
    return 'Estoque critico';
  }
  if (projection > 0 && availableStock > projection * 2) {
    return 'Estoque alto';
  }
  return 'Estoque normal';
}

function statusClassName(status) {
  const normalized = normalizeKey(status);
  if (normalized === 'semestoque') {
    return 'status-sem-estoque';
  }
  if (normalized === 'estoquebaixo') {
    return 'status-baixo';
  }
  if (normalized === 'estoquecritico') {
    return 'status-critico';
  }
  if (normalized === 'estoquealto') {
    return 'status-alto';
  }
  return 'status-normal';
}

function inferSector(product) {
  const groupSource = [product.group, product.subgroup]
    .join(' ')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();

  const source = [
    product.name,
    product.brand,
    product.group,
    product.subgroup,
    product.sku
  ]
    .join(' ')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();

  for (const rule of SECTOR_RULES) {
    if (rule.keywords.some((keyword) => groupSource.includes(keyword))) {
      return rule.name;
    }
    if (rule.keywords.some((keyword) => source.includes(keyword))) {
      return rule.name;
    }
  }

  return 'Outros';
}

function getVisibleProducts() {
  return state.products.filter((product) => {
    const search = normalizeKey(state.searchTerm);
    const searchableCode = normalizeKey(product.code);
    const searchableSku = normalizeKey(product.sku);
    const searchMatch = !search || searchableCode.includes(search) || searchableSku.includes(search);
    if (search) {
      return searchMatch;
    }
    const sectorMatch = state.selectedSector === 'Todos' || product.sector === state.selectedSector;
    const brandMatch = state.selectedBrand === 'Todas' || normalizeBrand(product.brand) === state.selectedBrand;
    const subgroupValue = product.subgroup || product.group || 'Sem subgrupo';
    const subgroupMatch = state.selectedSubgroup === 'Todos' || subgroupValue === state.selectedSubgroup;
    return searchMatch && sectorMatch && brandMatch && subgroupMatch;
  });
}

function getSectorScopedProducts() {
  if (state.selectedSector === 'Todos') {
    return state.products;
  }

  return state.products.filter((product) => product.sector === state.selectedSector);
}

function render() {
  const monthBuckets = getMonthBuckets();
  elements.monthHeader0.textContent = formatMonthLabel(monthBuckets[0]);
  elements.monthHeader1.textContent = formatMonthLabel(monthBuckets[1]);
  elements.monthHeader2.textContent = formatMonthLabel(monthBuckets[2]);
  elements.monthsLabel.textContent = monthBuckets.map(formatMonthLabel).join(' | ');

  if (!state.products.length) {
    elements.tableBody.innerHTML =
      '<tr><td colspan="18" class="empty-state">Nenhuma leitura consolidada ainda.</td></tr>';
    updateKpis([]);
    renderSectorChips([]);
    renderBrandOptions([]);
    renderSubgroupOptions([]);
    renderBrandSummary([]);
    return;
  }

  const sectorScopedProducts = getSectorScopedProducts();
  const visibleProducts = getVisibleProducts();
  renderSectorChips(state.products);
  renderBrandOptions(sectorScopedProducts);
  renderSubgroupOptions(sectorScopedProducts);
  renderBrandSummary(sectorScopedProducts, visibleProducts);

  if (!visibleProducts.length) {
    elements.tableBody.innerHTML =
      '<tr><td colspan="18" class="empty-state">Nenhum produto encontrado para o filtro selecionado.</td></tr>';
    updateKpis([]);
    return;
  }

  elements.tableBody.innerHTML = visibleProducts
    .map((product) => {
      return `
        <tr>
          <td>${escapeHtml(product.code)}</td>
          <td>${escapeHtml(product.sku)}</td>
          <td>${escapeHtml(product.name)}</td>
          <td>${escapeHtml(product.brand)}</td>
          <td>${escapeHtml(product.sector)}</td>
          <td>${escapeHtml(product.group)}</td>
          <td>${escapeHtml(product.subgroup)}</td>
          <td>${formatNumber(product.currentStock)}</td>
          <td>${formatNumber(product.reservedStock)}</td>
          <td>${formatNumber(product.availableStock)}</td>
          <td>${formatNumber(product.monthlySales[0])}</td>
          <td>${formatNumber(product.monthlySales[1])}</td>
          <td>${formatNumber(product.monthlySales[2])}</td>
          <td>${formatNumber(product.projection)}</td>
          <td>${formatNumber(product.averageSales)}</td>
          <td>${formatNumber(product.coverage)}</td>
          <td>${formatNumber(product.need)}</td>
          <td><span class="status-pill ${statusClassName(product.status)}">${escapeHtml(product.status)}</span></td>
        </tr>
      `;
    })
    .join('');

  updateKpis(visibleProducts);
}

function renderSectorChips(products) {
  const counts = new Map();
  products.forEach((product) => {
    counts.set(product.sector, (counts.get(product.sector) || 0) + 1);
  });

  const ordered = ['Gabinete', 'Placa-mae', 'Memoria', 'SSD', 'Placa de video', 'Perifericos', 'Outros']
    .filter((sector) => counts.has(sector))
    .map((sector) => `<span class="sector-chip">${escapeHtml(sector)}: ${formatNumber(counts.get(sector))}</span>`);

  elements.sectorChips.innerHTML = ordered.join('');
}

function renderBrandOptions(products) {
  const brands = Array.from(
    new Set(products.map((product) => normalizeBrand(product.brand)).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right, 'pt-BR'));

  if (state.selectedBrand !== 'Todas' && !brands.includes(state.selectedBrand)) {
    state.selectedBrand = 'Todas';
  }

  elements.brandFilter.innerHTML = ['<option value="Todas">Todas</option>']
    .concat(brands.map((brand) => `<option value="${escapeHtml(brand)}">${escapeHtml(brand)}</option>`))
    .join('');
  elements.brandFilter.value = state.selectedBrand;
}

function normalizeBrand(value) {
  const brand = sanitizeBrand(value);
  return brand || 'Sem marca';
}

function renderSubgroupOptions(products) {
  const subgroups = Array.from(
    new Set(products.map((product) => product.subgroup || product.group || 'Sem subgrupo').filter(Boolean))
  ).sort((left, right) => left.localeCompare(right, 'pt-BR'));

  if (state.selectedSubgroup !== 'Todos' && !subgroups.includes(state.selectedSubgroup)) {
    state.selectedSubgroup = 'Todos';
  }

  elements.subgroupFilter.innerHTML = ['<option value="Todos">Todos</option>']
    .concat(subgroups.map((subgroup) => `<option value="${escapeHtml(subgroup)}">${escapeHtml(subgroup)}</option>`))
    .join('');
  elements.subgroupFilter.value = state.selectedSubgroup;
}

function renderBrandSummary(sectorScopedProducts, visibleProducts) {
  if (!sectorScopedProducts.length) {
    elements.brandSummaryBody.innerHTML =
      '<tr><td colspan="5" class="empty-state">Importe as planilhas para gerar os totais por marca.</td></tr>';
    updateScopeSummary([]);
    return;
  }

  const brandMap = new Map();
  sectorScopedProducts.forEach((product) => {
    const brand = normalizeBrand(product.brand);
    const current = brandMap.get(brand) || {
      brand,
      products: 0,
      stock: 0,
      currentSales: 0,
      projection: 0
    };

    current.products += 1;
    current.stock += product.currentStock;
    current.currentSales += product.monthlySales[2];
    current.projection += product.projection;
    brandMap.set(brand, current);
  });

  const rows = Array.from(brandMap.values()).sort((left, right) => right.currentSales - left.currentSales);
  elements.brandSummaryBody.innerHTML = rows
    .map((row) => `
      <tr>
        <td>${escapeHtml(row.brand)}</td>
        <td>${formatNumber(row.products)}</td>
        <td>${formatNumber(row.stock)}</td>
        <td>${formatNumber(row.currentSales)}</td>
        <td>${formatNumber(row.projection)}</td>
      </tr>
    `)
    .join('');

  updateScopeSummary(visibleProducts, rows.length);
}

function updateScopeSummary(products, brandCount = 0) {
  const searchMode = Boolean(state.searchTerm);
  const sectorLabel = state.selectedSector === 'Todos' ? 'Todos os produtos' : state.selectedSector;
  const brandLabel = state.selectedBrand === 'Todas' ? 'todas as marcas' : state.selectedBrand;
  const subgroupLabel = state.selectedSubgroup === 'Todos' ? 'todos os subgrupos' : state.selectedSubgroup;
  if (searchMode) {
    elements.scopeTitle.textContent = `Resultado da busca`;
    elements.scopeSubtitle.textContent = `Busca atual por codigo/SKU: ${state.searchTerm}`;
  } else {
    elements.scopeTitle.textContent = `Totais de ${sectorLabel}`;
    elements.scopeSubtitle.textContent = `Recorte atual: ${sectorLabel} | Marca: ${brandLabel} | Subgrupo: ${subgroupLabel}`;
  }
  elements.scopeStock.textContent = formatNumber(products.reduce((sum, product) => sum + product.currentStock, 0));
  elements.scopeCurrentSales.textContent = formatNumber(products.reduce((sum, product) => sum + product.monthlySales[2], 0));
  elements.scopeProjection.textContent = formatNumber(products.reduce((sum, product) => sum + product.projection, 0));
  elements.scopeBrandCount.textContent = formatNumber(brandCount);
}

function updateKpis(products) {
  const productCount = products.length;
  const totalStock = products.reduce((sum, product) => sum + product.currentStock, 0);
  const currentSales = products.reduce((sum, product) => sum + product.monthlySales[2], 0);
  const projection = products.reduce((sum, product) => sum + product.projection, 0);
  const noStock = products.filter((product) => product.availableStock <= 0).length;
  const lowStock = products.filter((product) => product.status === 'Estoque baixo' || product.status === 'Estoque critico').length;

  elements.kpiProducts.textContent = formatNumber(productCount);
  elements.kpiStock.textContent = formatNumber(totalStock);
  elements.kpiCurrentSales.textContent = formatNumber(currentSales);
  elements.kpiProjection.textContent = formatNumber(projection);
  elements.kpiNoStock.textContent = formatNumber(noStock);
  elements.kpiLowStock.textContent = formatNumber(lowStock);
}

function formatNumber(value) {
  return NUMBER_FORMATTER.format(Number(value || 0));
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function setMessage(text, tone) {
  elements.messageBox.textContent = text;
  elements.messageBox.className = `message ${tone}`;
}

async function processFiles() {
  const stockFile = elements.stockFile.files[0];
  const salesFile = elements.salesFile.files[0];

  if (!stockFile || !salesFile) {
    setMessage('Selecione as duas planilhas antes de processar.', 'error');
    return;
  }

  try {
    setMessage('Lendo planilhas e consolidando base...', 'info');
    const [rawStockRows, rawSalesRows] = await Promise.all([readFileRows(stockFile), readFileRows(salesFile)]);
    state.stockRows = parseStockRows(rawStockRows);
    state.salesRows = parseSalesRows(rawSalesRows);
    state.products = consolidateProducts(state.stockRows, state.salesRows);

    if (!state.products.length) {
      setMessage('As planilhas foram lidas, mas nenhuma linha valida foi consolidada. Verifique os nomes das colunas.', 'error');
      render();
      return;
    }

    setMessage(
      `Leitura concluida: ${state.stockRows.length} linhas de estoque, ${state.salesRows.length} linhas de vendas e ${state.products.length} produtos consolidados.`,
      'success'
    );
    render();
  } catch (error) {
    console.error(error);
    setMessage(`Falha ao processar as planilhas: ${error.message}`, 'error');
  }
}

function clearAll() {
  elements.stockFile.value = '';
  elements.salesFile.value = '';
  state.stockRows = [];
  state.salesRows = [];
  state.products = [];
  setMessage('Leitura limpa. Envie novas planilhas para processar.', 'info');
  render();
}

elements.processButton.addEventListener('click', processFiles);
elements.clearButton.addEventListener('click', clearAll);
elements.sectorFilter.addEventListener('change', (event) => {
  state.selectedSector = event.target.value;
  state.selectedBrand = 'Todas';
  state.selectedSubgroup = 'Todos';
  render();
});
elements.searchButton.addEventListener('click', () => {
  state.searchTerm = elements.productSearch.value.trim();
  render();
});
elements.productSearch.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault();
    state.searchTerm = elements.productSearch.value.trim();
    render();
  }
});
elements.clearSearchButton.addEventListener('click', () => {
  elements.productSearch.value = '';
  state.searchTerm = '';
  render();
});
elements.brandFilter.addEventListener('change', (event) => {
  state.selectedBrand = event.target.value;
  render();
});
elements.subgroupFilter.addEventListener('change', (event) => {
  state.selectedSubgroup = event.target.value;
  render();
});

render();
