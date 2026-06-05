const {
  useEffect,
  useMemo,
  useRef,
  useState
} = React;
const DASHBOARD_XLSX = "./public/base-dashboard.xlsx";
const DASHBOARD_CSV = "./public/base-dashboard.csv";
const GOOGLE_SHEETS_SPREADSHEET_ID = "1JldFrcw8oaVAWXXhFyJCMVvm_Be9IXju90UAqpzSLXM";
const GOOGLE_SHEETS_GID = "27856229";
const GOOGLE_SHEETS_SOURCE_LABEL = "Google Sheets - BASE DASHBOARD";
const GOOGLE_SHEETS_PROXY = "./api/google-sheets";
const AUTO_REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const COLORS = ["#0f766e", "#2563eb", "#d97706", "#e11d48", "#7c3aed", "#059669", "#334155", "#be185d", "#0ea5e9", "#a16207"];
const EMPTY_FILTERS = {
  startDate: "",
  endDate: "",
  platforms: [],
  channels: [],
  products: [],
  categories: [],
  companies: [],
  customerTypes: [],
  search: ""
};
const fieldSynonyms = {
  date: ["data", "date", "data pedido", "data venda", "dt venda"],
  month: ["mes", "mês", "month"],
  day: ["dia", "day"],
  quantity: ["qtd", "quantidade", "qtde", "quantity", "units", "unidades"],
  discountValue: ["desc comp", "desc. comp.", "desconto", "discount", "discounts"],
  channel: ["origem", "canal", "sales channel", "canal venda", "source"],
  itemNo: ["n item", "nº item", "numero item", "item", "item no"],
  unitValue: ["valor unitario", "valor unitário", "preco unitario", "preço unitário", "unit price"],
  platform: ["plataforma", "marketplace", "sales platform", "canal marketplace"],
  customer: ["cliente", "customer", "comprador"],
  sku: ["sku", "codigo sku", "cod sku"],
  product: ["produto", "product", "nome produto"],
  acos: ["acos", "ad cost", "ads cost"],
  orderId: ["pedido", "pedido id", "order", "order id", "numero pedido", "n pedido"],
  company: ["empresa", "marca", "brand", "company"],
  discountRate: ["custo % comp", "custo comp", "discount rate", "taxa desconto"],
  totalValue: ["valor total", "total", "receita", "revenue", "gross revenue"],
  productCost: ["custo produto", "product cost", "cogs", "custo"],
  platformFeeRate: ["taxa plataforma", "marketplace fee", "commission rate"],
  effectiveFeeRate: ["taxa comp", "taxa comp.", "effective fee", "taxa efetiva"],
  shipping: ["frete", "shipping", "shipping value"],
  taxRate: ["imposto", "tax", "tax rate"],
  operationRate: ["operacao", "operação", "operation", "operation rate"],
  profitUnit: ["lucro unitario", "lucro unitário", "unit profit"],
  profitTotal: ["lucro total", "profit", "gross margin", "margin value"],
  dailyGoal: ["meta diaria", "meta diária", "daily goal"],
  monthlyGoal: ["meta mes", "meta mês", "monthly goal"],
  status: ["status", "order status", "situacao", "situação"],
  paymentMethod: ["pagamento", "forma pagamento", "payment method"],
  region: ["estado", "uf", "regiao", "região", "state", "region"],
  seller: ["seller", "vendedor", "consultor"],
  category: ["categoria", "category", "product category"]
};
const currencyFormatter = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});
const currencyFormatterDetailed = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});
const numberFormatter = new Intl.NumberFormat("pt-BR", {
  maximumFractionDigits: 0
});
const percentFormatter = new Intl.NumberFormat("pt-BR", {
  style: "percent",
  minimumFractionDigits: 1,
  maximumFractionDigits: 1
});
function normalizeText(value) {
  return String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[º°]/g, "").replace(/[._/\\-]+/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
}
function cleanLabel(value, fallback = "Sem informação") {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  return text || fallback;
}
function titleCaseFallback(value) {
  return value.toLowerCase().split(" ").filter(Boolean).map(part => part.charAt(0).toUpperCase() + part.slice(1)).join(" ");
}
function normalizeChannel(value) {
  const label = cleanLabel(value);
  const normalized = normalizeText(label);
  if (normalized === "traego pago" || normalized === "trafego pago") return "Tráfego Pago";
  if (normalized === "organico") return "Orgânico";
  return label;
}
function normalizePlatform(value) {
  const label = cleanLabel(value, "Sem plataforma");
  const normalized = normalizeText(label);
  if (normalized === "full ml c" || normalized === "full ml p" || normalized === "full ml") {
    return "FULL Mercado Livre";
  }
  if (normalized === "mercado livre c" || normalized === "mercado livre p") {
    return "Mercado Livre";
  }
  if (normalized === "tiktok") return "TikTok";
  return label;
}
function parseNumber(value, defaultValue = 0, asPercent = false) {
  if (value === null || value === undefined || value === "") return defaultValue;
  if (typeof value === "number" && Number.isFinite(value)) {
    return asPercent && Math.abs(value) > 1 ? value / 100 : value;
  }
  let text = String(value).trim();
  if (!text) return defaultValue;
  const hasPercent = text.includes("%");
  text = text.replace(/\s/g, "").replace(/[^\d,.\-]/g, "");
  if (!text || text === "-" || text === "," || text === ".") return defaultValue;
  const comma = text.lastIndexOf(",");
  const dot = text.lastIndexOf(".");
  if (comma > -1 && dot > -1) {
    if (comma > dot) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      text = text.replace(/,/g, "");
    }
  } else if (comma > -1) {
    text = text.replace(",", ".");
  }
  const parsed = Number(text);
  if (!Number.isFinite(parsed)) return defaultValue;
  if (asPercent || hasPercent) return Math.abs(parsed) > 1 ? parsed / 100 : parsed;
  return parsed;
}
function excelSerialToDate(serial) {
  if (!Number.isFinite(serial)) return null;
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate());
}
function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === "number") {
    return excelSerialToDate(value);
  }
  const text = String(value).trim();
  if (!text) return null;
  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
  }
  const brMatch = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (brMatch) {
    const year = Number(brMatch[3].length === 2 ? `20${brMatch[3]}` : brMatch[3]);
    return new Date(year, Number(brMatch[2]) - 1, Number(brMatch[1]));
  }
  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}
function inputDateToDate(value) {
  if (!value) return null;
  const [year, month, day] = value.split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
}
function formatInputDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}
function formatShortDate(date) {
  return date.toLocaleDateString("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit"
  });
}
function addDays(date, amount) {
  const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  copy.setDate(copy.getDate() + amount);
  return copy;
}
function daysBetween(start, end) {
  const ms = new Date(end.getFullYear(), end.getMonth(), end.getDate()).getTime() - new Date(start.getFullYear(), start.getMonth(), start.getDate()).getTime();
  return Math.round(ms / 86400000);
}
function formatCurrency(value) {
  return currencyFormatter.format(value || 0);
}
function formatCurrencyDetailed(value) {
  return currencyFormatterDetailed.format(value || 0);
}
function formatNumber(value) {
  return numberFormatter.format(value || 0);
}
function formatPercent(value) {
  if (!Number.isFinite(value)) return "0,0%";
  return percentFormatter.format(value);
}
function compactCurrency(value) {
  return formatCurrencyDetailed(value || 0);
}
function detectColumns(headers) {
  const normalizedHeaders = headers.map(header => ({
    original: header,
    normalized: normalizeText(header)
  }));
  const mapping = {};
  Object.keys(fieldSynonyms).forEach(field => {
    const options = fieldSynonyms[field].map(normalizeText);
    const exact = normalizedHeaders.find(header => options.includes(header.normalized));
    if (exact) {
      mapping[field] = exact.original;
      return;
    }
    const fuzzy = normalizedHeaders.find(header => options.some(option => header.normalized.includes(option) || option.includes(header.normalized)));
    if (fuzzy) mapping[field] = fuzzy.original;
  });
  return mapping;
}
function rowValue(row, column) {
  if (!column) return undefined;
  if (Object.prototype.hasOwnProperty.call(row, column)) return row[column];
  const normalizedColumn = normalizeText(column);
  const key = Object.keys(row).find(candidate => normalizeText(candidate) === normalizedColumn);
  return key ? row[key] : undefined;
}
function deriveCategory(productValue, skuValue, itemValue, explicit) {
  if (explicit && explicit !== "Sem informação") return explicit;
  const product = normalizeText(productValue);
  const sku = normalizeText(skuValue);
  const item = normalizeText(itemValue);

  // Spreadsheet transformation point: the current file has no product category column.
  // This fallback keeps category analytics usable and can be replaced by ERP/category rules later.
  if (["term preto", "term laranja", "term branco"].includes(product)) return "Termômetros";
  if (/^(tk|tp)\d+/.test(product) || /^(tk|tp)\d+/.test(sku)) return "Interruptores";
  if (product.includes("hub")) return "Acessórios";
  if (product.includes("robo") || sku.startsWith("x10") || sku.startsWith("x20") || sku.startsWith("s10")) {
    return "Robôs";
  }
  if (product.includes("cartao") || sku.includes("ct")) return "Cartões TAG";
  if (product.includes("term")) return "Terminais";
  if (sku.startsWith("tk") || sku.startsWith("tp")) return "Acessórios";
  if (item.includes("he")) return "Terminais";
  if (item.includes("sh") || item.includes("ush")) return "Acessórios";
  if (item.includes("sl") || item.includes("usl")) return "Fechaduras";
  return "Sem categoria";
}

// DataParser: converts raw spreadsheet rows into normalized sales records used by KPIs,
// filters, charts, tables and future API/database adapters.
function mapRowsToSales(rawRows) {
  if (!rawRows.length) return [];
  const headers = Object.keys(rawRows[0] || {});
  const mapping = detectColumns(headers);
  if (!mapping.date || !mapping.totalValue) {
    throw new Error("A base precisa ter ao menos uma coluna de data e uma coluna de valor total.");
  }
  const records = rawRows.map((row, index) => {
    const date = parseDate(rowValue(row, mapping.date));
    if (!date) return null;
    const quantity = Math.max(0, parseNumber(rowValue(row, mapping.quantity), 1));
    const unitValue = parseNumber(rowValue(row, mapping.unitValue), 0);
    const totalValueRaw = parseNumber(rowValue(row, mapping.totalValue), unitValue * quantity);
    const totalValue = Number.isFinite(totalValueRaw) ? totalValueRaw : unitValue * quantity;
    const grossRevenue = unitValue && quantity ? unitValue * quantity : totalValue;
    const product = cleanLabel(rowValue(row, mapping.product), "Produto sem nome");
    const sku = cleanLabel(rowValue(row, mapping.sku), "Sem SKU");
    const itemNo = cleanLabel(rowValue(row, mapping.itemNo), "Sem item");
    const orderId = cleanLabel(rowValue(row, mapping.orderId), `linha-${index + 1}`);
    const customer = cleanLabel(rowValue(row, mapping.customer), "Cliente não informado");
    const platform = normalizePlatform(rowValue(row, mapping.platform));
    const channel = normalizeChannel(rowValue(row, mapping.channel));
    const company = cleanLabel(rowValue(row, mapping.company), "Sem empresa");
    const explicitCategory = cleanLabel(rowValue(row, mapping.category), "");
    const category = deriveCategory(product, sku, itemNo, explicitCategory);
    const discountValueUnit = parseNumber(rowValue(row, mapping.discountValue), 0);
    const discountRate = parseNumber(rowValue(row, mapping.discountRate), unitValue ? discountValueUnit / unitValue : 0, true);
    const productCostUnit = parseNumber(rowValue(row, mapping.productCost), 0);
    const platformFeeRate = parseNumber(rowValue(row, mapping.platformFeeRate), 0, true);
    const effectiveFeeRate = parseNumber(rowValue(row, mapping.effectiveFeeRate), platformFeeRate, true);
    const shippingUnit = parseNumber(rowValue(row, mapping.shipping), 0);
    const taxRate = parseNumber(rowValue(row, mapping.taxRate), 0, true);
    const operationRate = parseNumber(rowValue(row, mapping.operationRate), 0, true);
    const acosRate = parseNumber(rowValue(row, mapping.acos), 0, true);
    const feeAmount = totalValue * effectiveFeeRate;
    const baseFeeAmount = totalValue * platformFeeRate;
    const shippingTotal = shippingUnit * quantity;
    const productCostTotal = productCostUnit * quantity;
    const taxAmount = totalValue * taxRate;
    const operationAmount = totalValue * operationRate;
    const adCostAmount = totalValue * acosRate;
    const fallbackProfitUnit = unitValue - productCostUnit - shippingUnit - unitValue * effectiveFeeRate - unitValue * taxRate - unitValue * operationRate - unitValue * acosRate;
    const profitUnit = parseNumber(rowValue(row, mapping.profitUnit), fallbackProfitUnit);
    const profitTotal = parseNumber(rowValue(row, mapping.profitTotal), profitUnit * quantity);
    return {
      id: `${orderId}-${index}`,
      date,
      dateKey: formatInputDate(date),
      monthLabel: date.toLocaleDateString("pt-BR", {
        month: "short",
        year: "numeric"
      }),
      quantity,
      unitValue,
      totalValue,
      grossRevenue,
      discountValue: discountValueUnit * quantity,
      discountRate,
      channel,
      itemNo,
      platform,
      customer,
      customerType: "Novo",
      sku,
      product,
      category,
      acosRate,
      orderId,
      company,
      productCostUnit,
      productCostTotal,
      platformFeeRate,
      effectiveFeeRate,
      feeAmount,
      baseFeeAmount,
      shippingUnit,
      shippingTotal,
      taxRate,
      taxAmount,
      operationRate,
      operationAmount,
      adCostAmount,
      profitUnit,
      profitTotal,
      dailyGoal: parseNumber(rowValue(row, mapping.dailyGoal), 0),
      monthlyGoal: parseNumber(rowValue(row, mapping.monthlyGoal), 0),
      status: cleanLabel(rowValue(row, mapping.status), "Não disponível"),
      paymentMethod: cleanLabel(rowValue(row, mapping.paymentMethod), "Não disponível"),
      region: cleanLabel(rowValue(row, mapping.region), "Não disponível"),
      seller: cleanLabel(rowValue(row, mapping.seller), "Não disponível")
    };
  }).filter(Boolean);
  return attachCustomerTypes(records);
}
function attachCustomerTypes(records) {
  const firstPurchase = new Map();
  records.forEach(record => {
    const key = normalizeText(record.customer);
    const current = firstPurchase.get(key);
    const time = record.date.getTime();
    if (current === undefined || time < current) firstPurchase.set(key, time);
  });
  return records.map(record => {
    const first = firstPurchase.get(normalizeText(record.customer));
    return {
      ...record,
      customerType: first === record.date.getTime() ? "Novo" : "Recorrente"
    };
  });
}
function parseCsv(text) {
  const cleaned = text.replace(/^\ufeff/, "");
  const firstLine = cleaned.split(/\r?\n/, 1)[0] || "";
  const delimiter = firstLine.split(";").length > firstLine.split(",").length ? ";" : ",";
  const rows = [];
  let row = [];
  let value = "";
  let quoted = false;
  for (let index = 0; index < cleaned.length; index += 1) {
    const char = cleaned[index];
    const next = cleaned[index + 1];
    if (char === '"' && quoted && next === '"') {
      value += '"';
      index += 1;
      continue;
    }
    if (char === '"') {
      quoted = !quoted;
      continue;
    }
    if (char === delimiter && !quoted) {
      row.push(value);
      value = "";
      continue;
    }
    if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(value);
      if (row.some(cell => cell.trim() !== "")) rows.push(row);
      row = [];
      value = "";
      continue;
    }
    value += char;
  }
  row.push(value);
  if (row.some(cell => cell.trim() !== "")) rows.push(row);
  const headers = rows.shift()?.map(header => header.trim()) || [];
  return rows.map(cells => {
    const output = {};
    headers.forEach((header, index) => {
      output[header] = cells[index] ?? "";
    });
    return output;
  });
}

// SpreadsheetUploader/DataParser: reads Excel workbooks in the browser and returns
// plain row objects before the mapping layer applies column normalization.
async function parseExcelBuffer(buffer) {
  if (typeof XLSX === "undefined") {
    throw new Error("Biblioteca XLSX indisponível.");
  }
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    raw: true
  });
  const sheetName = workbook.SheetNames.includes("Base de Dados") ? "Base de Dados" : workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(worksheet, {
    defval: null,
    raw: true
  });
}
function googleSheetsCsvUrl() {
  return `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEETS_SPREADSHEET_ID}/gviz/tq?tqx=out:csv&gid=${GOOGLE_SHEETS_GID}&cacheBust=${Date.now()}`;
}
function googleSheetsFetchSources() {
  const directSource = {
    url: googleSheetsCsvUrl(),
    source: GOOGLE_SHEETS_SOURCE_LABEL
  };
  const host = typeof window !== "undefined" ? window.location.hostname : "";
  const protocol = typeof window !== "undefined" ? window.location.protocol : "";
  const localHosts = ["localhost", "127.0.0.1", "::1"];
  if (protocol.startsWith("http") && localHosts.includes(host)) {
    return [{
      url: `${GOOGLE_SHEETS_PROXY}?cacheBust=${Date.now()}`,
      source: GOOGLE_SHEETS_SOURCE_LABEL
    }, directSource];
  }
  return [directSource];
}
function looksLikeHtml(text) {
  return /^<!doctype html|^<html/i.test(text.trim());
}
async function fetchCsvSource(url, source) {
  const response = await fetch(url, {
    cache: "no-store"
  });
  if (!response.ok) {
    throw new Error(`Fonte retornou HTTP ${response.status}.`);
  }
  const csvText = await response.text();
  if (!csvText.trim() || looksLikeHtml(csvText)) {
    throw new Error("A fonte nÃ£o retornou CSV vÃ¡lido.");
  }
  const rows = parseCsv(csvText);
  if (!rows.length) {
    throw new Error("A planilha nÃ£o possui linhas de dados.");
  }
  return {
    rows,
    source,
    mode: "google"
  };
}
async function loadGoogleSheetsRows() {
  const errors = [];
  for (const source of googleSheetsFetchSources()) {
    try {
      return await fetchCsvSource(source.url, source.source);
    } catch (error) {
      errors.push(error?.message || "erro desconhecido");
    }
  }
  throw new Error(`NÃ£o foi possÃ­vel carregar o Google Sheets. ${errors.join(" | ")}`);
}
async function loadInitialRows() {
  let googleSheetsError = "";
  try {
    return await loadGoogleSheetsRows();
  } catch (error) {
    googleSheetsError = error?.message || "erro desconhecido";
  }
  const fallbackWarning = `NÃ£o foi possÃ­vel carregar o Google Sheets em tempo real (${googleSheetsError}). Usando a base local como fallback.`;
  const embeddedCsv = window.__BASE_DASHBOARD_CSV__;
  if (embeddedCsv) {
    return {
      rows: parseCsv(embeddedCsv),
      source: "Fallback local - BASE DASHBOARD.xlsx",
      warning: fallbackWarning,
      mode: "local"
    };
  }
  try {
    const response = await fetch(DASHBOARD_XLSX);
    if (!response.ok) throw new Error("Arquivo XLSX inicial não encontrado.");
    const rows = await parseExcelBuffer(await response.arrayBuffer());
    return {
      rows,
      source: "Fallback local - public/base-dashboard.xlsx",
      warning: fallbackWarning,
      mode: "local"
    };
  } catch (error) {
    try {
      const response = await fetch(DASHBOARD_CSV);
      if (!response.ok) throw error;
      const rows = parseCsv(await response.text());
      return {
        rows,
        source: "Fallback local - BASE DASHBOARD.xlsx",
        warning: fallbackWarning,
        mode: "local"
      };
    } catch {
      if (embeddedCsv) {
        return {
          rows: parseCsv(embeddedCsv),
          source: "Fallback local - BASE DASHBOARD.xlsx",
          warning: fallbackWarning,
          mode: "local"
        };
      }
      throw error;
    }
  }
}
async function parseUploadedFile(file) {
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (extension === "csv") {
    return parseCsv(await file.text());
  }
  return parseExcelBuffer(await file.arrayBuffer());
}
function uniqueValues(records, getter) {
  return Array.from(new Set(records.map(getter).filter(Boolean))).sort((a, b) => a.localeCompare(b, "pt-BR"));
}
function matchesList(selected, value) {
  return selected.length === 0 || selected.includes(value);
}
function applyDimensionFilters(records, filters) {
  const search = normalizeText(filters.search);
  return records.filter(record => {
    if (!matchesList(filters.platforms, record.platform)) return false;
    if (!matchesList(filters.channels, record.channel)) return false;
    if (!matchesList(filters.products, record.product)) return false;
    if (!matchesList(filters.categories, record.category)) return false;
    if (!matchesList(filters.companies, record.company)) return false;
    if (!matchesList(filters.customerTypes, record.customerType)) return false;
    if (!search) return true;
    const haystack = normalizeText([record.product, record.sku, record.orderId, record.customer, record.platform, record.company, record.category].join(" "));
    return haystack.includes(search);
  });
}
function filterByDate(records, startDate, endDate) {
  const start = inputDateToDate(startDate);
  const end = inputDateToDate(endDate);
  return records.filter(record => {
    if (start && record.date < start) return false;
    if (end && record.date > end) return false;
    return true;
  });
}
function calcMetrics(records) {
  const orderSet = new Set(records.map(record => record.orderId));
  const revenue = sum(records, record => record.totalValue);
  const orders = orderSet.size || 0;
  const units = sum(records, record => record.quantity);
  const discounts = sum(records, record => record.discountValue);
  const shipping = sum(records, record => record.shippingTotal);
  const fees = sum(records, record => record.feeAmount);
  const taxes = sum(records, record => record.taxAmount);
  const operation = sum(records, record => record.operationAmount);
  const ads = sum(records, record => record.adCostAmount);
  const productCost = sum(records, record => record.productCostTotal);
  const profit = sum(records, record => record.profitTotal);
  const netRevenue = revenue - discounts - fees - shipping - taxes - operation - ads;
  return {
    revenue,
    grossRevenue: sum(records, record => record.grossRevenue),
    orders,
    units,
    aov: orders ? revenue / orders : 0,
    discounts,
    shipping,
    fees,
    taxes,
    operation,
    ads,
    productCost,
    netRevenue,
    profit,
    margin: revenue ? profit / revenue : 0
  };
}
function sum(items, getter) {
  return items.reduce((total, item) => total + (getter(item) || 0), 0);
}
function previousRange(startDate, endDate) {
  const start = inputDateToDate(startDate);
  const end = inputDateToDate(endDate);
  if (!start || !end) return {
    start: "",
    end: ""
  };
  const length = daysBetween(start, end) + 1;
  const previousEnd = addDays(start, -1);
  const previousStart = addDays(previousEnd, -(length - 1));
  return {
    start: formatInputDate(previousStart),
    end: formatInputDate(previousEnd)
  };
}
function getDateBounds(records) {
  if (!records.length) return null;
  const times = records.map(record => record.date.getTime());
  return {
    min: new Date(Math.min(...times)),
    max: new Date(Math.max(...times))
  };
}
function getDefaultRange(records) {
  const bounds = getDateBounds(records);
  if (!bounds) return {
    startDate: "",
    endDate: ""
  };
  const start = addDays(bounds.max, -29) < bounds.min ? bounds.min : addDays(bounds.max, -29);
  return {
    startDate: formatInputDate(start),
    endDate: formatInputDate(bounds.max)
  };
}
function getMonthOptions(records) {
  const buckets = new Map();
  records.forEach(record => {
    const key = `${record.date.getFullYear()}-${String(record.date.getMonth() + 1).padStart(2, "0")}`;
    const current = buckets.get(key) || {
      dates: []
    };
    current.dates.push(record.date);
    buckets.set(key, current);
  });
  return Array.from(buckets.entries()).sort(([a], [b]) => a.localeCompare(b)).map(([key, bucket]) => {
    const times = bucket.dates.map(date => date.getTime());
    const [year, month] = key.split("-").map(Number);
    const monthName = new Date(year, month - 1, 1).toLocaleDateString("pt-BR", {
      month: "short"
    }).replace(".", "");
    return {
      key,
      label: `${titleCaseFallback(monthName)}/${String(year).slice(-2)}`,
      startDate: formatInputDate(new Date(Math.min(...times))),
      endDate: formatInputDate(new Date(Math.max(...times)))
    };
  });
}
function getIsoWeek(date) {
  const utc = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(utc.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((utc.getTime() - yearStart.getTime()) / 86400000 + 1) / 7);
  return {
    year: utc.getUTCFullYear(),
    week
  };
}
function periodKey(date, granularity) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  if (granularity === "day") {
    return {
      key: formatInputDate(date),
      label: formatShortDate(date)
    };
  }
  if (granularity === "week") {
    const week = getIsoWeek(date);
    return {
      key: `${week.year}-W${String(week.week).padStart(2, "0")}`,
      label: `S${String(week.week).padStart(2, "0")} ${week.year}`
    };
  }
  if (granularity === "quarter") {
    return {
      key: `${year}-Q${Math.ceil(month / 3)}`,
      label: `T${Math.ceil(month / 3)} ${year}`
    };
  }
  if (granularity === "year") {
    return {
      key: String(year),
      label: String(year)
    };
  }
  return {
    key: `${year}-${String(month).padStart(2, "0")}`,
    label: date.toLocaleDateString("pt-BR", {
      month: "short",
      year: "2-digit"
    })
  };
}
function groupByPeriod(records, granularity) {
  const buckets = new Map();
  records.forEach(record => {
    const period = periodKey(record.date, granularity);
    const current = buckets.get(period.key) || {
      key: period.key,
      label: period.label,
      date: record.date,
      revenue: 0,
      grossRevenue: 0,
      units: 0,
      profit: 0,
      orders: new Set()
    };
    current.revenue += record.totalValue;
    current.grossRevenue += record.grossRevenue;
    current.units += record.quantity;
    current.profit += record.profitTotal;
    current.orders.add(record.orderId);
    if (!current.date || record.date < current.date) current.date = record.date;
    buckets.set(period.key, current);
  });
  return Array.from(buckets.values()).sort((a, b) => a.key.localeCompare(b.key)).map(bucket => ({
    ...bucket,
    ordersCount: bucket.orders.size
  }));
}
function groupMonthlyRevenueByYear(records) {
  if (!records.length) return [];
  const years = Array.from(new Set(records.map(record => record.date.getFullYear()))).sort((a, b) => a - b);
  const hasMultipleYears = years.length > 1;
  const revenueByMonth = new Map();
  records.forEach(record => {
    const key = `${record.date.getFullYear()}-${String(record.date.getMonth() + 1).padStart(2, "0")}`;
    revenueByMonth.set(key, (revenueByMonth.get(key) || 0) + record.totalValue);
  });
  return years.flatMap(year => Array.from({
    length: 12
  }, (_, monthIndex) => {
    const key = `${year}-${String(monthIndex + 1).padStart(2, "0")}`;
    const label = new Date(year, monthIndex, 1).toLocaleDateString("pt-BR", {
      month: "short",
      year: hasMultipleYears ? "2-digit" : undefined
    });
    return {
      key,
      label,
      revenue: revenueByMonth.get(key) || 0
    };
  }));
}
function groupByDimension(records, getter) {
  const buckets = new Map();
  records.forEach(record => {
    const name = getter(record);
    const current = buckets.get(name) || {
      key: name,
      label: name,
      revenue: 0,
      grossRevenue: 0,
      units: 0,
      profit: 0,
      orders: new Set()
    };
    current.revenue += record.totalValue;
    current.grossRevenue += record.grossRevenue;
    current.units += record.quantity;
    current.profit += record.profitTotal;
    current.orders.add(record.orderId);
    buckets.set(name, current);
  });
  return Array.from(buckets.values()).map(bucket => ({
    ...bucket,
    name: bucket.label,
    ordersCount: bucket.orders.size,
    margin: bucket.revenue ? bucket.profit / bucket.revenue : 0,
    aov: bucket.orders.size ? bucket.revenue / bucket.orders.size : 0
  }));
}
function delta(current, previous) {
  const amount = current - previous;
  const percent = previous ? amount / Math.abs(previous) : current ? 1 : 0;
  return {
    amount,
    percent,
    className: amount > 0 ? "up" : amount < 0 ? "down" : "flat"
  };
}
function chartOptions(valueFormatter, stacked = false) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: {
      mode: "index",
      intersect: false
    },
    plugins: {
      legend: {
        labels: {
          color: "#475569",
          boxWidth: 11,
          usePointStyle: true
        }
      },
      tooltip: {
        callbacks: {
          label: context => {
            const label = context.dataset.label ? `${context.dataset.label}: ` : "";
            const value = typeof context.raw === "number" ? context.raw : Number(context.parsed.y ?? context.parsed.x ?? 0);
            return `${label}${valueFormatter(value)}`;
          }
        }
      }
    },
    scales: {
      x: {
        stacked,
        grid: {
          display: false
        },
        ticks: {
          color: "#64748b",
          maxRotation: 0,
          autoSkip: true
        }
      },
      y: {
        stacked,
        beginAtZero: true,
        grid: {
          color: "rgba(148, 163, 184, 0.22)"
        },
        ticks: {
          color: "#64748b",
          callback: value => valueFormatter(Number(value))
        }
      }
    }
  };
}
let monthlyValueLabelsRegistered = false;
function registerMonthlyValueLabelsPlugin() {
  if (monthlyValueLabelsRegistered || typeof Chart === "undefined") return;
  Chart.register({
    id: "monthlyValueLabels",
    afterDatasetsDraw(chart, _args, options) {
      if (!options?.enabled) return;
      const {
        ctx,
        chartArea
      } = chart;
      ctx.save();
      ctx.fillStyle = options.color || "#334155";
      ctx.font = options.font || "700 10px Inter, Segoe UI, sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";
      chart.data.datasets.forEach((dataset, datasetIndex) => {
        const meta = chart.getDatasetMeta(datasetIndex);
        meta.data.forEach((element, index) => {
          const value = Number(dataset.data[index] || 0);
          if (!value || !element) return;
          const y = Math.max(chartArea.top + 14, element.y - 6);
          ctx.fillText(formatCurrencyDetailed(value), element.x, y);
        });
      });
      ctx.restore();
    }
  });
  monthlyValueLabelsRegistered = true;
}
function monthlyRevenueOptions(maxRevenue) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        display: false
      },
      monthlyValueLabels: {
        enabled: true
      },
      tooltip: {
        callbacks: {
          label: context => `Faturamento total: ${formatCurrencyDetailed(Number(context.raw || 0))}`
        }
      }
    },
    scales: {
      x: {
        grid: {
          display: false
        },
        ticks: {
          color: "#64748b",
          maxRotation: 0,
          autoSkip: false
        }
      },
      y: {
        beginAtZero: true,
        suggestedMax: maxRevenue ? maxRevenue * 1.18 : undefined,
        grid: {
          color: "rgba(148, 163, 184, 0.22)"
        },
        ticks: {
          color: "#64748b",
          callback: value => formatCurrencyDetailed(Number(value))
        }
      }
    }
  };
}
function horizontalCurrencyOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    indexAxis: "y",
    plugins: {
      legend: {
        display: false
      },
      tooltip: {
        callbacks: {
          label: context => `${context.dataset.label}: ${formatCurrency(Number(context.raw || 0))}`
        }
      }
    },
    scales: {
      x: {
        beginAtZero: true,
        grid: {
          color: "rgba(148, 163, 184, 0.22)"
        },
        ticks: {
          color: "#64748b",
          callback: value => compactCurrency(Number(value))
        }
      },
      y: {
        grid: {
          display: false
        },
        ticks: {
          color: "#64748b"
        }
      }
    }
  };
}
function makeLineData(labels, revenue, previous) {
  const datasets = [{
    label: "Receita",
    data: revenue,
    borderColor: COLORS[0],
    backgroundColor: "rgba(15, 118, 110, 0.14)",
    tension: 0.28,
    fill: true,
    pointRadius: 2
  }];
  if (previous) {
    datasets.push({
      label: "Período anterior",
      data: previous,
      borderColor: COLORS[3],
      backgroundColor: "rgba(225, 29, 72, 0.08)",
      borderDash: [6, 5],
      tension: 0.28,
      fill: false,
      pointRadius: 0
    });
  }
  return {
    labels,
    datasets
  };
}
function rangeDescription(start, end) {
  const startDate = inputDateToDate(start);
  const endDate = inputDateToDate(end);
  if (!startDate || !endDate) return "Período não definido";
  return `${formatShortDate(startDate)} a ${formatShortDate(endDate)}`;
}
function App() {
  const fileInputRef = useRef(null);
  const [records, setRecords] = useState([]);
  const [filters, setFilters] = useState(EMPTY_FILTERS);
  const [granularity, setGranularity] = useState("day");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [warning, setWarning] = useState("");
  const [sourceName, setSourceName] = useState(GOOGLE_SHEETS_SOURCE_LABEL);
  const [dataSourceMode, setDataSourceMode] = useState("google");
  const [lastSyncAt, setLastSyncAt] = useState(null);
  const [rankMode, setRankMode] = useState("best");
  const [comparison, setComparison] = useState({
    aStart: "",
    aEnd: "",
    bStart: "",
    bEnd: ""
  });
  useEffect(() => {
    let mounted = true;
    setLoading(true);
    loadInitialRows().then(({
      rows,
      source,
      warning: loadWarning,
      mode
    }) => {
      if (!mounted) return;
      const normalized = mapRowsToSales(rows);
      setRecords(normalized);
      setSourceName(source);
      setDataSourceMode(mode);
      setLastSyncAt(new Date());
      setWarning(loadWarning || "");
      const defaults = getDefaultRange(normalized);
      const previous = previousRange(defaults.startDate, defaults.endDate);
      setFilters(current => ({
        ...current,
        ...defaults
      }));
      setComparison({
        aStart: defaults.startDate,
        aEnd: defaults.endDate,
        bStart: previous.start,
        bEnd: previous.end
      });
    }).catch(loadError => {
      if (!mounted) return;
      setError(loadError?.message || "Não foi possível carregar a base inicial.");
    }).finally(() => {
      if (mounted) setLoading(false);
    });
    return () => {
      mounted = false;
    };
  }, []);
  async function handleFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setError("");
    setWarning("");
    try {
      const rows = await parseUploadedFile(file);
      const normalized = mapRowsToSales(rows);
      const defaults = getDefaultRange(normalized);
      const previous = previousRange(defaults.startDate, defaults.endDate);
      setRecords(normalized);
      setSourceName(file.name);
      setDataSourceMode("upload");
      setLastSyncAt(new Date());
      setFilters({
        ...EMPTY_FILTERS,
        ...defaults
      });
      setComparison({
        aStart: defaults.startDate,
        aEnd: defaults.endDate,
        bStart: previous.start,
        bEnd: previous.end
      });
    } catch (uploadError) {
      setError(uploadError?.message || "Não foi possível processar o arquivo enviado.");
    } finally {
      setLoading(false);
      event.target.value = "";
    }
  }
  async function handleGoogleRefresh() {
    setLoading(true);
    setError("");
    setWarning("");
    try {
      const {
        rows,
        source,
        warning: loadWarning,
        mode
      } = await loadGoogleSheetsRows();
      const normalized = mapRowsToSales(rows);
      const defaults = getDefaultRange(normalized);
      const previous = previousRange(defaults.startDate, defaults.endDate);
      setRecords(normalized);
      setSourceName(source);
      setDataSourceMode(mode);
      setLastSyncAt(new Date());
      setWarning(loadWarning || "");
      setFilters(current => current.startDate && current.endDate ? current : {
        ...current,
        ...defaults
      });
      setComparison(current => current.aStart && current.aEnd && current.bStart && current.bEnd ? current : {
        aStart: defaults.startDate,
        aEnd: defaults.endDate,
        bStart: previous.start,
        bEnd: previous.end
      });
    } catch (refreshError) {
      setError(refreshError?.message || "NÃ£o foi possÃ­vel atualizar os dados do Google Sheets.");
    } finally {
      setLoading(false);
    }
  }
  useEffect(() => {
    if (dataSourceMode === "upload") return undefined;
    let disposed = false;
    const interval = window.setInterval(async () => {
      try {
        const {
          rows,
          source,
          mode
        } = await loadGoogleSheetsRows();
        if (disposed) return;
        setRecords(mapRowsToSales(rows));
        setSourceName(source);
        setDataSourceMode(mode);
        setLastSyncAt(new Date());
        setWarning("");
      } catch (refreshError) {
        if (disposed) return;
        setWarning(`AtualizaÃ§Ã£o automÃ¡tica do Google Sheets falhou: ${refreshError?.message || "erro desconhecido"}`);
      }
    }, AUTO_REFRESH_INTERVAL_MS);
    return () => {
      disposed = true;
      window.clearInterval(interval);
    };
  }, [dataSourceMode]);
  const dateBounds = useMemo(() => getDateBounds(records), [records]);
  const dimensionFiltered = useMemo(() => applyDimensionFilters(records, filters), [records, filters]);
  const currentRecords = useMemo(() => filterByDate(dimensionFiltered, filters.startDate, filters.endDate), [dimensionFiltered, filters.startDate, filters.endDate]);
  const previous = useMemo(() => previousRange(filters.startDate, filters.endDate), [filters.startDate, filters.endDate]);
  const previousRecords = useMemo(() => filterByDate(dimensionFiltered, previous.start, previous.end), [dimensionFiltered, previous.start, previous.end]);
  const metrics = useMemo(() => calcMetrics(currentRecords), [currentRecords]);
  const previousMetrics = useMemo(() => calcMetrics(previousRecords), [previousRecords]);
  const monthOptions = useMemo(() => getMonthOptions(records), [records]);
  const options = useMemo(() => ({
    platforms: uniqueValues(records, record => record.platform),
    channels: uniqueValues(records, record => record.channel),
    products: uniqueValues(records, record => record.product),
    categories: uniqueValues(records, record => record.category),
    companies: uniqueValues(records, record => record.company),
    customerTypes: ["Novo", "Recorrente"]
  }), [records]);
  const periodBuckets = useMemo(() => groupByPeriod(currentRecords, granularity), [currentRecords, granularity]);
  const previousBuckets = useMemo(() => groupByPeriod(previousRecords, granularity), [previousRecords, granularity]);
  const monthlyRevenueBuckets = useMemo(() => groupMonthlyRevenueByYear(dimensionFiltered), [dimensionFiltered]);
  const weeklyBuckets = useMemo(() => groupByPeriod(currentRecords, "week"), [currentRecords]);
  const dailyBuckets = useMemo(() => groupByPeriod(currentRecords, "day"), [currentRecords]);
  const platformBuckets = useMemo(() => groupByDimension(currentRecords, record => record.platform).sort((a, b) => b.revenue - a.revenue), [currentRecords]);
  const categoryBuckets = useMemo(() => groupByDimension(currentRecords, record => record.category).sort((a, b) => b.revenue - a.revenue), [currentRecords]);
  const channelBuckets = useMemo(() => groupByDimension(currentRecords, record => record.channel).sort((a, b) => b.revenue - a.revenue), [currentRecords]);
  const productBuckets = useMemo(() => groupByDimension(currentRecords, record => `${record.product}||${record.sku}||${record.category}`).map(bucket => {
    const [product, sku, category] = bucket.name.split("||");
    return {
      ...bucket,
      product,
      sku,
      category
    };
  }).sort((a, b) => rankMode === "best" ? b.revenue - a.revenue : a.profit - b.profit), [currentRecords, rankMode]);
  const comparisonMetrics = useMemo(() => {
    const a = calcMetrics(filterByDate(dimensionFiltered, comparison.aStart, comparison.aEnd));
    const b = calcMetrics(filterByDate(dimensionFiltered, comparison.bStart, comparison.bEnd));
    return {
      a,
      b
    };
  }, [dimensionFiltered, comparison]);
  const previousRevenueByIndex = previousBuckets.map(bucket => bucket.revenue);
  const kpiDeltas = {
    revenue: delta(metrics.revenue, previousMetrics.revenue),
    orders: delta(metrics.orders, previousMetrics.orders),
    units: delta(metrics.units, previousMetrics.units),
    profit: delta(metrics.profit, previousMetrics.profit)
  };
  const lineChartData = makeLineData(periodBuckets.map(bucket => bucket.label), periodBuckets.map(bucket => bucket.revenue), previousRevenueByIndex.length ? previousRevenueByIndex : undefined);
  const ordersChartData = {
    labels: periodBuckets.map(bucket => bucket.label),
    datasets: [{
      type: "line",
      label: "Faturamento",
      data: periodBuckets.map(bucket => bucket.revenue),
      borderColor: COLORS[0],
      backgroundColor: "rgba(15, 118, 110, 0.12)",
      tension: 0.25,
      pointRadius: 2,
      yAxisID: "y"
    }, {
      label: "Unidades",
      data: periodBuckets.map(bucket => bucket.units),
      backgroundColor: "rgba(217, 119, 6, 0.68)",
      borderRadius: 6,
      yAxisID: "y1"
    }]
  };
  const platformChartData = {
    labels: platformBuckets.slice(0, 10).map(bucket => {
      const total = sum(platformBuckets, item => item.revenue);
      const percent = total ? bucket.revenue / total : 0;
      return `${bucket.name} (${formatPercent(percent)})`;
    }),
    datasets: [{
      label: "Receita",
      data: platformBuckets.slice(0, 10).map(bucket => bucket.revenue),
      backgroundColor: platformBuckets.slice(0, 10).map((_, index) => COLORS[index % COLORS.length]),
      borderRadius: 6
    }]
  };
  const categoryChartData = {
    labels: categoryBuckets.map(bucket => bucket.name),
    datasets: [{
      label: "Receita",
      data: categoryBuckets.map(bucket => bucket.revenue),
      backgroundColor: categoryBuckets.map((_, index) => COLORS[index % COLORS.length]),
      borderWidth: 0
    }]
  };
  const productChartData = {
    labels: productBuckets.slice(0, 12).map(bucket => bucket.product),
    datasets: [{
      label: rankMode === "best" ? "Receita" : "Lucro",
      data: productBuckets.slice(0, 12).map(bucket => rankMode === "best" ? bucket.revenue : bucket.profit),
      backgroundColor: productBuckets.slice(0, 12).map(bucket => bucket.profit >= 0 ? "rgba(5, 150, 105, 0.72)" : "rgba(225, 29, 72, 0.72)"),
      borderRadius: 6
    }]
  };
  const channelChartData = {
    labels: channelBuckets.map(bucket => bucket.name),
    datasets: [{
      label: "Receita",
      data: channelBuckets.map(bucket => bucket.revenue),
      backgroundColor: channelBuckets.map((_, index) => COLORS[index % COLORS.length]),
      borderRadius: 6
    }]
  };
  const monthlyChartData = {
    labels: monthlyRevenueBuckets.map(bucket => bucket.label),
    datasets: [{
      label: "Faturamento total",
      data: monthlyRevenueBuckets.map(bucket => bucket.revenue),
      backgroundColor: "rgba(15, 118, 110, 0.72)",
      borderRadius: 6
    }]
  };
  const weeklyChartData = {
    labels: weeklyBuckets.map(bucket => bucket.label),
    datasets: [{
      label: "Receita",
      data: weeklyBuckets.map(bucket => bucket.revenue),
      borderColor: COLORS[1],
      backgroundColor: "rgba(37, 99, 235, 0.12)",
      tension: 0.22,
      fill: true
    }]
  };
  const dailyTrendData = {
    labels: dailyBuckets.map(bucket => bucket.label),
    datasets: [{
      label: "Receita diária",
      data: dailyBuckets.map(bucket => bucket.revenue),
      borderColor: COLORS[2],
      backgroundColor: "rgba(217, 119, 6, 0.14)",
      tension: 0.18,
      fill: true,
      pointRadius: 1.5
    }]
  };
  const scatterData = {
    datasets: [{
      label: "Dias",
      data: dailyBuckets.map(bucket => ({
        x: bucket.ordersCount,
        y: bucket.revenue,
        label: bucket.label
      })),
      backgroundColor: "rgba(15, 118, 110, 0.72)",
      borderColor: COLORS[0]
    }]
  };
  const periodLabel = rangeDescription(filters.startDate, filters.endDate);
  const previousLabel = rangeDescription(previous.start, previous.end);
  const monthlyMaxRevenue = Math.max(...monthlyRevenueBuckets.map(bucket => bucket.revenue), 0);
  return /*#__PURE__*/React.createElement("div", {
    className: "app-shell"
  }, /*#__PURE__*/React.createElement("header", {
    className: "topbar"
  }, /*#__PURE__*/React.createElement("div", {
    className: "topbar-inner"
  }, /*#__PURE__*/React.createElement("div", {
    className: "brand-block"
  }, /*#__PURE__*/React.createElement("p", {
    className: "eyebrow"
  }, "E-commerce e marketplaces"), /*#__PURE__*/React.createElement("h1", {
    className: "brand-title"
  }, "Sales Dashboard"), /*#__PURE__*/React.createElement("div", {
    className: "brand-meta"
  }, /*#__PURE__*/React.createElement("span", null, sourceName), /*#__PURE__*/React.createElement("span", null, formatNumber(records.length), " linhas"), dateBounds && /*#__PURE__*/React.createElement("span", null, formatShortDate(dateBounds.min), " a ", formatShortDate(dateBounds.max)), lastSyncAt && /*#__PURE__*/React.createElement("span", null, "Sincronizado ", lastSyncAt.toLocaleTimeString("pt-BR", {
    hour: "2-digit",
    minute: "2-digit"
  })))), /*#__PURE__*/React.createElement("div", {
    className: "topbar-actions"
  }, /*#__PURE__*/React.createElement("span", {
    className: "status-pill"
  }, /*#__PURE__*/React.createElement("span", {
    className: `status-dot ${error ? "error" : loading ? "warn" : ""}`
  }), loading ? "Carregando" : error ? "Erro" : "Base ativa"), /*#__PURE__*/React.createElement("button", {
    className: "btn",
    onClick: handleGoogleRefresh,
    disabled: loading
  }, /*#__PURE__*/React.createElement("span", {
    "aria-hidden": "true"
  }, "\xE2\u2020\xBB"), "Atualizar Sheets"), /*#__PURE__*/React.createElement("button", {
    className: "btn btn-primary",
    onClick: () => fileInputRef.current?.click()
  }, /*#__PURE__*/React.createElement("span", {
    "aria-hidden": "true"
  }, "\u21A5"), "Carregar planilha"), /*#__PURE__*/React.createElement("input", {
    ref: fileInputRef,
    className: "hidden-file",
    type: "file",
    accept: ".xlsx,.xls,.csv",
    onChange: handleFileChange
  })))), /*#__PURE__*/React.createElement("main", {
    className: "main-grid"
  }, /*#__PURE__*/React.createElement(FiltersPanel, {
    filters: filters,
    setFilters: setFilters,
    options: options,
    monthOptions: monthOptions,
    comparison: comparison,
    setComparison: setComparison,
    dateBounds: dateBounds,
    reset: () => {
      const defaults = getDefaultRange(records);
      setFilters({
        ...EMPTY_FILTERS,
        ...defaults
      });
    }
  }), /*#__PURE__*/React.createElement("div", {
    className: "content-area"
  }, error && /*#__PURE__*/React.createElement("div", {
    className: "alert"
  }, error), warning && /*#__PURE__*/React.createElement("div", {
    className: "alert"
  }, warning), loading ? /*#__PURE__*/React.createElement("div", {
    className: "loading-state"
  }, "Processando dados da planilha") : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("section", {
    className: "section"
  }, /*#__PURE__*/React.createElement("div", {
    className: "section-header"
  }, /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("h2", {
    className: "section-title"
  }, "Resumo executivo"), /*#__PURE__*/React.createElement("p", {
    className: "section-subtitle"
  }, "Per\xEDodo atual: ", periodLabel, " \xB7 Comparativo: ", previousLabel)), /*#__PURE__*/React.createElement(GranularitySelector, {
    value: granularity,
    onChange: setGranularity
  })), /*#__PURE__*/React.createElement("div", {
    className: "kpi-grid"
  }, /*#__PURE__*/React.createElement(KPICard, {
    label: "Receita total",
    value: formatCurrency(metrics.revenue),
    delta: kpiDeltas.revenue,
    foot: `Dif.: ${formatCurrency(kpiDeltas.revenue.amount)}`
  }), /*#__PURE__*/React.createElement(KPICard, {
    label: "Pedidos",
    value: formatNumber(metrics.orders),
    delta: kpiDeltas.orders,
    foot: `Dif.: ${formatNumber(kpiDeltas.orders.amount)}`
  }), /*#__PURE__*/React.createElement(KPICard, {
    label: "Unidades vendidas",
    value: formatNumber(metrics.units),
    delta: kpiDeltas.units,
    foot: `Ticket médio: ${formatCurrency(metrics.aov)}`
  }), /*#__PURE__*/React.createElement(KPICard, {
    label: "Taxas marketplace",
    value: formatCurrency(metrics.fees),
    foot: `Frete: ${formatCurrency(metrics.shipping)}`
  }))), /*#__PURE__*/React.createElement(DateComparisonSelector, {
    comparison: comparison,
    setComparison: setComparison,
    metricsA: comparisonMetrics.a,
    metricsB: comparisonMetrics.b
  }), /*#__PURE__*/React.createElement("section", {
    className: "chart-grid"
  }, /*#__PURE__*/React.createElement(ChartCard, {
    title: "Receita ao longo do tempo",
    subtitle: "Atual versus per\xEDodo anterior de mesma dura\xE7\xE3o",
    type: "line",
    data: lineChartData,
    options: chartOptions(compactCurrency)
  }), /*#__PURE__*/React.createElement(ChartCard, {
    title: "Faturamento x unidades",
    subtitle: "Receita e volume agrupados pelo per\xEDodo selecionado",
    type: "bar",
    data: ordersChartData,
    options: dualAxisOptions()
  }), /*#__PURE__*/React.createElement(ChartCard, {
    title: "Vendas por plataforma",
    subtitle: "Top marketplaces e canais por receita",
    type: "bar",
    data: platformChartData,
    options: horizontalCurrencyOptions()
  }), /*#__PURE__*/React.createElement(ChartCard, {
    title: "Vendas por categoria",
    subtitle: "Categoria expl\xEDcita ou derivada do produto/SKU",
    type: "doughnut",
    data: categoryChartData,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "right",
          labels: {
            color: "#475569",
            boxWidth: 11,
            usePointStyle: true
          }
        },
        tooltip: {
          callbacks: {
            label: context => `${context.label}: ${formatCurrency(Number(context.raw || 0))}`
          }
        }
      }
    }
  }), /*#__PURE__*/React.createElement(ChartCard, {
    title: rankMode === "best" ? "Produtos mais vendidos" : "Produtos com pior lucro",
    subtitle: "Ranking por receita ou lucro no per\xEDodo atual",
    type: "bar",
    data: productChartData,
    options: horizontalCurrencyOptions(),
    action: /*#__PURE__*/React.createElement(RankToggle, {
      value: rankMode,
      onChange: setRankMode
    })
  }), /*#__PURE__*/React.createElement(ChartCard, {
    title: "Receita por canal",
    subtitle: "Distribui\xE7\xE3o de faturamento por origem da venda",
    type: "bar",
    data: channelChartData,
    options: horizontalCurrencyOptions()
  }), /*#__PURE__*/React.createElement(ChartCard, {
    title: "Faturamento mensal do ano",
    subtitle: "Receita total por m\xEAs, respeitando os filtros de plataforma, produto e canal",
    type: "bar",
    data: monthlyChartData,
    options: monthlyRevenueOptions(monthlyMaxRevenue),
    wide: true
  }), /*#__PURE__*/React.createElement(ChartCard, {
    title: "Comparativo semanal",
    subtitle: "Receita por semana ISO",
    type: "line",
    data: weeklyChartData,
    options: chartOptions(compactCurrency)
  }), /*#__PURE__*/React.createElement(ChartCard, {
    title: "Receita x pedidos",
    subtitle: "Cada ponto representa um dia de venda",
    type: "scatter",
    data: scatterData,
    options: scatterOptions()
  })), /*#__PURE__*/React.createElement("section", {
    className: "section"
  }, /*#__PURE__*/React.createElement(ChartCard, {
    title: "Tend\xEAncia di\xE1ria de vendas",
    subtitle: "Receita di\xE1ria dentro do per\xEDodo filtrado",
    type: "line",
    data: dailyTrendData,
    options: chartOptions(compactCurrency),
    short: true
  })), /*#__PURE__*/React.createElement("section", {
    className: "section"
  }, /*#__PURE__*/React.createElement(SalesHeatmap, {
    records: currentRecords
  })), /*#__PURE__*/React.createElement("section", {
    className: "table-grid"
  }, /*#__PURE__*/React.createElement(ProductRankingTable, {
    rows: productBuckets.slice(0, 14),
    rankMode: rankMode,
    setRankMode: setRankMode
  }), /*#__PURE__*/React.createElement(PlatformPerformanceTable, {
    rows: platformBuckets
  }))))));
}
function dualAxisOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          color: "#475569",
          boxWidth: 11,
          usePointStyle: true
        }
      },
      tooltip: {
        callbacks: {
          label: context => {
            const value = Number(context.raw || 0);
            return context.dataset.label === "Unidades" ? `Unidades: ${formatNumber(value)}` : `${context.dataset.label}: ${formatCurrency(value)}`;
          }
        }
      }
    },
    scales: {
      x: {
        grid: {
          display: false
        },
        ticks: {
          color: "#64748b"
        }
      },
      y: {
        beginAtZero: true,
        position: "left",
        grid: {
          color: "rgba(148, 163, 184, 0.22)"
        },
        ticks: {
          color: "#64748b",
          callback: value => compactCurrency(Number(value))
        }
      },
      y1: {
        beginAtZero: true,
        position: "right",
        grid: {
          drawOnChartArea: false
        },
        ticks: {
          color: "#64748b",
          callback: value => formatNumber(Number(value))
        }
      }
    }
  };
}
function scatterOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        display: false
      },
      tooltip: {
        callbacks: {
          label: context => {
            const raw = context.raw || {};
            return `${raw.label || "Dia"}: ${formatNumber(context.parsed.x)} pedidos · ${formatCurrency(context.parsed.y)}`;
          }
        }
      }
    },
    scales: {
      x: {
        beginAtZero: true,
        title: {
          display: true,
          text: "Pedidos",
          color: "#64748b"
        },
        grid: {
          color: "rgba(148, 163, 184, 0.18)"
        },
        ticks: {
          color: "#64748b"
        }
      },
      y: {
        beginAtZero: true,
        title: {
          display: true,
          text: "Receita",
          color: "#64748b"
        },
        grid: {
          color: "rgba(148, 163, 184, 0.22)"
        },
        ticks: {
          color: "#64748b",
          callback: value => compactCurrency(Number(value))
        }
      }
    }
  };
}
function FiltersPanel({
  filters,
  setFilters,
  options,
  monthOptions,
  comparison,
  setComparison,
  dateBounds,
  reset
}) {
  const min = dateBounds ? formatInputDate(dateBounds.min) : undefined;
  const max = dateBounds ? formatInputDate(dateBounds.max) : undefined;
  return /*#__PURE__*/React.createElement("aside", {
    className: "filters-panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-header"
  }, /*#__PURE__*/React.createElement("h2", {
    className: "panel-title"
  }, "Filtros"), /*#__PURE__*/React.createElement("p", {
    className: "panel-subtitle"
  }, "A sele\xE7\xE3o atual recalcula KPIs, gr\xE1ficos e tabelas em tempo real.")), /*#__PURE__*/React.createElement("div", {
    className: "filters-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("span", {
    className: "field-label"
  }, "Per\xEDodo"), /*#__PURE__*/React.createElement("div", {
    className: "field-row"
  }, /*#__PURE__*/React.createElement("input", {
    className: "input",
    type: "date",
    min: min,
    max: max,
    value: filters.startDate,
    onChange: event => setFilters(current => ({
      ...current,
      startDate: event.target.value
    }))
  }), /*#__PURE__*/React.createElement("input", {
    className: "input",
    type: "date",
    min: min,
    max: max,
    value: filters.endDate,
    onChange: event => setFilters(current => ({
      ...current,
      endDate: event.target.value
    }))
  }))), /*#__PURE__*/React.createElement(MonthComparisonFilter, {
    monthOptions: monthOptions,
    comparison: comparison,
    setComparison: setComparison
  }), /*#__PURE__*/React.createElement(MonthQuickFilter, {
    monthOptions: monthOptions,
    filters: filters,
    setFilters: setFilters,
    dateBounds: dateBounds
  }), /*#__PURE__*/React.createElement(MultiSelect, {
    label: "Plataforma",
    options: options.platforms,
    selected: filters.platforms,
    onChange: value => setFilters(current => ({
      ...current,
      platforms: value
    }))
  }), /*#__PURE__*/React.createElement(MultiSelect, {
    label: "Canal",
    options: options.channels,
    selected: filters.channels,
    onChange: value => setFilters(current => ({
      ...current,
      channels: value
    }))
  }), /*#__PURE__*/React.createElement(MultiSelect, {
    label: "Produto",
    options: options.products,
    selected: filters.products,
    onChange: value => setFilters(current => ({
      ...current,
      products: value
    })),
    searchable: true
  }), /*#__PURE__*/React.createElement(MultiSelect, {
    label: "Categoria",
    options: options.categories,
    selected: filters.categories,
    onChange: value => setFilters(current => ({
      ...current,
      categories: value
    }))
  }), /*#__PURE__*/React.createElement(MultiSelect, {
    label: "Empresa",
    options: options.companies,
    selected: filters.companies,
    onChange: value => setFilters(current => ({
      ...current,
      companies: value
    }))
  }), /*#__PURE__*/React.createElement(MultiSelect, {
    label: "Tipo cliente",
    options: options.customerTypes,
    selected: filters.customerTypes,
    onChange: value => setFilters(current => ({
      ...current,
      customerTypes: value
    }))
  }), /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("label", {
    htmlFor: "global-search"
  }, "Busca"), /*#__PURE__*/React.createElement("input", {
    id: "global-search",
    className: "input",
    value: filters.search,
    placeholder: "Pedido, SKU, cliente...",
    onChange: event => setFilters(current => ({
      ...current,
      search: event.target.value
    }))
  })), /*#__PURE__*/React.createElement("div", {
    className: "filter-footer"
  }, /*#__PURE__*/React.createElement("button", {
    className: "btn",
    type: "button",
    onClick: reset
  }, "Limpar"))));
}
function MonthComparisonFilter({
  monthOptions,
  comparison,
  setComparison
}) {
  if (!monthOptions.length) return null;
  const selectedA = monthOptions.find(month => month.startDate === comparison.aStart && month.endDate === comparison.aEnd)?.key || "";
  const selectedB = monthOptions.find(month => month.startDate === comparison.bStart && month.endDate === comparison.bEnd)?.key || "";
  function updateMonth(slot, key) {
    const selected = monthOptions.find(month => month.key === key);
    if (!selected) return;
    setComparison(current => slot === "a" ? {
      ...current,
      aStart: selected.startDate,
      aEnd: selected.endDate
    } : {
      ...current,
      bStart: selected.startDate,
      bEnd: selected.endDate
    });
  }
  return /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("span", {
    className: "field-label"
  }, "Comparativo mensal"), /*#__PURE__*/React.createElement("div", {
    className: "month-comparison-grid"
  }, /*#__PURE__*/React.createElement("label", {
    className: "select-field"
  }, /*#__PURE__*/React.createElement("span", null, "M\xEAs A"), /*#__PURE__*/React.createElement("select", {
    className: "input",
    value: selectedA,
    onChange: event => updateMonth("a", event.target.value)
  }, /*#__PURE__*/React.createElement("option", {
    value: "",
    disabled: true
  }, "Selecione"), monthOptions.map(month => /*#__PURE__*/React.createElement("option", {
    key: month.key,
    value: month.key
  }, month.label)))), /*#__PURE__*/React.createElement("label", {
    className: "select-field"
  }, /*#__PURE__*/React.createElement("span", null, "M\xEAs B"), /*#__PURE__*/React.createElement("select", {
    className: "input",
    value: selectedB,
    onChange: event => updateMonth("b", event.target.value)
  }, /*#__PURE__*/React.createElement("option", {
    value: "",
    disabled: true
  }, "Selecione"), monthOptions.map(month => /*#__PURE__*/React.createElement("option", {
    key: month.key,
    value: month.key
  }, month.label))))));
}
function MonthQuickFilter({
  monthOptions,
  filters,
  setFilters,
  dateBounds
}) {
  if (!monthOptions.length) return null;
  const fullStart = dateBounds ? formatInputDate(dateBounds.min) : "";
  const fullEnd = dateBounds ? formatInputDate(dateBounds.max) : "";
  const isAllActive = filters.startDate === fullStart && filters.endDate === fullEnd;
  return /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("span", {
    className: "field-label"
  }, "M\xEAs"), /*#__PURE__*/React.createElement("div", {
    className: "month-button-grid"
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: `month-button ${isAllActive ? "active" : ""}`,
    onClick: () => setFilters(current => ({
      ...current,
      startDate: fullStart,
      endDate: fullEnd
    }))
  }, "Todos"), monthOptions.map(month => {
    const active = filters.startDate === month.startDate && filters.endDate === month.endDate;
    return /*#__PURE__*/React.createElement("button", {
      type: "button",
      className: `month-button ${active ? "active" : ""}`,
      key: month.key,
      onClick: () => setFilters(current => ({
        ...current,
        startDate: month.startDate,
        endDate: month.endDate
      }))
    }, month.label);
  })));
}
function MultiSelect({
  label,
  options,
  selected,
  onChange,
  searchable = false
}) {
  const [search, setSearch] = useState("");
  const visibleOptions = options.filter(option => normalizeText(option).includes(normalizeText(search))).slice(0, 160);
  const summary = selected.length === 0 ? "Todos" : `${selected.length} selecionado${selected.length > 1 ? "s" : ""}`;
  function toggle(value) {
    if (selected.includes(value)) {
      onChange(selected.filter(item => item !== value));
    } else {
      onChange([...selected, value]);
    }
  }
  return /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("span", {
    className: "field-label"
  }, label), /*#__PURE__*/React.createElement("details", {
    className: "multi-select"
  }, /*#__PURE__*/React.createElement("summary", null, summary, /*#__PURE__*/React.createElement("span", {
    "aria-hidden": "true"
  }, "\u2304")), /*#__PURE__*/React.createElement("div", {
    className: "multi-select-body"
  }, searchable && /*#__PURE__*/React.createElement("input", {
    className: "input search-input",
    value: search,
    placeholder: "Filtrar op\xE7\xF5es",
    onChange: event => setSearch(event.target.value)
  }), /*#__PURE__*/React.createElement("label", {
    className: "check-row"
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: selected.length === 0,
    onChange: () => onChange([])
  }), /*#__PURE__*/React.createElement("span", null, "Todos")), visibleOptions.map(option => /*#__PURE__*/React.createElement("label", {
    className: "check-row",
    key: option,
    title: option
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: selected.includes(option),
    onChange: () => toggle(option)
  }), /*#__PURE__*/React.createElement("span", null, option))))));
}
function GranularitySelector({
  value,
  onChange
}) {
  const items = [["day", "Dia"], ["week", "Semana"], ["month", "Mês"], ["quarter", "Trimestre"], ["year", "Ano"]];
  return /*#__PURE__*/React.createElement("div", {
    className: "segmented",
    "aria-label": "Agrupamento"
  }, items.map(([key, label]) => /*#__PURE__*/React.createElement("button", {
    key: key,
    className: `segment ${value === key ? "active" : ""}`,
    onClick: () => onChange(key)
  }, label)));
}
function KPICard({
  label,
  value,
  delta: deltaValue,
  foot
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "kpi-card"
  }, /*#__PURE__*/React.createElement("div", {
    className: "kpi-label"
  }, /*#__PURE__*/React.createElement("span", null, label), deltaValue && /*#__PURE__*/React.createElement("span", {
    className: `delta ${deltaValue.className}`
  }, formatPercent(deltaValue.percent))), /*#__PURE__*/React.createElement("div", {
    className: "kpi-value mono"
  }, value), foot && /*#__PURE__*/React.createElement("div", {
    className: "kpi-foot"
  }, foot));
}
function DateComparisonSelector({
  comparison,
  setComparison,
  metricsA,
  metricsB
}) {
  const rows = [["Receita", metricsA.revenue, metricsB.revenue, formatCurrency], ["Pedidos", metricsA.orders, metricsB.orders, formatNumber], ["Unidades", metricsA.units, metricsB.units, formatNumber], ["Ticket médio", metricsA.aov, metricsB.aov, formatCurrency]];
  return /*#__PURE__*/React.createElement("section", {
    className: "comparison-panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "section-header"
  }, /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("h2", {
    className: "section-title"
  }, "Compara\xE7\xE3o personalizada"), /*#__PURE__*/React.createElement("p", {
    className: "section-subtitle"
  }, "Per\xEDodo A contra per\xEDodo B, mantendo os demais filtros ativos."))), /*#__PURE__*/React.createElement("div", {
    className: "comparison-controls"
  }, /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("label", null, "In\xEDcio A"), /*#__PURE__*/React.createElement("input", {
    className: "input",
    type: "date",
    value: comparison.aStart,
    onChange: event => setComparison(current => ({
      ...current,
      aStart: event.target.value
    }))
  })), /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("label", null, "Fim A"), /*#__PURE__*/React.createElement("input", {
    className: "input",
    type: "date",
    value: comparison.aEnd,
    onChange: event => setComparison(current => ({
      ...current,
      aEnd: event.target.value
    }))
  })), /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("label", null, "In\xEDcio B"), /*#__PURE__*/React.createElement("input", {
    className: "input",
    type: "date",
    value: comparison.bStart,
    onChange: event => setComparison(current => ({
      ...current,
      bStart: event.target.value
    }))
  })), /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("label", null, "Fim B"), /*#__PURE__*/React.createElement("input", {
    className: "input",
    type: "date",
    value: comparison.bEnd,
    onChange: event => setComparison(current => ({
      ...current,
      bEnd: event.target.value
    }))
  }))), /*#__PURE__*/React.createElement("div", {
    className: "table-scroll"
  }, /*#__PURE__*/React.createElement("table", null, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", null, "M\xE9trica"), /*#__PURE__*/React.createElement("th", null, "Per\xEDodo A"), /*#__PURE__*/React.createElement("th", null, "Per\xEDodo B"), /*#__PURE__*/React.createElement("th", null, "Diferen\xE7a"), /*#__PURE__*/React.createElement("th", null, "Varia\xE7\xE3o"))), /*#__PURE__*/React.createElement("tbody", null, rows.map(([label, a, b, formatter]) => {
    const rowDelta = delta(a, b);
    return /*#__PURE__*/React.createElement("tr", {
      key: label
    }, /*#__PURE__*/React.createElement("td", null, label), /*#__PURE__*/React.createElement("td", {
      className: "mono"
    }, formatter(a)), /*#__PURE__*/React.createElement("td", {
      className: "mono"
    }, formatter(b)), /*#__PURE__*/React.createElement("td", {
      className: "mono"
    }, formatter(a - b)), /*#__PURE__*/React.createElement("td", null, /*#__PURE__*/React.createElement("span", {
      className: `delta ${rowDelta.className}`
    }, formatPercent(rowDelta.percent))));
  })))));
}
function RankToggle({
  value,
  onChange
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "segmented"
  }, /*#__PURE__*/React.createElement("button", {
    className: `segment ${value === "best" ? "active" : ""}`,
    onClick: () => onChange("best")
  }, "Top"), /*#__PURE__*/React.createElement("button", {
    className: `segment ${value === "worst" ? "active" : ""}`,
    onClick: () => onChange("worst")
  }, "Piores"));
}
function ChartCard({
  title,
  subtitle,
  type,
  data,
  options,
  action,
  short = false,
  wide = false
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: `chart-card ${wide ? "wide" : ""}`
  }, /*#__PURE__*/React.createElement("div", {
    className: "chart-title-row"
  }, /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("h3", {
    className: "chart-title"
  }, title), /*#__PURE__*/React.createElement("p", {
    className: "chart-subtitle"
  }, subtitle)), action), /*#__PURE__*/React.createElement("div", {
    className: `chart-wrap ${short ? "short" : ""}`
  }, /*#__PURE__*/React.createElement(ChartCanvas, {
    type: type,
    data: data,
    options: options
  })));
}
function ChartCanvas({
  type,
  data,
  options
}) {
  const canvasRef = useRef(null);
  const chartRef = useRef(null);
  const configKey = JSON.stringify({
    type,
    data,
    options
  });
  useEffect(() => {
    if (typeof Chart === "undefined" || !canvasRef.current) return undefined;
    registerMonthlyValueLabelsPlugin();
    if (chartRef.current) chartRef.current.destroy();
    chartRef.current = new Chart(canvasRef.current, {
      type,
      data,
      options
    });
    return () => {
      if (chartRef.current) {
        chartRef.current.destroy();
        chartRef.current = null;
      }
    };
  }, [configKey]);
  if (typeof Chart === "undefined") {
    return /*#__PURE__*/React.createElement("div", {
      className: "chart-empty"
    }, "Biblioteca de gr\xE1ficos n\xE3o carregada.");
  }
  const hasData = data?.datasets?.some(dataset => Array.isArray(dataset.data) && dataset.data.length);
  if (!hasData) {
    return /*#__PURE__*/React.createElement("div", {
      className: "chart-empty"
    }, "Sem dados para o filtro atual.");
  }
  return /*#__PURE__*/React.createElement("canvas", {
    ref: canvasRef
  });
}
function SalesHeatmap({
  records
}) {
  const buckets = useMemo(() => {
    const grouped = new Map();
    records.forEach(record => grouped.set(record.dateKey, (grouped.get(record.dateKey) || 0) + record.totalValue));
    return grouped;
  }, [records]);
  const dates = useMemo(() => {
    if (!records.length) return [];
    const bounds = getDateBounds(records);
    if (!bounds) return [];
    const start = addDays(bounds.min, -((bounds.min.getDay() + 6) % 7));
    const end = addDays(bounds.max, 6 - (bounds.max.getDay() + 6) % 7);
    const output = [];
    for (let day = start; day <= end; day = addDays(day, 1)) output.push(day);
    return output;
  }, [records]);
  const maxRevenue = Math.max(...Array.from(buckets.values()), 0);
  const dayLabels = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"];
  return /*#__PURE__*/React.createElement("div", {
    className: "heatmap-card"
  }, /*#__PURE__*/React.createElement("div", {
    className: "chart-title-row"
  }, /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("h3", {
    className: "chart-title"
  }, "Heatmap de melhores dias"), /*#__PURE__*/React.createElement("p", {
    className: "chart-subtitle"
  }, "Intensidade por receita di\xE1ria no per\xEDodo filtrado"))), /*#__PURE__*/React.createElement("div", {
    className: "heatmap-grid"
  }, dayLabels.map(label => /*#__PURE__*/React.createElement("div", {
    className: "heatmap-day-label",
    key: label
  }, label)), dates.map(date => {
    const key = formatInputDate(date);
    const revenue = buckets.get(key) || 0;
    const intensity = maxRevenue ? Math.max(0.08, revenue / maxRevenue) : 0;
    const isEmpty = !buckets.has(key);
    const background = isEmpty ? undefined : `rgba(15, 118, 110, ${0.12 + intensity * 0.68})`;
    return /*#__PURE__*/React.createElement("div", {
      className: `heatmap-cell ${isEmpty ? "empty" : ""}`,
      key: key,
      style: {
        background
      },
      title: `${formatShortDate(date)} · ${formatCurrencyDetailed(revenue)}`
    }, /*#__PURE__*/React.createElement("span", {
      className: "heatmap-date"
    }, date.getDate()), /*#__PURE__*/React.createElement("span", {
      className: "heatmap-value"
    }, revenue ? compactCurrency(revenue) : ""));
  })));
}
function ProductRankingTable({
  rows,
  rankMode,
  setRankMode
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "table-card"
  }, /*#__PURE__*/React.createElement("div", {
    className: "table-head"
  }, /*#__PURE__*/React.createElement("h3", {
    className: "table-title"
  }, "Ranking de produtos"), /*#__PURE__*/React.createElement(RankToggle, {
    value: rankMode,
    onChange: setRankMode
  })), /*#__PURE__*/React.createElement("div", {
    className: "table-scroll product-table-scroll"
  }, /*#__PURE__*/React.createElement("table", {
    className: "compact-product-table"
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", null, "Produto"), /*#__PURE__*/React.createElement("th", null, "Categoria"), /*#__PURE__*/React.createElement("th", null, "Receita"), /*#__PURE__*/React.createElement("th", null, "Ped."), /*#__PURE__*/React.createElement("th", null, "Unid."))), /*#__PURE__*/React.createElement("tbody", null, rows.map(row => /*#__PURE__*/React.createElement("tr", {
    key: `${row.product}-${row.sku}-${row.category}`
  }, /*#__PURE__*/React.createElement("td", null, /*#__PURE__*/React.createElement("span", {
    className: "rank-name"
  }, /*#__PURE__*/React.createElement("strong", null, row.product), /*#__PURE__*/React.createElement("span", null, row.sku))), /*#__PURE__*/React.createElement("td", null, row.category), /*#__PURE__*/React.createElement("td", {
    className: "mono"
  }, formatCurrency(row.revenue)), /*#__PURE__*/React.createElement("td", {
    className: "mono"
  }, formatNumber(row.ordersCount)), /*#__PURE__*/React.createElement("td", {
    className: "mono"
  }, formatNumber(row.units))))))));
}
function PlatformPerformanceTable({
  rows
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "table-card"
  }, /*#__PURE__*/React.createElement("div", {
    className: "table-head"
  }, /*#__PURE__*/React.createElement("h3", {
    className: "table-title"
  }, "Performance por plataforma")), /*#__PURE__*/React.createElement("div", {
    className: "table-scroll platform-table-scroll"
  }, /*#__PURE__*/React.createElement("table", {
    className: "compact-platform-table"
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", null, "Plataforma"), /*#__PURE__*/React.createElement("th", null, "Receita"), /*#__PURE__*/React.createElement("th", null, "Ped."), /*#__PURE__*/React.createElement("th", null, "Unid."), /*#__PURE__*/React.createElement("th", null, "Ticket"))), /*#__PURE__*/React.createElement("tbody", null, rows.map(row => /*#__PURE__*/React.createElement("tr", {
    key: row.name
  }, /*#__PURE__*/React.createElement("td", null, row.name), /*#__PURE__*/React.createElement("td", {
    className: "mono"
  }, formatCurrency(row.revenue)), /*#__PURE__*/React.createElement("td", {
    className: "mono"
  }, formatNumber(row.ordersCount)), /*#__PURE__*/React.createElement("td", {
    className: "mono"
  }, formatNumber(row.units)), /*#__PURE__*/React.createElement("td", {
    className: "mono"
  }, formatCurrency(row.aov))))))));
}
ReactDOM.createRoot(document.getElementById("root")).render(/*#__PURE__*/React.createElement(App, null));
//# sourceURL=src/app.tsx
