declare const React: any;
declare const ReactDOM: any;
declare const Chart: any;
declare const XLSX: any;

const {
  useEffect,
  useMemo,
  useRef,
  useState,
} = React;

type RawRow = Record<string, any>;

type ColumnMapping = {
  date?: string;
  month?: string;
  day?: string;
  quantity?: string;
  discountValue?: string;
  channel?: string;
  itemNo?: string;
  unitValue?: string;
  platform?: string;
  customer?: string;
  sku?: string;
  product?: string;
  acos?: string;
  orderId?: string;
  company?: string;
  discountRate?: string;
  totalValue?: string;
  productCost?: string;
  platformFeeRate?: string;
  effectiveFeeRate?: string;
  shipping?: string;
  taxRate?: string;
  operationRate?: string;
  profitUnit?: string;
  profitTotal?: string;
  dailyGoal?: string;
  monthlyGoal?: string;
  status?: string;
  paymentMethod?: string;
  region?: string;
  seller?: string;
  category?: string;
};

type SaleRecord = {
  id: string;
  date: Date;
  dateKey: string;
  monthLabel: string;
  quantity: number;
  unitValue: number;
  totalValue: number;
  grossRevenue: number;
  discountValue: number;
  discountRate: number;
  channel: string;
  itemNo: string;
  platform: string;
  customer: string;
  customerType: "Novo" | "Recorrente";
  sku: string;
  product: string;
  category: string;
  acosRate: number;
  orderId: string;
  company: string;
  productCostUnit: number;
  productCostTotal: number;
  platformFeeRate: number;
  effectiveFeeRate: number;
  feeAmount: number;
  baseFeeAmount: number;
  shippingUnit: number;
  shippingTotal: number;
  taxRate: number;
  taxAmount: number;
  operationRate: number;
  operationAmount: number;
  adCostAmount: number;
  profitUnit: number;
  profitTotal: number;
  dailyGoal: number;
  monthlyGoal: number;
  status: string;
  paymentMethod: string;
  region: string;
  seller: string;
};

type Filters = {
  startDate: string;
  endDate: string;
  platforms: string[];
  channels: string[];
  products: string[];
  categories: string[];
  companies: string[];
  customerTypes: string[];
  search: string;
};

type MonthOption = {
  key: string;
  label: string;
  startDate: string;
  endDate: string;
};

type ComparisonRange = {
  aStart: string;
  aEnd: string;
  bStart: string;
  bEnd: string;
};

type DataSourceMode = "google" | "local" | "upload";

type DataLoadResult = {
  rows: RawRow[];
  source: string;
  warning?: string;
  mode: DataSourceMode;
};

type Metrics = {
  revenue: number;
  grossRevenue: number;
  orders: number;
  units: number;
  aov: number;
  discounts: number;
  shipping: number;
  fees: number;
  taxes: number;
  operation: number;
  ads: number;
  productCost: number;
  netRevenue: number;
  profit: number;
  margin: number;
};

type Bucket = {
  key: string;
  label: string;
  date?: Date;
  revenue: number;
  grossRevenue: number;
  units: number;
  profit: number;
  orders: Set<string>;
};

const DASHBOARD_XLSX = "./base-dashboard.xlsx";
const DASHBOARD_CSV = "./base-dashboard.csv";
const GOOGLE_SHEETS_SPREADSHEET_ID = "1JldFrcw8oaVAWXXhFyJCMVvm_Be9IXju90UAqpzSLXM";
const GOOGLE_SHEETS_GID = "27856229";
const GOOGLE_SHEETS_SOURCE_LABEL = "Google Sheets - BASE DASHBOARD";
const GOOGLE_SHEETS_PROXY = "./api/google-sheets";
const AUTO_REFRESH_INTERVAL_MS = 5 * 60 * 1000;

const COLORS = [
  "#0f766e",
  "#2563eb",
  "#d97706",
  "#e11d48",
  "#7c3aed",
  "#059669",
  "#334155",
  "#be185d",
  "#0ea5e9",
  "#a16207",
];

const EMPTY_FILTERS: Filters = {
  startDate: "",
  endDate: "",
  platforms: [],
  channels: [],
  products: [],
  categories: [],
  companies: [],
  customerTypes: [],
  search: "",
};

const fieldSynonyms: Record<keyof ColumnMapping, string[]> = {
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
  category: ["categoria", "category", "product category"],
};

const currencyFormatter = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const currencyFormatterDetailed = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const numberFormatter = new Intl.NumberFormat("pt-BR", {
  maximumFractionDigits: 0,
});

const percentFormatter = new Intl.NumberFormat("pt-BR", {
  style: "percent",
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});

function normalizeText(value: unknown): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[º°]/g, "")
    .replace(/[._/\\-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function cleanLabel(value: unknown, fallback = "Sem informação"): string {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  return text || fallback;
}

function titleCaseFallback(value: string): string {
  return value
    .toLowerCase()
    .split(" ")
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

function normalizeChannel(value: unknown): string {
  const label = cleanLabel(value);
  const normalized = normalizeText(label);
  if (normalized === "traego pago" || normalized === "trafego pago") return "Tráfego Pago";
  if (normalized === "organico") return "Orgânico";
  return label;
}

function normalizePlatform(value: unknown): string {
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

function parseNumber(value: unknown, defaultValue = 0, asPercent = false): number {
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

function excelSerialToDate(serial: number): Date | null {
  if (!Number.isFinite(serial)) return null;
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate());
}

function parseDate(value: unknown): Date | null {
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

function inputDateToDate(value: string): Date | null {
  if (!value) return null;
  const [year, month, day] = value.split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
}

function formatInputDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatShortDate(date: Date): string {
  return date.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit", year: "2-digit" });
}

function addDays(date: Date, amount: number): Date {
  const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  copy.setDate(copy.getDate() + amount);
  return copy;
}

function daysBetween(start: Date, end: Date): number {
  const ms = new Date(end.getFullYear(), end.getMonth(), end.getDate()).getTime()
    - new Date(start.getFullYear(), start.getMonth(), start.getDate()).getTime();
  return Math.round(ms / 86400000);
}

function formatCurrency(value: number): string {
  return currencyFormatter.format(value || 0);
}

function formatCurrencyDetailed(value: number): string {
  return currencyFormatterDetailed.format(value || 0);
}

function formatNumber(value: number): string {
  return numberFormatter.format(value || 0);
}

function formatPercent(value: number): string {
  if (!Number.isFinite(value)) return "0,0%";
  return percentFormatter.format(value);
}

function compactCurrency(value: number): string {
  return formatCurrencyDetailed(value || 0);
}

function detectColumns(headers: string[]): ColumnMapping {
  const normalizedHeaders = headers.map((header) => ({
    original: header,
    normalized: normalizeText(header),
  }));
  const mapping: ColumnMapping = {};

  (Object.keys(fieldSynonyms) as Array<keyof ColumnMapping>).forEach((field) => {
    const options = fieldSynonyms[field].map(normalizeText);
    const exact = normalizedHeaders.find((header) => options.includes(header.normalized));
    if (exact) {
      mapping[field] = exact.original;
      return;
    }
    const fuzzy = normalizedHeaders.find((header) =>
      options.some((option) => header.normalized.includes(option) || option.includes(header.normalized)),
    );
    if (fuzzy) mapping[field] = fuzzy.original;
  });

  return mapping;
}

function rowValue(row: RawRow, column?: string): unknown {
  if (!column) return undefined;
  if (Object.prototype.hasOwnProperty.call(row, column)) return row[column];
  const normalizedColumn = normalizeText(column);
  const key = Object.keys(row).find((candidate) => normalizeText(candidate) === normalizedColumn);
  return key ? row[key] : undefined;
}

function deriveCategory(productValue: string, skuValue: string, itemValue: string, explicit?: string): string {
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
function mapRowsToSales(rawRows: RawRow[]): SaleRecord[] {
  if (!rawRows.length) return [];
  const headers = Object.keys(rawRows[0] || {});
  const mapping = detectColumns(headers);

  if (!mapping.date || !mapping.totalValue) {
    throw new Error("A base precisa ter ao menos uma coluna de data e uma coluna de valor total.");
  }

  const records = rawRows
    .map((row, index) => {
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
      const fallbackProfitUnit =
        unitValue
        - productCostUnit
        - shippingUnit
        - unitValue * effectiveFeeRate
        - unitValue * taxRate
        - unitValue * operationRate
        - unitValue * acosRate;
      const profitUnit = parseNumber(rowValue(row, mapping.profitUnit), fallbackProfitUnit);
      const profitTotal = parseNumber(rowValue(row, mapping.profitTotal), profitUnit * quantity);

      return {
        id: `${orderId}-${index}`,
        date,
        dateKey: formatInputDate(date),
        monthLabel: date.toLocaleDateString("pt-BR", { month: "short", year: "numeric" }),
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
        customerType: "Novo" as const,
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
        seller: cleanLabel(rowValue(row, mapping.seller), "Não disponível"),
      };
    })
    .filter(Boolean) as SaleRecord[];

  return attachCustomerTypes(records);
}

function attachCustomerTypes(records: SaleRecord[]): SaleRecord[] {
  const firstPurchase = new Map<string, number>();
  records.forEach((record) => {
    const key = normalizeText(record.customer);
    const current = firstPurchase.get(key);
    const time = record.date.getTime();
    if (current === undefined || time < current) firstPurchase.set(key, time);
  });

  return records.map((record) => {
    const first = firstPurchase.get(normalizeText(record.customer));
    return {
      ...record,
      customerType: first === record.date.getTime() ? "Novo" : "Recorrente",
    };
  });
}

function parseCsv(text: string): RawRow[] {
  const cleaned = text.replace(/^\ufeff/, "");
  const firstLine = cleaned.split(/\r?\n/, 1)[0] || "";
  const delimiter = firstLine.split(";").length > firstLine.split(",").length ? ";" : ",";
  const rows: string[][] = [];
  let row: string[] = [];
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
      if (row.some((cell) => cell.trim() !== "")) rows.push(row);
      row = [];
      value = "";
      continue;
    }
    value += char;
  }
  row.push(value);
  if (row.some((cell) => cell.trim() !== "")) rows.push(row);

  const headers = rows.shift()?.map((header) => header.trim()) || [];
  return rows.map((cells) => {
    const output: RawRow = {};
    headers.forEach((header, index) => {
      output[header] = cells[index] ?? "";
    });
    return output;
  });
}

// SpreadsheetUploader/DataParser: reads Excel workbooks in the browser and returns
// plain row objects before the mapping layer applies column normalization.
async function parseExcelBuffer(buffer: ArrayBuffer): Promise<RawRow[]> {
  if (typeof XLSX === "undefined") {
    throw new Error("Biblioteca XLSX indisponível.");
  }
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    raw: true,
  });
  const sheetName = workbook.SheetNames.includes("Base de Dados")
    ? "Base de Dados"
    : workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(worksheet, {
    defval: null,
    raw: true,
  });
}

function googleSheetsCsvUrl(): string {
  return `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEETS_SPREADSHEET_ID}/gviz/tq?tqx=out:csv&gid=${GOOGLE_SHEETS_GID}&cacheBust=${Date.now()}`;
}

function googleSheetsFetchSources(): Array<{ url: string; source: string }> {
  const directSource = { url: googleSheetsCsvUrl(), source: GOOGLE_SHEETS_SOURCE_LABEL };
  const host = typeof window !== "undefined" ? window.location.hostname : "";
  const protocol = typeof window !== "undefined" ? window.location.protocol : "";
  const localHosts = ["localhost", "127.0.0.1", "::1"];

  if (protocol.startsWith("http") && localHosts.includes(host)) {
    return [
      { url: `${GOOGLE_SHEETS_PROXY}?cacheBust=${Date.now()}`, source: GOOGLE_SHEETS_SOURCE_LABEL },
      directSource,
    ];
  }

  return [directSource];
}

function looksLikeHtml(text: string): boolean {
  return /^<!doctype html|^<html/i.test(text.trim());
}

async function fetchCsvSource(url: string, source: string): Promise<DataLoadResult> {
  const response = await fetch(url, { cache: "no-store" });
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

  return { rows, source, mode: "google" };
}

async function loadGoogleSheetsRows(): Promise<DataLoadResult> {
  const errors: string[] = [];
  for (const source of googleSheetsFetchSources()) {
    try {
      return await fetchCsvSource(source.url, source.source);
    } catch (error: any) {
      errors.push(error?.message || "erro desconhecido");
    }
  }

  throw new Error(`NÃ£o foi possÃ­vel carregar o Google Sheets. ${errors.join(" | ")}`);
}

async function loadInitialRows(): Promise<DataLoadResult> {
  let googleSheetsError = "";
  try {
    return await loadGoogleSheetsRows();
  } catch (error: any) {
    googleSheetsError = error?.message || "erro desconhecido";
  }

  const fallbackWarning = `NÃ£o foi possÃ­vel carregar o Google Sheets em tempo real (${googleSheetsError}). Usando a base local como fallback.`;
  const embeddedCsv = (window as any).__BASE_DASHBOARD_CSV__;
  if (embeddedCsv) {
    return {
      rows: parseCsv(embeddedCsv),
      source: "Fallback local - BASE DASHBOARD.xlsx",
      warning: fallbackWarning,
      mode: "local",
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
      mode: "local",
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
        mode: "local",
      };
    } catch {
      if (embeddedCsv) {
        return {
          rows: parseCsv(embeddedCsv),
          source: "Fallback local - BASE DASHBOARD.xlsx",
          warning: fallbackWarning,
          mode: "local",
        };
      }
      throw error;
    }
  }
}

async function parseUploadedFile(file: File): Promise<RawRow[]> {
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (extension === "csv") {
    return parseCsv(await file.text());
  }
  return parseExcelBuffer(await file.arrayBuffer());
}

function uniqueValues(records: SaleRecord[], getter: (record: SaleRecord) => string): string[] {
  return Array.from(new Set(records.map(getter).filter(Boolean))).sort((a, b) => a.localeCompare(b, "pt-BR"));
}

function matchesList(selected: string[], value: string): boolean {
  return selected.length === 0 || selected.includes(value);
}

function applyDimensionFilters(records: SaleRecord[], filters: Filters): SaleRecord[] {
  const search = normalizeText(filters.search);
  return records.filter((record) => {
    if (!matchesList(filters.platforms, record.platform)) return false;
    if (!matchesList(filters.channels, record.channel)) return false;
    if (!matchesList(filters.products, record.product)) return false;
    if (!matchesList(filters.categories, record.category)) return false;
    if (!matchesList(filters.companies, record.company)) return false;
    if (!matchesList(filters.customerTypes, record.customerType)) return false;
    if (!search) return true;
    const haystack = normalizeText([
      record.product,
      record.sku,
      record.orderId,
      record.customer,
      record.platform,
      record.company,
      record.category,
    ].join(" "));
    return haystack.includes(search);
  });
}

function filterByDate(records: SaleRecord[], startDate: string, endDate: string): SaleRecord[] {
  const start = inputDateToDate(startDate);
  const end = inputDateToDate(endDate);
  return records.filter((record) => {
    if (start && record.date < start) return false;
    if (end && record.date > end) return false;
    return true;
  });
}

function calcMetrics(records: SaleRecord[]): Metrics {
  const orderSet = new Set(records.map((record) => record.orderId));
  const revenue = sum(records, (record) => record.totalValue);
  const orders = orderSet.size || 0;
  const units = sum(records, (record) => record.quantity);
  const discounts = sum(records, (record) => record.discountValue);
  const shipping = sum(records, (record) => record.shippingTotal);
  const fees = sum(records, (record) => record.feeAmount);
  const taxes = sum(records, (record) => record.taxAmount);
  const operation = sum(records, (record) => record.operationAmount);
  const ads = sum(records, (record) => record.adCostAmount);
  const productCost = sum(records, (record) => record.productCostTotal);
  const profit = sum(records, (record) => record.profitTotal);
  const netRevenue = revenue - discounts - fees - shipping - taxes - operation - ads;
  return {
    revenue,
    grossRevenue: sum(records, (record) => record.grossRevenue),
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
    margin: revenue ? profit / revenue : 0,
  };
}

function sum<T>(items: T[], getter: (item: T) => number): number {
  return items.reduce((total, item) => total + (getter(item) || 0), 0);
}

function previousRange(startDate: string, endDate: string): { start: string; end: string } {
  const start = inputDateToDate(startDate);
  const end = inputDateToDate(endDate);
  if (!start || !end) return { start: "", end: "" };
  const length = daysBetween(start, end) + 1;
  const previousEnd = addDays(start, -1);
  const previousStart = addDays(previousEnd, -(length - 1));
  return {
    start: formatInputDate(previousStart),
    end: formatInputDate(previousEnd),
  };
}

function getDateBounds(records: SaleRecord[]): { min: Date; max: Date } | null {
  if (!records.length) return null;
  const times = records.map((record) => record.date.getTime());
  return {
    min: new Date(Math.min(...times)),
    max: new Date(Math.max(...times)),
  };
}

function getDefaultRange(records: SaleRecord[]): { startDate: string; endDate: string } {
  const bounds = getDateBounds(records);
  if (!bounds) return { startDate: "", endDate: "" };
  const start = addDays(bounds.max, -29) < bounds.min ? bounds.min : addDays(bounds.max, -29);
  return {
    startDate: formatInputDate(start),
    endDate: formatInputDate(bounds.max),
  };
}

function getMonthOptions(records: SaleRecord[]): MonthOption[] {
  const buckets = new Map<string, { dates: Date[] }>();
  records.forEach((record) => {
    const key = `${record.date.getFullYear()}-${String(record.date.getMonth() + 1).padStart(2, "0")}`;
    const current = buckets.get(key) || {
      dates: [],
    };
    current.dates.push(record.date);
    buckets.set(key, current);
  });

  return Array.from(buckets.entries())
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, bucket]) => {
      const times = bucket.dates.map((date) => date.getTime());
      const [year, month] = key.split("-").map(Number);
      const monthName = new Date(year, month - 1, 1).toLocaleDateString("pt-BR", { month: "short" }).replace(".", "");
      return {
        key,
        label: `${titleCaseFallback(monthName)}/${String(year).slice(-2)}`,
        startDate: formatInputDate(new Date(Math.min(...times))),
        endDate: formatInputDate(new Date(Math.max(...times))),
      };
    });
}

function getIsoWeek(date: Date): { year: number; week: number } {
  const utc = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(utc.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((utc.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
  return { year: utc.getUTCFullYear(), week };
}

function periodKey(date: Date, granularity: string): { key: string; label: string } {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  if (granularity === "day") {
    return { key: formatInputDate(date), label: formatShortDate(date) };
  }
  if (granularity === "week") {
    const week = getIsoWeek(date);
    return { key: `${week.year}-W${String(week.week).padStart(2, "0")}`, label: `S${String(week.week).padStart(2, "0")} ${week.year}` };
  }
  if (granularity === "quarter") {
    return { key: `${year}-Q${Math.ceil(month / 3)}`, label: `T${Math.ceil(month / 3)} ${year}` };
  }
  if (granularity === "year") {
    return { key: String(year), label: String(year) };
  }
  return {
    key: `${year}-${String(month).padStart(2, "0")}`,
    label: date.toLocaleDateString("pt-BR", { month: "short", year: "2-digit" }),
  };
}

function groupByPeriod(records: SaleRecord[], granularity: string): Array<Bucket & { ordersCount: number }> {
  const buckets = new Map<string, Bucket>();
  records.forEach((record) => {
    const period = periodKey(record.date, granularity);
    const current = buckets.get(period.key) || {
      key: period.key,
      label: period.label,
      date: record.date,
      revenue: 0,
      grossRevenue: 0,
      units: 0,
      profit: 0,
      orders: new Set<string>(),
    };
    current.revenue += record.totalValue;
    current.grossRevenue += record.grossRevenue;
    current.units += record.quantity;
    current.profit += record.profitTotal;
    current.orders.add(record.orderId);
    if (!current.date || record.date < current.date) current.date = record.date;
    buckets.set(period.key, current);
  });

  return Array.from(buckets.values())
    .sort((a, b) => a.key.localeCompare(b.key))
    .map((bucket) => ({ ...bucket, ordersCount: bucket.orders.size }));
}

function groupMonthlyRevenueByYear(records: SaleRecord[]): Array<{ key: string; label: string; revenue: number }> {
  if (!records.length) return [];
  const years = Array.from(new Set(records.map((record) => record.date.getFullYear()))).sort((a, b) => a - b);
  const hasMultipleYears = years.length > 1;
  const revenueByMonth = new Map<string, number>();

  records.forEach((record) => {
    const key = `${record.date.getFullYear()}-${String(record.date.getMonth() + 1).padStart(2, "0")}`;
    revenueByMonth.set(key, (revenueByMonth.get(key) || 0) + record.totalValue);
  });

  return years.flatMap((year) =>
    Array.from({ length: 12 }, (_, monthIndex) => {
      const key = `${year}-${String(monthIndex + 1).padStart(2, "0")}`;
      const label = new Date(year, monthIndex, 1).toLocaleDateString("pt-BR", {
        month: "short",
        year: hasMultipleYears ? "2-digit" : undefined,
      });
      return {
        key,
        label,
        revenue: revenueByMonth.get(key) || 0,
      };
    }),
  );
}

function groupByDimension(records: SaleRecord[], getter: (record: SaleRecord) => string): Array<Bucket & { name: string; ordersCount: number; margin: number; aov: number }> {
  const buckets = new Map<string, Bucket>();
  records.forEach((record) => {
    const name = getter(record);
    const current = buckets.get(name) || {
      key: name,
      label: name,
      revenue: 0,
      grossRevenue: 0,
      units: 0,
      profit: 0,
      orders: new Set<string>(),
    };
    current.revenue += record.totalValue;
    current.grossRevenue += record.grossRevenue;
    current.units += record.quantity;
    current.profit += record.profitTotal;
    current.orders.add(record.orderId);
    buckets.set(name, current);
  });
  return Array.from(buckets.values()).map((bucket) => ({
    ...bucket,
    name: bucket.label,
    ordersCount: bucket.orders.size,
    margin: bucket.revenue ? bucket.profit / bucket.revenue : 0,
    aov: bucket.orders.size ? bucket.revenue / bucket.orders.size : 0,
  }));
}

function delta(current: number, previous: number): { amount: number; percent: number; className: string } {
  const amount = current - previous;
  const percent = previous ? amount / Math.abs(previous) : current ? 1 : 0;
  return {
    amount,
    percent,
    className: amount > 0 ? "up" : amount < 0 ? "down" : "flat",
  };
}

function chartOptions(valueFormatter: (value: number) => string, stacked = false): any {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: "index", intersect: false },
    plugins: {
      legend: {
        labels: {
          color: "#475569",
          boxWidth: 11,
          usePointStyle: true,
        },
      },
      tooltip: {
        callbacks: {
          label: (context: any) => {
            const label = context.dataset.label ? `${context.dataset.label}: ` : "";
            const value = typeof context.raw === "number"
              ? context.raw
              : Number(context.parsed.y ?? context.parsed.x ?? 0);
            return `${label}${valueFormatter(value)}`;
          },
        },
      },
    },
    scales: {
      x: {
        stacked,
        grid: { display: false },
        ticks: { color: "#64748b", maxRotation: 0, autoSkip: true },
      },
      y: {
        stacked,
        beginAtZero: true,
        grid: { color: "rgba(148, 163, 184, 0.22)" },
        ticks: {
          color: "#64748b",
          callback: (value: any) => valueFormatter(Number(value)),
        },
      },
    },
  };
}

let monthlyValueLabelsRegistered = false;

function registerMonthlyValueLabelsPlugin() {
  if (monthlyValueLabelsRegistered || typeof Chart === "undefined") return;
  Chart.register({
    id: "monthlyValueLabels",
    afterDatasetsDraw(chart: any, _args: any, options: any) {
      if (!options?.enabled) return;
      const { ctx, chartArea } = chart;
      ctx.save();
      ctx.fillStyle = options.color || "#334155";
      ctx.font = options.font || "700 10px Inter, Segoe UI, sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";
      chart.data.datasets.forEach((dataset: any, datasetIndex: number) => {
        const meta = chart.getDatasetMeta(datasetIndex);
        meta.data.forEach((element: any, index: number) => {
          const value = Number(dataset.data[index] || 0);
          if (!value || !element) return;
          const y = Math.max(chartArea.top + 14, element.y - 6);
          ctx.fillText(formatCurrencyDetailed(value), element.x, y);
        });
      });
      ctx.restore();
    },
  });
  monthlyValueLabelsRegistered = true;
}

function monthlyRevenueOptions(maxRevenue: number): any {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      monthlyValueLabels: { enabled: true },
      tooltip: {
        callbacks: {
          label: (context: any) => `Faturamento total: ${formatCurrencyDetailed(Number(context.raw || 0))}`,
        },
      },
    },
    scales: {
      x: {
        grid: { display: false },
        ticks: { color: "#64748b", maxRotation: 0, autoSkip: false },
      },
      y: {
        beginAtZero: true,
        suggestedMax: maxRevenue ? maxRevenue * 1.18 : undefined,
        grid: { color: "rgba(148, 163, 184, 0.22)" },
        ticks: {
          color: "#64748b",
          callback: (value: any) => formatCurrencyDetailed(Number(value)),
        },
      },
    },
  };
}

function horizontalCurrencyOptions(): any {
  return {
    responsive: true,
    maintainAspectRatio: false,
    indexAxis: "y",
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: (context: any) => `${context.dataset.label}: ${formatCurrency(Number(context.raw || 0))}`,
        },
      },
    },
    scales: {
      x: {
        beginAtZero: true,
        grid: { color: "rgba(148, 163, 184, 0.22)" },
        ticks: { color: "#64748b", callback: (value: any) => compactCurrency(Number(value)) },
      },
      y: {
        grid: { display: false },
        ticks: { color: "#64748b" },
      },
    },
  };
}

function makeLineData(labels: string[], revenue: number[], previous?: number[]): any {
  const datasets: any[] = [
    {
      label: "Receita",
      data: revenue,
      borderColor: COLORS[0],
      backgroundColor: "rgba(15, 118, 110, 0.14)",
      tension: 0.28,
      fill: true,
      pointRadius: 2,
    },
  ];
  if (previous) {
    datasets.push({
      label: "Período anterior",
      data: previous,
      borderColor: COLORS[3],
      backgroundColor: "rgba(225, 29, 72, 0.08)",
      borderDash: [6, 5],
      tension: 0.28,
      fill: false,
      pointRadius: 0,
    });
  }
  return { labels, datasets };
}

function rangeDescription(start: string, end: string): string {
  const startDate = inputDateToDate(start);
  const endDate = inputDateToDate(end);
  if (!startDate || !endDate) return "Período não definido";
  return `${formatShortDate(startDate)} a ${formatShortDate(endDate)}`;
}

function App() {
  const [records, setRecords] = useState<SaleRecord[]>([]);
  const [filters, setFilters] = useState<Filters>(EMPTY_FILTERS);
  const [granularity, setGranularity] = useState("day");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [warning, setWarning] = useState("");
  const [sourceName, setSourceName] = useState(GOOGLE_SHEETS_SOURCE_LABEL);
  const [dataSourceMode, setDataSourceMode] = useState<DataSourceMode>("google");
  const [lastSyncAt, setLastSyncAt] = useState<Date | null>(null);
  const [theme, setTheme] = useState<"light" | "dark">(() => {
    try {
      return window.localStorage.getItem("dashboard-theme") === "dark" ? "dark" : "light";
    } catch {
      return "light";
    }
  });
  const [rankMode, setRankMode] = useState<"best" | "worst">("best");
  const [comparison, setComparison] = useState({
    aStart: "",
    aEnd: "",
    bStart: "",
    bEnd: "",
  });

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    try {
      window.localStorage.setItem("dashboard-theme", theme);
    } catch {
      // Theme persistence is optional; the dashboard still works without localStorage.
    }
  }, [theme]);

  useEffect(() => {
    let mounted = true;
    setLoading(true);
    loadInitialRows()
      .then(({ rows, source, warning: loadWarning, mode }) => {
        if (!mounted) return;
        const normalized = mapRowsToSales(rows);
        setRecords(normalized);
        setSourceName(source);
        setDataSourceMode(mode);
        setLastSyncAt(new Date());
        setWarning(loadWarning || "");
        const defaults = getDefaultRange(normalized);
        const previous = previousRange(defaults.startDate, defaults.endDate);
        setFilters((current) => ({ ...current, ...defaults }));
        setComparison({
          aStart: defaults.startDate,
          aEnd: defaults.endDate,
          bStart: previous.start,
          bEnd: previous.end,
        });
      })
      .catch((loadError: any) => {
        if (!mounted) return;
        setError(loadError?.message || "Não foi possível carregar a base inicial.");
      })
      .finally(() => {
        if (mounted) setLoading(false);
      });
    return () => {
      mounted = false;
    };
  }, []);

  async function handleFileChange(event: any) {
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
      setFilters({ ...EMPTY_FILTERS, ...defaults });
      setComparison({
        aStart: defaults.startDate,
        aEnd: defaults.endDate,
        bStart: previous.start,
        bEnd: previous.end,
      });
    } catch (uploadError: any) {
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
      const { rows, source, warning: loadWarning, mode } = await loadGoogleSheetsRows();
      const normalized = mapRowsToSales(rows);
      const defaults = getDefaultRange(normalized);
      const previous = previousRange(defaults.startDate, defaults.endDate);
      setRecords(normalized);
      setSourceName(source);
      setDataSourceMode(mode);
      setLastSyncAt(new Date());
      setWarning(loadWarning || "");
      setFilters((current) => current.startDate && current.endDate ? current : { ...current, ...defaults });
      setComparison((current: ComparisonRange) => (
        current.aStart && current.aEnd && current.bStart && current.bEnd
          ? current
          : {
              aStart: defaults.startDate,
              aEnd: defaults.endDate,
              bStart: previous.start,
              bEnd: previous.end,
            }
      ));
    } catch (refreshError: any) {
      setError(refreshError?.message || "NÃ£o foi possÃ­vel atualizar os dados do Google Sheets.");
    } finally {
      setLoading(false);
    }
  }

  function handleExportPdf() {
    document.body.classList.add("pdf-export-mode");
    const cleanup = () => document.body.classList.remove("pdf-export-mode");
    window.addEventListener("afterprint", cleanup, { once: true });
    window.setTimeout(() => {
      window.print();
      window.setTimeout(cleanup, 1200);
    }, 50);
  }

  useEffect(() => {
    if (dataSourceMode === "upload") return undefined;
    let disposed = false;
    const interval = window.setInterval(async () => {
      try {
        const { rows, source, mode } = await loadGoogleSheetsRows();
        if (disposed) return;
        setRecords(mapRowsToSales(rows));
        setSourceName(source);
        setDataSourceMode(mode);
        setLastSyncAt(new Date());
        setWarning("");
      } catch (refreshError: any) {
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
  const currentRecords = useMemo(
    () => filterByDate(dimensionFiltered, filters.startDate, filters.endDate),
    [dimensionFiltered, filters.startDate, filters.endDate],
  );
  const previous = useMemo(() => previousRange(filters.startDate, filters.endDate), [filters.startDate, filters.endDate]);
  const previousRecords = useMemo(
    () => filterByDate(dimensionFiltered, previous.start, previous.end),
    [dimensionFiltered, previous.start, previous.end],
  );
  const metrics = useMemo(() => calcMetrics(currentRecords), [currentRecords]);
  const previousMetrics = useMemo(() => calcMetrics(previousRecords), [previousRecords]);
  const monthOptions = useMemo(() => getMonthOptions(records), [records]);

  const options = useMemo(() => ({
    platforms: uniqueValues(records, (record) => record.platform),
    channels: uniqueValues(records, (record) => record.channel),
    products: uniqueValues(records, (record) => record.product),
    categories: uniqueValues(records, (record) => record.category),
    companies: uniqueValues(records, (record) => record.company),
    customerTypes: ["Novo", "Recorrente"],
  }), [records]);

  const periodBuckets = useMemo(() => groupByPeriod(currentRecords, granularity), [currentRecords, granularity]);
  const previousBuckets = useMemo(() => groupByPeriod(previousRecords, granularity), [previousRecords, granularity]);
  const monthlyRevenueBuckets = useMemo(() => groupMonthlyRevenueByYear(dimensionFiltered), [dimensionFiltered]);
  const weeklyBuckets = useMemo(() => groupByPeriod(currentRecords, "week"), [currentRecords]);
  const dailyBuckets = useMemo(() => groupByPeriod(currentRecords, "day"), [currentRecords]);
  const platformBuckets = useMemo(
    () => groupByDimension(currentRecords, (record) => record.platform).sort((a, b) => b.revenue - a.revenue),
    [currentRecords],
  );
  const categoryBuckets = useMemo(
    () => groupByDimension(currentRecords, (record) => record.category).sort((a, b) => b.revenue - a.revenue),
    [currentRecords],
  );
  const channelBuckets = useMemo(
    () => groupByDimension(currentRecords, (record) => record.channel).sort((a, b) => b.revenue - a.revenue),
    [currentRecords],
  );
  const productBuckets = useMemo(
    () => groupByDimension(currentRecords, (record) => `${record.product}||${record.sku}||${record.category}`)
      .map((bucket) => {
        const [product, sku, category] = bucket.name.split("||");
        return { ...bucket, product, sku, category };
      })
      .sort((a, b) => rankMode === "best" ? b.revenue - a.revenue : a.revenue - b.revenue),
    [currentRecords, rankMode],
  );
  const comparisonMetrics = useMemo(() => {
    const a = calcMetrics(filterByDate(dimensionFiltered, comparison.aStart, comparison.aEnd));
    const b = calcMetrics(filterByDate(dimensionFiltered, comparison.bStart, comparison.bEnd));
    return { a, b };
  }, [dimensionFiltered, comparison]);

  const previousRevenueByIndex = previousBuckets.map((bucket) => bucket.revenue);
  const kpiDeltas = {
    revenue: delta(metrics.revenue, previousMetrics.revenue),
    orders: delta(metrics.orders, previousMetrics.orders),
    units: delta(metrics.units, previousMetrics.units),
    profit: delta(metrics.profit, previousMetrics.profit),
  };

  const lineChartData = makeLineData(
    periodBuckets.map((bucket) => bucket.label),
    periodBuckets.map((bucket) => bucket.revenue),
    previousRevenueByIndex.length ? previousRevenueByIndex : undefined,
  );

  const ordersChartData = {
    labels: periodBuckets.map((bucket) => bucket.label),
    datasets: [
      {
        type: "line",
        label: "Faturamento",
        data: periodBuckets.map((bucket) => bucket.revenue),
        borderColor: COLORS[0],
        backgroundColor: "rgba(15, 118, 110, 0.12)",
        tension: 0.25,
        pointRadius: 2,
        yAxisID: "y",
      },
      {
        label: "Unidades",
        data: periodBuckets.map((bucket) => bucket.units),
        backgroundColor: "rgba(217, 119, 6, 0.68)",
        borderRadius: 6,
        yAxisID: "y1",
      },
    ],
  };

  const platformChartData = {
    labels: platformBuckets.slice(0, 10).map((bucket) => {
      const total = sum(platformBuckets, (item) => item.revenue);
      const percent = total ? bucket.revenue / total : 0;
      return `${bucket.name} (${formatPercent(percent)})`;
    }),
    datasets: [
      {
        label: "Receita",
        data: platformBuckets.slice(0, 10).map((bucket) => bucket.revenue),
        backgroundColor: platformBuckets.slice(0, 10).map((_, index) => COLORS[index % COLORS.length]),
        borderRadius: 6,
      },
    ],
  };

  const categoryChartData = {
    labels: categoryBuckets.map((bucket) => bucket.name),
    datasets: [
      {
        label: "Receita",
        data: categoryBuckets.map((bucket) => bucket.revenue),
        backgroundColor: categoryBuckets.map((_, index) => COLORS[index % COLORS.length]),
        borderWidth: 0,
      },
    ],
  };

  const productChartData = {
    labels: productBuckets.slice(0, 12).map((bucket: any) => bucket.product),
    datasets: [
      {
        label: "Receita",
        data: productBuckets.slice(0, 12).map((bucket: any) => bucket.revenue),
        backgroundColor: productBuckets.slice(0, 12).map(() => rankMode === "best" ? "rgba(5, 150, 105, 0.72)" : "rgba(225, 29, 72, 0.72)"),
        borderRadius: 6,
      },
    ],
  };

  const channelChartData = {
    labels: channelBuckets.map((bucket) => bucket.name),
    datasets: [
      {
        label: "Receita",
        data: channelBuckets.map((bucket) => bucket.revenue),
        backgroundColor: channelBuckets.map((_, index) => COLORS[index % COLORS.length]),
        borderRadius: 6,
      },
    ],
  };

  const monthlyChartData = {
    labels: monthlyRevenueBuckets.map((bucket) => bucket.label),
    datasets: [
      {
        label: "Faturamento total",
        data: monthlyRevenueBuckets.map((bucket) => bucket.revenue),
        backgroundColor: "rgba(15, 118, 110, 0.72)",
        borderRadius: 6,
      },
    ],
  };

  const weeklyChartData = {
    labels: weeklyBuckets.map((bucket) => bucket.label),
    datasets: [
      {
        label: "Receita",
        data: weeklyBuckets.map((bucket) => bucket.revenue),
        borderColor: COLORS[1],
        backgroundColor: "rgba(37, 99, 235, 0.12)",
        tension: 0.22,
        fill: true,
      },
    ],
  };

  const dailyTrendData = {
    labels: dailyBuckets.map((bucket) => bucket.label),
    datasets: [
      {
        label: "Receita diária",
        data: dailyBuckets.map((bucket) => bucket.revenue),
        borderColor: COLORS[2],
        backgroundColor: "rgba(217, 119, 6, 0.14)",
        tension: 0.18,
        fill: true,
        pointRadius: 1.5,
      },
    ],
  };

  const scatterData = {
    datasets: [
      {
        label: "Dias",
        data: dailyBuckets.map((bucket) => ({
          x: bucket.ordersCount,
          y: bucket.revenue,
          label: bucket.label,
        })),
        backgroundColor: "rgba(15, 118, 110, 0.72)",
        borderColor: COLORS[0],
      },
    ],
  };

  const periodLabel = rangeDescription(filters.startDate, filters.endDate);
  const previousLabel = rangeDescription(previous.start, previous.end);
  const monthlyMaxRevenue = Math.max(...monthlyRevenueBuckets.map((bucket) => bucket.revenue), 0);

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="topbar-inner">
          <div className="brand-block">
            <h1 className="brand-title">Dashboard - Vendas E-commerce</h1>
            <div className="brand-meta">
              {dateBounds && <span>Data do documento: {formatShortDate(dateBounds.min)} a {formatShortDate(dateBounds.max)}</span>}
              {lastSyncAt && <span>Sincronizado: {lastSyncAt.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}</span>}
            </div>
          </div>
          <div className="brand-logos" aria-label="Logotipos Kuanttum e Primebras">
            <img className="brand-logo kuanttum-logo" src="./assets/logo-kuanttum.png" alt="Kuanttum" />
            <img className="brand-logo primebras-logo" src="./assets/logo-primebras.png" alt="Primebras" />
          </div>
          <div className="topbar-actions">
            <span className="status-pill">
              <span className={`status-dot ${error ? "error" : loading ? "warn" : ""}`} />
              {loading ? "Carregando" : error ? "Erro" : "Base ativa"}
            </span>
            <button className="btn" onClick={handleGoogleRefresh} disabled={loading}>
              <span aria-hidden="true">â†»</span>
              Atualizar
            </button>
            <button className="btn" onClick={handleExportPdf}>
              <span aria-hidden="true">PDF</span>
              Exportar PDF
            </button>
            <button className="btn btn-primary" onClick={() => setTheme(theme === "dark" ? "light" : "dark")}>
              <span aria-hidden="true">{theme === "dark" ? "Claro" : "Escuro"}</span>
              {theme === "dark" ? "Modo claro" : "Modo escuro"}
            </button>
          </div>
        </div>
      </header>

      <main className="main-grid">
        <FiltersPanel
          filters={filters}
          setFilters={setFilters}
          options={options}
          monthOptions={monthOptions}
          comparison={comparison}
          setComparison={setComparison}
          dateBounds={dateBounds}
          reset={() => {
            const defaults = getDefaultRange(records);
            setFilters({ ...EMPTY_FILTERS, ...defaults });
          }}
        />

        <div className="content-area">
          {error && <div className="alert">{error}</div>}
          {warning && <div className="alert">{warning}</div>}
          {loading ? (
            <div className="loading-state">Processando dados da planilha</div>
          ) : (
            <>
              <section className="section">
                <div className="section-header">
                  <div>
                    <h2 className="section-title">Resumo executivo</h2>
                    <p className="section-subtitle">Período atual: {periodLabel} · Comparativo: {previousLabel}</p>
                  </div>
                  <GranularitySelector value={granularity} onChange={setGranularity} />
                </div>
                <div className="kpi-grid">
                  <KPICard label="Receita total" value={formatCurrency(metrics.revenue)} delta={kpiDeltas.revenue} foot={`Dif.: ${formatCurrency(kpiDeltas.revenue.amount)}`} />
                  <KPICard label="Pedidos" value={formatNumber(metrics.orders)} delta={kpiDeltas.orders} foot={`Dif.: ${formatNumber(kpiDeltas.orders.amount)}`} />
                  <KPICard label="Unidades vendidas" value={formatNumber(metrics.units)} delta={kpiDeltas.units} foot={`Ticket médio: ${formatCurrency(metrics.aov)}`} />
                  <KPICard label="Taxas marketplace" value={formatCurrency(metrics.fees)} foot={`Frete: ${formatCurrency(metrics.shipping)}`} />
                </div>
              </section>

              <DateComparisonSelector
                comparison={comparison}
                setComparison={setComparison}
                metricsA={comparisonMetrics.a}
                metricsB={comparisonMetrics.b}
              />

              <section className="chart-grid">
                <ChartCard
                  title="Receita ao longo do tempo"
                  subtitle="Atual versus período anterior de mesma duração"
                  type="line"
                  data={lineChartData}
                  options={chartOptions(compactCurrency)}
                />
                <ChartCard
                  title="Faturamento x unidades"
                  subtitle="Receita e volume agrupados pelo período selecionado"
                  type="bar"
                  data={ordersChartData}
                  options={dualAxisOptions()}
                />
                <ChartCard
                  title="Vendas por plataforma"
                  subtitle="Top marketplaces e canais por receita"
                  type="bar"
                  data={platformChartData}
                  options={horizontalCurrencyOptions()}
                />
                <ChartCard
                  title="Vendas por categoria"
                  subtitle="Categoria explícita ou derivada do produto/SKU"
                  type="doughnut"
                  data={categoryChartData}
                  options={{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                      legend: { position: "right", labels: { color: "#475569", boxWidth: 11, usePointStyle: true } },
                      tooltip: { callbacks: { label: (context: any) => `${context.label}: ${formatCurrency(Number(context.raw || 0))}` } },
                    },
                  }}
                />
                <ChartCard
                  title={rankMode === "best" ? "Produtos mais vendidos" : "Produtos com menor faturamento"}
                  subtitle="Ranking por receita no período atual"
                  type="bar"
                  data={productChartData}
                  options={horizontalCurrencyOptions()}
                  action={<RankToggle value={rankMode} onChange={setRankMode} />}
                />
                <ChartCard
                  title="Receita por canal"
                  subtitle="Distribuição de faturamento por origem da venda"
                  type="bar"
                  data={channelChartData}
                  options={horizontalCurrencyOptions()}
                />
                <ChartCard
                  title="Faturamento mensal do ano"
                  subtitle="Receita total por mês, respeitando os filtros de plataforma, produto e canal"
                  type="bar"
                  data={monthlyChartData}
                  options={monthlyRevenueOptions(monthlyMaxRevenue)}
                  wide
                />
                <ChartCard
                  title="Comparativo semanal"
                  subtitle="Receita por semana ISO"
                  type="line"
                  data={weeklyChartData}
                  options={chartOptions(compactCurrency)}
                />
                <ChartCard
                  title="Receita x pedidos"
                  subtitle="Cada ponto representa um dia de venda"
                  type="scatter"
                  data={scatterData}
                  options={scatterOptions()}
                />
              </section>

              <section className="section">
                <ChartCard
                  title="Tendência diária de vendas"
                  subtitle="Receita diária dentro do período filtrado"
                  type="line"
                  data={dailyTrendData}
                  options={chartOptions(compactCurrency)}
                  short
                />
              </section>

              <section className="section">
                <SalesHeatmap records={currentRecords} />
              </section>

              <section className="table-grid">
                <ProductRankingTable rows={productBuckets.slice(0, 14)} rankMode={rankMode} setRankMode={setRankMode} />
                <PlatformPerformanceTable rows={platformBuckets} />
              </section>
            </>
          )}
        </div>
      </main>
    </div>
  );
}

function dualAxisOptions(): any {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { labels: { color: "#475569", boxWidth: 11, usePointStyle: true } },
      tooltip: {
        callbacks: {
          label: (context: any) => {
            const value = Number(context.raw || 0);
            return context.dataset.label === "Unidades"
              ? `Unidades: ${formatNumber(value)}`
              : `${context.dataset.label}: ${formatCurrency(value)}`;
          },
        },
      },
    },
    scales: {
      x: { grid: { display: false }, ticks: { color: "#64748b" } },
      y: {
        beginAtZero: true,
        position: "left",
        grid: { color: "rgba(148, 163, 184, 0.22)" },
        ticks: { color: "#64748b", callback: (value: any) => compactCurrency(Number(value)) },
      },
      y1: {
        beginAtZero: true,
        position: "right",
        grid: { drawOnChartArea: false },
        ticks: { color: "#64748b", callback: (value: any) => formatNumber(Number(value)) },
      },
    },
  };
}

function scatterOptions(): any {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: (context: any) => {
            const raw = context.raw || {};
            return `${raw.label || "Dia"}: ${formatNumber(context.parsed.x)} pedidos · ${formatCurrency(context.parsed.y)}`;
          },
        },
      },
    },
    scales: {
      x: {
        beginAtZero: true,
        title: { display: true, text: "Pedidos", color: "#64748b" },
        grid: { color: "rgba(148, 163, 184, 0.18)" },
        ticks: { color: "#64748b" },
      },
      y: {
        beginAtZero: true,
        title: { display: true, text: "Receita", color: "#64748b" },
        grid: { color: "rgba(148, 163, 184, 0.22)" },
        ticks: { color: "#64748b", callback: (value: any) => compactCurrency(Number(value)) },
      },
    },
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
  reset,
}: {
  filters: Filters;
  setFilters: (updater: any) => void;
  options: Record<string, string[]>;
  monthOptions: MonthOption[];
  comparison: ComparisonRange;
  setComparison: (updater: any) => void;
  dateBounds: { min: Date; max: Date } | null;
  reset: () => void;
}) {
  const min = dateBounds ? formatInputDate(dateBounds.min) : undefined;
  const max = dateBounds ? formatInputDate(dateBounds.max) : undefined;

  return (
    <aside className="filters-panel">
      <div className="panel-header">
        <h2 className="panel-title">Filtros</h2>
        <p className="panel-subtitle">A seleção atual recalcula KPIs, gráficos e tabelas em tempo real.</p>
      </div>
      <div className="filters-body">
        <div className="field">
          <span className="field-label">Período</span>
          <div className="field-row">
            <input
              className="input"
              type="date"
              min={min}
              max={max}
              value={filters.startDate}
              onChange={(event) => setFilters((current: Filters) => ({ ...current, startDate: event.target.value }))}
            />
            <input
              className="input"
              type="date"
              min={min}
              max={max}
              value={filters.endDate}
              onChange={(event) => setFilters((current: Filters) => ({ ...current, endDate: event.target.value }))}
            />
          </div>
        </div>

        <MonthComparisonFilter
          monthOptions={monthOptions}
          comparison={comparison}
          setComparison={setComparison}
        />

        <MonthQuickFilter
          monthOptions={monthOptions}
          filters={filters}
          setFilters={setFilters}
          dateBounds={dateBounds}
        />

        <MultiSelect label="Plataforma" options={options.platforms} selected={filters.platforms} onChange={(value) => setFilters((current: Filters) => ({ ...current, platforms: value }))} />
        <MultiSelect label="Canal" options={options.channels} selected={filters.channels} onChange={(value) => setFilters((current: Filters) => ({ ...current, channels: value }))} />
        <MultiSelect label="Produto" options={options.products} selected={filters.products} onChange={(value) => setFilters((current: Filters) => ({ ...current, products: value }))} searchable />
        <MultiSelect label="Categoria" options={options.categories} selected={filters.categories} onChange={(value) => setFilters((current: Filters) => ({ ...current, categories: value }))} />
        <MultiSelect label="Empresa" options={options.companies} selected={filters.companies} onChange={(value) => setFilters((current: Filters) => ({ ...current, companies: value }))} />
        <MultiSelect label="Tipo cliente" options={options.customerTypes} selected={filters.customerTypes} onChange={(value) => setFilters((current: Filters) => ({ ...current, customerTypes: value }))} />

        <div className="field">
          <label htmlFor="global-search">Busca</label>
          <input
            id="global-search"
            className="input"
            value={filters.search}
            placeholder="Pedido, SKU, cliente..."
            onChange={(event) => setFilters((current: Filters) => ({ ...current, search: event.target.value }))}
          />
        </div>

        <div className="filter-footer">
          <button className="btn" type="button" onClick={reset}>Limpar</button>
        </div>
      </div>
    </aside>
  );
}

function MonthComparisonFilter({
  monthOptions,
  comparison,
  setComparison,
}: {
  monthOptions: MonthOption[];
  comparison: ComparisonRange;
  setComparison: (updater: any) => void;
}) {
  if (!monthOptions.length) return null;

  const selectedA = monthOptions.find((month) => month.startDate === comparison.aStart && month.endDate === comparison.aEnd)?.key || "";
  const selectedB = monthOptions.find((month) => month.startDate === comparison.bStart && month.endDate === comparison.bEnd)?.key || "";

  function updateMonth(slot: "a" | "b", key: string) {
    const selected = monthOptions.find((month) => month.key === key);
    if (!selected) return;
    setComparison((current: ComparisonRange) => slot === "a"
      ? { ...current, aStart: selected.startDate, aEnd: selected.endDate }
      : { ...current, bStart: selected.startDate, bEnd: selected.endDate });
  }

  return (
    <div className="field">
      <span className="field-label">Comparativo mensal</span>
      <div className="month-comparison-grid">
        <label className="select-field">
          <span>Mês A</span>
          <select className="input" value={selectedA} onChange={(event) => updateMonth("a", event.target.value)}>
            <option value="" disabled>Selecione</option>
            {monthOptions.map((month) => (
              <option key={month.key} value={month.key}>{month.label}</option>
            ))}
          </select>
        </label>
        <label className="select-field">
          <span>Mês B</span>
          <select className="input" value={selectedB} onChange={(event) => updateMonth("b", event.target.value)}>
            <option value="" disabled>Selecione</option>
            {monthOptions.map((month) => (
              <option key={month.key} value={month.key}>{month.label}</option>
            ))}
          </select>
        </label>
      </div>
    </div>
  );
}

function MonthQuickFilter({
  monthOptions,
  filters,
  setFilters,
  dateBounds,
}: {
  monthOptions: MonthOption[];
  filters: Filters;
  setFilters: (updater: any) => void;
  dateBounds: { min: Date; max: Date } | null;
}) {
  if (!monthOptions.length) return null;
  const fullStart = dateBounds ? formatInputDate(dateBounds.min) : "";
  const fullEnd = dateBounds ? formatInputDate(dateBounds.max) : "";
  const isAllActive = filters.startDate === fullStart && filters.endDate === fullEnd;

  return (
    <div className="field">
      <span className="field-label">Mês</span>
      <div className="month-button-grid">
        <button
          type="button"
          className={`month-button ${isAllActive ? "active" : ""}`}
          onClick={() => setFilters((current: Filters) => ({ ...current, startDate: fullStart, endDate: fullEnd }))}
        >
          Todos
        </button>
        {monthOptions.map((month) => {
          const active = filters.startDate === month.startDate && filters.endDate === month.endDate;
          return (
            <button
              type="button"
              className={`month-button ${active ? "active" : ""}`}
              key={month.key}
              onClick={() => setFilters((current: Filters) => ({ ...current, startDate: month.startDate, endDate: month.endDate }))}
            >
              {month.label}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function MultiSelect({
  label,
  options,
  selected,
  onChange,
  searchable = false,
}: {
  label: string;
  options: string[];
  selected: string[];
  onChange: (value: string[]) => void;
  searchable?: boolean;
}) {
  const [search, setSearch] = useState("");
  const visibleOptions = options.filter((option) => normalizeText(option).includes(normalizeText(search))).slice(0, 160);
  const summary = selected.length === 0 ? "Todos" : `${selected.length} selecionado${selected.length > 1 ? "s" : ""}`;

  function toggle(value: string) {
    if (selected.includes(value)) {
      onChange(selected.filter((item) => item !== value));
    } else {
      onChange([...selected, value]);
    }
  }

  return (
    <div className="field">
      <span className="field-label">{label}</span>
      <details className="multi-select">
        <summary>{summary}<span aria-hidden="true">⌄</span></summary>
        <div className="multi-select-body">
          {searchable && (
            <input
              className="input search-input"
              value={search}
              placeholder="Filtrar opções"
              onChange={(event) => setSearch(event.target.value)}
            />
          )}
          <label className="check-row">
            <input type="checkbox" checked={selected.length === 0} onChange={() => onChange([])} />
            <span>Todos</span>
          </label>
          {visibleOptions.map((option) => (
            <label className="check-row" key={option} title={option}>
              <input type="checkbox" checked={selected.includes(option)} onChange={() => toggle(option)} />
              <span>{option}</span>
            </label>
          ))}
        </div>
      </details>
    </div>
  );
}

function GranularitySelector({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  const items = [
    ["day", "Dia"],
    ["week", "Semana"],
    ["month", "Mês"],
    ["quarter", "Trimestre"],
    ["year", "Ano"],
  ];
  return (
    <div className="segmented" aria-label="Agrupamento">
      {items.map(([key, label]) => (
        <button key={key} className={`segment ${value === key ? "active" : ""}`} onClick={() => onChange(key)}>
          {label}
        </button>
      ))}
    </div>
  );
}

function KPICard({
  label,
  value,
  delta: deltaValue,
  foot,
}: {
  label: string;
  value: string;
  delta?: { percent: number; className: string };
  foot?: string;
}) {
  return (
    <div className="kpi-card">
      <div className="kpi-label">
        <span>{label}</span>
        {deltaValue && <span className={`delta ${deltaValue.className}`}>{formatPercent(deltaValue.percent)}</span>}
      </div>
      <div className="kpi-value mono">{value}</div>
      {foot && <div className="kpi-foot">{foot}</div>}
    </div>
  );
}

function DateComparisonSelector({
  comparison,
  setComparison,
  metricsA,
  metricsB,
}: {
  comparison: ComparisonRange;
  setComparison: (updater: any) => void;
  metricsA: Metrics;
  metricsB: Metrics;
}) {
  const rows = [
    ["Receita", metricsA.revenue, metricsB.revenue, formatCurrency],
    ["Pedidos", metricsA.orders, metricsB.orders, formatNumber],
    ["Unidades", metricsA.units, metricsB.units, formatNumber],
    ["Ticket médio", metricsA.aov, metricsB.aov, formatCurrency],
  ] as Array<[string, number, number, (value: number) => string]>;

  return (
    <section className="comparison-panel">
      <div className="section-header">
        <div>
          <h2 className="section-title">Comparação personalizada</h2>
          <p className="section-subtitle">Período A contra período B, mantendo os demais filtros ativos.</p>
        </div>
      </div>
      <div className="comparison-controls">
        <div className="field">
          <label>Início A</label>
          <input className="input" type="date" value={comparison.aStart} onChange={(event) => setComparison((current: any) => ({ ...current, aStart: event.target.value }))} />
        </div>
        <div className="field">
          <label>Fim A</label>
          <input className="input" type="date" value={comparison.aEnd} onChange={(event) => setComparison((current: any) => ({ ...current, aEnd: event.target.value }))} />
        </div>
        <div className="field">
          <label>Início B</label>
          <input className="input" type="date" value={comparison.bStart} onChange={(event) => setComparison((current: any) => ({ ...current, bStart: event.target.value }))} />
        </div>
        <div className="field">
          <label>Fim B</label>
          <input className="input" type="date" value={comparison.bEnd} onChange={(event) => setComparison((current: any) => ({ ...current, bEnd: event.target.value }))} />
        </div>
      </div>
      <div className="table-scroll">
        <table>
          <thead>
            <tr>
              <th>Métrica</th>
              <th>Período A</th>
              <th>Período B</th>
              <th>Diferença</th>
              <th>Variação</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(([label, a, b, formatter]) => {
              const rowDelta = delta(a, b);
              return (
                <tr key={label}>
                  <td>{label}</td>
                  <td className="mono">{formatter(a)}</td>
                  <td className="mono">{formatter(b)}</td>
                  <td className="mono">{formatter(a - b)}</td>
                  <td><span className={`delta ${rowDelta.className}`}>{formatPercent(rowDelta.percent)}</span></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function RankToggle({ value, onChange }: { value: string; onChange: (value: "best" | "worst") => void }) {
  return (
    <div className="segmented">
      <button className={`segment ${value === "best" ? "active" : ""}`} onClick={() => onChange("best")}>Top</button>
      <button className={`segment ${value === "worst" ? "active" : ""}`} onClick={() => onChange("worst")}>Piores</button>
    </div>
  );
}

function ChartCard({
  title,
  subtitle,
  type,
  data,
  options,
  action,
  short = false,
  wide = false,
}: {
  title: string;
  subtitle: string;
  type: string;
  data: any;
  options: any;
  action?: any;
  short?: boolean;
  wide?: boolean;
}) {
  return (
    <div className={`chart-card ${wide ? "wide" : ""}`}>
      <div className="chart-title-row">
        <div>
          <h3 className="chart-title">{title}</h3>
          <p className="chart-subtitle">{subtitle}</p>
        </div>
        {action}
      </div>
      <div className={`chart-wrap ${short ? "short" : ""}`}>
        <ChartCanvas type={type} data={data} options={options} />
      </div>
    </div>
  );
}

function ChartCanvas({ type, data, options }: { type: string; data: any; options: any }) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const chartRef = useRef<any>(null);
  const configKey = JSON.stringify({ type, data, options });

  useEffect(() => {
    if (typeof Chart === "undefined" || !canvasRef.current) return undefined;
    registerMonthlyValueLabelsPlugin();
    if (chartRef.current) chartRef.current.destroy();
    chartRef.current = new Chart(canvasRef.current, {
      type,
      data,
      options,
    });
    return () => {
      if (chartRef.current) {
        chartRef.current.destroy();
        chartRef.current = null;
      }
    };
  }, [configKey]);

  if (typeof Chart === "undefined") {
    return <div className="chart-empty">Biblioteca de gráficos não carregada.</div>;
  }

  const hasData = data?.datasets?.some((dataset: any) => Array.isArray(dataset.data) && dataset.data.length);
  if (!hasData) {
    return <div className="chart-empty">Sem dados para o filtro atual.</div>;
  }

  return <canvas ref={canvasRef} />;
}

function SalesHeatmap({ records }: { records: SaleRecord[] }) {
  const buckets = useMemo(() => {
    const grouped = new Map<string, number>();
    records.forEach((record) => grouped.set(record.dateKey, (grouped.get(record.dateKey) || 0) + record.totalValue));
    return grouped;
  }, [records]);

  const dates = useMemo(() => {
    if (!records.length) return [];
    const bounds = getDateBounds(records);
    if (!bounds) return [];
    const start = addDays(bounds.min, -((bounds.min.getDay() + 6) % 7));
    const end = addDays(bounds.max, 6 - ((bounds.max.getDay() + 6) % 7));
    const output: Date[] = [];
    for (let day = start; day <= end; day = addDays(day, 1)) output.push(day);
    return output;
  }, [records]);

  const maxRevenue = Math.max(...Array.from(buckets.values()), 0);
  const dayLabels = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"];

  return (
    <div className="heatmap-card">
      <div className="chart-title-row">
        <div>
          <h3 className="chart-title">Heatmap de melhores dias</h3>
          <p className="chart-subtitle">Intensidade por receita diária no período filtrado</p>
        </div>
      </div>
      <div className="heatmap-grid">
        {dayLabels.map((label) => <div className="heatmap-day-label" key={label}>{label}</div>)}
        {dates.map((date) => {
          const key = formatInputDate(date);
          const revenue = buckets.get(key) || 0;
          const intensity = maxRevenue ? Math.max(0.08, revenue / maxRevenue) : 0;
          const isEmpty = !buckets.has(key);
          const background = isEmpty
            ? undefined
            : `rgba(15, 118, 110, ${0.12 + intensity * 0.68})`;
          return (
            <div
              className={`heatmap-cell ${isEmpty ? "empty" : ""}`}
              key={key}
              style={{ background }}
              title={`${formatShortDate(date)} · ${formatCurrencyDetailed(revenue)}`}
            >
              <span className="heatmap-date">{date.getDate()}</span>
              <span className="heatmap-value">{revenue ? compactCurrency(revenue) : ""}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ProductRankingTable({
  rows,
  rankMode,
  setRankMode,
}: {
  rows: any[];
  rankMode: "best" | "worst";
  setRankMode: (value: "best" | "worst") => void;
}) {
  return (
    <div className="table-card">
      <div className="table-head">
        <h3 className="table-title">Ranking de produtos</h3>
        <RankToggle value={rankMode} onChange={setRankMode} />
      </div>
      <div className="table-scroll product-table-scroll">
        <table className="compact-product-table">
          <thead>
            <tr>
              <th>Produto</th>
              <th>Categoria</th>
              <th>Receita</th>
              <th>Ped.</th>
              <th>Unid.</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={`${row.product}-${row.sku}-${row.category}`}>
                <td>
                  <span className="rank-name">
                    <strong>{row.product}</strong>
                    <span>{row.sku}</span>
                  </span>
                </td>
                <td>{row.category}</td>
                <td className="mono">{formatCurrency(row.revenue)}</td>
                <td className="mono">{formatNumber(row.ordersCount)}</td>
                <td className="mono">{formatNumber(row.units)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function PlatformPerformanceTable({ rows }: { rows: any[] }) {
  return (
    <div className="table-card">
      <div className="table-head">
        <h3 className="table-title">Performance por plataforma</h3>
      </div>
      <div className="table-scroll platform-table-scroll">
        <table className="compact-platform-table">
          <thead>
            <tr>
              <th>Plataforma</th>
              <th>Receita</th>
              <th>Ped.</th>
              <th>Unid.</th>
              <th>Ticket</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row.name}>
                <td>{row.name}</td>
                <td className="mono">{formatCurrency(row.revenue)}</td>
                <td className="mono">{formatNumber(row.ordersCount)}</td>
                <td className="mono">{formatNumber(row.units)}</td>
                <td className="mono">{formatCurrency(row.aov)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<App />);
