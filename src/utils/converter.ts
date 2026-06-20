import type { ConversionResult, ExcelCell, ExcelRow } from "../types";

export const EXCISE_TAX_COLUMN =
  "Congo DRC Telecommunication Excise Tax_Final";
export const DATE_FACTURE_COLUMN = "Date facture";
export const COMMENT_B_COLUMN = "Commentaire B";
export const OLD_DATE_COLUMN = "Old Date";

const EXCISE_TAX_COLUMN_CANDIDATES = [
  EXCISE_TAX_COLUMN,
  "Congo DRC Telecommunication Excise Tax Final",
  "congo_drc_telecommunication_excise_tax_final",
  "telecommunication_excise_tax_final",
  "telecom_excise_tax_final",
  "excise_tax_final",
  "Excise Tax Final",
  "Telecommunication Excise Tax",
  "Telecom Excise Tax",
];

const DATE_FACTURE_COLUMN_CANDIDATES = [
  DATE_FACTURE_COLUMN,
  "Date de facture",
  "Date facture original",
  "Original invoice date",
  "original_invoice_date",
  "invoice_date",
  "Invoice Date",
  "Billing Date",
];

type ConversionOptions = {
  normalizationDate?: Date;
};

type DexyFieldKey =
  | "lineNumber"
  | "itemReference"
  | "itemDescription"
  | "itemType"
  | "taxGroup"
  | "quantity"
  | "unitPrice"
  | "unitPriceMode";

type InvoiceGroup = {
  invoiceNumber: string;
  rows: ExcelRow[];
  taxTotal: number;
};

const LINE_NUMBER_FALLBACK_COLUMN = "N° ligne";

const INVOICE_COLUMN_CANDIDATES = [
  "Numéro de facture",
  "Numero de facture",
  "Numéro facture",
  "Numero facture",
  "N° facture",
  "N° de facture",
  "No facture",
  "Nº facture",
  "Facture",
  "Invoice Number",
  "Invoice No",
  "Invoice",
  "Billing Invoice",
];

const DEXY_FIELD_CANDIDATES: Record<DexyFieldKey, string[]> = {
  lineNumber: [
    "Numéro de ligne",
    "Numero de ligne",
    "Numéro ligne",
    "Numero ligne",
    "Numéro ligne facture",
    "Numero ligne facture",
    "N° ligne",
    "N° de ligne",
    "N° ligne facture",
    "N° de ligne facture",
    "No ligne",
    "No de ligne",
    "No ligne facture",
    "Nº ligne",
    "Nº de ligne",
    "Ligne facture",
    "Ligne article",
    "Num ligne",
    "Num. ligne",
    "N ligne",
    "Line Number",
    "Line No",
    "Invoice Line Number",
    "Invoice Line No",
    "Line Item Number",
    "Line",
  ],
  itemReference: [
    "Référence article",
    "Reference article",
    "Réf. article",
    "Article Reference",
    "Item Reference",
    "SKU",
  ],
  itemDescription: [
    "Désignation article",
    "Designation article",
    "Libellé article",
    "Libelle article",
    "Article Description",
    "Item Description",
    "Description",
  ],
  itemType: ["Type article", "Article Type", "Item Type", "Line Type", "Type"],
  taxGroup: [
    "Groupe taxation",
    "Groupe taxe",
    "Tax Group",
    "Taxation Group",
  ],
  quantity: ["Quantité", "Quantite", "Quantity", "Qty"],
  unitPrice: ["Prix unitaire", "Unit Price", "Prix unitaire HT"],
  unitPriceMode: [
    "Mode prix unitaire [HT | TTC]",
    "Mode prix unitaire HT TTC",
    "Mode prix unitaire",
    "Mode prix",
    "Unit Price Mode",
    "Price Mode",
  ],
};

const DEFAULT_DEXY_COLUMNS: Record<DexyFieldKey, string> = {
  lineNumber: LINE_NUMBER_FALLBACK_COLUMN,
  itemReference: "Référence article",
  itemDescription: "Désignation article",
  itemType: "Type article",
  taxGroup: "Groupe taxation",
  quantity: "Quantité",
  unitPrice: "Prix unitaire",
  unitPriceMode: "Mode prix unitaire",
};

const TAX_LINE_VALUES = {
  itemReference: "ACCISE",
  itemDescription: "Congo DRC Telecommunication Excise Tax (10%)",
  itemType: "TAX",
  taxGroup: "L",
  quantity: 1,
  unitPriceMode: "HT",
} as const;

export function convertStarlinkToDexy(
  rows: ExcelRow[],
  sourceColumns?: string[],
  options: ConversionOptions = {},
): ConversionResult {
  if (rows.length === 0) {
    throw new Error("Le fichier ne contient aucune ligne à convertir.");
  }

  const columns = normalizeColumns(sourceColumns ?? collectColumns(rows));
  const taxColumn = findExciseTaxColumn(columns);
  const invoiceColumn = findColumn(columns, INVOICE_COLUMN_CANDIDATES);
  const missingColumns = [
    !invoiceColumn ? "Numéro facture" : null,
    !taxColumn ? EXCISE_TAX_COLUMN : null,
  ].filter((column): column is string => Boolean(column));

  if (!taxColumn || !invoiceColumn) {
    throw new Error(
      `Les colonnes obligatoires suivantes sont absentes de la feuille : ${missingColumns.join(
        ", ",
      )}.`,
    );
  }

  const dateColumn =
    findColumn(columns, DATE_FACTURE_COLUMN_CANDIDATES) ?? DATE_FACTURE_COLUMN;
  const normalizationDate = toDateOnly(options.normalizationDate ?? new Date());
  const dexyColumns = resolveDexyColumns(columns);
  const outputColumns = buildOutputColumns(
    columns,
    taxColumn,
    dateColumn,
    dexyColumns,
  );
  const groups = groupRowsByInvoice(rows, invoiceColumn, taxColumn);
  const warnings: string[] = [];
  const technicalRows: ExcelRow[] = [];
  let totalTax = 0;
  let taxRowsAdded = 0;

  for (const group of groups) {
    group.rows.forEach((row, index) => {
      const outputRow = createTraceableRow(
        row,
        taxColumn,
        invoiceColumn,
        dateColumn,
        group.invoiceNumber,
        normalizationDate,
      );
      outputRow[dexyColumns.lineNumber] = index + 1;
      technicalRows.push(outputRow);
    });

    const roundedTax = roundCurrency(group.taxTotal);

    if (roundedTax !== 0) {
      taxRowsAdded += 1;
      totalTax += roundedTax;
      technicalRows.push(
        createTaxRow(
          group,
          taxColumn,
          invoiceColumn,
          dateColumn,
          dexyColumns,
          normalizationDate,
          group.rows.length + 1,
        ),
      );
    }

    if (group.invoiceNumber === "") {
      warnings.push(
        "Au moins une ligne n'a pas de numéro de facture ; elle a été regroupée sous une facture vide.",
      );
    }
  }

  const outputRows = technicalRows.map((row) =>
    stripColumn(row, OLD_DATE_COLUMN),
  );

  return {
    rows: outputRows,
    columns: outputColumns,
    summary: {
      invoicesProcessed: groups.length,
      originalRows: rows.length,
      taxRowsAdded,
      totalTax: roundCurrency(totalTax),
    },
    invoiceColumn,
    taxColumn,
    normalizationDate,
    warnings: Array.from(new Set(warnings)),
  };
}

function groupRowsByInvoice(
  rows: ExcelRow[],
  invoiceColumn: string,
  taxColumn: string,
): InvoiceGroup[] {
  const groupsByInvoice = new Map<string, InvoiceGroup>();

  for (const row of rows) {
    const invoiceNumber = normalizeInvoiceNumber(row[invoiceColumn]);
    const existingGroup = groupsByInvoice.get(invoiceNumber);
    const group =
      existingGroup ??
      {
        invoiceNumber,
        rows: [],
        taxTotal: 0,
      };

    group.rows.push({ ...row });
    group.taxTotal += parseAmount(row[taxColumn]);
    groupsByInvoice.set(invoiceNumber, group);
  }

  return Array.from(groupsByInvoice.values());
}

function createTaxRow(
  group: InvoiceGroup,
  taxColumn: string,
  invoiceColumn: string,
  dateColumn: string,
  dexyColumns: Record<DexyFieldKey, string>,
  normalizationDate: Date,
  lineNumber: number,
): ExcelRow {
  const baseRow = createTraceableRow(
    group.rows[0] ?? {},
    taxColumn,
    invoiceColumn,
    dateColumn,
    group.invoiceNumber,
    normalizationDate,
  );
  const taxTotal = roundCurrency(group.taxTotal);

  return {
    ...baseRow,
    [dexyColumns.lineNumber]: lineNumber,
    [dexyColumns.itemReference]: TAX_LINE_VALUES.itemReference,
    [dexyColumns.itemDescription]: TAX_LINE_VALUES.itemDescription,
    [dexyColumns.itemType]: TAX_LINE_VALUES.itemType,
    [dexyColumns.taxGroup]: TAX_LINE_VALUES.taxGroup,
    [dexyColumns.quantity]: TAX_LINE_VALUES.quantity,
    [dexyColumns.unitPrice]: taxTotal,
    [dexyColumns.unitPriceMode]: TAX_LINE_VALUES.unitPriceMode,
  };
}

function createTraceableRow(
  row: ExcelRow,
  taxColumn: string,
  invoiceColumn: string,
  dateColumn: string,
  invoiceNumber: string,
  normalizationDate: Date,
): ExcelRow {
  const outputRow = stripColumn(row, taxColumn);
  const oldDate = row[dateColumn];
  const oldDateDisplay = formatDateForTrace(oldDate);
  const resolvedInvoiceNumber =
    invoiceNumber || normalizeInvoiceNumber(row[invoiceColumn]);

  outputRow[dateColumn] = normalizationDate;
  outputRow[DATE_FACTURE_COLUMN] = normalizationDate;
  outputRow[OLD_DATE_COLUMN] = oldDateDisplay;
  outputRow[COMMENT_B_COLUMN] =
    `Origine : ${resolvedInvoiceNumber} - ${oldDateDisplay}`;

  return outputRow;
}

function resolveDexyColumns(columns: string[]): Record<DexyFieldKey, string> {
  const resolvedColumns = Object.fromEntries(
    (Object.keys(DEXY_FIELD_CANDIDATES) as DexyFieldKey[]).map((key) => [
      key,
      findColumn(columns, DEXY_FIELD_CANDIDATES[key]) ??
        DEFAULT_DEXY_COLUMNS[key],
    ]),
  ) as Record<DexyFieldKey, string>;

  resolvedColumns.lineNumber =
    findColumn(columns, DEXY_FIELD_CANDIDATES.lineNumber) ??
    findLikelyLineNumberColumn(columns) ??
    DEFAULT_DEXY_COLUMNS.lineNumber;

  return resolvedColumns;
}

function buildOutputColumns(
  sourceColumns: string[],
  taxColumn: string,
  dateColumn: string,
  dexyColumns: Record<DexyFieldKey, string>,
): string[] {
  const outputColumns = sourceColumns.filter(
    (column) => column !== taxColumn && column !== OLD_DATE_COLUMN,
  );

  for (const column of Object.values(dexyColumns)) {
    if (!outputColumns.includes(column)) {
      outputColumns.push(column);
    }
  }

  for (const column of [dateColumn, DATE_FACTURE_COLUMN, COMMENT_B_COLUMN]) {
    if (!outputColumns.includes(column)) {
      outputColumns.push(column);
    }
  }

  return outputColumns;
}

function collectColumns(rows: ExcelRow[]): string[] {
  const columns: string[] = [];

  for (const row of rows) {
    for (const column of Object.keys(row)) {
      if (!columns.includes(column)) {
        columns.push(column);
      }
    }
  }

  return columns;
}

function normalizeColumns(columns: string[]): string[] {
  return columns
    .map((column) => column.trim())
    .filter((column, index, allColumns) => {
      return column !== "" && allColumns.indexOf(column) === index;
    });
}

function findColumn(columns: string[], candidates: string[]): string | undefined {
  const normalizedCandidates = candidates.map(normalizeHeader);

  return columns.find((column) =>
    normalizedCandidates.includes(normalizeHeader(column)),
  );
}

function findExciseTaxColumn(columns: string[]): string | undefined {
  const exactMatch = findColumn(columns, EXCISE_TAX_COLUMN_CANDIDATES);

  if (exactMatch) {
    return exactMatch;
  }

  const looseMatches = columns.filter((column) => {
    const normalizedColumn = normalizeHeader(column);
    const words = normalizedColumn.split(" ").filter(Boolean);
    const hasExciseTax = words.includes("excise") && words.includes("tax");
    const hasFinal = words.includes("final");
    const hasTelecomSignal =
      words.includes("telecommunication") ||
      words.includes("telecom") ||
      normalizedColumn.includes("congo drc");

    return hasExciseTax && (hasFinal || hasTelecomSignal);
  });

  if (looseMatches.length === 1) {
    return looseMatches[0];
  }

  return (
    looseMatches.find((column) =>
      normalizeHeader(column).split(" ").includes("final"),
    ) ?? looseMatches[0]
  );
}

function findLikelyLineNumberColumn(columns: string[]): string | undefined {
  const invoiceLikeColumns = new Set(
    INVOICE_COLUMN_CANDIDATES.map(normalizeHeader),
  );

  return columns.find((column) => {
    const normalizedColumn = normalizeHeader(column);
    const words = normalizedColumn.split(" ").filter(Boolean);
    const hasLineSignal =
      words.includes("ligne") ||
      words.includes("line") ||
      normalizedColumn.includes("ligne facture") ||
      normalizedColumn.includes("invoice line");
    const hasNumberSignal =
      words.includes("numero") ||
      words.includes("num") ||
      words.includes("no") ||
      words.includes("n") ||
      words.includes("number");

    return (
      hasLineSignal &&
      (hasNumberSignal || normalizedColumn.includes("ligne")) &&
      !invoiceLikeColumns.has(normalizedColumn)
    );
  });
}

function normalizeHeader(header: string): string {
  return header
    .replace(/\u00a0/g, " ")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\u200b-\u200d\ufeff]/g, "")
    .replace(/[°º]/g, "o")
    .replace(/[\[\]()/|]+/g, " ")
    .replace(/[_\s-]+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeInvoiceNumber(value: ExcelCell): string {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  return String(value).trim();
}

function parseAmount(value: ExcelCell): number {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  if (typeof value !== "string") {
    return 0;
  }

  const normalizedValue = value
    .replace(/\s/g, "")
    .replace(/[A-Za-z$€CDF]+/g, "")
    .replace(/^\((.*)\)$/, "-$1");
  const lastCommaIndex = normalizedValue.lastIndexOf(",");
  const lastDotIndex = normalizedValue.lastIndexOf(".");
  let decimalValue = normalizedValue;

  if (lastCommaIndex > -1 && lastDotIndex > -1) {
    decimalValue =
      lastCommaIndex > lastDotIndex
        ? normalizedValue.replace(/\./g, "").replace(",", ".")
        : normalizedValue.replace(/,/g, "");
  } else if (lastCommaIndex > -1) {
    decimalValue = normalizedValue.replace(",", ".");
  }

  const parsedValue = Number.parseFloat(decimalValue);
  return Number.isFinite(parsedValue) ? parsedValue : 0;
}

function formatDateForTrace(value: ExcelCell): string {
  const dateValue = parseDateValue(value);

  if (dateValue) {
    return formatDate(dateValue);
  }

  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function parseDateValue(value: ExcelCell): Date | null {
  if (value instanceof Date) {
    return toDateOnly(value);
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialToDate(value);
  }

  if (typeof value !== "string") {
    return null;
  }

  const dmyDate = parseDayMonthYearDate(value);

  if (dmyDate) {
    return dmyDate;
  }

  const parsedTimestamp = Date.parse(value);
  return Number.isNaN(parsedTimestamp)
    ? null
    : toDateOnly(new Date(parsedTimestamp));
}

function parseDayMonthYearDate(value: string): Date | null {
  const match = value.trim().match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);

  if (!match) {
    return null;
  }

  const day = Number.parseInt(match[1], 10);
  const month = Number.parseInt(match[2], 10);
  const yearInput = Number.parseInt(match[3], 10);
  const year = yearInput < 100 ? 2000 + yearInput : yearInput;

  if (
    !Number.isInteger(day) ||
    !Number.isInteger(month) ||
    !Number.isInteger(year) ||
    day < 1 ||
    day > 31 ||
    month < 1 ||
    month > 12
  ) {
    return null;
  }

  return toDateOnly(new Date(year, month - 1, day));
}

function excelSerialToDate(value: number): Date | null {
  if (value <= 0) {
    return null;
  }

  const milliseconds = Math.round((value - 25569) * 86400000);
  return toDateOnly(new Date(milliseconds));
}

function formatDate(value: Date): string {
  const day = String(value.getDate()).padStart(2, "0");
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const year = value.getFullYear();
  return `${day}/${month}/${year}`;
}

function toDateOnly(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

function roundCurrency(value: number): number {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function stripColumn(row: ExcelRow, columnToRemove: string): ExcelRow {
  const { [columnToRemove]: _removedColumn, ...remainingColumns } = row;
  return { ...remainingColumns };
}
