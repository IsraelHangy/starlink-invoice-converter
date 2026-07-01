import * as XLSX from "xlsx";
import type { ExcelCell, ExcelRow, ImportedWorkbook } from "../types";
import {
  CREDIT_NOTE_ORIGIN_ERP_DATE_COLUMN,
  CREDIT_NOTE_ORIGIN_ERP_INVOICE_COLUMN,
  DATE_FACTURE_COLUMN,
  EXCISE_TAX_COLUMN,
  OLD_DATE_COLUMN,
} from "./converter";

type SheetMatrix = (string | number | boolean | Date | null | undefined)[][];
type TemplateColumnFormat = {
  style?: XLSX.CellObject["s"];
  numberFormat?: string;
  sampleType?: XLSX.ExcelDataType;
  dateAsSerial?: boolean;
  decimals?: number;
};

const HEADER_SCAN_LIMIT = 50;
const DEXY_TEMPLATE_URL =
  "/templates/TemplateImport_Dexy_CD_STARLINK_TEST.xlsx";
const DEXY_TEMPLATE_SHEET = "LIGNES";
const INVOICE_DATE_COLUMN = DATE_FACTURE_COLUMN;
const UNIT_PRICE_COLUMN = "Prix Unitaire";
const ORIGINAL_PRICE_COLUMN = "Prix Hors remise";
const B2F_CURRENCY_NAME_COLUMN = "B2F Devise [Nom]";
const B2F_EXCHANGE_RATE_DATE_COLUMN = "B2F Devise [Date cours]";
const B2F_EXCHANGE_RATE_COLUMN = "B2F Devise [Taux de change]";
const IMPORT_CURRENCY_COLUMN = "Devise Import";
const IMPORT_EXCHANGE_RATE_COLUMN = "Taux Devise Import vers CDF";
const EXTRA_EXPORT_COLUMNS = [
  CREDIT_NOTE_ORIGIN_ERP_INVOICE_COLUMN,
  CREDIT_NOTE_ORIGIN_ERP_DATE_COLUMN,
];
const INVOICE_NUMBER_COLUMN_CANDIDATES = [
  "Numéro facture",
  "Numero facture",
  "Numéro de facture",
  "Numero de facture",
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
const SOURCE_DOCUMENT_NUMBER_COLUMN_CANDIDATES = [
  "Numéro Document",
  "Numero Document",
  "Numéro de Document",
  "Numero de Document",
  "Numéro Documents",
  "Numero Documents",
  "Numéro de Documents",
  "Numero de Documents",
  "Document Number",
  "Document No",
];
const B2F_CURRENCY_COLUMN_CANDIDATES = [
  B2F_CURRENCY_NAME_COLUMN,
  "B2F Devise Nom",
  "Devise Import",
  "Currency",
  "Transaction Currency",
  "transaction_currency",
];
const PAYMENT_METHOD_DEFAULT = "VIREMENT";
const OPERATOR_CODE_DEFAULT = "ADM";
const OPERATOR_NAME_DEFAULT = "Administrateur";
const PAYMENT_METHOD_COLUMN_CANDIDATES = [
  "Méthode de paiement",
  "Methode de paiement",
  "Mode de paiement",
  "Mode paiement",
  "Moyen de paiement",
  "Payment Method",
];
const OPERATOR_CODE_COLUMN_CANDIDATES = [
  "Opérateur [Code]",
  "Operateur [Code]",
  "Code opérateur",
  "Code operateur",
  "Operator Code",
];
const OPERATOR_NAME_COLUMN_CANDIDATES = [
  "Opérateur [Nom]",
  "Operateur [Nom]",
  "Nom opérateur",
  "Nom operateur",
  "Operator Name",
];

export async function readExcelFile(file: File): Promise<ImportedWorkbook> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, {
    type: "array",
    cellDates: true,
  });

  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("Le fichier Excel ne contient aucune feuille.");
  }

  const worksheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json<
    (string | number | boolean | Date | null | undefined)[]
  >(worksheet, {
    header: 1,
    defval: null,
    raw: true,
    blankrows: false,
  });
  const headerIndex = detectHeaderRowIndex(matrix);
  const columns = extractColumns(matrix, headerIndex);
  const rows = extractRows(matrix, headerIndex, columns);

  return {
    fileName: file.name,
    sheetName,
    uploadedAt: new Date(),
    rows,
    columns,
  };
}

export function exportExcelFile(
  rows: ExcelRow[],
  columns: string[],
  fileName: string,
): Promise<void> {
  return exportExcelFileFromTemplate(rows, columns, fileName);
}

async function exportExcelFileFromTemplate(
  rows: ExcelRow[],
  fallbackColumns: string[],
  fileName: string,
): Promise<void> {
  const response = await fetch(DEXY_TEMPLATE_URL);

  if (!response.ok) {
    throw new Error("Template Excel DEXY introuvable pour l'export.");
  }

  const templateData = await response.arrayBuffer();
  const workbook = XLSX.read(templateData, {
    type: "array",
    cellDates: true,
    cellNF: true,
    cellStyles: true,
  });
  const sheetName = workbook.SheetNames.includes(DEXY_TEMPLATE_SHEET)
    ? DEXY_TEMPLATE_SHEET
    : workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("Le template DEXY ne contient aucune feuille.");
  }

  const worksheet = workbook.Sheets[sheetName];
  const headerIndex = detectHeaderRowIndexFromColumns(worksheet, fallbackColumns);
  const templateColumns = extractColumnsFromWorksheet(worksheet, headerIndex);
  const exportColumns = resolveExportColumns(templateColumns, fallbackColumns);
  const templateColumnIndexes = createTemplateColumnIndexes(templateColumns);
  const columnFormats = exportColumns.map((column, index) =>
    getTemplateColumnFormat(
      worksheet,
      headerIndex + 1,
      resolveTemplateColumnIndex(column, index, templateColumnIndexes),
      column,
    ),
  );

  writeHeaderRow(
    worksheet,
    headerIndex,
    exportColumns,
    templateColumnIndexes,
  );
  clearDataRows(worksheet, headerIndex, exportColumns.length);
  writeRowsToTemplate(worksheet, rows, exportColumns, columnFormats, headerIndex);
  updateWorksheetRange(worksheet, headerIndex, exportColumns.length, rows.length);

  XLSX.writeFile(workbook, fileName, {
    bookType: "xlsx",
    cellDates: true,
  });
}

function resolveExportColumns(
  templateColumns: string[],
  fallbackColumns: string[],
): string[] {
  const standardColumns = templateColumns.filter(
    (column) => !isExciseTaxExportColumn(column),
  );
  const exportColumns = [...standardColumns];

  appendExtraExportColumns(exportColumns, fallbackColumns);

  if (!hasColumn(fallbackColumns, OLD_DATE_COLUMN)) {
    return exportColumns;
  }

  if (hasColumn(exportColumns, OLD_DATE_COLUMN)) {
    return exportColumns;
  }

  const invoiceDateIndex = exportColumns.findIndex((column) =>
    isSameColumn(column, INVOICE_DATE_COLUMN),
  );
  const insertIndex =
    invoiceDateIndex >= 0 ? invoiceDateIndex + 1 : exportColumns.length;

  exportColumns.splice(insertIndex, 0, OLD_DATE_COLUMN);
  return exportColumns;
}

function appendExtraExportColumns(
  exportColumns: string[],
  fallbackColumns: string[],
): void {
  for (const extraColumn of EXTRA_EXPORT_COLUMNS) {
    const fallbackColumn = fallbackColumns.find((column) =>
      isSameColumn(column, extraColumn),
    );

    if (fallbackColumn && !hasColumn(exportColumns, fallbackColumn)) {
      exportColumns.push(fallbackColumn);
    }
  }
}

function createTemplateColumnIndexes(columns: string[]): Map<string, number> {
  const indexes = new Map<string, number>();

  columns.forEach((column, index) => {
    const normalizedColumn = normalizeHeader(column);
    const comparableColumn = normalizeComparableHeader(column);

    if (!indexes.has(normalizedColumn)) {
      indexes.set(normalizedColumn, index);
    }

    if (!indexes.has(comparableColumn)) {
      indexes.set(comparableColumn, index);
    }
  });

  return indexes;
}

function resolveTemplateColumnIndex(
  column: string,
  fallbackIndex: number,
  templateColumnIndexes: Map<string, number>,
): number {
  if (isSameColumn(column, OLD_DATE_COLUMN)) {
    return (
      templateColumnIndexes.get(normalizeHeader(INVOICE_DATE_COLUMN)) ??
      templateColumnIndexes.get(normalizeComparableHeader(INVOICE_DATE_COLUMN)) ??
      fallbackIndex
    );
  }

  return (
    templateColumnIndexes.get(normalizeHeader(column)) ??
    templateColumnIndexes.get(normalizeComparableHeader(column)) ??
    fallbackIndex
  );
}

function writeHeaderRow(
  worksheet: XLSX.WorkSheet,
  headerIndex: number,
  exportColumns: string[],
  templateColumnIndexes: Map<string, number>,
): void {
  const range = XLSX.utils.decode_range(worksheet["!ref"] ?? "A1:A1");
  const maxColumn = Math.max(range.e.c, exportColumns.length - 1);
  const originalHeaderCells = new Map<number, XLSX.CellObject>();

  for (let columnIndex = 0; columnIndex <= range.e.c; columnIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: headerIndex, c: columnIndex });
    const cell = worksheet[address];

    if (cell) {
      originalHeaderCells.set(columnIndex, { ...cell });
    }
  }

  exportColumns.forEach((column, columnIndex) => {
    const sourceColumnIndex = resolveTemplateColumnIndex(
      column,
      columnIndex,
      templateColumnIndexes,
    );
    const address = XLSX.utils.encode_cell({ r: headerIndex, c: columnIndex });
    const sourceCell = originalHeaderCells.get(sourceColumnIndex);

    worksheet[address] = {
      ...(sourceCell?.s ? { s: sourceCell.s } : {}),
      t: "s",
      v: column,
    };
  });

  for (
    let columnIndex = exportColumns.length;
    columnIndex <= maxColumn;
    columnIndex += 1
  ) {
    delete worksheet[XLSX.utils.encode_cell({ r: headerIndex, c: columnIndex })];
  }
}

function detectHeaderRowIndexFromColumns(
  worksheet: XLSX.WorkSheet,
  fallbackColumns: string[],
): number {
  const matrix = XLSX.utils.sheet_to_json<SheetMatrix[number]>(worksheet, {
    header: 1,
    defval: null,
    raw: true,
    blankrows: false,
  });
  const fallbackHeaders = new Set(fallbackColumns.map(normalizeHeader));
  const templateHeaderIndex = matrix.findIndex((row) =>
    row.some((cell) => fallbackHeaders.has(normalizeHeader(cell))),
  );

  return templateHeaderIndex >= 0 ? templateHeaderIndex : 0;
}

function extractColumnsFromWorksheet(
  worksheet: XLSX.WorkSheet,
  headerIndex: number,
): string[] {
  const range = XLSX.utils.decode_range(worksheet["!ref"] ?? "A1:A1");
  const columns: string[] = [];

  for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: headerIndex, c: columnIndex });
    const header = formatHeader(worksheet[address]?.v);

    if (header) {
      columns.push(header);
    }
  }

  return columns;
}

function getTemplateColumnFormat(
  worksheet: XLSX.WorkSheet,
  sampleRowIndex: number,
  columnIndex: number,
  columnName: string,
): TemplateColumnFormat {
  const sampleAddress = XLSX.utils.encode_cell({
    r: sampleRowIndex,
    c: columnIndex,
  });
  const sampleCell = worksheet[sampleAddress];

  return applyTemplateColumnOverride(columnName, {
    style: sampleCell?.s,
    numberFormat: sampleCell?.z,
    sampleType: sampleCell?.t,
  });
}

function clearDataRows(
  worksheet: XLSX.WorkSheet,
  headerIndex: number,
  columnCount: number,
): void {
  const range = XLSX.utils.decode_range(worksheet["!ref"] ?? "A1:A1");

  for (let rowIndex = headerIndex + 1; rowIndex <= range.e.r; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      delete worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
    }
  }
}

function writeRowsToTemplate(
  worksheet: XLSX.WorkSheet,
  rows: ExcelRow[],
  templateColumns: string[],
  columnFormats: TemplateColumnFormat[],
  headerIndex: number,
): void {
  const rowColumnLookup = createRowColumnLookup(rows);

  rows.forEach((row, rowOffset) => {
    const rowIndex = headerIndex + 1 + rowOffset;

    templateColumns.forEach((column, columnIndex) => {
      const value = getValueForTemplateColumn(row, column, rowColumnLookup);
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const cell = createCell(value, columnFormats[columnIndex]);

      if (cell) {
        worksheet[address] = cell;
      }
    });
  });
}

function updateWorksheetRange(
  worksheet: XLSX.WorkSheet,
  headerIndex: number,
  columnCount: number,
  rowCount: number,
): void {
  const endRow = Math.max(headerIndex + rowCount, headerIndex);
  const endColumn = Math.max(columnCount - 1, 0);
  const ref = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: endRow, c: endColumn },
  });

  worksheet["!ref"] = ref;

  if (worksheet["!autofilter"]) {
    worksheet["!autofilter"].ref = ref;
  }
}

function createRowColumnLookup(rows: ExcelRow[]): Map<string, string> {
  const lookup = new Map<string, string>();

  rows.forEach((row) => {
    Object.keys(row).forEach((column) => {
      const normalizedColumn = normalizeHeader(column);
      const comparableColumn = normalizeComparableHeader(column);

      if (!lookup.has(normalizedColumn)) {
        lookup.set(normalizedColumn, column);
      }

      if (!lookup.has(comparableColumn)) {
        lookup.set(comparableColumn, column);
      }
    });
  });

  return lookup;
}

function getValueForTemplateColumn(
  row: ExcelRow,
  templateColumn: string,
  rowColumnLookup: Map<string, string>,
): ExcelCell {
  const normalizedTemplateColumn = normalizeHeader(templateColumn);
  const comparableTemplateColumn = normalizeComparableHeader(templateColumn);

  const value = getMappedValueForTemplateColumn(
    row,
    templateColumn,
    normalizedTemplateColumn,
    comparableTemplateColumn,
    rowColumnLookup,
  );

  if (isSameColumn(normalizedTemplateColumn, B2F_EXCHANGE_RATE_COLUMN)) {
    return getB2fExchangeRateValue(row, value, rowColumnLookup);
  }

  if (isSameColumn(normalizedTemplateColumn, IMPORT_CURRENCY_COLUMN)) {
    return getImportCurrencyValue(row, value, rowColumnLookup);
  }

  if (isSameColumn(normalizedTemplateColumn, IMPORT_EXCHANGE_RATE_COLUMN)) {
    return getImportExchangeRateValue(row, value, rowColumnLookup);
  }

  if (
    isSameColumn(normalizedTemplateColumn, B2F_EXCHANGE_RATE_DATE_COLUMN)
  ) {
    return getB2fExchangeRateDateValue(row, value, rowColumnLookup);
  }

  if (
    isColumnCandidate(normalizedTemplateColumn, INVOICE_NUMBER_COLUMN_CANDIDATES) &&
    isBlankValue(value)
  ) {
    return getMappedValueForKnownColumns(
      row,
      SOURCE_DOCUMENT_NUMBER_COLUMN_CANDIDATES,
      rowColumnLookup,
    );
  }

  if (
    isColumnCandidate(normalizedTemplateColumn, PAYMENT_METHOD_COLUMN_CANDIDATES) &&
    isBlankValue(value)
  ) {
    return PAYMENT_METHOD_DEFAULT;
  }

  if (
    isColumnCandidate(normalizedTemplateColumn, OPERATOR_CODE_COLUMN_CANDIDATES) &&
    isBlankValue(value)
  ) {
    return OPERATOR_CODE_DEFAULT;
  }

  if (
    isColumnCandidate(normalizedTemplateColumn, OPERATOR_NAME_COLUMN_CANDIDATES) &&
    isBlankValue(value)
  ) {
    return OPERATOR_NAME_DEFAULT;
  }

  if (
    isSameColumn(normalizedTemplateColumn, ORIGINAL_PRICE_COLUMN) &&
    isBlankValue(value)
  ) {
    return 0;
  }

  return value;
}

function getB2fExchangeRateValue(
  row: ExcelRow,
  value: ExcelCell,
  rowColumnLookup: Map<string, string>,
): ExcelCell {
  if (isForeignCurrencyRow(row, rowColumnLookup)) {
    return isBlankValue(value) ? null : value;
  }

  return 1;
}

function getB2fExchangeRateDateValue(
  row: ExcelRow,
  value: ExcelCell,
  rowColumnLookup: Map<string, string>,
): ExcelCell {
  if (isForeignCurrencyRow(row, rowColumnLookup)) {
    return isBlankValue(value) ? null : value;
  }

  if (isBlankValue(value)) {
    return getMappedValueForKnownColumn(row, INVOICE_DATE_COLUMN, rowColumnLookup);
  }

  return value;
}

function isForeignCurrencyRow(
  row: ExcelRow,
  rowColumnLookup: Map<string, string>,
): boolean {
  const currencyCode = getInvoiceCurrencyCode(row, rowColumnLookup);

  return currencyCode !== "" && currencyCode !== "CDF";
}

function getImportCurrencyValue(
  row: ExcelRow,
  value: ExcelCell,
  rowColumnLookup: Map<string, string>,
): ExcelCell {
  const currencyCode = getInvoiceCurrencyCode(row, rowColumnLookup);

  if (currencyCode !== "") {
    return currencyCode;
  }

  return isBlankValue(value) ? null : value;
}

function getImportExchangeRateValue(
  row: ExcelRow,
  value: ExcelCell,
  rowColumnLookup: Map<string, string>,
): ExcelCell {
  const currencyCode = getInvoiceCurrencyCode(row, rowColumnLookup);

  if (currencyCode === "CDF") {
    return 1;
  }

  if (currencyCode === "") {
    return isBlankValue(value) ? null : value;
  }

  const b2fExchangeRate = getMappedValueForKnownColumn(
    row,
    B2F_EXCHANGE_RATE_COLUMN,
    rowColumnLookup,
  );

  return isBlankValue(b2fExchangeRate) ? null : b2fExchangeRate;
}

function getInvoiceCurrencyCode(
  row: ExcelRow,
  rowColumnLookup: Map<string, string>,
): string {
  const currency = getMappedValueForKnownColumns(
    row,
    B2F_CURRENCY_COLUMN_CANDIDATES,
    rowColumnLookup,
  );
  return formatCurrencyCode(currency);
}

function formatCurrencyCode(value: ExcelCell): string {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim().toUpperCase();
}

function getMappedValueForTemplateColumn(
  row: ExcelRow,
  templateColumn: string,
  normalizedTemplateColumn: string,
  comparableTemplateColumn: string,
  rowColumnLookup: Map<string, string>,
): ExcelCell {
  if (Object.prototype.hasOwnProperty.call(row, templateColumn)) {
    return row[templateColumn];
  }

  if (isCreditNoteOriginErpColumn(templateColumn)) {
    return getExactMappedValueForTemplateColumn(
      row,
      normalizedTemplateColumn,
      comparableTemplateColumn,
      rowColumnLookup,
    );
  }

  const rowColumn =
    rowColumnLookup.get(normalizedTemplateColumn) ??
    rowColumnLookup.get(comparableTemplateColumn) ??
    findLooseRowColumn(normalizedTemplateColumn, rowColumnLookup) ??
    findLooseRowColumn(comparableTemplateColumn, rowColumnLookup);

  return rowColumn ? row[rowColumn] : null;
}

function getExactMappedValueForTemplateColumn(
  row: ExcelRow,
  normalizedTemplateColumn: string,
  comparableTemplateColumn: string,
  rowColumnLookup: Map<string, string>,
): ExcelCell {
  const rowColumn =
    rowColumnLookup.get(normalizedTemplateColumn) ??
    rowColumnLookup.get(comparableTemplateColumn);

  return rowColumn ? row[rowColumn] : null;
}

function isCreditNoteOriginErpColumn(columnName: string): boolean {
  return (
    isSameColumn(columnName, CREDIT_NOTE_ORIGIN_ERP_INVOICE_COLUMN) ||
    isSameColumn(columnName, CREDIT_NOTE_ORIGIN_ERP_DATE_COLUMN)
  );
}

function getMappedValueForKnownColumn(
  row: ExcelRow,
  columnName: string,
  rowColumnLookup: Map<string, string>,
): ExcelCell {
  const normalizedColumn = normalizeHeader(columnName);
  const comparableColumn = normalizeComparableHeader(columnName);
  const rowColumn =
    rowColumnLookup.get(normalizedColumn) ??
    rowColumnLookup.get(comparableColumn) ??
    findLooseRowColumn(normalizedColumn, rowColumnLookup) ??
    findLooseRowColumn(comparableColumn, rowColumnLookup);

  return rowColumn ? row[rowColumn] : null;
}

function getMappedValueForKnownColumns(
  row: ExcelRow,
  columnNames: string[],
  rowColumnLookup: Map<string, string>,
): ExcelCell {
  for (const columnName of columnNames) {
    const value = getMappedValueForKnownColumn(row, columnName, rowColumnLookup);

    if (!isBlankValue(value)) {
      return value;
    }
  }

  return null;
}

function findLooseRowColumn(
  templateColumn: string,
  rowColumnLookup: Map<string, string>,
): string | undefined {
  for (const [rowColumn, originalColumn] of rowColumnLookup) {
    if (
      templateColumn.startsWith(`${rowColumn} `) ||
      rowColumn.startsWith(`${templateColumn} `)
    ) {
      return originalColumn;
    }
  }

  return undefined;
}

function createCell(
  value: ExcelCell,
  columnFormat: TemplateColumnFormat,
): XLSX.CellObject | null {
  if (value === null || value === undefined || value === "") {
    return createBlankCell(columnFormat);
  }

  const baseCell: Partial<XLSX.CellObject> = {};

  if (columnFormat.style) {
    baseCell.s = columnFormat.style;
  }

  if (columnFormat.numberFormat) {
    baseCell.z = columnFormat.numberFormat;
  }

  if (columnFormat.dateAsSerial) {
    const dateValue = parseDateValue(value);

    if (dateValue) {
      return {
        ...baseCell,
        t: "n",
        v: dateToExcelSerial(dateValue),
      };
    }
  }

  if (columnFormat.numberFormat === "@") {
    return {
      ...baseCell,
      t: "s",
      v: String(value),
    };
  }

  if (value instanceof Date) {
    return {
      ...baseCell,
      t: "d",
      v: toDateOnly(value),
    };
  }

  if (columnFormat.sampleType === "d") {
    const dateValue = parseDateValue(value);

    if (dateValue) {
      return {
        ...baseCell,
        t: "d",
        v: dateValue,
      };
    }
  }

  if (typeof value === "number") {
    return {
      ...baseCell,
      t: "n",
      v: normalizeNumberForColumn(value, columnFormat),
    };
  }

  if (columnFormat.sampleType === "n") {
    const numberValue = parseNumberValue(value);

    if (numberValue !== null) {
      return {
        ...baseCell,
        t: "n",
        v: normalizeNumberForColumn(numberValue, columnFormat),
      };
    }
  }

  if (typeof value === "boolean") {
    return {
      ...baseCell,
      t: "b",
      v: value,
    };
  }

  return {
    ...baseCell,
    t: "s",
    v: String(value),
  };
}

function createBlankCell(
  columnFormat: TemplateColumnFormat,
): XLSX.CellObject | null {
  if (!columnFormat.style && !columnFormat.numberFormat) {
    return null;
  }

  return {
    ...(columnFormat.style ? { s: columnFormat.style } : {}),
    ...(columnFormat.numberFormat ? { z: columnFormat.numberFormat } : {}),
    t: "z",
  };
}

function parseNumberValue(value: ExcelCell): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value !== "string") {
    return null;
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
  return Number.isFinite(parsedValue) ? parsedValue : null;
}

function parseDateValue(value: ExcelCell): Date | null {
  if (value instanceof Date) {
    return toDateOnly(value);
  }

  if (typeof value === "number") {
    const parsedDate = XLSX.SSF.parse_date_code(value);

    if (!parsedDate) {
      return null;
    }

    return new Date(
      parsedDate.y,
      parsedDate.m - 1,
      parsedDate.d,
    );
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

function applyTemplateColumnOverride(
  columnName: string,
  columnFormat: TemplateColumnFormat,
): TemplateColumnFormat {
  if (isSameColumn(columnName, OLD_DATE_COLUMN)) {
    return {
      ...columnFormat,
      numberFormat: "@",
      sampleType: "s",
      dateAsSerial: false,
    };
  }

  if (isSameColumn(columnName, INVOICE_DATE_COLUMN)) {
    return {
      ...columnFormat,
      numberFormat: "dd/mm/yyyy",
      sampleType: "n",
      dateAsSerial: true,
    };
  }

  if (
    isSameColumn(columnName, B2F_EXCHANGE_RATE_DATE_COLUMN) ||
    isSameColumn(columnName, CREDIT_NOTE_ORIGIN_ERP_DATE_COLUMN)
  ) {
    return {
      ...columnFormat,
      numberFormat: "dd/mm/yyyy",
      sampleType: "n",
      dateAsSerial: true,
    };
  }

  if (
    isSameColumn(columnName, UNIT_PRICE_COLUMN) ||
    isSameColumn(columnName, ORIGINAL_PRICE_COLUMN) ||
    isSameColumn(columnName, B2F_EXCHANGE_RATE_COLUMN)
  ) {
    return {
      ...columnFormat,
      numberFormat: "0.00",
      sampleType: "n",
      decimals: 2,
    };
  }

  return columnFormat;
}

function normalizeNumberForColumn(
  value: number,
  columnFormat: TemplateColumnFormat,
): number {
  if (columnFormat.decimals === undefined) {
    return value;
  }

  const factor = 10 ** columnFormat.decimals;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function isSameColumn(left: unknown, right: unknown): boolean {
  return normalizeComparableHeader(left) === normalizeComparableHeader(right);
}

function hasColumn(columns: string[], columnName: string): boolean {
  return columns.some((column) => isSameColumn(column, columnName));
}

function isColumnCandidate(columnName: string, candidates: string[]): boolean {
  return candidates.some((candidate) => isSameColumn(columnName, candidate));
}

function isExciseTaxExportColumn(columnName: string): boolean {
  const normalizedColumn = normalizeHeader(columnName);

  return (
    isSameColumn(columnName, EXCISE_TAX_COLUMN) ||
    isSameColumn(columnName, "Excise Tax_Local") ||
    (normalizedColumn.includes("congo drc telecommunication excise tax") &&
      normalizedColumn.includes("final")) ||
    (normalizedColumn.includes("excise tax") &&
      normalizedColumn.includes("local"))
  );
}

function toDateOnly(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

function dateToExcelSerial(value: Date): number {
  return (Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()) -
    Date.UTC(1899, 11, 30)) /
    86400000;
}

function isBlankValue(value: ExcelCell): boolean {
  return value === null || value === undefined || value === "";
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

  return new Date(year, month - 1, day);
}

function detectHeaderRowIndex(matrix: SheetMatrix): number {
  const rowsToScan = matrix.slice(0, HEADER_SCAN_LIMIT);
  const taxHeader = normalizeHeader(EXCISE_TAX_COLUMN);
  const exactTaxHeaderIndex = rowsToScan.findIndex((row) =>
    row.some((cell) => normalizeHeader(cell) === taxHeader),
  );

  if (exactTaxHeaderIndex >= 0) {
    return exactTaxHeaderIndex;
  }

  const likelyTaxHeaderIndex = rowsToScan.findIndex((row) =>
    row.some((cell) => {
      const header = normalizeHeader(cell);
      return (
        (header.includes("congo drc telecommunication excise tax") &&
          header.includes("final")) ||
        (header.includes("excise tax") && header.includes("local"))
      );
    }),
  );

  if (likelyTaxHeaderIndex >= 0) {
    return likelyTaxHeaderIndex;
  }

  const firstDenseRowIndex = rowsToScan.findIndex(
    (row) => row.filter((cell) => normalizeHeader(cell) !== "").length >= 3,
  );

  return firstDenseRowIndex >= 0 ? firstDenseRowIndex : 0;
}

function extractColumns(matrix: SheetMatrix, headerIndex: number): string[] {
  const headerRow = matrix[headerIndex] ?? [];
  const columns = headerRow.map((cell, index) => {
    const header = formatHeader(cell);
    return header || `Colonne ${index + 1}`;
  });

  return makeUniqueColumns(columns);
}

function extractRows(
  matrix: SheetMatrix,
  headerIndex: number,
  columns: string[],
): ExcelRow[] {
  return matrix
    .slice(headerIndex + 1)
    .filter((row) => row.some((cell) => cell !== null && cell !== undefined && cell !== ""))
    .map((row) => {
      return columns.reduce<ExcelRow>((record, column, index) => {
        record[column] = row[index] ?? null;
        return record;
      }, {});
    });
}

function makeUniqueColumns(columns: string[]): string[] {
  const seen = new Map<string, number>();

  return columns.map((column) => {
    const count = seen.get(column) ?? 0;
    seen.set(column, count + 1);
    return count === 0 ? column : `${column}_${count + 1}`;
  });
}

function formatHeader(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeHeader(value: unknown): string {
  return formatHeader(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\u200b-\u200d\ufeff]/g, "")
    .replace(/[°º]/g, "o")
    .replace(/[\[\]()/|]+/g, " ")
    .replace(/[_\s-]+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeComparableHeader(value: unknown): string {
  return normalizeHeader(value)
    .replace(/\b(de|du|des|d)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
