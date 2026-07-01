export type ExcelCell = string | number | boolean | Date | null | undefined;

export type ExcelRow = Record<string, ExcelCell>;

export interface ImportedWorkbook {
  fileName: string;
  sheetName: string;
  uploadedAt: Date;
  rows: ExcelRow[];
  columns: string[];
}

export interface ReferenceWorkbook {
  rows: ExcelRow[];
  columns: string[];
}

export interface ConversionSummary {
  invoicesProcessed: number;
  originalRows: number;
  taxRowsAdded: number;
  totalTax: number;
}

export interface ConversionResult {
  rows: ExcelRow[];
  columns: string[];
  summary: ConversionSummary;
  invoiceColumn: string;
  taxColumn: string;
  normalizationDate: Date;
  warnings: string[];
}
