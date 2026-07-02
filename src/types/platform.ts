export type PlatformDocumentKind = "sales" | "credit_note" | "mixed";

export type PlatformDocumentType = "FV" | "FA";

export type PlatformBatchStatus = "draft" | "processing" | "success" | "failed";

export type PlatformCurrencyMode = "CDF" | "USD" | "EUR" | "MULTI";

export type PlatformReferenceType = "RAN" | "COR";

export interface PlatformOrganization {
  id: string;
  name: string;
  slug: string;
}

export interface PlatformConversionTemplate {
  id: string;
  organizationId: string;
  name: string;
  sourceCompany: string;
  documentType: PlatformDocumentKind;
  sourceColumns: string[];
  dexyColumns: string[];
  rules: Record<string, unknown>;
  isActive: boolean;
}

export interface PlatformConversionBatch {
  id: string;
  organizationId: string;
  templateId?: string;
  sourceCompany: string;
  documentType: PlatformDocumentKind;
  status: PlatformBatchStatus;
  sourceFileName?: string;
  outputFileName?: string;
  currencyMode?: PlatformCurrencyMode;
  invoiceCount: number;
  sourceRowCount: number;
  taxRowCount: number;
  totalTax: number;
  errorMessage?: string;
  createdAt: string;
  completedAt?: string;
}

export interface PlatformCertifiedInvoice {
  id?: string;
  organizationId: string;
  batchId?: string;
  sourceCompany: string;
  invoiceNumber: string;
  normalizedInvoiceNumber?: string;
  dgiCode?: string;
  documentType: PlatformDocumentType;
  invoiceDate?: string;
  originalInvoiceDate?: string;
  currency: string;
  exchangeRate?: number;
  exchangeRateDate?: string;
  totalHT?: number;
  totalTTC?: number;
  totalTax?: number;
}

export interface PlatformCreditNoteLink {
  id?: string;
  organizationId: string;
  creditNoteInvoiceId: string;
  originalInvoiceId: string;
  referenceType: PlatformReferenceType;
  referenceDescription: "ANNULATION" | "CORRECTION";
}

export interface PlatformConversionEvent {
  organizationId: string;
  batchId?: string;
  level: "info" | "warning" | "error";
  message: string;
  details?: Record<string, unknown>;
}
