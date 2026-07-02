import type {
  PlatformCertifiedInvoice,
  PlatformConversionBatch,
  PlatformConversionEvent,
  PlatformCreditNoteLink,
} from "../types/platform";
import { isSupabaseConfigured, supabaseClient } from "./supabaseClient";

export interface PlatformRepository {
  createConversionBatch(
    batch: Omit<PlatformConversionBatch, "id" | "createdAt">,
  ): Promise<PlatformConversionBatch | null>;
  saveCertifiedInvoices(invoices: PlatformCertifiedInvoice[]): Promise<void>;
  saveCreditNoteLinks(links: PlatformCreditNoteLink[]): Promise<void>;
  logConversionEvent(event: PlatformConversionEvent): Promise<void>;
}

class NoopPlatformRepository implements PlatformRepository {
  async createConversionBatch(): Promise<PlatformConversionBatch | null> {
    return null;
  }

  async saveCertifiedInvoices(): Promise<void> {
    return;
  }

  async saveCreditNoteLinks(): Promise<void> {
    return;
  }

  async logConversionEvent(): Promise<void> {
    return;
  }
}

class SupabasePlatformRepository implements PlatformRepository {
  async createConversionBatch(
    batch: Omit<PlatformConversionBatch, "id" | "createdAt">,
  ): Promise<PlatformConversionBatch | null> {
    if (!supabaseClient) {
      return null;
    }

    const { data, error } = await supabaseClient
      .from("conversion_batches")
      .insert(toBatchRecord(batch))
      .select()
      .single();

    if (error) {
      throw new Error(error.message);
    }

    return data ? fromBatchRecord(data) : null;
  }

  async saveCertifiedInvoices(
    invoices: PlatformCertifiedInvoice[],
  ): Promise<void> {
    if (!supabaseClient || invoices.length === 0) {
      return;
    }

    const { error } = await supabaseClient
      .from("certified_invoices")
      .upsert(invoices.map(toCertifiedInvoiceRecord), {
        onConflict:
          "organization_id,source_company,invoice_number,document_type",
      });

    if (error) {
      throw new Error(error.message);
    }
  }

  async saveCreditNoteLinks(links: PlatformCreditNoteLink[]): Promise<void> {
    if (!supabaseClient || links.length === 0) {
      return;
    }

    const { error } = await supabaseClient
      .from("credit_note_links")
      .upsert(links.map(toCreditNoteLinkRecord), {
        onConflict: "credit_note_invoice_id",
      });

    if (error) {
      throw new Error(error.message);
    }
  }

  async logConversionEvent(event: PlatformConversionEvent): Promise<void> {
    if (!supabaseClient) {
      return;
    }

    const { error } = await supabaseClient
      .from("conversion_events")
      .insert(toConversionEventRecord(event));

    if (error) {
      throw new Error(error.message);
    }
  }
}

function toBatchRecord(
  batch: Omit<PlatformConversionBatch, "id" | "createdAt">,
): Record<string, unknown> {
  return {
    organization_id: batch.organizationId,
    template_id: batch.templateId ?? null,
    source_company: batch.sourceCompany,
    document_type: batch.documentType,
    status: batch.status,
    source_file_name: batch.sourceFileName ?? null,
    output_file_name: batch.outputFileName ?? null,
    currency_mode: batch.currencyMode ?? null,
    invoice_count: batch.invoiceCount,
    source_row_count: batch.sourceRowCount,
    tax_row_count: batch.taxRowCount,
    total_tax: batch.totalTax,
    error_message: batch.errorMessage ?? null,
    completed_at: batch.completedAt ?? null,
  };
}

function fromBatchRecord(record: Record<string, unknown>): PlatformConversionBatch {
  return {
    id: String(record.id),
    organizationId: String(record.organization_id),
    templateId: toOptionalString(record.template_id),
    sourceCompany: String(record.source_company),
    documentType: record.document_type as PlatformConversionBatch["documentType"],
    status: record.status as PlatformConversionBatch["status"],
    sourceFileName: toOptionalString(record.source_file_name),
    outputFileName: toOptionalString(record.output_file_name),
    currencyMode: record.currency_mode as PlatformConversionBatch["currencyMode"],
    invoiceCount: Number(record.invoice_count ?? 0),
    sourceRowCount: Number(record.source_row_count ?? 0),
    taxRowCount: Number(record.tax_row_count ?? 0),
    totalTax: Number(record.total_tax ?? 0),
    errorMessage: toOptionalString(record.error_message),
    createdAt: String(record.created_at),
    completedAt: toOptionalString(record.completed_at),
  };
}

function toCertifiedInvoiceRecord(
  invoice: PlatformCertifiedInvoice,
): Record<string, unknown> {
  return {
    id: invoice.id,
    organization_id: invoice.organizationId,
    batch_id: invoice.batchId ?? null,
    source_company: invoice.sourceCompany,
    invoice_number: invoice.invoiceNumber,
    normalized_invoice_number: invoice.normalizedInvoiceNumber ?? null,
    dgi_code: invoice.dgiCode ?? null,
    document_type: invoice.documentType,
    invoice_date: invoice.invoiceDate ?? null,
    original_invoice_date: invoice.originalInvoiceDate ?? null,
    currency: invoice.currency,
    exchange_rate: invoice.exchangeRate ?? null,
    exchange_rate_date: invoice.exchangeRateDate ?? null,
    total_ht: invoice.totalHT ?? null,
    total_ttc: invoice.totalTTC ?? null,
    total_tax: invoice.totalTax ?? null,
  };
}

function toCreditNoteLinkRecord(
  link: PlatformCreditNoteLink,
): Record<string, unknown> {
  return {
    id: link.id,
    organization_id: link.organizationId,
    credit_note_invoice_id: link.creditNoteInvoiceId,
    original_invoice_id: link.originalInvoiceId,
    reference_type: link.referenceType,
    reference_description: link.referenceDescription,
  };
}

function toConversionEventRecord(
  event: PlatformConversionEvent,
): Record<string, unknown> {
  return {
    organization_id: event.organizationId,
    batch_id: event.batchId ?? null,
    level: event.level,
    message: event.message,
    details: event.details ?? {},
  };
}

function toOptionalString(value: unknown): string | undefined {
  return value === null || value === undefined ? undefined : String(value);
}

export const platformRepository: PlatformRepository =
  isSupabaseConfigured
    ? new SupabasePlatformRepository()
    : new NoopPlatformRepository();
