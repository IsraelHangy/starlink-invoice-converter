import type {
  PlatformCertifiedInvoice,
  PlatformConversionBatch,
  PlatformConversionEvent,
  PlatformCreditNoteLink,
} from "../types/platform";

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

export const platformRepository: PlatformRepository =
  new NoopPlatformRepository();
