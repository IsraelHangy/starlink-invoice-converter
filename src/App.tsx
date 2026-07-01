import { useEffect, useMemo, useRef, useState } from "react";
import type { ChangeEvent } from "react";
import {
  AlertCircle,
  ArrowRightLeft,
  CheckCircle2,
  DollarSign,
  Download,
  FileSearch,
  Loader2,
  Play,
  Upload,
  X,
} from "lucide-react";
import UploadCard from "./components/UploadCard";
import PreviewTable from "./components/PreviewTable";
import SummaryCard from "./components/SummaryCard";
import type { ConversionResult, ImportedWorkbook } from "./types";
import { exportExcelFile } from "./utils/excel";
import {
  convertStarlinkToDexyInWorker,
  type ExchangeRateUpdateRequest,
  readExcelFileInWorker,
} from "./utils/processingWorkerClient";

type PreviewMode = "before" | "after";

const EXPORT_CURRENCY_COLUMN_CANDIDATES = [
  "B2F Devise [Nom]",
  "B2F Devise Nom",
  "Devise Import",
  "Currency",
  "Transaction Currency",
  "transaction_currency",
];

const BCC_USD_RATE_CACHE_KEY = "dexy_converter_last_bcc_usd_rate";

export default function App() {
  const [workbook, setWorkbook] = useState<ImportedWorkbook | null>(null);
  const [conversion, setConversion] = useState<ConversionResult | null>(null);
  const [previewMode, setPreviewMode] = useState<PreviewMode>("before");
  const [error, setError] = useState<string | null>(null);
  const [isReadingWorkbook, setIsReadingWorkbook] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [readingFileSize, setReadingFileSize] = useState<number | null>(null);
  const [showProcessingStatus, setShowProcessingStatus] = useState(false);
  const [referenceWorkbook, setReferenceWorkbook] =
    useState<ImportedWorkbook | null>(null);
  const [creditNoteReferencePromptOpen, setCreditNoteReferencePromptOpen] =
    useState(false);
  const [exchangeRatePromptOpen, setExchangeRatePromptOpen] = useState(false);
  const [dgiRateInput, setDgiRateInput] = useState("");
  const [dgiRateError, setDgiRateError] = useState<string | null>(null);
  const [dgiRateMeta, setDgiRateMeta] = useState<string | null>(null);
  const [dgiRatePlaceholder, setDgiRatePlaceholder] = useState(
    "En attente du taux DGI du jour",
  );
  const [isFetchingDgiRate, setIsFetchingDgiRate] = useState(false);

  const canConvert =
    Boolean(workbook?.rows.length) &&
    !isReadingWorkbook &&
    !isConverting &&
    !isExporting;
  const canExport =
    Boolean(conversion?.rows.length) &&
    !isReadingWorkbook &&
    !isConverting &&
    !isExporting;
  const isProcessing = isReadingWorkbook || isConverting || isExporting;
  const processingDelay = getProcessingDelay({
    fileSize: readingFileSize,
    isConverting,
    isExporting,
    isReadingWorkbook,
    rowCount: workbook?.rows.length ?? 0,
  });
  const processingMessage = getProcessingMessage({
    fileSize: readingFileSize,
    isReadingWorkbook,
    isConverting,
    isExporting,
    rowCount: workbook?.rows.length ?? 0,
  });

  useEffect(() => {
    if (!isProcessing) {
      setShowProcessingStatus(false);
      return;
    }

    setShowProcessingStatus(false);
    const timeoutId = window.setTimeout(() => {
      setShowProcessingStatus(true);
    }, processingDelay);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [isProcessing, processingDelay]);

  const preview = useMemo(() => {
    if (previewMode === "after" && conversion) {
      return {
        title: "Aperçu après conversion",
        rows: conversion.rows,
        columns: conversion.columns,
        emptyMessage: "Convertissez un fichier pour voir le résultat DEXY.",
      };
    }

    return {
      title: "Aperçu avant conversion",
      rows: workbook?.rows ?? [],
      columns: workbook?.columns ?? [],
      emptyMessage: "Importez un fichier Excel Starlink pour afficher l'aperçu.",
    };
  }, [conversion, previewMode, workbook]);

  const handleWorkbookLoaded = (loadedWorkbook: ImportedWorkbook) => {
    setWorkbook(loadedWorkbook);
    setConversion(null);
    setReferenceWorkbook(null);
    setCreditNoteReferencePromptOpen(false);
    setPreviewMode("before");
    setError(null);
  };

  const handleConvert = async () => {
    if (!workbook) {
      setError("Importez d'abord un fichier Excel Starlink.");
      return;
    }

    const documentKind = getWorkbookDocumentKind(workbook.rows);

    if (documentKind === "mixed") {
      setError(
        "Ce fichier contient a la fois des ventes (FV) et des avoirs (FA). Importez les ventes et les avoirs dans deux fichiers separes.",
      );
      return;
    }

    if (documentKind === "credit" && !referenceWorkbook) {
      setCreditNoteReferencePromptOpen(true);
      setError(null);
      return;
    }

    if (hasUsdRows(workbook.rows)) {
      setDgiRateInput("");
      setDgiRateError(null);
      setDgiRateMeta(null);
      setDgiRatePlaceholder(getCachedBccRatePlaceholder());
      setExchangeRatePromptOpen(true);
      setError(null);
      void loadBccUsdRate();
      return;
    }

    await runConversion();
  };

  const handleContinueAfterReference = () => {
    if (!referenceWorkbook) {
      setError("Importez le template reference DEXY avant de convertir l'avoir.");
      return;
    }

    setCreditNoteReferencePromptOpen(false);

    if (hasUsdRows(workbook?.rows ?? [])) {
      setDgiRateInput("");
      setDgiRateError(null);
      setDgiRateMeta(null);
      setDgiRatePlaceholder(getCachedBccRatePlaceholder());
      setExchangeRatePromptOpen(true);
      setError(null);
      void loadBccUsdRate();
      return;
    }

    void runConversion();
  };

  const runConversion = async (
    exchangeRateUpdate?: ExchangeRateUpdateRequest,
  ) => {
    if (!workbook) {
      setError("Importez d'abord un fichier Excel Starlink.");
      return;
    }

    setIsConverting(true);
    setError(null);

    try {
      const convertedWorkbook = await convertStarlinkToDexyInWorker(
        workbook.rows,
        workbook.columns,
        workbook.uploadedAt,
        exchangeRateUpdate,
        getWorkbookDocumentKind(workbook.rows) === "credit" && referenceWorkbook
          ? {
              rows: referenceWorkbook.rows,
              columns: referenceWorkbook.columns,
            }
          : undefined,
      );
      setConversion(convertedWorkbook);
      setPreviewMode("after");
      setError(null);
    } catch (caughtError) {
      const message =
        caughtError instanceof Error
          ? caughtError.message
          : "Conversion impossible.";
      setConversion(null);
      setPreviewMode("before");
      setError(message);

      if (shouldResetCreditNoteReference(message, workbook.rows)) {
        setReferenceWorkbook(null);
        setCreditNoteReferencePromptOpen(false);
      }
    } finally {
      setIsConverting(false);
    }
  };

  const handleKeepTemplateRates = () => {
    setExchangeRatePromptOpen(false);
    setDgiRateError(null);
    void runConversion();
  };

  const loadBccUsdRate = async () => {
    setIsFetchingDgiRate(true);

    try {
      const bccRate = await fetchBccUsdRate();
      setDgiRateInput(formatRateInput(bccRate.rate));
      setDgiRatePlaceholder("Taux DGI du jour");
      const publishedAt = formatPublishedAtInFrench(bccRate.publishedAt);
      setDgiRateMeta(
        publishedAt
          ? `Taux DGI du jour : ${formatRateInput(bccRate.rate)} - ${publishedAt}`
          : `Taux DGI du jour : ${formatRateInput(bccRate.rate)}`,
      );
      cacheBccRate(bccRate);
      setDgiRateError(null);
    } catch {
      setDgiRateMeta(
        "Taux DGI indisponible automatiquement. Vous pouvez le renseigner manuellement.",
      );
    } finally {
      setIsFetchingDgiRate(false);
    }
  };

  const handleApplyDgiRate = () => {
    const rate = parseRateInput(dgiRateInput);

    if (rate === null || rate <= 0) {
      setDgiRateError("Renseignez un taux DGI valide avant de continuer.");
      return;
    }

    setExchangeRatePromptOpen(false);
    setDgiRateError(null);
    void runConversion({
      currency: "USD",
      rate,
      rateDate: new Date(),
    });
  };

  const handleExport = async () => {
    if (!conversion) {
      setError("Convertissez le fichier avant l'export.");
      return;
    }

    setIsExporting(true);

    try {
      await exportExcelFile(
        conversion.rows,
        conversion.columns,
        buildDexyExportFileName(conversion.rows, new Date()),
      );
      setError(null);
    } catch (caughtError) {
      setError(
        caughtError instanceof Error
          ? caughtError.message
          : "Export DEXY impossible.",
      );
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <main className="min-h-screen bg-[#f4f7fb]">
      <div className="border-b border-slate-200 bg-white">
        <div className="mx-auto w-full max-w-5xl px-4 py-8 text-center sm:px-6 lg:px-8">
          <header className="flex flex-col items-center">
            <div className="flex w-full items-center justify-center gap-3 sm:gap-8">
              <img
                className="h-14 w-auto max-w-[48%] object-contain sm:h-24 sm:max-w-[340px]"
                src="/assets/dexy-converter-logo.png"
                alt="DEXY CD Converter"
              />
              <img
                className="h-16 w-auto max-w-[45%] object-contain sm:h-24 sm:max-w-[330px]"
                src="/assets/mti-logo.jpg"
                alt="MTI Integrated Business Solution"
              />
            </div>

            <h1 className="mt-6 text-3xl font-bold text-ink sm:text-4xl">
              Starlink to DEXY Converter
            </h1>
            <p className="mt-3 max-w-2xl text-sm leading-6 text-slate-600 sm:text-base">
              Convertissez un fichier Excel Starlink en fichier DEXY CD SFE
              conforme, avec une ligne TAX consolidée par facture.
            </p>

            <div className="mt-6 flex w-full flex-col justify-center gap-3 sm:w-auto sm:flex-row">
            <button
              className="inline-flex min-h-11 items-center justify-center gap-2 whitespace-nowrap rounded-md bg-[#5f55d6] px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-[#5147c4] disabled:cursor-not-allowed disabled:bg-slate-400 sm:min-w-36"
              type="button"
              onClick={() => void handleConvert()}
              disabled={!canConvert}
              title="Convertir le fichier Starlink importe"
            >
              {isConverting ? (
                <Loader2 className="animate-spin" aria-hidden="true" size={17} />
              ) : (
                <Play aria-hidden="true" size={17} />
              )}
              {isConverting ? "Conversion..." : "Convertir"}
            </button>
            <button
              className="inline-flex min-h-11 items-center justify-center gap-2 whitespace-nowrap rounded-md bg-[#b0003a] px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-[#960033] disabled:cursor-not-allowed disabled:bg-slate-400 sm:min-w-40"
              type="button"
              onClick={() => void handleExport()}
              disabled={!canExport || isExporting}
              title="Exporter le fichier Excel DEXY"
            >
              <Download aria-hidden="true" size={17} />
              {isExporting ? "Export..." : "Export .xlsx"}
            </button>
            </div>
          </header>
        </div>
      </div>

      <div className="mx-auto flex w-full max-w-7xl flex-col gap-6 px-4 py-6 sm:px-6 lg:px-8">

        {error ? (
          <ValidationErrorAlert message={error} />
        ) : null}

        {showProcessingStatus && processingMessage ? (
          <ProcessingStatus message={processingMessage} />
        ) : null}

        <div className="grid gap-6 lg:grid-cols-[minmax(0,1.25fr)_minmax(340px,0.75fr)]">
          <UploadCard
            workbook={workbook}
            onReadFile={readExcelFileInWorker}
            onWorkbookLoaded={handleWorkbookLoaded}
            onError={setError}
            onReadingChange={(isReading, file) => {
              setIsReadingWorkbook(isReading);
              setReadingFileSize(isReading ? file?.size ?? null : null);
            }}
          />
          <SummaryCard
            summary={conversion?.summary}
            warnings={conversion?.warnings ?? []}
          />
        </div>

        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div className="inline-flex w-fit rounded-lg border border-slate-200 bg-white p-1 shadow-sm">
            <PreviewToggle
              active={previewMode === "before"}
              onClick={() => setPreviewMode("before")}
              label="Avant"
            />
            <PreviewToggle
              active={previewMode === "after"}
              onClick={() => setPreviewMode("after")}
              label="Après"
              disabled={!conversion || isConverting}
            />
          </div>

          {conversion ? (
            <p className="flex items-center gap-2 text-sm text-slate-600">
              <ArrowRightLeft aria-hidden="true" size={16} />
              Colonne facture : {conversion.invoiceColumn}
            </p>
          ) : null}
        </div>

      <PreviewTable
        title={preview.title}
        rows={preview.rows}
        columns={preview.columns}
        emptyMessage={preview.emptyMessage}
      />
      </div>

      {creditNoteReferencePromptOpen ? (
        <CreditNoteReferencePrompt
          referenceWorkbook={referenceWorkbook}
          onCancel={() => setCreditNoteReferencePromptOpen(false)}
          onContinue={handleContinueAfterReference}
          onReadReferenceFile={readExcelFileInWorker}
          onReferenceLoaded={(loadedReferenceWorkbook) => {
            setReferenceWorkbook(loadedReferenceWorkbook);
            setError(null);
          }}
          onError={setError}
        />
      ) : null}

      {exchangeRatePromptOpen ? (
        <ExchangeRatePrompt
          detectedInvoices={countInvoicesByCurrency(workbook?.rows ?? [], "USD")}
          error={dgiRateError}
          hasCdfRows={hasCdfRows(workbook?.rows ?? [])}
          isFetchingRate={isFetchingDgiRate}
          isRateValid={isValidRateInput(dgiRateInput)}
          ratePlaceholder={dgiRatePlaceholder}
          rateInput={dgiRateInput}
          rateMeta={dgiRateMeta}
          onApply={handleApplyDgiRate}
          onCancel={() => setExchangeRatePromptOpen(false)}
          onKeepTemplate={handleKeepTemplateRates}
          onRateInputChange={(value) => {
            setDgiRateInput(value);
            setDgiRateError(null);
          }}
        />
      ) : null}
    </main>
  );
}

interface PreviewToggleProps {
  active: boolean;
  label: string;
  onClick: () => void;
  disabled?: boolean;
}

function ProcessingStatus({ message }: { message: string }) {
  return (
    <div
      className="relative overflow-hidden rounded-lg border border-[#cbc5ff] bg-[#f5f3ff] px-4 py-4 text-base text-[#29224f] shadow-sm sm:px-5"
      role="status"
      aria-live="polite"
    >
      <div className="absolute inset-x-0 top-0 h-1 overflow-hidden bg-[#e1ddff]">
        <span className="dexy-progress-bar block h-full rounded-full bg-[#5f55d6]" />
      </div>
      <div className="flex items-start gap-4">
        <div className="relative mt-0.5 flex h-10 w-10 flex-none items-center justify-center rounded-full bg-white shadow-sm ring-1 ring-[#ddd8ff]">
          <span className="dexy-processing-ring absolute inset-1 rounded-full border-2 border-[#d9d4ff] border-t-[#5f55d6]" />
          <span className="dexy-processing-dot h-2.5 w-2.5 rounded-full bg-[#0f9af0]" />
        </div>
        <div className="min-w-0">
          <p className="text-lg font-semibold text-[#171039]">Traitement en cours</p>
          <p className="mt-1 leading-6">{message}</p>
          <p className="mt-2 text-sm font-medium text-[#5f55d6]">
            Merci de patienter, la page reste active pendant la transformation.
          </p>
        </div>
      </div>
    </div>
  );
}

function ValidationErrorAlert({ message }: { message: string }) {
  return (
    <div
      className="relative overflow-hidden rounded-lg border border-[#fecdd3] bg-[#fff1f2] px-4 py-4 text-base text-[#7f1d1d] shadow-sm sm:px-5"
      role="alert"
    >
      <div className="absolute inset-x-0 top-0 h-1 bg-[#fb7185]" />
      <div className="flex items-start gap-4">
        <div className="mt-0.5 flex h-10 w-10 flex-none items-center justify-center rounded-full bg-white shadow-sm ring-1 ring-[#fecdd3]">
          <AlertCircle
            aria-hidden="true"
            className="h-5 w-5 text-[#e11d48]"
            strokeWidth={2.3}
          />
        </div>
        <div className="min-w-0">
          <p className="text-lg font-semibold text-[#881337]">Erreur de validation</p>
          <p className="mt-1 leading-6 text-[#7f1d1d]">{message}</p>
          <p className="mt-2 text-sm font-medium text-[#be123c]">
            Corrigez le fichier ou importez le bon template référence, puis relancez
            la transformation.
          </p>
        </div>
      </div>
    </div>
  );
}

interface ExchangeRatePromptProps {
  detectedInvoices: number;
  error: string | null;
  hasCdfRows: boolean;
  isFetchingRate: boolean;
  isRateValid: boolean;
  ratePlaceholder: string;
  rateInput: string;
  rateMeta: string | null;
  onApply: () => void;
  onCancel: () => void;
  onKeepTemplate: () => void;
  onRateInputChange: (value: string) => void;
}

interface CreditNoteReferencePromptProps {
  referenceWorkbook: ImportedWorkbook | null;
  onCancel: () => void;
  onContinue: () => void;
  onReadReferenceFile: (file: File) => Promise<ImportedWorkbook>;
  onReferenceLoaded: (workbook: ImportedWorkbook) => void;
  onError: (message: string) => void;
}

function CreditNoteReferencePrompt({
  referenceWorkbook,
  onCancel,
  onContinue,
  onReadReferenceFile,
  onReferenceLoaded,
  onError,
}: CreditNoteReferencePromptProps) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [isReadingReference, setIsReadingReference] = useState(false);
  const [localError, setLocalError] = useState<string | null>(null);

  const handleReferenceFile = async (file?: File) => {
    if (!file) {
      return;
    }

    if (!/\.(xlsx|xls|xlsm|csv)$/i.test(file.name)) {
      const message =
        "Format non supporte. Importez le template reference DEXY en Excel.";
      setLocalError(message);
      onError(message);
      return;
    }

    setIsReadingReference(true);
    setLocalError(null);

    try {
      const loadedReferenceWorkbook = await onReadReferenceFile(file);
      onReferenceLoaded(loadedReferenceWorkbook);
    } catch (caughtError) {
      const message =
        caughtError instanceof Error
          ? caughtError.message
          : "Lecture du template reference impossible.";
      setLocalError(message);
      onError(message);
    } finally {
      setIsReadingReference(false);
    }
  };

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    void handleReferenceFile(event.target.files?.[0]);
    event.target.value = "";
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-950/45 px-3 py-4 sm:items-center">
      <section
        aria-modal="true"
        className="w-full max-w-2xl overflow-hidden rounded-2xl bg-white shadow-2xl"
        role="dialog"
      >
        <header className="border-b border-[#746bff] px-5 py-4 sm:px-6">
          <div className="flex items-start gap-4">
            <button
              aria-label="Fermer"
              className="flex h-12 w-12 flex-none items-center justify-center rounded-full bg-[#f0efff] text-ink transition hover:bg-[#e2dfff]"
              type="button"
              onClick={onCancel}
            >
              <X aria-hidden="true" size={22} />
            </button>
            <div className="min-w-0 border-l-2 border-[#5f55d6] pl-4">
              <h2 className="text-lg font-bold text-ink">
                Reference DEXY des factures de vente
              </h2>
              <p className="mt-1 text-sm font-medium text-slate-400">
                Requise pour convertir les avoirs FA
              </p>
            </div>
          </div>
        </header>

        <div className="px-5 py-6 sm:px-6 sm:py-8">
          <p className="text-base leading-7 text-slate-700 sm:text-lg">
            Ce fichier contient des avoirs. Importez le template reference DEXY
            contenant les factures de vente certifiees afin de renseigner le
            Code DEF DGI, la facture d'origine et le type de reference.
          </p>

          <div className="mt-5 rounded-xl border border-dashed border-[#c8c1ff] bg-[#f7f5ff] p-5 text-center">
            <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-xl bg-[#5f55d6] text-white">
              {isReadingReference ? (
                <Loader2 className="animate-spin" aria-hidden="true" size={24} />
              ) : referenceWorkbook ? (
                <CheckCircle2 aria-hidden="true" size={24} />
              ) : (
                <FileSearch aria-hidden="true" size={24} />
              )}
            </div>
            <p className="mt-3 text-sm font-semibold text-ink">
              {isReadingReference
                ? "Lecture du template reference..."
                : referenceWorkbook
                  ? "Template reference DEXY charge"
                  : "Selectionnez le template reference DEXY"}
            </p>
            {referenceWorkbook ? (
              <p className="mt-2 text-xs leading-5 text-slate-500">
                {referenceWorkbook.fileName} - {referenceWorkbook.rows.length}{" "}
                ligne{referenceWorkbook.rows.length > 1 ? "s" : ""}
              </p>
            ) : null}
            <button
              className="mt-4 inline-flex items-center gap-2 rounded-md bg-[#5f55d6] px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-[#5147c4] disabled:cursor-not-allowed disabled:bg-slate-400"
              type="button"
              disabled={isReadingReference}
              onClick={() => inputRef.current?.click()}
            >
              <Upload aria-hidden="true" size={16} />
              {referenceWorkbook ? "Remplacer le template" : "Importer le template"}
            </button>
            <input
              ref={inputRef}
              className="hidden"
              type="file"
              accept=".xlsx,.xls,.xlsm,.csv"
              onChange={handleInputChange}
            />
          </div>

          {localError ? (
            <p className="mt-4 rounded-lg border border-[#8f2146] bg-[#210922] px-4 py-3 text-sm font-semibold text-[#ffb0bf]">
              {localError}
            </p>
          ) : null}
        </div>

        <footer className="flex flex-col gap-3 border-t border-[#746bff] px-5 py-4 sm:flex-row sm:px-6">
          <button
            className="inline-flex h-11 flex-1 items-center justify-center rounded-lg border border-slate-300 bg-white px-5 text-sm font-bold text-ink transition hover:bg-slate-50"
            type="button"
            onClick={onCancel}
          >
            Annuler
          </button>
          <button
            className="inline-flex h-11 flex-1 items-center justify-center rounded-lg bg-[#5f55d6] px-5 text-sm font-bold text-white transition hover:bg-[#5147c4] disabled:cursor-not-allowed disabled:bg-slate-400"
            type="button"
            disabled={!referenceWorkbook || isReadingReference}
            onClick={onContinue}
          >
            Continuer la conversion
          </button>
        </footer>
      </section>
    </div>
  );
}

function ExchangeRatePrompt({
  detectedInvoices,
  error,
  hasCdfRows,
  isFetchingRate,
  isRateValid,
  ratePlaceholder,
  rateInput,
  rateMeta,
  onApply,
  onCancel,
  onKeepTemplate,
  onRateInputChange,
}: ExchangeRatePromptProps) {
  const todayLabel = formatExportDate(new Date());

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-950/45 px-3 py-4 sm:items-center">
      <section
        aria-modal="true"
        className="w-full max-w-2xl overflow-hidden rounded-2xl bg-white shadow-2xl"
        role="dialog"
      >
        <header className="border-b border-[#746bff] px-5 py-4 sm:px-6">
          <div className="flex items-start gap-4">
            <button
              aria-label="Fermer"
              className="flex h-12 w-12 flex-none items-center justify-center rounded-full bg-[#f0efff] text-ink transition hover:bg-[#e2dfff]"
              type="button"
              onClick={onCancel}
            >
              <X aria-hidden="true" size={22} />
            </button>
            <div className="min-w-0 border-l-2 border-[#5f55d6] pl-4">
              <h2 className="text-lg font-bold text-ink">
                Mise à jour des taux de change
              </h2>
              <p className="mt-1 text-sm font-medium text-slate-400">
                Demande de confirmation
              </p>
            </div>
          </div>
        </header>

        <div className="px-5 py-6 sm:px-6 sm:py-8">
          <p className="text-base leading-7 text-slate-700 sm:text-lg">
            Voulez-vous mettre à jour les taux de change des factures USD avec
            le taux DGI du jour ?
          </p>

          <div className="mt-5 grid gap-3 sm:grid-cols-[minmax(0,0.75fr)_minmax(0,1.25fr)]">
            <div className="inline-flex items-center gap-2 rounded-lg bg-[#ede9ff] px-4 py-3 text-sm font-bold text-[#5f55d6]">
              <DollarSign aria-hidden="true" size={18} />
              {detectedInvoices} facture{detectedInvoices > 1 ? "s" : ""} USD
            </div>
            <div className="rounded-lg bg-[#fff4e5] px-4 py-3 text-sm font-semibold text-[#9a4a00]">
              {isFetchingRate
                ? "Récupération du taux DGI en cours..."
                : `Date cours appliquée : ${todayLabel}`}
            </div>
          </div>

          <label className="mt-6 block text-sm font-semibold text-slate-700">
            Taux DGI USD vers CDF
            <input
              className="mt-2 h-12 w-full rounded-lg border border-slate-300 px-4 text-base font-semibold text-ink outline-none transition focus:border-[#5f55d6] focus:ring-4 focus:ring-[#5f55d6]/15"
              inputMode="decimal"
              placeholder={ratePlaceholder}
              type="text"
              value={rateInput}
              onChange={(event) => onRateInputChange(event.target.value)}
            />
          </label>

          {rateMeta ? (
            <p className="mt-3 rounded-lg bg-slate-100 px-4 py-3 text-sm font-semibold text-slate-600">
              {rateMeta}
            </p>
          ) : null}

          <p className="mt-3 text-sm leading-6 text-slate-500">
            {hasCdfRows
              ? "Si vous choisissez Oui, seules les lignes USD seront mises à jour. Les lignes CDF gardent le taux 1."
              : "Si vous choisissez Oui, seules les lignes USD seront mises à jour."}
          </p>

          {error ? (
            <p className="mt-3 rounded-lg border border-[#8f2146] bg-[#210922] px-4 py-3 text-sm font-semibold text-[#ffb0bf]">
              {error}
            </p>
          ) : null}
        </div>

        <footer className="flex flex-col gap-3 border-t border-[#746bff] px-5 py-4 sm:flex-row sm:px-6">
          <button
            className="inline-flex h-11 flex-1 items-center justify-center rounded-lg border border-slate-300 bg-white px-5 text-sm font-bold text-ink transition hover:bg-slate-50"
            type="button"
            onClick={onKeepTemplate}
          >
            NON
          </button>
          <button
            className="inline-flex h-11 flex-1 items-center justify-center rounded-lg bg-[#5f55d6] px-5 text-sm font-bold text-white transition hover:bg-[#5147c4] disabled:cursor-not-allowed disabled:bg-slate-400"
            type="button"
            disabled={isFetchingRate || !isRateValid}
            onClick={onApply}
          >
            {isFetchingRate ? "Chargement..." : "OUI"}
          </button>
        </footer>
      </section>
    </div>
  );
}

function getProcessingMessage({
  fileSize,
  isReadingWorkbook,
  isConverting,
  isExporting,
  rowCount,
}: {
  fileSize: number | null;
  isReadingWorkbook: boolean;
  isConverting: boolean;
  isExporting: boolean;
  rowCount: number;
}): string | null {
  if (isReadingWorkbook) {
    const volumeLabel = getFileVolumeLabel(fileSize);
    return `Lecture du fichier Excel en cours. La durée dépend du volume du fichier${volumeLabel ? ` (${volumeLabel})` : ""}.`;
  }

  if (isConverting) {
    const volumeHint =
      rowCount >= 250
        ? "Ce volume peut prendre un peu plus de temps."
        : "Pour un petit volume, cette étape doit rester rapide.";

    return `Transformation DEXY en cours : regroupement des factures, recalcul des lignes et génération des lignes TAX. ${volumeHint}`;
  }

  if (isExporting) {
    return "Génération du fichier Excel DEXY conforme. Le téléchargement va démarrer automatiquement.";
  }

  return null;
}

function getProcessingDelay({
  fileSize,
  isReadingWorkbook,
  isConverting,
  isExporting,
  rowCount,
}: {
  fileSize: number | null;
  isReadingWorkbook: boolean;
  isConverting: boolean;
  isExporting: boolean;
  rowCount: number;
}): number {
  if (isExporting) {
    return 300;
  }

  if (isConverting) {
    if (rowCount >= 250) {
      return 120;
    }

    if (rowCount >= 80) {
      return 220;
    }

    return 550;
  }

  if (isReadingWorkbook) {
    if (fileSize !== null && fileSize >= 2_000_000) {
      return 120;
    }

    if (fileSize !== null && fileSize >= 700_000) {
      return 220;
    }
  }

  return 550;
}

function getFileVolumeLabel(fileSize: number | null): string | null {
  if (fileSize === null) {
    return null;
  }

  if (fileSize >= 2_000_000) {
    return "fichier volumineux";
  }

  if (fileSize >= 700_000) {
    return "fichier moyen";
  }

  return "petit fichier";
}

type WorkbookDocumentKind = "sales" | "credit" | "mixed" | "unknown";

function getWorkbookDocumentKind(rows: ConversionResult["rows"]): WorkbookDocumentKind {
  const documentTypes = new Set<string>();

  for (const row of rows) {
    const typeValue = getRowValue(row, "Type facture");

    if (typeValue === null || typeValue === undefined || typeValue === "") {
      continue;
    }

    documentTypes.add(String(typeValue).trim().toUpperCase());
  }

  const hasSales = documentTypes.has("FV");
  const hasCredit = documentTypes.has("FA");

  if (hasSales && hasCredit) {
    return "mixed";
  }

  if (hasCredit) {
    return "credit";
  }

  if (hasSales) {
    return "sales";
  }

  return "unknown";
}

function shouldResetCreditNoteReference(
  message: string,
  rows: ConversionResult["rows"],
): boolean {
  if (getWorkbookDocumentKind(rows) !== "credit") {
    return false;
  }

  const normalizedMessage = message
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase();

  return (
    normalizedMessage.includes("conversion des avoirs impossible") ||
    (normalizedMessage.includes("template reference dexy") &&
      normalizedMessage.includes("code def dgi")) ||
    (normalizedMessage.includes("template reference dexy") &&
      normalizedMessage.includes("facture de vente"))
  );
}

function hasUsdRows(rows: ConversionResult["rows"]): boolean {
  return rows.some((row) => getRowCurrency(row) === "USD");
}

function hasCdfRows(rows: ConversionResult["rows"]): boolean {
  return rows.some((row) => getRowCurrency(row) === "CDF");
}

function countInvoicesByCurrency(
  rows: ConversionResult["rows"],
  currency: string,
): number {
  const invoices = new Set<string>();
  const normalizedCurrency = currency.toUpperCase();

  for (const row of rows) {
    if (getRowCurrency(row) !== normalizedCurrency) {
      continue;
    }

    const invoiceNumber =
      getRowValue(row, "Numéro Document") ??
      getRowValue(row, "Numero Document") ??
      getRowValue(row, "Numéro facture") ??
      getRowValue(row, "Numero facture");

    if (invoiceNumber !== null && invoiceNumber !== undefined && invoiceNumber !== "") {
      invoices.add(String(invoiceNumber));
    }
  }

  return invoices.size;
}

function getRowValue(
  row: ConversionResult["rows"][number],
  candidate: string,
): unknown {
  const column = Object.keys(row).find((rowColumn) =>
    isSameExportColumn(rowColumn, candidate),
  );

  return column ? row[column] : null;
}

function parseRateInput(value: unknown): number | null {
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

function formatRateInput(value: number | null): string {
  if (value === null) {
    return "";
  }

  return String(value).replace(".", ",");
}

function isValidRateInput(value: unknown): boolean {
  const rate = parseRateInput(value);
  return rate !== null && rate > 0;
}

type BccUsdRate = {
  currency: "USD";
  rate: number;
  publishedAt: string | null;
  sourceUrl: string;
};

type CachedBccRate = {
  rate: number;
  savedAt: string;
  publishedAt: string | null;
};

function getCachedBccRatePlaceholder(): string {
  const cachedRate = readCachedBccRate();

  if (!cachedRate) {
    return "En attente du taux DGI du jour";
  }

  const savedDate = formatCachedRateDate(cachedRate.savedAt);
  const dateSuffix = savedDate ? ` (${savedDate})` : "";
  return `Dernier taux DGI connu${dateSuffix} : ${formatRateInput(cachedRate.rate)}`;
}

function readCachedBccRate(): CachedBccRate | null {
  try {
    const rawValue = window.localStorage.getItem(BCC_USD_RATE_CACHE_KEY);

    if (!rawValue) {
      return null;
    }

    const parsedValue = JSON.parse(rawValue) as Partial<CachedBccRate>;

    if (
      typeof parsedValue.rate !== "number" ||
      !Number.isFinite(parsedValue.rate) ||
      typeof parsedValue.savedAt !== "string"
    ) {
      return null;
    }

    return {
      rate: parsedValue.rate,
      savedAt: parsedValue.savedAt,
      publishedAt: parsedValue.publishedAt ?? null,
    };
  } catch {
    return null;
  }
}

function cacheBccRate(rate: BccUsdRate): void {
  try {
    const cachedRate: CachedBccRate = {
      rate: rate.rate,
      savedAt: new Date().toISOString(),
      publishedAt: rate.publishedAt,
    };
    window.localStorage.setItem(
      BCC_USD_RATE_CACHE_KEY,
      JSON.stringify(cachedRate),
    );
  } catch {
    // Le cache est seulement un confort visuel : la conversion ne depend pas de lui.
  }
}

function formatCachedRateDate(value: string): string {
  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return "";
  }

  return formatLongFrenchDate(date, false);
}

function formatPublishedAtInFrench(value: string | null): string {
  if (!value) {
    return "";
  }

  const parsedDate = parseEnglishPublishedAt(value);

  if (!parsedDate) {
    return value;
  }

  return formatLongFrenchDate(parsedDate, true);
}

function parseEnglishPublishedAt(value: string): Date | null {
  const match = value.match(
    /^(?:[A-Za-z]+,\s*)?(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})(?:\s+(\d{1,2}):(\d{2}):(\d{2}))?$/,
  );

  if (!match) {
    return null;
  }

  const [, dayText, monthText, yearText, hourText, minuteText, secondText] =
    match;
  const month = ENGLISH_MONTH_INDEX[monthText.toLowerCase()];

  if (month === undefined) {
    return null;
  }

  return new Date(
    Number(yearText),
    month,
    Number(dayText),
    Number(hourText ?? 0),
    Number(minuteText ?? 0),
    Number(secondText ?? 0),
  );
}

function formatLongFrenchDate(value: Date, includeTime: boolean): string {
  const formatter = new Intl.DateTimeFormat("fr-FR", {
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric",
    ...(includeTime
      ? {
          hour: "2-digit",
          minute: "2-digit",
          second: "2-digit",
          hour12: false,
        }
      : {}),
  });
  const formattedDate = formatter.format(value).replace(/\s+/g, " ").trim();

  return formattedDate.charAt(0).toUpperCase() + formattedDate.slice(1);
}

const ENGLISH_MONTH_INDEX: Record<string, number> = {
  january: 0,
  february: 1,
  march: 2,
  april: 3,
  may: 4,
  june: 5,
  july: 6,
  august: 7,
  september: 8,
  october: 9,
  november: 10,
  december: 11,
};

async function fetchBccUsdRate(): Promise<BccUsdRate> {
  const response = await fetch("/api/bcc-rate", {
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error("Taux DGI indisponible.");
  }

  const payload = (await response.json()) as Partial<BccUsdRate>;

  if (
    payload.currency !== "USD" ||
    typeof payload.rate !== "number" ||
    !Number.isFinite(payload.rate)
  ) {
    throw new Error("Réponse DGI invalide.");
  }

  return {
    currency: "USD",
    rate: payload.rate,
    publishedAt: payload.publishedAt ?? null,
    sourceUrl: payload.sourceUrl ?? "",
  };
}

function formatExportDate(value: Date): string {
  const day = String(value.getDate()).padStart(2, "0");
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const year = value.getFullYear();
  return `${day}-${month}-${year}`;
}

function buildDexyExportFileName(
  rows: ConversionResult["rows"],
  exportDate: Date,
): string {
  const currencies = collectExportCurrencies(rows);
  const currencyLabel =
    currencies.length > 1 ? "Multi_Devises" : currencies[0] ?? "CDF";
  const documentKind = getWorkbookDocumentKind(rows);
  const documentLabel = documentKind === "credit" ? "_AVOIR" : "";
  return `Dexy_${currencyLabel}${documentLabel}_${formatExportDate(exportDate)}.xlsx`;
}

function collectExportCurrencies(rows: ConversionResult["rows"]): string[] {
  const currencies: string[] = [];

  for (const row of rows) {
    const currency = getRowCurrency(row);

    if (currency && !currencies.includes(currency)) {
      currencies.push(currency);
    }
  }

  return currencies;
}

function getRowCurrency(row: ConversionResult["rows"][number]): string {
  for (const candidate of EXPORT_CURRENCY_COLUMN_CANDIDATES) {
    const column = Object.keys(row).find((rowColumn) =>
      isSameExportColumn(rowColumn, candidate),
    );
    const value = column ? row[column] : null;
    const currency = formatExportCurrency(value);

    if (currency) {
      return currency;
    }
  }

  return "";
}

function formatExportCurrency(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value)
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function isSameExportColumn(left: unknown, right: unknown): boolean {
  return normalizeExportColumn(left) === normalizeExportColumn(right);
}

function normalizeExportColumn(value: unknown): string {
  return String(value)
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

function PreviewToggle({
  active,
  label,
  onClick,
  disabled = false,
}: PreviewToggleProps) {
  return (
    <button
      className={`rounded-md px-4 py-2 text-sm font-semibold transition disabled:cursor-not-allowed disabled:text-slate-400 ${
        active ? "bg-ink text-white shadow-sm" : "text-slate-600 hover:bg-slate-100"
      }`}
      type="button"
      onClick={onClick}
      disabled={disabled}
    >
      {label}
    </button>
  );
}
