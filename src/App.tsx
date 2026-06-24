import { useEffect, useMemo, useState } from "react";
import {
  AlertCircle,
  ArrowRightLeft,
  DollarSign,
  Download,
  Loader2,
  Play,
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
  const [exchangeRatePromptOpen, setExchangeRatePromptOpen] = useState(false);
  const [dgiRateInput, setDgiRateInput] = useState("");
  const [dgiRateError, setDgiRateError] = useState<string | null>(null);
  const [dgiRateMeta, setDgiRateMeta] = useState<string | null>(null);
  const [dgiRatePlaceholder, setDgiRatePlaceholder] = useState(
    "En attente du taux BCC du jour",
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
    setPreviewMode("before");
    setError(null);
  };

  const handleConvert = async () => {
    if (!workbook) {
      setError("Importez d'abord un fichier Excel Starlink.");
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
      );
      setConversion(convertedWorkbook);
      setPreviewMode("after");
      setError(null);
    } catch (caughtError) {
      setConversion(null);
      setPreviewMode("before");
      setError(
        caughtError instanceof Error
          ? caughtError.message
          : "Conversion impossible.",
      );
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
      setDgiRatePlaceholder("Taux BCC du jour");
      setDgiRateMeta(
        bccRate.publishedAt
          ? `Taux BCC : ${formatRateInput(bccRate.rate)} - ${bccRate.publishedAt}`
          : `Taux BCC : ${formatRateInput(bccRate.rate)}`,
      );
      cacheBccRate(bccRate);
      setDgiRateError(null);
    } catch {
      setDgiRateMeta(
        "Taux BCC indisponible automatiquement. Vous pouvez le renseigner manuellement.",
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
      className="relative overflow-hidden rounded-lg border border-[#cbc5ff] bg-[#f5f3ff] px-4 py-4 text-sm text-[#29224f] shadow-sm sm:px-5"
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
          <p className="font-semibold text-[#171039]">Traitement en cours</p>
          <p className="mt-1 leading-6">{message}</p>
          <p className="mt-2 text-xs font-medium text-[#5f55d6]">
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
      className="flex gap-4 rounded-2xl border border-[#8f2146] bg-[#210922] px-5 py-5 text-[#ff9db0] shadow-sm sm:px-6"
      role="alert"
    >
      <AlertCircle
        aria-hidden="true"
        className="mt-0.5 h-7 w-7 flex-none text-[#ff9db0]"
        strokeWidth={2.2}
      />
      <div className="min-w-0">
        <p className="text-lg font-bold leading-7 text-[#ffb0bf]">
          Erreur de validation :
        </p>
        <p className="mt-1 text-base font-medium leading-7 text-[#ffb0bf] sm:text-lg">
          {message}
        </p>
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
                ? "Récupération du taux BCC en cours..."
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
    return "En attente du taux BCC du jour";
  }

  const savedDate = formatCachedRateDate(cachedRate.savedAt);
  const dateSuffix = savedDate ? ` (${savedDate})` : "";
  return `Dernier taux BCC connu${dateSuffix} : ${formatRateInput(cachedRate.rate)}`;
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

  return formatExportDate(date);
}

async function fetchBccUsdRate(): Promise<BccUsdRate> {
  const response = await fetch("/api/bcc-rate", {
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error("Taux BCC indisponible.");
  }

  const payload = (await response.json()) as Partial<BccUsdRate>;

  if (
    payload.currency !== "USD" ||
    typeof payload.rate !== "number" ||
    !Number.isFinite(payload.rate)
  ) {
    throw new Error("Réponse BCC invalide.");
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
  return `Dexy_${currencyLabel}_${formatExportDate(exportDate)}.xlsx`;
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
