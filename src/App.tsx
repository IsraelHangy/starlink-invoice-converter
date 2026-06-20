import { useMemo, useState } from "react";
import {
  AlertCircle,
  ArrowRightLeft,
  Download,
  Play,
} from "lucide-react";
import UploadCard from "./components/UploadCard";
import PreviewTable from "./components/PreviewTable";
import SummaryCard from "./components/SummaryCard";
import type { ConversionResult, ImportedWorkbook } from "./types";
import { convertStarlinkToDexy } from "./utils/converter";
import { exportExcelFile } from "./utils/excel";

type PreviewMode = "before" | "after";

export default function App() {
  const [workbook, setWorkbook] = useState<ImportedWorkbook | null>(null);
  const [conversion, setConversion] = useState<ConversionResult | null>(null);
  const [previewMode, setPreviewMode] = useState<PreviewMode>("before");
  const [error, setError] = useState<string | null>(null);
  const [isExporting, setIsExporting] = useState(false);

  const canConvert = Boolean(workbook?.rows.length);
  const canExport = Boolean(conversion?.rows.length);

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

  const handleConvert = () => {
    if (!workbook) {
      setError("Importez d'abord un fichier Excel Starlink.");
      return;
    }

    try {
      const convertedWorkbook = convertStarlinkToDexy(
        workbook.rows,
        workbook.columns,
        { normalizationDate: workbook.uploadedAt },
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
    }
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
        `Dexy_conforme_${formatExportDate(new Date())}.xlsx`,
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
            <div className="flex w-full flex-col items-center justify-center gap-5 sm:flex-row sm:gap-8">
              <img
                className="h-20 w-auto max-w-[280px] object-contain sm:h-24 sm:max-w-[340px]"
                src="/assets/dexy-converter-logo.png"
                alt="DEXY CD Converter"
              />
              <img
                className="h-16 w-auto max-w-[240px] object-contain sm:h-20 sm:max-w-[300px]"
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
              onClick={handleConvert}
              disabled={!canConvert}
              title="Convertir le fichier Starlink importe"
            >
              <Play aria-hidden="true" size={17} />
              Convertir
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

        <div className="grid gap-6 lg:grid-cols-[minmax(0,1.25fr)_minmax(340px,0.75fr)]">
          <UploadCard
            workbook={workbook}
            onWorkbookLoaded={handleWorkbookLoaded}
            onError={setError}
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
              disabled={!conversion}
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
    </main>
  );
}

interface PreviewToggleProps {
  active: boolean;
  label: string;
  onClick: () => void;
  disabled?: boolean;
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

function formatExportDate(value: Date): string {
  const day = String(value.getDate()).padStart(2, "0");
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const year = value.getFullYear();
  return `${day}-${month}-${year}`;
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
