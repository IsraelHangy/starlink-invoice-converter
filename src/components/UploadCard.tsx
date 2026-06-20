import { ChangeEvent, DragEvent, useRef, useState } from "react";
import { CheckCircle2, FileSpreadsheet, Upload } from "lucide-react";
import type { ImportedWorkbook } from "../types";
import { readExcelFile } from "../utils/excel";

interface UploadCardProps {
  workbook: ImportedWorkbook | null;
  onWorkbookLoaded: (workbook: ImportedWorkbook) => void;
  onError: (message: string) => void;
  onReadingChange?: (isReading: boolean) => void;
}

export default function UploadCard({
  workbook,
  onWorkbookLoaded,
  onError,
  onReadingChange,
}: UploadCardProps) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isReading, setIsReading] = useState(false);

  const handleFile = async (file?: File) => {
    if (!file) {
      return;
    }

    if (!/\.(xlsx|xls|xlsm|csv)$/i.test(file.name)) {
      onError("Format non supporté. Importez un fichier .xlsx, .xls, .xlsm ou .csv.");
      return;
    }

    setIsReading(true);
    onReadingChange?.(true);

    try {
      const importedWorkbook = await readExcelFile(file);
      onWorkbookLoaded(importedWorkbook);
    } catch (error) {
      onError(error instanceof Error ? error.message : "Lecture du fichier impossible.");
    } finally {
      setIsReading(false);
      onReadingChange?.(false);
    }
  };

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    void handleFile(event.target.files?.[0]);
  };

  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(false);
    void handleFile(event.dataTransfer.files?.[0]);
  };

  return (
    <section className="rounded-lg border border-slate-200 bg-white p-5 shadow-soft">
      <div className="flex items-start justify-between gap-4">
        <div className="flex min-w-0 items-start gap-3">
        <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-lg bg-[#5f55d6] text-white shadow-sm">
          <FileSpreadsheet aria-hidden="true" size={22} />
        </div>
        <div className="min-w-0">
          <h2 className="text-lg font-semibold text-ink">Import Excel Starlink</h2>
          <p className="mt-1 text-sm leading-6 text-slate-600">
            Chargez le fichier source contenant la colonne taxe d'accise télécom.
          </p>
        </div>
        </div>
        {workbook ? (
          <span className="inline-flex shrink-0 items-center gap-1.5 rounded-md bg-[#f4f2ff] px-2.5 py-1 text-xs font-semibold text-[#5f55d6]">
            <CheckCircle2 aria-hidden="true" size={14} />
            Importé
          </span>
        ) : null}
      </div>

      <div
        className={`mt-5 flex min-h-44 flex-col items-center justify-center rounded-lg border-2 border-dashed px-5 py-7 text-center transition ${
          isDragging
            ? "border-[#5f55d6] bg-[#f4f2ff] shadow-inner"
            : workbook
              ? "border-[#c8c1ff] bg-[#f4f2ff]"
              : "border-slate-300 bg-slate-50"
        }`}
        onDragOver={(event) => {
          event.preventDefault();
          setIsDragging(true);
        }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={handleDrop}
      >
        {workbook ? (
          <CheckCircle2 className="text-[#5f55d6]" aria-hidden="true" size={32} />
        ) : (
          <Upload className="text-[#5f55d6]" aria-hidden="true" size={30} />
        )}
        <p className="mt-3 text-sm font-medium text-ink">
          {isReading
            ? "Lecture et analyse du fichier Excel en cours..."
            : workbook
              ? "Fichier Starlink chargé"
              : "Glissez un fichier ici ou choisissez-le manuellement."}
        </p>
        {isReading ? (
          <p className="mt-2 max-w-md text-xs leading-5 text-slate-500">
            Les fichiers avec beaucoup de factures peuvent prendre quelques
            minutes. Gardez cette page ouverte jusqu'à la fin du traitement.
          </p>
        ) : null}
        <button
          className="mt-4 inline-flex items-center gap-2 rounded-md bg-[#5f55d6] px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-[#5147c4] disabled:cursor-not-allowed disabled:bg-slate-400"
          type="button"
          onClick={() => inputRef.current?.click()}
          disabled={isReading}
        >
          <Upload aria-hidden="true" size={16} />
          {isReading ? "Lecture..." : workbook ? "Remplacer le fichier" : "Choisir un fichier"}
        </button>
        <input
          ref={inputRef}
          className="hidden"
          type="file"
          accept=".xlsx,.xls,.xlsm,.csv"
          onChange={handleInputChange}
        />
      </div>

      {workbook ? (
        <dl className="mt-4 grid grid-cols-1 gap-3 rounded-lg border border-slate-200 bg-[#f8fafc] p-4 text-sm sm:grid-cols-3">
          <div>
            <dt className="font-medium text-slate-500">Fichier</dt>
            <dd className="mt-1 truncate font-semibold text-ink">{workbook.fileName}</dd>
          </div>
          <div>
            <dt className="font-medium text-slate-500">Feuille</dt>
            <dd className="mt-1 truncate font-semibold text-ink">{workbook.sheetName}</dd>
          </div>
          <div>
            <dt className="font-medium text-slate-500">Lignes</dt>
            <dd className="mt-1 font-semibold text-ink">{workbook.rows.length}</dd>
          </div>
        </dl>
      ) : null}
    </section>
  );
}
