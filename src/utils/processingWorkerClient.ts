import type { ConversionResult, ExcelRow, ImportedWorkbook } from "../types";

export type ExchangeRateUpdateRequest = {
  currency: string;
  rate: number;
  rateDate: Date;
};

type WorkerTaskPayload =
  | {
      type: "read-workbook";
      file: File;
    }
  | {
      type: "convert-workbook";
      rows: ExcelRow[];
      columns: string[];
      normalizationDate: Date;
      exchangeRateUpdate?: ExchangeRateUpdateRequest;
    };

type WorkerRequest = WorkerTaskPayload & {
  id: number;
};

type WorkerResponse =
  | {
      id: number;
      ok: true;
      type: "read-workbook";
      workbook: ImportedWorkbook;
    }
  | {
      id: number;
      ok: true;
      type: "convert-workbook";
      conversion: ConversionResult;
    }
  | {
      id: number;
      ok: false;
      message: string;
    };

type PendingRequest = {
  resolve: (value: ImportedWorkbook | ConversionResult) => void;
  reject: (error: Error) => void;
};

let processingWorker: Worker | null = null;
let nextRequestId = 1;
const pendingRequests = new Map<number, PendingRequest>();

export function readExcelFileInWorker(file: File): Promise<ImportedWorkbook> {
  return runWorkerTask<ImportedWorkbook>({
    type: "read-workbook",
    file,
  });
}

export function convertStarlinkToDexyInWorker(
  rows: ExcelRow[],
  columns: string[],
  normalizationDate: Date,
  exchangeRateUpdate?: ExchangeRateUpdateRequest,
): Promise<ConversionResult> {
  return runWorkerTask<ConversionResult>({
    type: "convert-workbook",
    rows,
    columns,
    normalizationDate,
    exchangeRateUpdate,
  });
}

function runWorkerTask<T extends ImportedWorkbook | ConversionResult>(
  payload: WorkerTaskPayload,
): Promise<T> {
  const worker = getProcessingWorker();
  const id = nextRequestId;
  nextRequestId += 1;

  return new Promise<T>((resolve, reject) => {
    pendingRequests.set(id, {
      resolve: (value) => resolve(value as T),
      reject,
    });

    worker.postMessage({
      ...payload,
      id,
    } satisfies WorkerRequest);
  });
}

function getProcessingWorker(): Worker {
  if (processingWorker) {
    return processingWorker;
  }

  processingWorker = new Worker(
    new URL("../workers/processingWorker.ts", import.meta.url),
    { type: "module" },
  );

  processingWorker.addEventListener(
    "message",
    (event: MessageEvent<WorkerResponse>) => {
      const response = event.data;
      const pendingRequest = pendingRequests.get(response.id);

      if (!pendingRequest) {
        return;
      }

      pendingRequests.delete(response.id);

      if (!response.ok) {
        pendingRequest.reject(new Error(response.message));
        return;
      }

      if (response.type === "read-workbook") {
        pendingRequest.resolve(response.workbook);
        return;
      }

      pendingRequest.resolve(response.conversion);
    },
  );

  processingWorker.addEventListener("error", (event) => {
    rejectPendingRequests(
      new Error(event.message || "Traitement interrompu dans le navigateur."),
    );
  });

  processingWorker.addEventListener("messageerror", () => {
    rejectPendingRequests(
      new Error("Le navigateur n'a pas pu transmettre les donnees au worker."),
    );
  });

  return processingWorker;
}

function rejectPendingRequests(error: Error): void {
  pendingRequests.forEach((request) => {
    request.reject(error);
  });
  pendingRequests.clear();
  processingWorker?.terminate();
  processingWorker = null;
}
