import type { ConversionResult, ExcelRow, ImportedWorkbook } from "../types";
import { convertStarlinkToDexy } from "../utils/converter";
import { readExcelFile } from "../utils/excel";

type ReadWorkbookRequest = {
  id: number;
  type: "read-workbook";
  file: File;
};

type ConvertWorkbookRequest = {
  id: number;
  type: "convert-workbook";
  rows: ExcelRow[];
  columns: string[];
  normalizationDate: Date;
  exchangeRateUpdate?: {
    currency: string;
    rate: number;
    rateDate: Date;
  };
};

type WorkerRequest = ReadWorkbookRequest | ConvertWorkbookRequest;

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

self.addEventListener("message", (event: MessageEvent<WorkerRequest>) => {
  void handleMessage(event.data);
});

async function handleMessage(message: WorkerRequest): Promise<void> {
  try {
    if (message.type === "read-workbook") {
      const workbook = await readExcelFile(message.file);
      postResponse({
        id: message.id,
        ok: true,
        type: "read-workbook",
        workbook,
      });
      return;
    }

    const conversion = convertStarlinkToDexy(message.rows, message.columns, {
      normalizationDate: message.normalizationDate,
      exchangeRateUpdate: message.exchangeRateUpdate,
    });

    postResponse({
      id: message.id,
      ok: true,
      type: "convert-workbook",
      conversion,
    });
  } catch (caughtError) {
    postResponse({
      id: message.id,
      ok: false,
      message:
        caughtError instanceof Error
          ? caughtError.message
          : "Traitement impossible.",
    });
  }
}

function postResponse(response: WorkerResponse): void {
  self.postMessage(response);
}
