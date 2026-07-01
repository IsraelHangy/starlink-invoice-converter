import type {
  ConversionResult,
  ExcelCell,
  ExcelRow,
  ReferenceWorkbook,
} from "../types";

export const EXCISE_TAX_COLUMN =
  "Congo DRC Telecommunication Excise Tax_Final";
export const DATE_FACTURE_COLUMN = "Date facture";
export const COMMENT_A_COLUMN = "Commentaire A";
export const COMMENT_B_COLUMN = "Commentaire B";
export const COMMENT_C_COLUMN = "Commentaire C";
export const COMMENT_D_COLUMN = "Commentaire D";
export const COMMENT_E_COLUMN = "Commentaire E";
export const COMMENT_F_COLUMN = "Commentaire F";
export const COMMENT_G_COLUMN = "Commentaire G";
export const COMMENT_H_COLUMN = "Commentaire H";
export const OLD_DATE_COLUMN = "Old Date";

const TYPE_FACTURE_COLUMN = "Type facture";
const SALES_DOCUMENT_TYPE = "FV";
const CREDIT_NOTE_DOCUMENT_TYPE = "FA";
const CREDIT_NOTE_ORIGIN_DGI_COLUMN = "Facture origine [FA | EA]";
const CREDIT_NOTE_REFERENCE_TYPE_COLUMN = "Type référence [FA | EA]";
const CREDIT_NOTE_REFERENCE_DESCRIPTION_COLUMN =
  "Description de référence [FA | EA]";
export const CREDIT_NOTE_ORIGIN_ERP_INVOICE_COLUMN =
  "Numéro Facture origne [ERP]";
export const CREDIT_NOTE_ORIGIN_ERP_DATE_COLUMN =
  "Date Facture origine [ERP]";

const EXCISE_TAX_COLUMN_CANDIDATES = [
  EXCISE_TAX_COLUMN,
  "Excise Tax_Local",
  "Excise Tax Local",
  "excise_tax_local",
  "Congo DRC Telecommunication Excise Tax Final",
  "congo_drc_telecommunication_excise_tax_final",
  "telecommunication_excise_tax_final",
  "telecom_excise_tax_final",
  "excise_tax_final",
  "Excise Tax Final",
  "Telecommunication Excise Tax",
  "Telecom Excise Tax",
];

const DATE_FACTURE_COLUMN_CANDIDATES = [
  DATE_FACTURE_COLUMN,
  "Date de facture",
  "Date facture original",
  "Original invoice date",
  "original_invoice_date",
  "invoice_date",
  "Invoice Date",
  "Billing Date",
];

const TYPOLOGY_COLUMN_CANDIDATES = [
  "Typologie client",
  "Typologie Client",
  "Client typology",
  "Customer typology",
  "account_type",
];

const NIF_CLIENT_COLUMN_CANDIDATES = [
  "NIF Client",
  "NIF client",
  "NIF",
  "Numéro fiscal client",
  "Numero fiscal client",
  "Tax ID",
  "Customer Tax ID",
];

const TYPE_FACTURE_COLUMN_CANDIDATES = [
  TYPE_FACTURE_COLUMN,
  "Type de facture",
  "Invoice Type",
  "Document Type",
];

const TOTAL_HT_COLUMN_CANDIDATES = [
  "Total HT",
  "Total hors taxe",
  "Total Hors Taxe",
  "Total tax basis",
  "transaction_tax_basis",
];

const TOTAL_TTC_COLUMN_CANDIDATES = [
  "Total TTC",
  "Total toutes taxes",
  "Total gross amount",
  "transaction_gross_amount",
];

const REFERENCE_FACTURE_COLUMN_CANDIDATES = [
  "Facture",
  "Numéro facture",
  "Numero facture",
  "Numéro Document",
  "Numero Document",
  "Invoice Number",
  "Invoice",
];

const REFERENCE_DATE_COLUMN_CANDIDATES = [
  "Date",
  DATE_FACTURE_COLUMN,
  "Date facture origine",
  "Invoice Date",
];

const REFERENCE_CODE_DEF_DGI_COLUMN_CANDIDATES = [
  "Code DEF DGI",
  "Code DEF",
  "DEF DGI",
  "Code DGI",
];

const B2F_CURRENCY_COLUMN_CANDIDATES = [
  "B2F Devise [Nom]",
  "B2F Devise Nom",
  "Devise Import",
  "Currency",
  "Transaction Currency",
  "transaction_currency",
];

const B2F_EXCHANGE_RATE_DATE_COLUMN_CANDIDATES = [
  "B2F Devise [Date cours]",
  "B2F Devise Date cours",
];

const B2F_EXCHANGE_RATE_COLUMN_CANDIDATES = [
  "B2F Devise [Taux de change]",
  "B2F Devise Taux de change",
];

const IMPORT_CURRENCY_COLUMN = "Devise Import";
const IMPORT_EXCHANGE_RATE_COLUMN = "Taux Devise Import vers CDF";
const EXCHANGE_RATE_COLUMN_CANDIDATES = [
  ...B2F_EXCHANGE_RATE_COLUMN_CANDIDATES,
  IMPORT_EXCHANGE_RATE_COLUMN,
];
const DEFAULT_AMOUNT_TOLERANCE = 0.05;
const CDF_AMOUNT_TOLERANCE = 5;
const MAX_INFERRED_RATE_DRIFT_RATIO = 0.02;
const MAX_INFERRED_RATE_PAIR_DRIFT_RATIO = 0.002;

type ConversionOptions = {
  normalizationDate?: Date;
  exchangeRateUpdate?: {
    currency: string;
    rate: number;
    rateDate: Date;
  };
  referenceWorkbook?: ReferenceWorkbook;
};

type DexyFieldKey =
  | "lineNumber"
  | "itemReference"
  | "itemDescription"
  | "itemType"
  | "taxGroup"
  | "quantity"
  | "unitPrice"
  | "unitPriceMode";

type InvoiceGroup = {
  invoiceNumber: string;
  documentType: string;
  rows: ExcelRow[];
  taxTotal: number;
  totalHT: number;
  totalTTC: number;
  currencyCode: string;
  exchangeRate?: number;
};

type InvoiceGroupingResult = {
  groups: InvoiceGroup[];
  skippedRowsWithoutInvoice: number;
};

type ReferenceInvoice = {
  invoiceNumber: string;
  codeDefDgi: string;
  date: ExcelCell;
  totalHT: number;
  totalTTC: number;
  currencyCode: string;
  exchangeRate?: number;
  exchangeRateDate?: ExcelCell;
};

type CreditNoteContext = {
  isCreditNote: boolean;
  reference?: ReferenceInvoice;
  referenceType?: "RAN" | "COR";
  referenceDescription?: "ANNULATION" | "CORRECTION";
  inferredExchangeRate?: number;
};

type CreditNoteReferenceMatch = {
  isFullCancellation: boolean;
  inferredExchangeRate?: number;
};

type CreditNoteResolution = {
  contexts: Map<string, CreditNoteContext>;
  totalCreditNotes: number;
  matchedCreditNotes: number;
  missingCreditNotes: number;
};

const LINE_NUMBER_FALLBACK_COLUMN = "N° ligne";

const INVOICE_COLUMN_CANDIDATES = [
  "Numéro Document",
  "Numero Document",
  "Numéro de Document",
  "Numero de Document",
  "Numéro Documents",
  "Numero Documents",
  "Numéro de Documents",
  "Numero de Documents",
  "Document Number",
  "Document No",
  "Numéro de facture",
  "Numero de facture",
  "Numéro facture",
  "Numero facture",
  "N° facture",
  "N° de facture",
  "No facture",
  "Nº facture",
  "Facture",
  "Invoice Number",
  "Invoice No",
  "Invoice",
  "Billing Invoice",
];

const DEXY_FIELD_CANDIDATES: Record<DexyFieldKey, string[]> = {
  lineNumber: [
    "Numéro de ligne",
    "Numero de ligne",
    "Numéro ligne",
    "Numero ligne",
    "Numéro ligne facture",
    "Numero ligne facture",
    "N° ligne",
    "N° de ligne",
    "N° ligne facture",
    "N° de ligne facture",
    "No ligne",
    "No de ligne",
    "No ligne facture",
    "Nº ligne",
    "Nº de ligne",
    "Ligne facture",
    "Ligne article",
    "Num ligne",
    "Num. ligne",
    "N ligne",
    "Line Number",
    "Line No",
    "Invoice Line Number",
    "Invoice Line No",
    "Line Item Number",
    "Line",
  ],
  itemReference: [
    "Référence article",
    "Reference article",
    "Réf. article",
    "Article Reference",
    "Item Reference",
    "SKU",
  ],
  itemDescription: [
    "Désignation article",
    "Designation article",
    "Libellé article",
    "Libelle article",
    "Article Description",
    "Item Description",
    "Description",
  ],
  itemType: ["Type article", "Article Type", "Item Type", "Line Type", "Type"],
  taxGroup: [
    "Groupe taxation",
    "Groupe taxe",
    "Tax Group",
    "Taxation Group",
  ],
  quantity: ["Quantité", "Quantite", "Quantity", "Qty"],
  unitPrice: ["Prix unitaire", "Unit Price", "Prix unitaire HT"],
  unitPriceMode: [
    "Mode prix unitaire [HT | TTC]",
    "Mode prix unitaire HT TTC",
    "Mode prix unitaire",
    "Mode prix",
    "Unit Price Mode",
    "Price Mode",
  ],
};

const DEFAULT_DEXY_COLUMNS: Record<DexyFieldKey, string> = {
  lineNumber: LINE_NUMBER_FALLBACK_COLUMN,
  itemReference: "Référence article",
  itemDescription: "Désignation article",
  itemType: "Type article",
  taxGroup: "Groupe taxation",
  quantity: "Quantité",
  unitPrice: "Prix unitaire",
  unitPriceMode: "Mode prix unitaire",
};

const TAX_LINE_VALUES = {
  itemReference: "ACCISE",
  itemDescription: "Congo DRC Telecommunication Excise Tax (10%)",
  itemType: "TAX",
  taxGroup: "L",
  quantity: 1,
  unitPriceMode: "HT",
} as const;

const DEFAULT_EXISTING_COLUMN_VALUES: Array<{
  candidates: string[];
  value: string;
}> = [
  {
    candidates: [
      "Méthode de paiement",
      "Methode de paiement",
      "Mode de paiement",
      "Mode paiement",
      "Moyen de paiement",
      "Payment Method",
    ],
    value: "VIREMENT",
  },
  {
    candidates: [
      "Opérateur [Code]",
      "Operateur [Code]",
      "Code opérateur",
      "Code operateur",
      "Operator Code",
    ],
    value: "ADM",
  },
  {
    candidates: [
      "Opérateur [Nom]",
      "Operateur [Nom]",
      "Nom opérateur",
      "Nom operateur",
      "Operator Name",
    ],
    value: "Administrateur",
  },
];

export function convertStarlinkToDexy(
  rows: ExcelRow[],
  sourceColumns?: string[],
  options: ConversionOptions = {},
): ConversionResult {
  if (rows.length === 0) {
    throw new Error("Le fichier ne contient aucune ligne à convertir.");
  }

  const columns = normalizeColumns(sourceColumns ?? collectColumns(rows));
  const taxColumn = findExciseTaxColumn(columns);
  const invoiceColumn = findColumn(columns, INVOICE_COLUMN_CANDIDATES);
  const documentTypeColumn = findColumn(columns, TYPE_FACTURE_COLUMN_CANDIDATES);
  const totalHTColumn = findColumn(columns, TOTAL_HT_COLUMN_CANDIDATES);
  const totalTTCColumn = findColumn(columns, TOTAL_TTC_COLUMN_CANDIDATES);
  const currencyColumn = findColumn(columns, B2F_CURRENCY_COLUMN_CANDIDATES);
  const exchangeRateColumn = findColumn(columns, EXCHANGE_RATE_COLUMN_CANDIDATES);
  const documentTypeSummary = summarizeDocumentTypes(rows, documentTypeColumn);
  const missingColumns = [
    !invoiceColumn ? "Numéro facture" : null,
    !taxColumn ? `${EXCISE_TAX_COLUMN} ou Excise Tax_Local` : null,
  ].filter((column): column is string => Boolean(column));

  if (!taxColumn || !invoiceColumn) {
    throw new Error(
      `Les colonnes obligatoires suivantes sont absentes de la feuille : ${missingColumns.join(
        ", ",
      )}.`,
    );
  }

  if (documentTypeSummary.hasSales && documentTypeSummary.hasCreditNotes) {
    throw new Error(
      "Ce fichier contient a la fois des ventes (FV) et des avoirs (FA). Importez les ventes et les avoirs dans deux fichiers separes.",
    );
  }

  if (documentTypeSummary.hasCreditNotes && !options.referenceWorkbook) {
    throw new Error(
      "Le fichier contient des avoirs (FA). Importez le template reference DEXY contenant les factures de vente certifiees avant de convertir.",
    );
  }

  const dateColumn =
    findColumn(columns, DATE_FACTURE_COLUMN_CANDIDATES) ?? DATE_FACTURE_COLUMN;
  const normalizationDate = toDateOnly(options.normalizationDate ?? new Date());
  const dexyColumns = resolveDexyColumns(columns);
  const outputColumns = buildOutputColumns(
    columns,
    taxColumn,
    dateColumn,
    dexyColumns,
  );
  const { groups, skippedRowsWithoutInvoice } = groupRowsByInvoice(
    rows,
    invoiceColumn,
    taxColumn,
    documentTypeColumn,
    totalHTColumn,
    totalTTCColumn,
    currencyColumn,
    exchangeRateColumn,
  );
  const warnings: string[] = [];
  const referenceInvoices = documentTypeSummary.hasCreditNotes
    ? buildReferenceInvoiceIndex(options.referenceWorkbook)
    : new Map<string, ReferenceInvoice>();
  const technicalRows: ExcelRow[] = [];
  let totalTax = 0;
  let taxRowsAdded = 0;

  if (groups.length === 0) {
    throw new Error(
      "Aucune facture valide n'a été trouvée : les lignes doivent contenir un numéro de facture.",
    );
  }

  const creditNoteResolution = documentTypeSummary.hasCreditNotes
    ? resolveCreditNoteContexts(groups, referenceInvoices)
    : null;

  if (
    creditNoteResolution &&
    creditNoteResolution.totalCreditNotes > 0 &&
    creditNoteResolution.matchedCreditNotes === 0
  ) {
    throw new Error(
      "Conversion des avoirs impossible : aucune facture d'avoir n'a retrouvé sa facture de vente certifiée avec Code DEF DGI dans le template référence DEXY. Importez un template référence contenant les ventes certifiées correspondant à ces avoirs, puis relancez la conversion.",
    );
  }

  if (creditNoteResolution && creditNoteResolution.missingCreditNotes > 0) {
    warnings.push(
      `Références DEXY partiellement trouvées : ${formatCount(
        creditNoteResolution.missingCreditNotes,
        "avoir introuvable",
        "avoirs introuvables",
      )} sur ${creditNoteResolution.totalCreditNotes}. Les champs d'origine FA restent vides uniquement pour les avoirs non retrouvés.`,
    );
  }

  for (const group of groups) {
    const creditNoteContext =
      creditNoteResolution?.contexts.get(getCreditNoteContextKey(group)) ??
      buildCreditNoteContext(group, referenceInvoices);

    group.rows.forEach((row, index) => {
      const outputRow = createTraceableRow(
        row,
        taxColumn,
        invoiceColumn,
        dateColumn,
        group.invoiceNumber,
        normalizationDate,
        options.exchangeRateUpdate,
        dexyColumns,
        creditNoteContext,
      );
      outputRow[dexyColumns.lineNumber] = index + 1;
      technicalRows.push(outputRow);
    });

    const roundedTax = roundCurrency(group.taxTotal);
    const exportedTax = creditNoteContext.isCreditNote
      ? Math.abs(roundedTax)
      : roundedTax;

    if (roundedTax !== 0) {
      taxRowsAdded += 1;
      totalTax += exportedTax;
      technicalRows.push(
        createTaxRow(
          group,
          taxColumn,
          invoiceColumn,
          dateColumn,
          dexyColumns,
          normalizationDate,
          options.exchangeRateUpdate,
          group.rows.length + 1,
          creditNoteContext,
        ),
      );
    }

  }

  if (skippedRowsWithoutInvoice > 0) {
    warnings.push(
      `${skippedRowsWithoutInvoice} ${skippedRowsWithoutInvoice > 1 ? "lignes sans numéro de facture ont été ignorées" : "ligne sans numéro de facture a été ignorée"} afin de ne pas créer de facture vide ni de ligne TAX en trop.`,
    );
  }

  const outputRows = technicalRows.map((row) =>
    stripColumn(row, OLD_DATE_COLUMN),
  );

  return {
    rows: outputRows,
    columns: outputColumns,
    summary: {
      invoicesProcessed: groups.length,
      originalRows: rows.length,
      taxRowsAdded,
      totalTax: roundCurrency(totalTax),
    },
    invoiceColumn,
    taxColumn,
    normalizationDate,
    warnings: Array.from(new Set(warnings)),
  };
}

function groupRowsByInvoice(
  rows: ExcelRow[],
  invoiceColumn: string,
  taxColumn: string,
  documentTypeColumn?: string,
  totalHTColumn?: string,
  totalTTCColumn?: string,
  currencyColumn?: string,
  exchangeRateColumn?: string,
): InvoiceGroupingResult {
  const groupsByInvoice = new Map<string, InvoiceGroup>();
  let skippedRowsWithoutInvoice = 0;

  for (const row of rows) {
    const invoiceNumber = normalizeInvoiceNumber(row[invoiceColumn]);

    if (invoiceNumber === "") {
      skippedRowsWithoutInvoice += 1;
      continue;
    }

    const existingGroup = groupsByInvoice.get(invoiceNumber);
    const group =
      existingGroup ??
      {
        invoiceNumber,
        documentType: normalizeDocumentType(
          documentTypeColumn ? row[documentTypeColumn] : null,
        ),
        rows: [],
        taxTotal: 0,
        totalHT: totalHTColumn ? parseAmount(row[totalHTColumn]) : 0,
        totalTTC: totalTTCColumn ? parseAmount(row[totalTTCColumn]) : 0,
        currencyCode: currencyColumn ? normalizeCurrencyCode(row[currencyColumn]) : "",
        exchangeRate: exchangeRateColumn
          ? parseExchangeRate(row[exchangeRateColumn])
          : undefined,
      };

    group.rows.push({ ...row });
    group.taxTotal += parseAmount(row[taxColumn]);
    if (!group.documentType && documentTypeColumn) {
      group.documentType = normalizeDocumentType(row[documentTypeColumn]);
    }
    if (group.totalHT === 0 && totalHTColumn) {
      group.totalHT = parseAmount(row[totalHTColumn]);
    }
    if (group.totalTTC === 0 && totalTTCColumn) {
      group.totalTTC = parseAmount(row[totalTTCColumn]);
    }
    if (!group.currencyCode && currencyColumn) {
      group.currencyCode = normalizeCurrencyCode(row[currencyColumn]);
    }
    if (!group.exchangeRate && exchangeRateColumn) {
      group.exchangeRate = parseExchangeRate(row[exchangeRateColumn]);
    }
    groupsByInvoice.set(invoiceNumber, group);
  }

  return {
    groups: Array.from(groupsByInvoice.values()),
    skippedRowsWithoutInvoice,
  };
}

function buildReferenceInvoiceIndex(
  referenceWorkbook?: ReferenceWorkbook,
): Map<string, ReferenceInvoice> {
  if (!referenceWorkbook) {
    return new Map();
  }

  const referenceColumns = normalizeColumns(
    referenceWorkbook.columns.length > 0
      ? referenceWorkbook.columns
      : collectColumns(referenceWorkbook.rows),
  );
  const typeColumn = findColumn(referenceColumns, TYPE_FACTURE_COLUMN_CANDIDATES);
  const invoiceColumn = findColumn(
    referenceColumns,
    REFERENCE_FACTURE_COLUMN_CANDIDATES,
  );
  const codeDefColumn = findColumn(
    referenceColumns,
    REFERENCE_CODE_DEF_DGI_COLUMN_CANDIDATES,
  );
  const dateColumn = findColumn(referenceColumns, REFERENCE_DATE_COLUMN_CANDIDATES);
  const totalHTColumn = findColumn(referenceColumns, TOTAL_HT_COLUMN_CANDIDATES);
  const totalTTCColumn = findColumn(referenceColumns, TOTAL_TTC_COLUMN_CANDIDATES);
  const currencyColumn = findColumn(referenceColumns, B2F_CURRENCY_COLUMN_CANDIDATES);
  const exchangeRateColumn = findColumn(
    referenceColumns,
    EXCHANGE_RATE_COLUMN_CANDIDATES,
  );
  const exchangeRateDateColumn = findColumn(
    referenceColumns,
    B2F_EXCHANGE_RATE_DATE_COLUMN_CANDIDATES,
  );
  const missingColumns = [
    !typeColumn ? "Type facture" : null,
    !invoiceColumn ? "Facture" : null,
    !codeDefColumn ? "Code DEF DGI" : null,
    !dateColumn ? "Date" : null,
    !totalHTColumn ? "Total HT" : null,
    !totalTTCColumn ? "Total TTC" : null,
  ].filter((column): column is string => Boolean(column));

  if (
    !typeColumn ||
    !invoiceColumn ||
    !codeDefColumn ||
    !dateColumn ||
    !totalHTColumn ||
    !totalTTCColumn
  ) {
    throw new Error(
      `Le template reference DEXY ne contient pas les colonnes requises : ${missingColumns.join(
        ", ",
      )}.`,
    );
  }

  const referenceInvoices = new Map<string, ReferenceInvoice>();

  for (const row of referenceWorkbook.rows) {
    if (normalizeDocumentType(row[typeColumn]) !== SALES_DOCUMENT_TYPE) {
      continue;
    }

    const invoiceNumber = normalizeInvoiceNumber(row[invoiceColumn]);
    const key = normalizeReferenceInvoiceKey(invoiceNumber);
    const codeDefDgi = formatCellAsText(row[codeDefColumn]);

    if (!key || !codeDefDgi) {
      continue;
    }

    referenceInvoices.set(key, {
      invoiceNumber,
      codeDefDgi,
      date: row[dateColumn],
      totalHT: parseAmount(row[totalHTColumn]),
      totalTTC: parseAmount(row[totalTTCColumn]),
      currencyCode: currencyColumn
        ? normalizeCurrencyCode(row[currencyColumn])
        : "",
      exchangeRate: exchangeRateColumn
        ? parseExchangeRate(row[exchangeRateColumn])
        : undefined,
      exchangeRateDate: exchangeRateDateColumn
        ? row[exchangeRateDateColumn]
        : undefined,
    });
  }

  if (referenceInvoices.size === 0) {
    throw new Error(
      "Le template référence DEXY ne contient aucune facture de vente (FV) exploitable avec un Code DEF DGI.",
    );
  }

  return referenceInvoices;
}

function resolveCreditNoteContexts(
  groups: InvoiceGroup[],
  referenceInvoices: Map<string, ReferenceInvoice>,
): CreditNoteResolution {
  const contexts = new Map<string, CreditNoteContext>();
  let totalCreditNotes = 0;
  let matchedCreditNotes = 0;
  let missingCreditNotes = 0;

  for (const group of groups) {
    if (group.documentType !== CREDIT_NOTE_DOCUMENT_TYPE) {
      continue;
    }

    totalCreditNotes += 1;
    const context = buildCreditNoteContext(group, referenceInvoices);
    contexts.set(getCreditNoteContextKey(group), context);

    if (context.reference?.codeDefDgi) {
      matchedCreditNotes += 1;
    } else {
      missingCreditNotes += 1;
    }
  }

  return {
    contexts,
    totalCreditNotes,
    matchedCreditNotes,
    missingCreditNotes,
  };
}

function getCreditNoteContextKey(group: InvoiceGroup): string {
  return normalizeReferenceInvoiceKey(group.invoiceNumber);
}

function buildCreditNoteContext(
  group: InvoiceGroup,
  referenceInvoices: Map<string, ReferenceInvoice>,
): CreditNoteContext {
  if (group.documentType !== CREDIT_NOTE_DOCUMENT_TYPE) {
    return {
      isCreditNote: false,
    };
  }

  const reference = referenceInvoices.get(
    normalizeReferenceInvoiceKey(group.invoiceNumber),
  );

  if (!reference) {
    return {
      isCreditNote: true,
    };
  }

  const referenceMatch = resolveCreditNoteReferenceMatch(group, reference);

  return {
    isCreditNote: true,
    reference,
    referenceType: referenceMatch.isFullCancellation ? "RAN" : "COR",
    referenceDescription: referenceMatch.isFullCancellation
      ? "ANNULATION"
      : "CORRECTION",
    inferredExchangeRate: referenceMatch.inferredExchangeRate,
  };
}

function resolveCreditNoteReferenceMatch(
  group: InvoiceGroup,
  reference: ReferenceInvoice,
): CreditNoteReferenceMatch {
  if (
    amountPairMatches(
      group.totalHT,
      group.totalTTC,
      reference.totalHT,
      reference.totalTTC,
    )
  ) {
    return { isFullCancellation: true };
  }

  const rates = uniquePositiveRates([
    reference.exchangeRate,
    group.exchangeRate,
  ]);

  for (const rate of rates) {
    if (
      amountPairMatches(
        group.totalHT * rate,
        group.totalTTC * rate,
        reference.totalHT,
        reference.totalTTC,
        CDF_AMOUNT_TOLERANCE,
      )
    ) {
      return { isFullCancellation: true };
    }

    if (
      amountPairMatches(
        group.totalHT,
        group.totalTTC,
        reference.totalHT * rate,
        reference.totalTTC * rate,
        CDF_AMOUNT_TOLERANCE,
      )
    ) {
      return { isFullCancellation: true };
    }
  }

  const inferredExchangeRate = inferReferenceExchangeRate(group, reference);

  return {
    isFullCancellation: Boolean(inferredExchangeRate),
    inferredExchangeRate,
  };
}

function amountPairMatches(
  leftHT: number,
  leftTTC: number,
  rightHT: number,
  rightTTC: number,
  tolerance = DEFAULT_AMOUNT_TOLERANCE,
): boolean {
  return (
    amountsMatch(leftHT, rightHT, tolerance) &&
    amountsMatch(leftTTC, rightTTC, tolerance)
  );
}

function amountsMatch(
  left: number,
  right: number,
  tolerance = DEFAULT_AMOUNT_TOLERANCE,
): boolean {
  return (
    Math.abs(roundCurrency(Math.abs(left)) - roundCurrency(Math.abs(right))) <=
    tolerance
  );
}

function uniquePositiveRates(rates: Array<number | undefined>): number[] {
  const uniqueRates: number[] = [];

  for (const rate of rates) {
    if (!rate || rate <= 0 || !Number.isFinite(rate)) {
      continue;
    }

    if (!uniqueRates.some((existingRate) => amountsMatch(existingRate, rate))) {
      uniqueRates.push(rate);
    }
  }

  return uniqueRates;
}

function inferReferenceExchangeRate(
  group: InvoiceGroup,
  reference: ReferenceInvoice,
): number | undefined {
  if (!isForeignCurrencyGroup(group)) {
    return undefined;
  }

  const htRate = inferExchangeRate(group.totalHT, reference.totalHT);
  const ttcRate = inferExchangeRate(group.totalTTC, reference.totalTTC);

  if (!htRate || !ttcRate) {
    return undefined;
  }

  const inferredRate = (htRate + ttcRate) / 2;

  if (!ratesAreClose(htRate, ttcRate, MAX_INFERRED_RATE_PAIR_DRIFT_RATIO)) {
    return undefined;
  }

  if (
    group.exchangeRate &&
    !ratesAreClose(
      inferredRate,
      group.exchangeRate,
      MAX_INFERRED_RATE_DRIFT_RATIO,
    )
  ) {
    return undefined;
  }

  return roundExchangeRate(inferredRate);
}

function isForeignCurrencyGroup(group: InvoiceGroup): boolean {
  return (
    Boolean(group.currencyCode && group.currencyCode !== "CDF") ||
    Boolean(group.exchangeRate && group.exchangeRate > 1)
  );
}

function inferExchangeRate(
  foreignAmount: number,
  cdfAmount: number,
): number | undefined {
  const normalizedForeignAmount = Math.abs(foreignAmount);
  const normalizedCdfAmount = Math.abs(cdfAmount);

  if (normalizedForeignAmount <= 0 || normalizedCdfAmount <= 0) {
    return undefined;
  }

  const rate = normalizedCdfAmount / normalizedForeignAmount;

  return Number.isFinite(rate) && rate > 1 ? rate : undefined;
}

function ratesAreClose(
  left: number,
  right: number,
  maxDriftRatio: number,
): boolean {
  const denominator = Math.max(Math.abs(left), Math.abs(right), 1);
  return Math.abs(left - right) / denominator <= maxDriftRatio;
}

function roundExchangeRate(rate: number): number {
  return Math.round(rate * 10000) / 10000;
}

function formatCount(count: number, singular: string, plural: string): string {
  return `${count} ${count > 1 ? plural : singular}`;
}

function createTaxRow(
  group: InvoiceGroup,
  taxColumn: string,
  invoiceColumn: string,
  dateColumn: string,
  dexyColumns: Record<DexyFieldKey, string>,
  normalizationDate: Date,
  exchangeRateUpdate: ConversionOptions["exchangeRateUpdate"],
  lineNumber: number,
  creditNoteContext: CreditNoteContext,
): ExcelRow {
  const baseRow = createTraceableRow(
    group.rows[0] ?? {},
    taxColumn,
    invoiceColumn,
    dateColumn,
    group.invoiceNumber,
    normalizationDate,
    exchangeRateUpdate,
    dexyColumns,
    creditNoteContext,
  );
  const taxTotal = roundCurrency(group.taxTotal);
  const unitPrice = creditNoteContext.isCreditNote ? Math.abs(taxTotal) : taxTotal;

  return {
    ...baseRow,
    [dexyColumns.lineNumber]: lineNumber,
    [dexyColumns.itemReference]: TAX_LINE_VALUES.itemReference,
    [dexyColumns.itemDescription]: TAX_LINE_VALUES.itemDescription,
    [dexyColumns.itemType]: TAX_LINE_VALUES.itemType,
    [dexyColumns.taxGroup]: TAX_LINE_VALUES.taxGroup,
    [dexyColumns.quantity]: TAX_LINE_VALUES.quantity,
    [dexyColumns.unitPrice]: unitPrice,
    [dexyColumns.unitPriceMode]: TAX_LINE_VALUES.unitPriceMode,
  };
}

function createTraceableRow(
  row: ExcelRow,
  taxColumn: string,
  invoiceColumn: string,
  dateColumn: string,
  invoiceNumber: string,
  normalizationDate: Date,
  exchangeRateUpdate?: ConversionOptions["exchangeRateUpdate"],
  dexyColumns?: Record<DexyFieldKey, string>,
  creditNoteContext?: CreditNoteContext,
): ExcelRow {
  const outputRow = stripColumn(row, taxColumn);
  const oldDate = row[dateColumn];
  const oldDateDisplay = formatDateForTrace(oldDate);
  const oldMonthYearDisplay = formatMonthYearForTrace(oldDate);
  const resolvedInvoiceNumber =
    invoiceNumber || normalizeInvoiceNumber(row[invoiceColumn]);

  outputRow[dateColumn] = normalizationDate;
  outputRow[DATE_FACTURE_COLUMN] = normalizationDate;
  outputRow[OLD_DATE_COLUMN] = oldDateDisplay;
  outputRow[COMMENT_A_COLUMN] = "";
  outputRow[COMMENT_B_COLUMN] = oldMonthYearDisplay;
  outputRow[COMMENT_C_COLUMN] = "Original Invoice :";
  outputRow[COMMENT_D_COLUMN] = oldDateDisplay;
  outputRow[COMMENT_E_COLUMN] =
    formatOriginalInvoiceReference(resolvedInvoiceNumber);
  applyDefaultValuesToExistingColumns(outputRow);
  applyNifRuleByClientTypology(outputRow);
  applyExchangeRateUpdate(outputRow, exchangeRateUpdate);
  applyCreditNoteValues(outputRow, dexyColumns, creditNoteContext);

  return outputRow;
}

function applyCreditNoteValues(
  row: ExcelRow,
  dexyColumns?: Record<DexyFieldKey, string>,
  creditNoteContext?: CreditNoteContext,
): void {
  if (!creditNoteContext?.isCreditNote) {
    return;
  }

  if (dexyColumns) {
    setPositiveNumericRowValue(row, dexyColumns.unitPrice);
  }

  if (!creditNoteContext.reference) {
    return;
  }

  setRowValue(
    row,
    CREDIT_NOTE_ORIGIN_DGI_COLUMN,
    creditNoteContext.reference.codeDefDgi,
  );
  setRowValue(
    row,
    CREDIT_NOTE_REFERENCE_TYPE_COLUMN,
    creditNoteContext.referenceType ?? "",
  );
  setRowValue(
    row,
    CREDIT_NOTE_REFERENCE_DESCRIPTION_COLUMN,
    creditNoteContext.referenceDescription ?? "",
  );
  setRowValue(
    row,
    CREDIT_NOTE_ORIGIN_ERP_INVOICE_COLUMN,
    creditNoteContext.reference.invoiceNumber,
  );
  setRowValue(
    row,
    CREDIT_NOTE_ORIGIN_ERP_DATE_COLUMN,
    creditNoteContext.reference.date,
  );
  applyReferenceExchangeRate(
    row,
    creditNoteContext.reference,
    creditNoteContext.inferredExchangeRate,
  );
}

function applyReferenceExchangeRate(
  row: ExcelRow,
  reference: ReferenceInvoice,
  inferredExchangeRate?: number,
): void {
  if (reference.currencyCode) {
    setRowValue(row, IMPORT_CURRENCY_COLUMN, reference.currencyCode);
  }

  const exchangeRate = reference.exchangeRate ?? inferredExchangeRate;

  if (!exchangeRate) {
    return;
  }

  setCandidateRowValue(
    row,
    B2F_EXCHANGE_RATE_COLUMN_CANDIDATES,
    exchangeRate,
  );
  setRowValue(row, IMPORT_EXCHANGE_RATE_COLUMN, exchangeRate);

  if (reference.exchangeRateDate) {
    setCandidateRowValue(
      row,
      B2F_EXCHANGE_RATE_DATE_COLUMN_CANDIDATES,
      reference.exchangeRateDate,
    );
  }
}

function setPositiveNumericRowValue(row: ExcelRow, columnName: string): void {
  const existingColumn = findExistingRowColumn(row, [columnName]) ?? columnName;
  const numericValue = parseAmount(row[existingColumn]);

  if (numericValue < 0) {
    row[existingColumn] = Math.abs(numericValue);
  }
}

function applyDefaultValuesToExistingColumns(row: ExcelRow): void {
  for (const { candidates, value } of DEFAULT_EXISTING_COLUMN_VALUES) {
    const existingColumn = Object.keys(row).find((column) =>
      candidates.some((candidate) => isSameColumn(column, candidate)),
    );

    if (existingColumn && isBlankValue(row[existingColumn])) {
      row[existingColumn] = value;
    }
  }
}

function formatOriginalInvoiceReference(invoiceNumber: string): string {
  const normalizedInvoiceNumber = invoiceNumber.trim();

  if (normalizedInvoiceNumber === "") {
    return "";
  }

  if (/^INV-DF-/i.test(normalizedInvoiceNumber)) {
    return normalizedInvoiceNumber;
  }

  return `INV-DF-${normalizedInvoiceNumber}`;
}

function applyNifRuleByClientTypology(row: ExcelRow): void {
  const typologyColumn = findExistingRowColumn(row, TYPOLOGY_COLUMN_CANDIDATES);

  if (!typologyColumn) {
    return;
  }

  const typology = formatCellAsText(row[typologyColumn]).toUpperCase();

  if (typology !== "PP") {
    return;
  }

  for (const nifColumn of findExistingRowColumns(row, NIF_CLIENT_COLUMN_CANDIDATES)) {
    row[nifColumn] = "";
  }
}

function findExistingRowColumn(
  row: ExcelRow,
  candidates: string[],
): string | undefined {
  return Object.keys(row).find((column) =>
    candidates.some((candidate) => isSameColumn(column, candidate)),
  );
}

function findExistingRowColumns(row: ExcelRow, candidates: string[]): string[] {
  return Object.keys(row).filter((column) =>
    candidates.some((candidate) => isSameColumn(column, candidate)),
  );
}

function formatCellAsText(value: ExcelCell): string {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function applyExchangeRateUpdate(
  row: ExcelRow,
  exchangeRateUpdate?: ConversionOptions["exchangeRateUpdate"],
): void {
  const currency = getRowCurrencyCode(row);

  if (currency === "CDF") {
    setRowValue(row, IMPORT_CURRENCY_COLUMN, "CDF");
    setRowValue(row, IMPORT_EXCHANGE_RATE_COLUMN, 1);
    setCandidateRowValue(row, B2F_EXCHANGE_RATE_COLUMN_CANDIDATES, 1);
    return;
  }

  if (!exchangeRateUpdate || currency !== exchangeRateUpdate.currency) {
    return;
  }

  setCandidateRowValue(row, B2F_EXCHANGE_RATE_DATE_COLUMN_CANDIDATES, exchangeRateUpdate.rateDate);
  setCandidateRowValue(row, B2F_EXCHANGE_RATE_COLUMN_CANDIDATES, exchangeRateUpdate.rate);
  setRowValue(row, IMPORT_CURRENCY_COLUMN, exchangeRateUpdate.currency);
  setRowValue(row, IMPORT_EXCHANGE_RATE_COLUMN, exchangeRateUpdate.rate);
}

function getRowCurrencyCode(row: ExcelRow): string {
  for (const candidate of B2F_CURRENCY_COLUMN_CANDIDATES) {
    const column = findExistingRowColumn(row, [candidate]);

    if (!column) {
      continue;
    }

    const currency = formatCellAsText(row[column]).toUpperCase();

    if (currency) {
      return currency;
    }
  }

  return "";
}

function setCandidateRowValue(
  row: ExcelRow,
  candidates: string[],
  value: ExcelCell,
): void {
  const column = findExistingRowColumn(row, candidates) ?? candidates[0];
  row[column] = value;
}

function setRowValue(row: ExcelRow, column: string, value: ExcelCell): void {
  const existingColumn = findExistingRowColumn(row, [column]) ?? column;
  row[existingColumn] = value;
}

function resolveDexyColumns(columns: string[]): Record<DexyFieldKey, string> {
  const resolvedColumns = Object.fromEntries(
    (Object.keys(DEXY_FIELD_CANDIDATES) as DexyFieldKey[]).map((key) => [
      key,
      findColumn(columns, DEXY_FIELD_CANDIDATES[key]) ??
        DEFAULT_DEXY_COLUMNS[key],
    ]),
  ) as Record<DexyFieldKey, string>;

  resolvedColumns.lineNumber =
    findColumn(columns, DEXY_FIELD_CANDIDATES.lineNumber) ??
    findLikelyLineNumberColumn(columns) ??
    DEFAULT_DEXY_COLUMNS.lineNumber;

  return resolvedColumns;
}

function buildOutputColumns(
  sourceColumns: string[],
  taxColumn: string,
  dateColumn: string,
  dexyColumns: Record<DexyFieldKey, string>,
): string[] {
  const outputColumns = sourceColumns.filter(
    (column) => column !== taxColumn && column !== OLD_DATE_COLUMN,
  );

  for (const column of Object.values(dexyColumns)) {
    if (!outputColumns.includes(column)) {
      outputColumns.push(column);
    }
  }

  for (const column of [
    dateColumn,
    DATE_FACTURE_COLUMN,
    COMMENT_A_COLUMN,
    COMMENT_B_COLUMN,
    COMMENT_C_COLUMN,
    COMMENT_D_COLUMN,
    COMMENT_E_COLUMN,
    COMMENT_F_COLUMN,
    COMMENT_G_COLUMN,
    COMMENT_H_COLUMN,
    CREDIT_NOTE_ORIGIN_ERP_INVOICE_COLUMN,
    CREDIT_NOTE_ORIGIN_ERP_DATE_COLUMN,
  ]) {
    if (!outputColumns.includes(column)) {
      outputColumns.push(column);
    }
  }

  return outputColumns;
}

function collectColumns(rows: ExcelRow[]): string[] {
  const columns: string[] = [];

  for (const row of rows) {
    for (const column of Object.keys(row)) {
      if (!columns.includes(column)) {
        columns.push(column);
      }
    }
  }

  return columns;
}

function normalizeColumns(columns: string[]): string[] {
  return columns
    .map((column) => column.trim())
    .filter((column, index, allColumns) => {
      return column !== "" && allColumns.indexOf(column) === index;
    });
}

function findColumn(columns: string[], candidates: string[]): string | undefined {
  const normalizedCandidates = candidates.map(normalizeHeader);

  return columns.find((column) =>
    normalizedCandidates.includes(normalizeHeader(column)),
  );
}

function summarizeDocumentTypes(
  rows: ExcelRow[],
  documentTypeColumn?: string,
): { hasSales: boolean; hasCreditNotes: boolean } {
  if (!documentTypeColumn) {
    return {
      hasSales: false,
      hasCreditNotes: false,
    };
  }

  let hasSales = false;
  let hasCreditNotes = false;

  for (const row of rows) {
    const documentType = normalizeDocumentType(row[documentTypeColumn]);

    if (documentType === SALES_DOCUMENT_TYPE) {
      hasSales = true;
    }

    if (documentType === CREDIT_NOTE_DOCUMENT_TYPE) {
      hasCreditNotes = true;
    }
  }

  return {
    hasSales,
    hasCreditNotes,
  };
}

function normalizeDocumentType(value: ExcelCell): string {
  return formatCellAsText(value).toUpperCase();
}

function normalizeCurrencyCode(value: ExcelCell): string {
  return formatCellAsText(value).toUpperCase();
}

function parseExchangeRate(value: ExcelCell): number | undefined {
  const parsedValue = parseAmount(value);
  return parsedValue > 0 ? parsedValue : undefined;
}

function findExciseTaxColumn(columns: string[]): string | undefined {
  const exactMatch = findColumn(columns, EXCISE_TAX_COLUMN_CANDIDATES);

  if (exactMatch) {
    return exactMatch;
  }

  const looseMatches = columns.filter((column) => {
    const normalizedColumn = normalizeHeader(column);
    const words = normalizedColumn.split(" ").filter(Boolean);
    const hasExciseTax = words.includes("excise") && words.includes("tax");
    const hasFinal = words.includes("final");
    const hasLocal = words.includes("local");
    const hasTelecomSignal =
      words.includes("telecommunication") ||
      words.includes("telecom") ||
      normalizedColumn.includes("congo drc");

    return hasExciseTax && (hasFinal || hasLocal || hasTelecomSignal);
  });

  if (looseMatches.length === 1) {
    return looseMatches[0];
  }

  return (
    looseMatches.find((column) =>
      normalizeHeader(column).split(" ").includes("final"),
    ) ?? looseMatches[0]
  );
}

function findLikelyLineNumberColumn(columns: string[]): string | undefined {
  const invoiceLikeColumns = new Set(
    INVOICE_COLUMN_CANDIDATES.map(normalizeHeader),
  );

  return columns.find((column) => {
    const normalizedColumn = normalizeHeader(column);
    const words = normalizedColumn.split(" ").filter(Boolean);
    const hasLineSignal =
      words.includes("ligne") ||
      words.includes("line") ||
      normalizedColumn.includes("ligne facture") ||
      normalizedColumn.includes("invoice line");
    const hasNumberSignal =
      words.includes("numero") ||
      words.includes("num") ||
      words.includes("no") ||
      words.includes("n") ||
      words.includes("number");

    return (
      hasLineSignal &&
      (hasNumberSignal || normalizedColumn.includes("ligne")) &&
      !invoiceLikeColumns.has(normalizedColumn)
    );
  });
}

function normalizeHeader(header: string): string {
  return header
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

function normalizeInvoiceNumber(value: ExcelCell): string {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  return String(value).trim();
}

function normalizeReferenceInvoiceKey(value: ExcelCell): string {
  return normalizeInvoiceNumber(value)
    .replace(/^INV-DF-/i, "")
    .replace(/-CN-\d+$/i, "")
    .trim()
    .toUpperCase();
}

function parseAmount(value: ExcelCell): number {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  if (typeof value !== "string") {
    return 0;
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
  return Number.isFinite(parsedValue) ? parsedValue : 0;
}

function formatDateForTrace(value: ExcelCell): string {
  const dateValue = parseDateValue(value);

  if (dateValue) {
    return formatDate(dateValue);
  }

  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function formatMonthYearForTrace(value: ExcelCell): string {
  const dateValue = parseDateValue(value);

  if (!dateValue) {
    return "";
  }

  const month = FRENCH_TRACE_MONTHS[dateValue.getMonth()];
  return `${month}-${dateValue.getFullYear()}`;
}

const FRENCH_TRACE_MONTHS = [
  "Janvier",
  "Février",
  "Mars",
  "Avril",
  "Mai",
  "Juin",
  "Juillet",
  "Août",
  "Septembre",
  "Octobre",
  "Novembre",
  "Décembre",
] as const;

function parseDateValue(value: ExcelCell): Date | null {
  if (value instanceof Date) {
    return toDateOnly(value);
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialToDate(value);
  }

  if (typeof value !== "string") {
    return null;
  }

  const dmyDate = parseDayMonthYearDate(value);

  if (dmyDate) {
    return dmyDate;
  }

  const parsedTimestamp = Date.parse(value);
  return Number.isNaN(parsedTimestamp)
    ? null
    : toDateOnly(new Date(parsedTimestamp));
}

function parseDayMonthYearDate(value: string): Date | null {
  const match = value.trim().match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);

  if (!match) {
    return null;
  }

  const day = Number.parseInt(match[1], 10);
  const month = Number.parseInt(match[2], 10);
  const yearInput = Number.parseInt(match[3], 10);
  const year = yearInput < 100 ? 2000 + yearInput : yearInput;

  if (
    !Number.isInteger(day) ||
    !Number.isInteger(month) ||
    !Number.isInteger(year) ||
    day < 1 ||
    day > 31 ||
    month < 1 ||
    month > 12
  ) {
    return null;
  }

  return toDateOnly(new Date(year, month - 1, day));
}

function excelSerialToDate(value: number): Date | null {
  if (value <= 0) {
    return null;
  }

  const milliseconds = Math.round((value - 25569) * 86400000);
  return toDateOnly(new Date(milliseconds));
}

function formatDate(value: Date): string {
  const day = String(value.getDate()).padStart(2, "0");
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const year = value.getFullYear();
  return `${day}/${month}/${year}`;
}

function toDateOnly(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

function roundCurrency(value: number): number {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function isSameColumn(left: unknown, right: unknown): boolean {
  return normalizeHeader(String(left)) === normalizeHeader(String(right));
}

function isBlankValue(value: ExcelCell): boolean {
  return value === null || value === undefined || value === "";
}

function stripColumn(row: ExcelRow, columnToRemove: string): ExcelRow {
  const { [columnToRemove]: _removedColumn, ...remainingColumns } = row;
  return { ...remainingColumns };
}
