import { CheckCircle2, ReceiptText, Sigma, Table2 } from "lucide-react";
import type { ReactNode } from "react";
import type { ConversionSummary } from "../types";

interface SummaryCardProps {
  summary?: ConversionSummary;
  warnings: string[];
}

export default function SummaryCard({ summary, warnings }: SummaryCardProps) {
  return (
    <section className="rounded-lg border border-slate-200 bg-white p-5 shadow-soft">
      <div className="flex items-start justify-between gap-4">
        <div className="flex min-w-0 items-center gap-3">
        <div className="flex h-11 w-11 items-center justify-center rounded-lg bg-[#b0003a] text-white shadow-sm">
          <ReceiptText aria-hidden="true" size={22} />
        </div>
        <div className="min-w-0">
          <h2 className="text-lg font-semibold text-ink">Résumé DEXY</h2>
          <p className="mt-1 text-sm text-slate-600">
            Totaux calculés après conversion.
          </p>
        </div>
        </div>
        {summary ? (
          <span className="inline-flex shrink-0 items-center gap-1.5 rounded-md bg-[#ecfdf5] px-2.5 py-1 text-xs font-semibold text-[#047857]">
            <CheckCircle2 aria-hidden="true" size={14} />
            Prêt
          </span>
        ) : null}
      </div>

      <div className="mt-5 grid grid-cols-2 gap-3">
        <Metric
          label="Factures"
          value={summary ? summary.invoicesProcessed : "-"}
          icon={<ReceiptText aria-hidden="true" size={18} />}
        />
        <Metric
          label="Lignes source"
          value={summary ? summary.originalRows : "-"}
          icon={<Table2 aria-hidden="true" size={18} />}
        />
        <Metric
          label="Lignes TAX"
          value={summary ? summary.taxRowsAdded : "-"}
          icon={<Sigma aria-hidden="true" size={18} />}
        />
        <Metric
          label="Total taxe"
          value={summary ? formatAmount(summary.totalTax) : "-"}
          icon={<Sigma aria-hidden="true" size={18} />}
        />
      </div>

      {warnings.length > 0 ? (
        <div className="mt-4 rounded-lg border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800">
          {warnings.map((warning) => (
            <p key={warning}>{warning}</p>
          ))}
        </div>
      ) : null}
    </section>
  );
}

interface MetricProps {
  label: string;
  value: string | number;
  icon: ReactNode;
}

function Metric({ label, value, icon }: MetricProps) {
  return (
    <div className="rounded-lg border border-slate-200 bg-[#f8fafc] p-4">
      <div className="flex items-center justify-between gap-3 text-slate-500">
        <span className="text-xs font-semibold uppercase tracking-wide">{label}</span>
        <span className="text-[#1298e8]">{icon}</span>
      </div>
      <p className="mt-2 break-words text-xl font-bold text-ink">{value}</p>
    </div>
  );
}

function formatAmount(value: number): string {
  return new Intl.NumberFormat("fr-FR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}
