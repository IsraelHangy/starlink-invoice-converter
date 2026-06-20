import type { ExcelCell, ExcelRow } from "../types";

interface PreviewTableProps {
  title: string;
  rows: ExcelRow[];
  columns: string[];
  emptyMessage: string;
}

const MAX_PREVIEW_ROWS = 30;

export default function PreviewTable({
  title,
  rows,
  columns,
  emptyMessage,
}: PreviewTableProps) {
  const visibleRows = rows.slice(0, MAX_PREVIEW_ROWS);
  const visibleColumns = columns;
  const hiddenRows = Math.max(rows.length - visibleRows.length, 0);
  const hiddenColumns = 0;

  return (
    <section className="rounded-lg border border-slate-200 bg-white shadow-soft">
      <div className="flex flex-col gap-2 border-b border-slate-200 px-5 py-4 sm:flex-row sm:items-center sm:justify-between">
        <h2 className="text-lg font-semibold text-ink">{title}</h2>
        {rows.length > 0 ? (
          <p className="rounded-md bg-slate-100 px-3 py-1 text-sm font-medium text-slate-600">
            {rows.length} lignes, {columns.length} colonnes
          </p>
        ) : null}
      </div>

      {rows.length === 0 ? (
        <div className="px-5 py-12 text-center text-sm text-slate-500">
          {emptyMessage}
        </div>
      ) : (
        <>
          <div className="max-h-[520px] overflow-auto">
            <table className="w-max min-w-full border-collapse text-left text-sm">
              <thead className="sticky top-0 z-10 bg-ink text-xs uppercase text-white">
                <tr>
                  {visibleColumns.map((column) => (
                    <th
                      key={column}
                      className="min-w-36 border-b border-r border-slate-700 px-3 py-3 font-semibold"
                    >
                      <span className="line-clamp-2">{column}</span>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {visibleRows.map((row, rowIndex) => (
                  <tr
                    key={`${rowIndex}-${String(row[visibleColumns[0]] ?? "")}`}
                    className={rowIndex % 2 === 0 ? "bg-white" : "bg-[#f8fafc]"}
                  >
                    {visibleColumns.map((column) => (
                      <td
                        key={column}
                        className="max-w-64 border-b border-r border-slate-100 px-3 py-2 align-top text-slate-700"
                      >
                        <span className="line-clamp-3 break-words">
                          {formatCell(row[column])}
                        </span>
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {(hiddenRows > 0 || hiddenColumns > 0) && (
            <p className="border-t border-slate-200 px-5 py-3 text-sm text-slate-500">
              Aperçu limité : {hiddenRows} lignes et {hiddenColumns} colonnes
              supplémentaires ne sont pas affichées.
            </p>
          )}
        </>
      )}
    </section>
  );
}

function formatCell(value: ExcelCell): string {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    return value.toLocaleDateString("fr-FR");
  }

  return String(value);
}
