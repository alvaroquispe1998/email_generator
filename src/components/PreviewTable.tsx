type PreviewTableProps = {
  columns: string[];
  rows: Array<Record<string, string | number>>;
  emptyState?: string;
};

export function PreviewTable({ columns, rows, emptyState }: PreviewTableProps) {
  if (columns.length === 0) {
    return (
      <div className="rounded-2xl border border-white/60 bg-white/60 p-4 text-sm text-ink/70">
        {emptyState ?? "Sin datos para mostrar."}
      </div>
    );
  }

  return (
    <div className="overflow-x-auto rounded-2xl border border-white/60 bg-white/70 shadow-sm">
      <table className="min-w-full text-sm">
        <thead className="bg-sand/80 text-left">
          <tr>
            {columns.map((column) => (
              <th
                key={column}
                className="whitespace-nowrap px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70"
              >
                {column}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, index) => (
            <tr key={index} className="border-t border-white/60">
              {columns.map((column) => {
                const value = row[column];
                const cellValue =
                  value === undefined || value === null || value === "" ? "-" : String(value);
                return (
                  <td key={column} className="whitespace-nowrap px-4 py-2 text-ink/90">
                    {cellValue}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
