import type { MappingConfig, MappingRule } from "@/lib/csv";

type MappingEditorProps = {
  outlookHeaders: string[];
  columns: string[];
  mapping: MappingConfig;
  onChange: (next: MappingConfig) => void;
  onReset: () => void;
};

const GENERATED_LABELS: Record<string, { label: string; value: string }> = {
  "Nombre de usuario": {
    label: "Generado: usuario institucional",
    value: "username"
  },
  "Nombre para mostrar": {
    label: "Generado: Apellido + Nombre",
    value: "displayName"
  }
};

function getRule(mapping: MappingConfig, header: string): MappingRule {
  return mapping[header] ?? { type: "fixed", value: "" };
}

export function MappingEditor({
  outlookHeaders,
  columns,
  mapping,
  onChange,
  onReset
}: MappingEditorProps) {
  const updateRule = (header: string, rule: MappingRule) => {
    onChange({ ...mapping, [header]: rule });
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h2 className="text-xl font-semibold text-ink">Configuracion de mapeo</h2>
          <p className="text-sm text-ink/70">
            Asigna cada columna de Outlook a un campo del Excel o a un valor fijo.
          </p>
        </div>
        <button
          type="button"
          onClick={onReset}
          className="rounded-full border border-ink/20 bg-white/70 px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40"
        >
          Restaurar defaults
        </button>
      </div>

      <div className="grid gap-3">
        {outlookHeaders.map((header) => {
          const rule = getRule(mapping, header);
          const generated = GENERATED_LABELS[header];
          const canGenerate = Boolean(generated);

          return (
            <div
              key={header}
              className="grid gap-2 rounded-2xl border border-white/60 bg-white/70 p-3 sm:grid-cols-[240px_160px_1fr]"
            >
              <div className="text-sm font-semibold text-ink">{header}</div>

              <select
                value={rule.type}
                onChange={(event) => {
                  const nextType = event.target.value as MappingRule["type"];
                  if (nextType === "generated" && generated) {
                    updateRule(header, { type: "generated", value: generated.value });
                    return;
                  }
                  if (nextType === "column") {
                    updateRule(header, { type: "column", value: columns[0] ?? "" });
                    return;
                  }
                  updateRule(header, { type: "fixed", value: "" });
                }}
                className="rounded-xl border border-ink/20 bg-white px-3 py-2 text-sm"
              >
                <option value="column">Columna</option>
                <option value="fixed">Valor fijo</option>
                {canGenerate && <option value="generated">Generado</option>}
              </select>

              <div className="flex items-center gap-2">
                {rule.type === "column" && (
                  <select
                    value={rule.value}
                    onChange={(event) =>
                      updateRule(header, { type: "column", value: event.target.value })
                    }
                    className="w-full rounded-xl border border-ink/20 bg-white px-3 py-2 text-sm"
                  >
                    <option value="">Sin asignar</option>
                    {columns.map((column) => (
                      <option key={column} value={column}>
                        {column}
                      </option>
                    ))}
                  </select>
                )}

                {rule.type === "fixed" && (
                  <input
                    value={rule.value}
                    onChange={(event) =>
                      updateRule(header, { type: "fixed", value: event.target.value })
                    }
                    className="w-full rounded-xl border border-ink/20 bg-white px-3 py-2 text-sm"
                    placeholder="Escribe un valor fijo..."
                  />
                )}

                {rule.type === "generated" && generated && (
                  <div className="text-sm text-ink/70">{generated.label}</div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
