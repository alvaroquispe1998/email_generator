"use client";

import { useEffect, useMemo, useState } from "react";
import Papa from "papaparse";
import { MappingEditor } from "@/components/MappingEditor";
import { PreviewTable } from "@/components/PreviewTable";
import {
  OUTLOOK_HEADERS,
  buildGeneratedUsername,
  buildOutputRow,
  buildOutputRows,
  type MappingConfig
} from "@/lib/csv";
import {
  buildUsernameWithSecondName,
  digitsOnly,
  stripAccents,
  toCleanString
} from "@/lib/normalization";
import { parseWorkbook, type DataRow, type WorkbookData } from "@/lib/xlsx";

const STORAGE_MAPPING = "outlook-mapping-v1";
const STORAGE_REQUIRED = "outlook-required-v1";
const EMPTY_SET = new Set<string>();

type RequiredConfig = {
  dni: boolean;
  celular: boolean;
  codigo: boolean;
};

type ValidationIssue = {
  rowNumber: number;
  missing: string[];
};

type OutlookDirectory = {
  fileName: string;
  rowCount: number;
  dnis: Set<string>;
  emails: Set<string>;
  error: string | null;
};

type ExistingDniMatch = {
  rowNumber: number;
  dni: string;
  nombre: string;
  apellido: string;
};

type EmailConflictRow = {
  rowNumber: number;
  nombre: string;
  apellido: string;
  generatedEmail: string;
};

const DEFAULT_REQUIRED: RequiredConfig = {
  dni: true,
  celular: true,
  codigo: true
};

const GENERATED_BY_HEADER: Record<string, string> = {
  "Nombre de usuario": "username",
  "Nombre para mostrar": "displayName"
};

function normalizeKey(value: string): string {
  return stripAccents(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function normalizeValue(value: string): string {
  return stripAccents(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function normalizeEmail(value: string): string {
  const cleaned = stripAccents(toCleanString(value)).toLowerCase();
  if (!cleaned) {
    return "";
  }
  const [localRaw, domainRaw] = cleaned.split("@");
  const local = (localRaw || "").replace(/[^a-z0-9.]/g, "");
  const domain = (domainRaw || "").replace(/[^a-z0-9.]/g, "");
  if (!domainRaw) {
    return local;
  }
  return `${local}@${domain}`;
}

function findColumn(columns: string[], candidates: string[]): string {
  const normalizedCandidates = candidates.map(normalizeKey);
  const found = columns.find((column) => {
    const normalizedColumn = normalizeKey(column);
    return normalizedCandidates.some((candidate) =>
      normalizedColumn.includes(candidate)
    );
  });
  return found ?? "";
}

function buildDefaultMapping(columns: string[]): MappingConfig {
  const nombre = findColumn(columns, ["NOMBRES", "NOMBRES COMPLETOS", "NOMBRE"]);
  const apellido = findColumn(columns, ["APELLIDOS", "APELLIDOS COMPLETOS", "APELLIDO"]);
  const celular = findColumn(columns, ["NUMERO DE CELULAR", "CELULAR", "TELEFONO MOVIL"]);
  const dni = findColumn(columns, ["DNI", "DOCUMENTO"]);
  const codigo = findColumn(columns, ["CODIGO DE ESTUDIANTE", "CODIGO ESTUDIANTE", "CODIGO"]);
  const correo = findColumn(columns, ["CORREO PERSONAL", "EMAIL PERSONAL", "MAIL"]);

  return {
    "Nombre de usuario": { type: "generated", value: GENERATED_BY_HEADER["Nombre de usuario"] },
    Nombre: { type: "column", value: nombre },
    Apellido: { type: "column", value: apellido },
    "Nombre para mostrar": { type: "generated", value: GENERATED_BY_HEADER["Nombre para mostrar"] },
    Puesto: { type: "fixed", value: "Estudiante" },
    Departamento: { type: "fixed", value: "" },
    "Número del trabajo": { type: "fixed", value: "" },
    "Teléfono de la oficina": { type: "fixed", value: "" },
    "Teléfono móvil": { type: "column", value: celular },
    Fax: { type: "column", value: dni },
    "Dirección de correo electrónico alternativa": { type: "column", value: correo },
    Dirección: { type: "fixed", value: "" },
    Ciudad: { type: "fixed", value: "" },
    "Estado o provincia": { type: "fixed", value: "" },
    "Código postal": { type: "column", value: codigo },
    "País o región": { type: "fixed", value: "Peru" }
  };
}

function getRowNameParts(row: DataRow, mapping: MappingConfig): { nombre: string; apellido: string } {
  const output = buildOutputRow(row, mapping);
  return {
    nombre: toCleanString(output["Nombre"]),
    apellido: toCleanString(output["Apellido"])
  };
}

function getRowDni(row: DataRow, mapping: MappingConfig): string {
  const output = buildOutputRow(row, mapping);
  return digitsOnly(toCleanString(output["Fax"]));
}

function getGeneratedEmail(row: DataRow, mapping: MappingConfig): string {
  return buildGeneratedUsername(row, mapping);
}

function getAlternateEmail(nombre: string, apellido: string): string {
  return buildUsernameWithSecondName(nombre, apellido);
}

function rowHasIngresoCondition(row: DataRow, conditionColumn: string): boolean {
  if (!conditionColumn) {
    return true;
  }
  const value = toCleanString(row[conditionColumn]);
  return normalizeValue(value) === "INGRESO";
}

function mergeMapping(defaults: MappingConfig, stored?: MappingConfig | null): MappingConfig {
  if (!stored) {
    return defaults;
  }
  const merged: MappingConfig = { ...defaults };
  OUTLOOK_HEADERS.forEach((header) => {
    if (stored[header]) {
      merged[header] = stored[header];
    }
  });
  return merged;
}

function sanitizeMapping(
  mapping: MappingConfig,
  columns: string[],
  defaults: MappingConfig
): MappingConfig {
  const sanitized: MappingConfig = { ...defaults };
  OUTLOOK_HEADERS.forEach((header) => {
    const rule = mapping[header];
    if (!rule) {
      return;
    }
    if (rule.type === "column") {
      sanitized[header] = columns.includes(rule.value)
        ? rule
        : { type: "column", value: "" };
      return;
    }
    if (rule.type === "generated") {
      const allowed = GENERATED_BY_HEADER[header];
      sanitized[header] = allowed && rule.value === allowed ? rule : defaults[header];
      return;
    }
    if (rule.type === "fixed") {
      sanitized[header] = rule;
    }
  });
  return sanitized;
}

function validateRows(
  rows: DataRow[],
  mapping: MappingConfig,
  required: RequiredConfig,
  conditionColumn: string,
  existingDnis: Set<string>
): ValidationIssue[] {
  return rows
    .map((row) => {
      if (!rowHasIngresoCondition(row, conditionColumn)) {
        return null;
      }
      const output = buildOutputRow(row, mapping);
      const dniValue = digitsOnly(toCleanString(output["Fax"]));
      if (dniValue && existingDnis.has(dniValue)) {
        return null;
      }
      const missing: string[] = [];

      if (required.dni && !toCleanString(output["Fax"])) {
        missing.push("DNI");
      }
      if (required.celular && !toCleanString(output["Teléfono móvil"])) {
        missing.push("Celular");
      }
      if (required.codigo && !toCleanString(output["Código postal"])) {
        missing.push("Código estudiante");
      }

      if (missing.length === 0) {
        return null;
      }
      return { rowNumber: row.__rowNumber, missing };
    })
    .filter(Boolean) as ValidationIssue[];
}

function rowMeetsRequired(
  row: DataRow,
  mapping: MappingConfig,
  required: RequiredConfig,
  conditionColumn: string,
  existingDnis: Set<string>
): boolean {
  if (!rowHasIngresoCondition(row, conditionColumn)) {
    return false;
  }
  const output = buildOutputRow(row, mapping);
  const dniValue = digitsOnly(toCleanString(output["Fax"]));
  if (dniValue && existingDnis.has(dniValue)) {
    return false;
  }

  if (required.dni && !toCleanString(output["Fax"])) {
    return false;
  }
  if (required.celular && !toCleanString(output["Teléfono móvil"])) {
    return false;
  }
  if (required.codigo && !toCleanString(output["Código postal"])) {
    return false;
  }

  return true;
}

function filterRowsForExport(
  rows: DataRow[],
  mapping: MappingConfig,
  required: RequiredConfig,
  conditionColumn: string,
  existingDnis: Set<string>
): DataRow[] {
  return rows.filter((row) =>
    rowMeetsRequired(row, mapping, required, conditionColumn, existingDnis)
  );
}

function filterRowsByEmail(
  rows: DataRow[],
  mapping: MappingConfig,
  emailOverrides: Record<string, string>,
  existingEmails: Set<string>,
  plannedEmailCounts: Map<string, number>
): DataRow[] {
  const seen = new Set<string>();
  return rows.filter((row) => {
    const key = String(row.__rowNumber);
    const override = emailOverrides[key];
    const email = normalizeEmail(override || getGeneratedEmail(row, mapping));
    if (!email) {
      return true;
    }
    if (existingEmails.has(email)) {
      return false;
    }
    const count = plannedEmailCounts.get(email) ?? 0;
    if (count > 1) {
      if (seen.has(email)) {
        return false;
      }
      seen.add(email);
    }
    return true;
  });
}

export default function HomePage() {
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [columns, setColumns] = useState<string[]>([]);
  const [rows, setRows] = useState<DataRow[]>([]);
  const [mapping, setMapping] = useState<MappingConfig>(() => buildDefaultMapping([]));
  const [required, setRequired] = useState<RequiredConfig>(DEFAULT_REQUIRED);
  const [csvPreview, setCsvPreview] = useState<Record<string, string>[]>([]);
  const [csvText, setCsvText] = useState<string>("");
  const [generatedRows, setGeneratedRows] = useState<Record<string, string>[]>([]);
  const [isParsing, setIsParsing] = useState(false);
  const [storedMapping, setStoredMapping] = useState<MappingConfig | null>(null);
  const [hasUserMapping, setHasUserMapping] = useState(false);
  const [outlookDirectory, setOutlookDirectory] = useState<OutlookDirectory | null>(null);
  const [emailOverrides, setEmailOverrides] = useState<Record<string, string>>({});

  useEffect(() => {
    const rawMapping = localStorage.getItem(STORAGE_MAPPING);
    if (rawMapping) {
      try {
        setStoredMapping(JSON.parse(rawMapping));
      } catch {
        setStoredMapping(null);
      }
    }
    const rawRequired = localStorage.getItem(STORAGE_REQUIRED);
    if (rawRequired) {
      try {
        setRequired({ ...DEFAULT_REQUIRED, ...JSON.parse(rawRequired) });
      } catch {
        setRequired(DEFAULT_REQUIRED);
      }
    }
  }, []);

  useEffect(() => {
    if (workbook && sheetName) {
      const sheet =
        workbook.sheets.find((item) => item.name === sheetName) ??
        workbook.sheets[0];
      setColumns(sheet?.columns ?? []);
      setRows(sheet?.rows ?? []);
    } else {
      setColumns([]);
      setRows([]);
    }
  }, [workbook, sheetName]);

  useEffect(() => {
    if (!columns.length) {
      return;
    }
    const defaults = buildDefaultMapping(columns);
    if (!hasUserMapping) {
      const base = storedMapping ? mergeMapping(defaults, storedMapping) : defaults;
      setMapping(sanitizeMapping(base, columns, defaults));
      return;
    }
    setMapping((prev) => sanitizeMapping(prev ?? defaults, columns, defaults));
  }, [columns, storedMapping, hasUserMapping]);

  useEffect(() => {
    localStorage.setItem(STORAGE_MAPPING, JSON.stringify(mapping));
  }, [mapping]);

  useEffect(() => {
    localStorage.setItem(STORAGE_REQUIRED, JSON.stringify(required));
  }, [required]);

  const conditionColumn = useMemo(
    () => findColumn(columns, ["CONDICION", "CONDICIÓN"]),
    [columns]
  );
  const existingDnis = outlookDirectory?.dnis ?? EMPTY_SET;
  const existingEmails = outlookDirectory?.emails ?? EMPTY_SET;

  const validationIssues = useMemo(
    () => validateRows(rows, mapping, required, conditionColumn, existingDnis),
    [rows, mapping, required, conditionColumn, existingDnis]
  );
  const baseExportRows = useMemo(
    () => filterRowsForExport(rows, mapping, required, conditionColumn, existingDnis),
    [rows, mapping, required, conditionColumn, existingDnis]
  );
  const totalCsvParts = useMemo(() => {
    if (!generatedRows.length) {
      return 0;
    }
    return Math.ceil(generatedRows.length / 249);
  }, [generatedRows.length]);
  const existingDniMatches = useMemo(() => {
    if (!existingDnis.size) {
      return [] as ExistingDniMatch[];
    }
    return rows
      .filter((row) => rowHasIngresoCondition(row, conditionColumn))
      .map((row) => {
        const dni = getRowDni(row, mapping);
        if (!dni || !existingDnis.has(dni)) {
          return null;
        }
        const { nombre, apellido } = getRowNameParts(row, mapping);
        return {
          rowNumber: row.__rowNumber,
          dni,
          nombre,
          apellido
        };
      })
      .filter(Boolean) as ExistingDniMatch[];
  }, [rows, mapping, conditionColumn, existingDnis]);
  const emailConflictRows = useMemo(() => {
    if (!existingEmails.size) {
      return [] as EmailConflictRow[];
    }
    return baseExportRows
      .map((row) => {
        const key = String(row.__rowNumber);
        const override = emailOverrides[key];
        const generatedEmail = normalizeEmail(getGeneratedEmail(row, mapping));
        const effectiveEmail = normalizeEmail(override || generatedEmail);
        if (!effectiveEmail || !existingEmails.has(effectiveEmail)) {
          return null;
        }
        const { nombre, apellido } = getRowNameParts(row, mapping);
        return {
          rowNumber: row.__rowNumber,
          nombre,
          apellido,
          generatedEmail
        };
      })
      .filter(Boolean) as EmailConflictRow[];
  }, [baseExportRows, mapping, emailOverrides, existingEmails]);
  const plannedEmailCounts = useMemo(() => {
    const counts = new Map<string, number>();
    baseExportRows.forEach((row) => {
      const key = String(row.__rowNumber);
      const override = emailOverrides[key];
      const email = normalizeEmail(override || getGeneratedEmail(row, mapping));
      if (!email) {
        return;
      }
      counts.set(email, (counts.get(email) ?? 0) + 1);
    });
    return counts;
  }, [baseExportRows, emailOverrides, mapping]);
  const exportRows = useMemo(
    () =>
      filterRowsByEmail(
        baseExportRows,
        mapping,
        emailOverrides,
        existingEmails,
        plannedEmailCounts
      ),
    [baseExportRows, mapping, emailOverrides, existingEmails, plannedEmailCounts]
  );

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    setIsParsing(true);
    setCsvPreview([]);
    setCsvText("");
    setHasUserMapping(false);
    setEmailOverrides({});
    setGeneratedRows([]);
    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseWorkbook(buffer);
      setWorkbook(parsed);
      setSheetName(parsed.sheets[0]?.name ?? "");
    } finally {
      setIsParsing(false);
    }
  };

  const handleOutlookCsvChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    try {
      const text = await file.text();
      const parsed = Papa.parse<Record<string, string>>(text, {
        header: true,
        skipEmptyLines: true
      });
      const fields = parsed.meta.fields ?? [];
      const faxField = findColumn(fields, ["Fax", "FAX"]);
      const upnField = findColumn(fields, [
        "User principal name",
        "UserPrincipalName",
        "User principalname"
      ]);

      if (!faxField || !upnField) {
        setOutlookDirectory({
          fileName: file.name,
          rowCount: parsed.data.length,
          dnis: new Set(),
          emails: new Set(),
          error: "No se encontraron las columnas Fax y User principal name en el CSV."
        });
        return;
      }

      const dnis = new Set<string>();
      const emails = new Set<string>();
      parsed.data.forEach((row) => {
        const dni = digitsOnly(toCleanString(row[faxField] ?? ""));
        if (dni) {
          dnis.add(dni);
        }
        const email = normalizeEmail(row[upnField] ?? "");
        if (email) {
          emails.add(email);
        }
      });

      setOutlookDirectory({
        fileName: file.name,
        rowCount: parsed.data.length,
        dnis,
        emails,
        error: null
      });
    } catch {
      setOutlookDirectory({
        fileName: file.name,
        rowCount: 0,
        dnis: new Set(),
        emails: new Set(),
        error: "No se pudo leer el CSV de Outlook."
      });
    }
  };

  const handleGenerate = () => {
    const outputRows = buildOutputRows(exportRows, mapping).map((row, index) => {
      const sourceRow = exportRows[index];
      const key = String(sourceRow.__rowNumber);
      const override = emailOverrides[key];
      if (override && toCleanString(override)) {
        row["Nombre de usuario"] = normalizeEmail(override);
      }
      return row;
    });
    setGeneratedRows(outputRows);
    setCsvPreview(outputRows.slice(0, 20));
    const csvBody = Papa.unparse(
      {
        fields: OUTLOOK_HEADERS,
        data: outputRows.map((row) => OUTLOOK_HEADERS.map((header) => row[header] ?? ""))
      },
      {
        delimiter: ",",
        newline: "\r\n"
      }
    );
    setCsvText(`\ufeff${csvBody}`);
  };

  const handleDownload = () => {
    if (!generatedRows.length) {
      return;
    }
    const chunkSize = 249;
    const totalParts = Math.ceil(generatedRows.length / chunkSize);

    for (let index = 0; index < totalParts; index += 1) {
      const start = index * chunkSize;
      const chunk = generatedRows.slice(start, start + chunkSize);
      const csvBody = Papa.unparse(
        {
          fields: OUTLOOK_HEADERS,
          data: chunk.map((row) => OUTLOOK_HEADERS.map((header) => row[header] ?? ""))
        },
        {
          delimiter: ",",
          newline: "\r\n"
        }
      );
      const csvTextChunk = `\ufeff${csvBody}`;
      const blob = new Blob([csvTextChunk], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download =
        totalParts === 1
          ? "contactos_outlook.csv"
          : `contactos_outlook_parte_${index + 1}.csv`;
      window.setTimeout(() => {
        anchor.click();
        URL.revokeObjectURL(url);
      }, index * 150);
    }
  };

  const handleReset = () => {
    const defaults = buildDefaultMapping(columns);
    setMapping(defaults);
    setRequired(DEFAULT_REQUIRED);
    setHasUserMapping(true);
  };

  return (
    <main className="px-4 py-10">
      <div className="mx-auto flex max-w-6xl flex-col gap-10">
        <header className="space-y-3">
          <span className="inline-flex w-fit items-center rounded-full border border-ink/15 bg-white/60 px-3 py-1 text-xs font-semibold uppercase tracking-[0.2em] text-ink/70">
            Outlook CSV - UAI
          </span>
          <h1 className="text-3xl font-semibold text-ink sm:text-4xl">
            Generador de contactos para Outlook
          </h1>
          <p className="max-w-2xl text-base text-ink/75">
            Carga un Excel, configura el mapeo y descarga un CSV listo para importar
            en Outlook (español).
          </p>
        </header>

        <section className="rounded-3xl border border-white/60 bg-white/70 p-6 shadow-sm">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-xl font-semibold text-ink">1. Cargar XLSX</h2>
              <p className="text-sm text-ink/70">
                Selecciona el archivo de informe con datos de estudiantes.
              </p>
            </div>
            <label className="inline-flex cursor-pointer items-center gap-2 rounded-full border border-ink/20 bg-white px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40">
              <input
                type="file"
                accept=".xlsx"
                className="hidden"
                onChange={handleFileChange}
              />
              Subir XLSX
            </label>
          </div>

          {isParsing && (
            <div className="mt-4 rounded-2xl border border-ink/10 bg-sand/70 px-4 py-3 text-sm text-ink/70">
              Procesando archivo...
            </div>
          )}

          {workbook && workbook.sheets.length > 1 && (
            <div className="mt-4 flex flex-col gap-2 sm:flex-row sm:items-center">
              <span className="text-sm font-semibold text-ink">Hoja:</span>
              <select
                value={sheetName}
                onChange={(event) => setSheetName(event.target.value)}
                className="rounded-xl border border-ink/20 bg-white px-3 py-2 text-sm"
              >
                {workbook.sheets.map((sheet) => (
                  <option key={sheet.name} value={sheet.name}>
                    {sheet.name}
                  </option>
                ))}
              </select>
            </div>
          )}

          <div className="mt-6">
            <h3 className="text-base font-semibold text-ink">Preview XLSX (primeras 20 filas)</h3>
            <div className="mt-3">
              <PreviewTable
                columns={columns}
                rows={rows.slice(0, 20)}
                emptyState="Carga un XLSX para visualizar datos."
              />
            </div>
          </div>
        </section>

        <section className="rounded-3xl border border-white/60 bg-white/70 p-6 shadow-sm">
          <MappingEditor
            outlookHeaders={OUTLOOK_HEADERS}
            columns={columns}
            mapping={mapping}
            onChange={(next) => {
              setHasUserMapping(true);
              setMapping(next);
            }}
            onReset={handleReset}
          />
        </section>

        <section className="rounded-3xl border border-white/60 bg-white/70 p-6 shadow-sm">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-xl font-semibold text-ink">2. Validar con CSV de Outlook</h2>
              <p className="text-sm text-ink/70">
                Sube el CSV exportado de Outlook para detectar DNI y correos existentes.
              </p>
            </div>
            <label className="inline-flex cursor-pointer items-center gap-2 rounded-full border border-ink/20 bg-white px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40">
              <input
                type="file"
                accept=".csv"
                className="hidden"
                onChange={handleOutlookCsvChange}
              />
              Subir CSV Outlook
            </label>
          </div>

          {!outlookDirectory && (
            <div className="mt-4 rounded-2xl border border-ink/10 bg-sand/70 px-4 py-3 text-sm text-ink/70">
              Aún no has cargado el CSV de Outlook. Esta validación es opcional.
            </div>
          )}

          {outlookDirectory && (
            <div className="mt-4 space-y-3">
              <div className="rounded-2xl border border-ink/10 bg-sand/70 px-4 py-3 text-sm text-ink/80">
                <div>Archivo: {outlookDirectory.fileName}</div>
                <div>Registros leídos: {outlookDirectory.rowCount}</div>
                <div>DNI detectados: {outlookDirectory.dnis.size}</div>
                <div>Correos detectados: {outlookDirectory.emails.size}</div>
              </div>
              {outlookDirectory.error && (
                <div className="rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
                  {outlookDirectory.error}
                </div>
              )}
            </div>
          )}

          {outlookDirectory && !outlookDirectory.error && (
            <div className="mt-6 space-y-6">
              <div>
                <h3 className="text-base font-semibold text-ink">
                  Personal ya existente (por DNI)
                </h3>
                <p className="text-sm text-ink/70">
                  Estas filas no se contarán en el CSV final.
                </p>
                <div className="mt-3">
                  {existingDniMatches.length === 0 ? (
                    <div className="rounded-2xl border border-ink/10 bg-white/70 px-4 py-3 text-sm text-ink/70">
                      No se encontraron coincidencias de DNI.
                    </div>
                  ) : (
                    <div className="overflow-x-auto rounded-2xl border border-white/60 bg-white/70">
                      <table className="min-w-full text-sm">
                        <thead className="bg-sand/80 text-left">
                          <tr>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              Fila
                            </th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              DNI
                            </th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              Apellido
                            </th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              Nombre
                            </th>
                          </tr>
                        </thead>
                        <tbody>
                          {existingDniMatches.map((match) => (
                            <tr key={match.rowNumber} className="border-t border-white/60">
                              <td className="px-4 py-2 text-ink/90">{match.rowNumber}</td>
                              <td className="px-4 py-2 text-ink/90">{match.dni}</td>
                              <td className="px-4 py-2 text-ink/90">{match.apellido || "-"}</td>
                              <td className="px-4 py-2 text-ink/90">{match.nombre || "-"}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </div>

              <div>
                <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <h3 className="text-base font-semibold text-ink">Correos ya usados</h3>
                    <p className="text-sm text-ink/70">
                      Edita el correo si el generado ya existe. Se valida disponibilidad.
                    </p>
                  </div>
                  {emailConflictRows.length > 0 && (
                    <button
                      type="button"
                      onClick={() => {
                        setEmailOverrides((prev) => {
                          const updated = { ...prev };
                          emailConflictRows.forEach((conflict) => {
                            const key = String(conflict.rowNumber);
                            if (updated[key]) {
                              return;
                            }
                            const alternate = getAlternateEmail(conflict.nombre, conflict.apellido);
                            if (!alternate) {
                              return;
                            }
                            if (normalizeEmail(alternate) === conflict.generatedEmail) {
                              return;
                            }
                            updated[key] = alternate;
                          });
                          return updated;
                        });
                      }}
                      className="rounded-full border border-ink/20 bg-white/80 px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40"
                    >
                      Usar segundo nombre + primer apellido
                    </button>
                  )}
                </div>
                <div className="mt-3">
                  {emailConflictRows.length === 0 ? (
                    <div className="rounded-2xl border border-ink/10 bg-white/70 px-4 py-3 text-sm text-ink/70">
                      No hay conflictos con correos existentes.
                    </div>
                  ) : (
                    <div className="overflow-x-auto rounded-2xl border border-white/60 bg-white/70">
                      <table className="min-w-full text-sm">
                        <thead className="bg-sand/80 text-left">
                          <tr>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              Fila
                            </th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              Apellido
                            </th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              Nombre
                            </th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              Generado
                            </th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              Correo final
                            </th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                              Estado
                            </th>
                          </tr>
                        </thead>
                        <tbody>
                          {emailConflictRows.map((conflict) => {
                            const key = String(conflict.rowNumber);
                            const currentValue =
                              emailOverrides[key] ?? conflict.generatedEmail;
                            const normalized = normalizeEmail(currentValue);
                            const isUsed = normalized
                              ? existingEmails.has(normalized)
                              : false;
                            const isDuplicate = normalized
                              ? (plannedEmailCounts.get(normalized) ?? 0) > 1
                              : false;
                            const status = !normalized
                              ? "Sin correo"
                              : isUsed
                              ? "En uso"
                              : isDuplicate
                              ? "Duplicado en la lista (no se exporta)"
                              : "Disponible";

                            return (
                              <tr key={conflict.rowNumber} className="border-t border-white/60">
                                <td className="px-4 py-2 text-ink/90">
                                  {conflict.rowNumber}
                                </td>
                                <td className="px-4 py-2 text-ink/90">
                                  {conflict.apellido || "-"}
                                </td>
                                <td className="px-4 py-2 text-ink/90">
                                  {conflict.nombre || "-"}
                                </td>
                                <td className="px-4 py-2 text-ink/70">
                                  {conflict.generatedEmail}
                                </td>
                                <td className="px-4 py-2">
                                  <input
                                    value={currentValue}
                                    onChange={(event) => {
                                      const nextValue = event.target.value;
                                      const cleanedValue = toCleanString(nextValue);
                                      setEmailOverrides((prev) => {
                                        const updated = { ...prev };
                                        if (
                                          !cleanedValue ||
                                          normalizeEmail(cleanedValue) === conflict.generatedEmail
                                        ) {
                                          delete updated[key];
                                        } else {
                                          updated[key] = cleanedValue;
                                        }
                                        return updated;
                                      });
                                    }}
                                    className="w-full min-w-[240px] rounded-xl border border-ink/20 bg-white px-3 py-2 text-sm"
                                  />
                                </td>
                                <td className="px-4 py-2 text-ink/90">{status}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}
        </section>

        <section className="rounded-3xl border border-white/60 bg-white/70 p-6 shadow-sm">
          <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
            <div>
              <h2 className="text-xl font-semibold text-ink">Validaciones y errores</h2>
              <p className="text-sm text-ink/70">
                Define que campos son obligatorios y revisa filas con faltantes.
              </p>
            </div>
            <div className="flex flex-wrap gap-3 text-sm text-ink/80">
              <label className="flex items-center gap-2 rounded-full border border-ink/10 bg-white/70 px-3 py-1">
                <input
                  type="checkbox"
                  checked={required.dni}
                  onChange={(event) =>
                    setRequired((prev) => ({ ...prev, dni: event.target.checked }))
                  }
                />
                DNI
              </label>
              <label className="flex items-center gap-2 rounded-full border border-ink/10 bg-white/70 px-3 py-1">
                <input
                  type="checkbox"
                  checked={required.celular}
                  onChange={(event) =>
                    setRequired((prev) => ({ ...prev, celular: event.target.checked }))
                  }
                />
                Celular
              </label>
              <label className="flex items-center gap-2 rounded-full border border-ink/10 bg-white/70 px-3 py-1">
                <input
                  type="checkbox"
                  checked={required.codigo}
                  onChange={(event) =>
                    setRequired((prev) => ({ ...prev, codigo: event.target.checked }))
                  }
                />
                Código estudiante
              </label>
            </div>
          </div>

          <div className="mt-4 rounded-2xl border border-ink/10 bg-sand/70 px-4 py-3 text-sm text-ink/70">
            {validationIssues.length === 0
              ? "No hay filas con problemas según las reglas actuales."
              : `${validationIssues.length} filas con datos incompletos.`}
          </div>

          {validationIssues.length > 0 && (
            <div className="mt-4 overflow-x-auto rounded-2xl border border-white/60 bg-white/70">
              <table className="min-w-full text-sm">
                <thead className="bg-sand/80 text-left">
                  <tr>
                    <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                      Fila
                    </th>
                    <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">
                      Faltantes
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {validationIssues.map((issue) => (
                    <tr key={issue.rowNumber} className="border-t border-white/60">
                      <td className="px-4 py-2 text-ink/90">{issue.rowNumber}</td>
                      <td className="px-4 py-2 text-ink/90">
                        {issue.missing.join(", ")}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>

        <section className="rounded-3xl border border-white/60 bg-white/70 p-6 shadow-sm">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-xl font-semibold text-ink">Generar CSV</h2>
              <p className="text-sm text-ink/70">
                Genera el CSV con el orden exacto de Outlook y revisa el preview.
              </p>
              <p className="text-sm text-ink/70">
                Filas exportables: {exportRows.length} de {rows.length} según condición INGRESÓ, campos obligatorios, sin DNI existente, sin correos ya usados y sin duplicados internos.
              </p>
              {totalCsvParts > 0 && (
                <p className="text-sm text-ink/70">
                  Descarga en {totalCsvParts} archivo(s) de hasta 249 filas cada uno.
                </p>
              )}
            </div>
            <div className="flex flex-wrap gap-3">
              <button
                type="button"
                onClick={handleGenerate}
                className="rounded-full bg-coral px-5 py-2 text-sm font-semibold text-white shadow-sm transition hover:brightness-95"
              >
                Generar CSV
              </button>
              <button
                type="button"
                onClick={handleDownload}
                disabled={!csvText}
                className="rounded-full border border-ink/20 bg-white/80 px-5 py-2 text-sm font-semibold text-ink transition hover:border-ink/40 disabled:cursor-not-allowed disabled:opacity-50"
              >
                Descargar contactos_outlook.csv
              </button>
            </div>
          </div>

          <div className="mt-6">
            <h3 className="text-base font-semibold text-ink">Preview CSV (primeras 20 filas)</h3>
            <div className="mt-3">
              <PreviewTable
                columns={OUTLOOK_HEADERS}
                rows={csvPreview}
                emptyState="Genera el CSV para ver el preview."
              />
            </div>
          </div>
        </section>
      </div>
    </main>
  );
}
