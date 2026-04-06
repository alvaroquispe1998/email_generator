"use client";

import { useEffect, useMemo, useState } from "react";
import Papa from "papaparse";
import { MappingEditor } from "@/components/MappingEditor";
import { PreviewTable } from "@/components/PreviewTable";
import {
  MATERNAL_SURNAME_HEADER,
  MAPPING_HEADERS,
  OUTLOOK_HEADERS,
  PATERNAL_SURNAME_HEADER,
  buildAlternateGeneratedUsername,
  buildGeneratedUsername,
  buildOutputRow,
  buildOutputRows,
  type MappingConfig
} from "@/lib/csv";
import {
  buildPlannedEmailCounts,
  filterRowsByEmail,
  filterRowsForExport,
  findExistingStudent,
  findInternalDuplicateRows,
  type InternalDuplicateRow,
  type RequiredConfig,
  type ValidationIssue,
  validateRows
} from "@/lib/export-validation";
import { findBestColumn, normalizeLookupKey } from "@/lib/matching";
import { digitsOnly, toCleanString } from "@/lib/normalization";
import {
  buildCorrelativeEmail,
  buildOutlookDirectoryData,
  normalizeEmail,
  type OutlookDirectoryData,
  type OutlookIdentifierConflict
} from "@/lib/outlook";
import { parseWorkbook, type DataRow, type WorkbookData } from "@/lib/xlsx";

const STORAGE_MAPPING = "outlook-mapping-v1";
const STORAGE_REQUIRED = "outlook-required-v1";
const EMPTY_SET = new Set<string>();
const EMPTY_DIRECTORY_MAP = new Map<string, string>();

type OutlookDirectory = OutlookDirectoryData & { fileName: string };

type ExistingStudentMatch = {
  rowNumber: number;
  dni: string;
  codigo: string;
  nombre: string;
  apellido: string;
  matchReason: string;
  existingEmail: string;
};

type EmailConflictRow = {
  rowNumber: number;
  row: DataRow;
  nombre: string;
  apellido: string;
  generatedEmail: string;
};

const DEFAULT_REQUIRED: RequiredConfig = { dni: true, celular: true, codigo: true };

const GENERATED_BY_HEADER: Record<string, string> = {
  "Nombre de usuario": "username",
  Apellido: "fullSurname",
  "Nombre para mostrar": "displayName"
};

const POSTAL_HEADER =
  OUTLOOK_HEADERS.find((header) => normalizeLookupKey(header) === "CODIGOPOSTAL") ??
  "Código postal";
const MOBILE_HEADER =
  OUTLOOK_HEADERS.find((header) => normalizeLookupKey(header) === "TELEFONOMOVIL") ??
  "Teléfono móvil";

function buildDefaultMapping(columns: string[]): MappingConfig {
  const nombre = findBestColumn(columns, ["NOMBRES", "NOMBRES COMPLETOS", "NOMBRE"]);
  const apellidoPaterno = findBestColumn(columns, ["A_PATERNO", "APELLIDO PATERNO", "APELLIDO_PATERNO"]);
  const apellidoMaterno = findBestColumn(columns, ["A_MATERNO", "APELLIDO MATERNO", "APELLIDO_MATERNO"]);
  const apellido = findBestColumn(columns, ["APELLIDOS", "APELLIDOS COMPLETOS", "APELLIDO"]);
  const celular = findBestColumn(columns, ["NUMERO DE CELULAR", "CELULAR", "TELEFONO MOVIL"]);
  const dni = findBestColumn(columns, ["DNI", "DOCUMENTO"]);
  const codigo = findBestColumn(columns, ["CODIGO DE ESTUDIANTE", "CODIGO ESTUDIANTE", "CODIGO"]);
  const correo = findBestColumn(columns, ["CORREO PERSONAL", "EMAIL PERSONAL", "MAIL"]);
  const usesStructuredSurnames = Boolean(apellidoPaterno || apellidoMaterno);

  return {
    "Nombre de usuario": { type: "generated", value: GENERATED_BY_HEADER["Nombre de usuario"] },
    Nombre: { type: "column", value: nombre },
    [PATERNAL_SURNAME_HEADER]: { type: "column", value: apellidoPaterno },
    [MATERNAL_SURNAME_HEADER]: { type: "column", value: apellidoMaterno },
    Apellido: usesStructuredSurnames
      ? { type: "generated", value: GENERATED_BY_HEADER.Apellido }
      : { type: "column", value: apellido },
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
    "País o región": { type: "fixed", value: "Perú" }
  };
}

function getRowNameParts(row: DataRow, mapping: MappingConfig) {
  const output = buildOutputRow(row, mapping);
  return {
    nombre: toCleanString(output["Nombre"]),
    apellido: toCleanString(output["Apellido"])
  };
}

function getRowStudentCode(row: DataRow, mapping: MappingConfig): string {
  return toCleanString(buildOutputRow(row, mapping)[POSTAL_HEADER]);
}

function getGeneratedEmail(row: DataRow, mapping: MappingConfig): string {
  return buildGeneratedUsername(row, mapping);
}

function getAlternateEmail(row: DataRow, mapping: MappingConfig): string {
  return buildAlternateGeneratedUsername(row, mapping);
}

function mergeMapping(defaults: MappingConfig, stored?: MappingConfig | null): MappingConfig {
  if (!stored) {
    return defaults;
  }
  const merged: MappingConfig = { ...defaults };
  MAPPING_HEADERS.forEach((header) => {
    if (!stored[header]) {
      return;
    }
    if (
      header === "Apellido" &&
      defaults[header]?.type === "generated" &&
      defaults[header]?.value === GENERATED_BY_HEADER.Apellido
    ) {
      merged[header] = defaults[header];
      return;
    }
    merged[header] = stored[header];
  });
  return merged;
}

function sanitizeMapping(mapping: MappingConfig, columns: string[], defaults: MappingConfig) {
  const sanitized: MappingConfig = { ...defaults };
  MAPPING_HEADERS.forEach((header) => {
    const rule = mapping[header];
    if (!rule) {
      return;
    }
    if (rule.type === "column") {
      sanitized[header] = columns.includes(rule.value) ? rule : { type: "column", value: "" };
      return;
    }
    if (rule.type === "generated") {
      const allowed = GENERATED_BY_HEADER[header];
      sanitized[header] = allowed && rule.value === allowed ? rule : defaults[header];
      return;
    }
    sanitized[header] = rule;
  });
  return sanitized;
}

function formatConflictKind(kind: OutlookIdentifierConflict["kind"]): string {
  return kind === "dni" ? "DNI" : "Código";
}

export default function HomePage() {
  const [workbook, setWorkbook] = useState<WorkbookData | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [columns, setColumns] = useState<string[]>([]);
  const [rows, setRows] = useState<DataRow[]>([]);
  const [mapping, setMapping] = useState<MappingConfig>(() => buildDefaultMapping([]));
  const [required, setRequired] = useState<RequiredConfig>(DEFAULT_REQUIRED);
  const [csvPreview, setCsvPreview] = useState<Record<string, string>[]>([]);
  const [csvText, setCsvText] = useState("");
  const [generatedRows, setGeneratedRows] = useState<Record<string, string>[]>([]);
  const [isParsing, setIsParsing] = useState(false);
  const [storedMapping, setStoredMapping] = useState<MappingConfig | null>(null);
  const [hasUserMapping, setHasUserMapping] = useState(false);
  const [outlookDirectory, setOutlookDirectory] = useState<OutlookDirectory | null>(null);
  const [emailOverrides, setEmailOverrides] = useState<Record<string, string>>({});
  const [pendingBulkAction, setPendingBulkAction] = useState<null | "secondName" | "correlative">(null);

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
      const sheet = workbook.sheets.find((item) => item.name === sheetName) ?? workbook.sheets[0];
      setColumns(sheet?.columns ?? []);
      setRows(sheet?.rows ?? []);
      return;
    }
    setColumns([]);
    setRows([]);
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

  const conditionColumn = useMemo(() => findBestColumn(columns, ["CONDICION"]), [columns]);
  const existingEmails = outlookDirectory?.emails ?? EMPTY_SET;
  const emailByDni = outlookDirectory?.emailByDni ?? EMPTY_DIRECTORY_MAP;
  const emailByCode = outlookDirectory?.emailByCode ?? EMPTY_DIRECTORY_MAP;

  const validationIssues = useMemo(
    () =>
      validateRows(
        rows,
        mapping,
        required,
        conditionColumn,
        MOBILE_HEADER,
        POSTAL_HEADER,
        emailByDni,
        emailByCode,
        getGeneratedEmail
      ),
    [rows, mapping, required, conditionColumn, emailByDni, emailByCode]
  );

  const baseExportRows = useMemo(
    () =>
      filterRowsForExport(
        rows,
        mapping,
        required,
        conditionColumn,
        MOBILE_HEADER,
        POSTAL_HEADER,
        emailByDni,
        emailByCode,
        getGeneratedEmail
      ),
    [rows, mapping, required, conditionColumn, emailByDni, emailByCode]
  );

  const existingStudentMatches = useMemo(() => {
    if (!emailByDni.size && !emailByCode.size) {
      return [] as ExistingStudentMatch[];
    }
    return rows
      .map((row) => {
        const output = buildOutputRow(row, mapping);
        const existingStudent = findExistingStudent(output, POSTAL_HEADER, emailByDni, emailByCode);
        if (!existingStudent) {
          return null;
        }
        return {
          rowNumber: row.__rowNumber,
          dni: digitsOnly(toCleanString(output["Fax"])),
          codigo: getRowStudentCode(row, mapping),
          nombre: toCleanString(output["Nombre"]),
          apellido: toCleanString(output["Apellido"]),
          matchReason: existingStudent.matchReason,
          existingEmail: existingStudent.existingEmail
        };
      })
      .filter(Boolean) as ExistingStudentMatch[];
  }, [rows, mapping, emailByDni, emailByCode]);

  const emailConflictRows = useMemo(() => {
    if (!existingEmails.size) {
      return [] as EmailConflictRow[];
    }
    return baseExportRows
      .map((row) => {
        const generatedEmail = normalizeEmail(getGeneratedEmail(row, mapping));
        if (!generatedEmail || !existingEmails.has(generatedEmail)) {
          return null;
        }
        const { nombre, apellido } = getRowNameParts(row, mapping);
        return { rowNumber: row.__rowNumber, row, nombre, apellido, generatedEmail };
      })
      .filter(Boolean) as EmailConflictRow[];
  }, [baseExportRows, mapping, existingEmails]);

  const plannedEmailCounts = useMemo(
    () => buildPlannedEmailCounts(baseExportRows, mapping, emailOverrides, getGeneratedEmail),
    [baseExportRows, mapping, emailOverrides]
  );

  const internalDuplicateRows = useMemo(
    () =>
      findInternalDuplicateRows(
        baseExportRows,
        mapping,
        emailOverrides,
        existingEmails,
        plannedEmailCounts,
        getGeneratedEmail
      ),
    [baseExportRows, mapping, emailOverrides, existingEmails, plannedEmailCounts]
  );

  const exportRows = useMemo(
    () =>
      filterRowsByEmail(
        baseExportRows,
        mapping,
        emailOverrides,
        existingEmails,
        plannedEmailCounts,
        getGeneratedEmail
      ),
    [baseExportRows, mapping, emailOverrides, existingEmails, plannedEmailCounts]
  );

  const totalCsvParts = useMemo(() => (generatedRows.length ? Math.ceil(generatedRows.length / 249) : 0), [generatedRows.length]);

  const buildUsedEmailSet = (overrides: Record<string, string>) => {
    const used = new Set<string>();
    existingEmails.forEach((email) => used.add(email));
    baseExportRows.forEach((row) => {
      const key = String(row.__rowNumber);
      const email = normalizeEmail(overrides[key] || getGeneratedEmail(row, mapping));
      if (email) {
        used.add(email);
      }
    });
    return used;
  };

  const applySecondNameOverrides = () => {
    setEmailOverrides(() => {
      const updated: Record<string, string> = {};
      emailConflictRows.forEach((conflict) => {
        const alternate = getAlternateEmail(conflict.row, mapping);
        if (!alternate || normalizeEmail(alternate) === conflict.generatedEmail) {
          return;
        }
        updated[String(conflict.rowNumber)] = alternate;
      });
      return updated;
    });
  };

  const applyCorrelativeOverrides = () => {
    setEmailOverrides(() => {
      const updated: Record<string, string> = {};
      const used = buildUsedEmailSet(updated);
      emailConflictRows.forEach((conflict) => {
        const candidate = buildCorrelativeEmail(conflict.generatedEmail, used);
        if (!candidate) {
          return;
        }
        updated[String(conflict.rowNumber)] = candidate;
        used.add(normalizeEmail(candidate));
      });
      return updated;
    });
  };

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
      const parsed = parseWorkbook(await file.arrayBuffer());
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
      const parsed = Papa.parse<Record<string, string>>(await file.text(), {
        header: true,
        skipEmptyLines: true
      });
      setOutlookDirectory({
        fileName: file.name,
        ...buildOutlookDirectoryData(parsed.data, parsed.meta.fields ?? [])
      });
    } catch {
      setOutlookDirectory({
        fileName: file.name,
        rowCount: 0,
        emails: new Set(),
        emailByDni: new Map(),
        emailByCode: new Map(),
        inconsistencies: [],
        error: "No se pudo leer el CSV de Outlook."
      });
    }
  };

  const buildCurrentOutputRows = () =>
    buildOutputRows(exportRows, mapping).map((row, index) => {
      const key = String(exportRows[index].__rowNumber);
      const override = emailOverrides[key];
      if (override && toCleanString(override)) {
        row["Nombre de usuario"] = normalizeEmail(override);
      }
      return row;
    });

  const handleGenerate = () => {
    const outputRows = buildCurrentOutputRows();
    setGeneratedRows(outputRows);
    setCsvPreview(outputRows.slice(0, 20));
    const csvBody = Papa.unparse(
      {
        fields: OUTLOOK_HEADERS,
        data: outputRows.map((row) => OUTLOOK_HEADERS.map((header) => row[header] ?? ""))
      },
      { delimiter: ",", newline: "\r\n" }
    );
    setCsvText(`\ufeff${csvBody}`);
  };

  const handleDownload = () => {
    const outputRows = buildCurrentOutputRows();
    if (!outputRows.length) {
      return;
    }
    const chunkSize = 249;
    const totalParts = Math.ceil(outputRows.length / chunkSize);
    for (let index = 0; index < totalParts; index += 1) {
      const chunk = outputRows.slice(index * chunkSize, index * chunkSize + chunkSize);
      const csvBody = Papa.unparse(
        {
          fields: OUTLOOK_HEADERS,
          data: chunk.map((row) => OUTLOOK_HEADERS.map((header) => row[header] ?? ""))
        },
        { delimiter: ",", newline: "\r\n" }
      );
      const blob = new Blob([`\ufeff${csvBody}`], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = totalParts === 1 ? "contactos_outlook.csv" : `contactos_outlook_parte_${index + 1}.csv`;
      window.setTimeout(() => {
        anchor.click();
        URL.revokeObjectURL(url);
      }, index * 150);
    }
  };

  const handleDownloadInconsistencies = () => {
    if (!outlookDirectory?.inconsistencies.length) {
      return;
    }

    const csvBody = Papa.unparse(
      {
        fields: ["Tipo", "Valor", "Correos"],
        data: outlookDirectory.inconsistencies.map((conflict) => [
          formatConflictKind(conflict.kind),
          conflict.value,
          conflict.emails.join(", ")
        ])
      },
      {
        delimiter: ",",
        newline: "\r\n"
      }
    );

    const blob = new Blob([`\ufeff${csvBody}`], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "inconsistencias_outlook.csv";
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const handleReset = () => {
    setMapping(buildDefaultMapping(columns));
    setRequired(DEFAULT_REQUIRED);
    setHasUserMapping(true);
  };

  return (
    <main className="px-4 py-10">
      <div className="mx-auto max-w-6xl space-y-8">
        <header className="space-y-3">
          <span className="inline-flex w-fit items-center rounded-full border border-ink/15 bg-white/60 px-3 py-1 text-xs font-semibold uppercase tracking-[0.2em] text-ink/70">
            UAI
          </span>
          <div>
            <h1 className="text-3xl font-semibold text-ink">Generador de contactos para Outlook</h1>
            <p className="mt-2 max-w-3xl text-sm text-ink/70">
              Carga un Excel, configura el mapeo y descarga un CSV listo para importar.
            </p>
          </div>
        </header>

        <section className="rounded-3xl border border-white/60 bg-white/70 p-6 shadow-sm">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-xl font-semibold text-ink">1. Cargar XLSX</h2>
              <p className="text-sm text-ink/70">Selecciona el archivo de informe con datos de estudiantes.</p>
            </div>
            <label className="inline-flex cursor-pointer items-center gap-2 rounded-full border border-ink/20 bg-white px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40">
              <input type="file" accept=".xlsx" onChange={handleFileChange} className="hidden" />
              Subir XLSX
            </label>
          </div>

          {isParsing && (
            <div className="mt-4 rounded-2xl border border-ink/10 bg-sand/70 px-4 py-3 text-sm text-ink/70">
              Leyendo archivo...
            </div>
          )}

          {workbook && workbook.sheets.length > 1 && (
            <div className="mt-6">
              <label className="text-sm font-semibold text-ink">Hoja</label>
              <select
                value={sheetName}
                onChange={(event) => setSheetName(event.target.value)}
                className="mt-2 rounded-xl border border-ink/20 bg-white px-3 py-2 text-sm"
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
            mappingHeaders={MAPPING_HEADERS}
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
                Sube el CSV exportado de Outlook para detectar alumnos con correo por DNI o código y correos ya usados.
              </p>
            </div>
            <label className="inline-flex cursor-pointer items-center gap-2 rounded-full border border-ink/20 bg-white px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40">
              <input type="file" accept=".csv" onChange={handleOutlookCsvChange} className="hidden" />
              Subir CSV de Outlook
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
                <div>DNI con correo detectados: {outlookDirectory.emailByDni.size}</div>
                <div>Códigos con correo detectados: {outlookDirectory.emailByCode.size}</div>
                <div>Correos detectados: {outlookDirectory.emails.size}</div>
                <div>Conflictos internos del directorio: {outlookDirectory.inconsistencies.length}</div>
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
                <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <h3 className="text-base font-semibold text-ink">Inconsistencias en el CSV de Outlook</h3>
                    <p className="text-sm text-ink/70">
                      Se muestran identificadores repetidos con correos distintos dentro del mismo CSV.
                    </p>
                  </div>
                  {outlookDirectory.inconsistencies.length > 0 && (
                    <button
                      type="button"
                      onClick={handleDownloadInconsistencies}
                      className="rounded-full border border-ink/20 bg-white/80 px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40"
                    >
                      Exportar inconsistencias
                    </button>
                  )}
                </div>
                <div className="mt-3">
                  {outlookDirectory.inconsistencies.length === 0 ? (
                    <div className="rounded-2xl border border-ink/10 bg-white/70 px-4 py-3 text-sm text-ink/70">
                      No se detectaron inconsistencias por DNI o código.
                    </div>
                  ) : (
                    <div className="overflow-x-auto rounded-2xl border border-white/60 bg-white/70">
                      <table className="min-w-full text-sm">
                        <thead className="bg-sand/80 text-left">
                          <tr>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Tipo</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Valor</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Correos</th>
                          </tr>
                        </thead>
                        <tbody>
                          {outlookDirectory.inconsistencies.map((conflict) => (
                            <tr key={`${conflict.kind}-${conflict.value}`} className="border-t border-white/60">
                              <td className="px-4 py-2 text-ink/90">{formatConflictKind(conflict.kind)}</td>
                              <td className="px-4 py-2 text-ink/90">{conflict.value}</td>
                              <td className="px-4 py-2 text-ink/90">{conflict.emails.join(", ")}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </div>

              <div>
                <h3 className="text-base font-semibold text-ink">Estudiantes con correo existente</h3>
                <p className="text-sm text-ink/70">
                  Estas filas no se mostrarán como exportables ni se incluirán en el CSV final.
                </p>
                <div className="mt-3">
                  {existingStudentMatches.length === 0 ? (
                    <div className="rounded-2xl border border-ink/10 bg-white/70 px-4 py-3 text-sm text-ink/70">
                      No se encontraron estudiantes con correo existente.
                    </div>
                  ) : (
                    <div className="overflow-x-auto rounded-2xl border border-white/60 bg-white/70">
                      <table className="min-w-full text-sm">
                        <thead className="bg-sand/80 text-left">
                          <tr>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Fila</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">DNI</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Código estudiante</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Apellido</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Nombre</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Coincidencia</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Correo existente</th>
                          </tr>
                        </thead>
                        <tbody>
                          {existingStudentMatches.map((match) => (
                            <tr key={match.rowNumber} className="border-t border-white/60">
                              <td className="px-4 py-2 text-ink/90">{match.rowNumber}</td>
                              <td className="px-4 py-2 text-ink/90">{match.dni}</td>
                              <td className="px-4 py-2 text-ink/90">{match.codigo || "-"}</td>
                              <td className="px-4 py-2 text-ink/90">{match.apellido || "-"}</td>
                              <td className="px-4 py-2 text-ink/90">{match.nombre || "-"}</td>
                              <td className="px-4 py-2 text-ink/90">{match.matchReason}</td>
                              <td className="px-4 py-2 text-ink/90">{match.existingEmail || "-"}</td>
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
                    <div className="flex flex-wrap gap-3">
                      <button
                        type="button"
                        onClick={() => {
                          const hasChanges = Object.keys(emailOverrides).length > 0;
                          if (hasChanges) {
                            setPendingBulkAction("secondName");
                            return;
                          }
                          applySecondNameOverrides();
                        }}
                        className="rounded-full border border-ink/20 bg-white/80 px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40"
                      >
                        Usar segundo nombre
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          const hasChanges = Object.keys(emailOverrides).length > 0;
                          if (hasChanges) {
                            setPendingBulkAction("correlative");
                            return;
                          }
                          applyCorrelativeOverrides();
                        }}
                        className="rounded-full border border-ink/20 bg-white/80 px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40"
                      >
                        Usar correlativo (+2)
                      </button>
                    </div>
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
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Fila</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Apellido</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Nombre</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Generado</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Correo final</th>
                            <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Estado</th>
                          </tr>
                        </thead>
                        <tbody>
                          {emailConflictRows.map((conflict) => {
                            const key = String(conflict.rowNumber);
                            const currentValue = emailOverrides[key] ?? conflict.generatedEmail;
                            const normalized = normalizeEmail(currentValue);
                            const isUsed = normalized ? existingEmails.has(normalized) : false;
                            const isDuplicate = normalized ? (plannedEmailCounts.get(normalized) ?? 0) > 1 : false;
                            const status = !normalized
                              ? "Sin correo"
                              : isUsed
                              ? "En uso"
                              : isDuplicate
                              ? "Duplicado en la lista"
                              : "Disponible";

                            return (
                              <tr key={conflict.rowNumber} className="border-t border-white/60">
                                <td className="px-4 py-2 text-ink/90">{conflict.rowNumber}</td>
                                <td className="px-4 py-2 text-ink/90">{conflict.apellido || "-"}</td>
                                <td className="px-4 py-2 text-ink/90">{conflict.nombre || "-"}</td>
                                <td className="px-4 py-2 text-ink/70">{conflict.generatedEmail}</td>
                                <td className="px-4 py-2">
                                  <input
                                    value={currentValue}
                                    onChange={(event) => {
                                      const cleanedValue = toCleanString(event.target.value);
                                      setEmailOverrides((prev) => {
                                        const updated = { ...prev };
                                        if (!cleanedValue || normalizeEmail(cleanedValue) === conflict.generatedEmail) {
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
              <h2 className="text-xl font-semibold text-ink">3. Validaciones y errores</h2>
              <p className="text-sm text-ink/70">Define qué campos son obligatorios y revisa filas con faltantes.</p>
            </div>
            <div className="flex flex-wrap gap-3 text-sm text-ink/80">
              <label className="flex items-center gap-2 rounded-full border border-ink/10 bg-white/70 px-3 py-1">
                <input
                  type="checkbox"
                  checked={required.dni}
                  onChange={(event) => setRequired((prev) => ({ ...prev, dni: event.target.checked }))}
                />
                DNI
              </label>
              <label className="flex items-center gap-2 rounded-full border border-ink/10 bg-white/70 px-3 py-1">
                <input
                  type="checkbox"
                  checked={required.celular}
                  onChange={(event) => setRequired((prev) => ({ ...prev, celular: event.target.checked }))}
                />
                Celular
              </label>
              <label className="flex items-center gap-2 rounded-full border border-ink/10 bg-white/70 px-3 py-1">
                <input
                  type="checkbox"
                  checked={required.codigo}
                  onChange={(event) => setRequired((prev) => ({ ...prev, codigo: event.target.checked }))}
                />
                Código estudiante
              </label>
            </div>
          </div>

          <div className="mt-4 rounded-2xl border border-ink/10 bg-sand/70 px-4 py-3 text-sm text-ink/70">
            {validationIssues.length === 0
              ? "No hay filas con problemas según las reglas actuales."
              : `${validationIssues.length} filas con datos incompletos o sin correo institucional.`}
          </div>

          {validationIssues.length > 0 && (
            <div className="mt-4 overflow-x-auto rounded-2xl border border-white/60 bg-white/70">
              <table className="min-w-full text-sm">
                <thead className="bg-sand/80 text-left">
                  <tr>
                    <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Fila</th>
                    <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Problemas</th>
                  </tr>
                </thead>
                <tbody>
                  {validationIssues.map((issue: ValidationIssue) => (
                    <tr key={issue.rowNumber} className="border-t border-white/60">
                      <td className="px-4 py-2 text-ink/90">{issue.rowNumber}</td>
                      <td className="px-4 py-2 text-ink/90">{issue.issues.join(", ")}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          <div className="mt-6">
            <h3 className="text-base font-semibold text-ink">Duplicados internos</h3>
            <p className="text-sm text-ink/70">
              Estas filas compiten por el mismo correo final dentro de la exportación.
            </p>
            <div className="mt-3">
              {internalDuplicateRows.length === 0 ? (
                <div className="rounded-2xl border border-ink/10 bg-white/70 px-4 py-3 text-sm text-ink/70">
                  No hay duplicados internos de correo.
                </div>
              ) : (
                <div className="overflow-x-auto rounded-2xl border border-white/60 bg-white/70">
                  <table className="min-w-full text-sm">
                    <thead className="bg-sand/80 text-left">
                      <tr>
                        <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Fila</th>
                        <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Apellido</th>
                        <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Nombre</th>
                        <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Generado</th>
                        <th className="px-4 py-3 text-xs font-semibold uppercase tracking-wide text-ink/70">Correo final duplicado</th>
                      </tr>
                    </thead>
                    <tbody>
                      {internalDuplicateRows.map((duplicate: InternalDuplicateRow) => (
                        <tr key={`${duplicate.rowNumber}-${duplicate.finalEmail}`} className="border-t border-white/60">
                          <td className="px-4 py-2 text-ink/90">{duplicate.rowNumber}</td>
                          <td className="px-4 py-2 text-ink/90">{duplicate.apellido || "-"}</td>
                          <td className="px-4 py-2 text-ink/90">{duplicate.nombre || "-"}</td>
                          <td className="px-4 py-2 text-ink/70">{duplicate.generatedEmail || "-"}</td>
                          <td className="px-4 py-2 text-ink/90">{duplicate.finalEmail}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        </section>

        <section className="rounded-3xl border border-white/60 bg-white/70 p-6 shadow-sm">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <h2 className="text-xl font-semibold text-ink">4. Generar CSV</h2>
              <p className="text-sm text-ink/70">Genera el CSV con el orden exacto de Outlook y revisa el preview.</p>
              <p className="text-sm text-ink/70">
                Filas exportables: {exportRows.length} de {rows.length} según condición INGRESO, campos obligatorios, sin alumnos con correo existente por DNI o código, sin correos ya usados, sin correo vacío y sin duplicados internos.
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

      {pendingBulkAction && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-ink/30 px-4">
          <div className="w-full max-w-md rounded-2xl border border-ink/10 bg-white p-6 shadow-lg">
            <div className="flex items-start justify-between gap-4">
              <div>
                <h4 className="text-lg font-semibold text-ink">Cambiar sugerencias</h4>
                <p className="mt-2 text-sm text-ink/70">
                  Estás a punto de reemplazar todos los correos editados por nuevas sugerencias.
                  Si quieres conservar tus cambios, puedes cancelar.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setPendingBulkAction(null)}
                className="text-sm font-semibold text-ink/60 transition hover:text-ink"
                aria-label="Cerrar"
              >
                X
              </button>
            </div>
            <div className="mt-5 flex flex-wrap justify-end gap-3">
              <button
                type="button"
                onClick={() => setPendingBulkAction(null)}
                className="rounded-full border border-ink/20 bg-white px-4 py-2 text-sm font-semibold text-ink transition hover:border-ink/40"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={() => {
                  if (pendingBulkAction === "secondName") {
                    applySecondNameOverrides();
                  } else {
                    applyCorrelativeOverrides();
                  }
                  setPendingBulkAction(null);
                }}
                className="rounded-full bg-coral px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:brightness-95"
              >
                Reemplazar sugerencias
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
