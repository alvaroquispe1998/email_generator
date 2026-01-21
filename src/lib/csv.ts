import Papa from "papaparse";
import { buildUsername, digitsOnly, toCleanString } from "./normalization";
import type { DataRow } from "./xlsx";

export const OUTLOOK_HEADERS = [
  "Nombre de usuario",
  "Nombre",
  "Apellido",
  "Nombre para mostrar",
  "Puesto",
  "Departamento",
  "Número del trabajo",
  "Teléfono de la oficina",
  "Teléfono móvil",
  "Fax",
  "Dirección de correo electrónico alternativa",
  "Dirección",
  "Ciudad",
  "Estado o provincia",
  "Código postal",
  "País o región"
];

export type MappingRule = {
  type: "column" | "fixed" | "generated";
  value: string;
};

export type MappingConfig = Record<string, MappingRule>;

const GENERATED_USERNAME = "username";
const GENERATED_DISPLAY = "displayName";
const HEADER_MOBILE = "Tel\u00e9fono m\u00f3vil";

function resolveBaseValue(
  rule: MappingRule | undefined,
  row: DataRow
): string {
  if (!rule) {
    return "";
  }
  if (rule.type === "fixed") {
    return toCleanString(rule.value);
  }
  if (rule.type === "column") {
    return toCleanString(row[rule.value] ?? "");
  }
  return "";
}

export function buildGeneratedUsername(row: DataRow, mapping: MappingConfig): string {
  const nombre = resolveBaseValue(mapping["Nombre"], row);
  const apellido = resolveBaseValue(mapping["Apellido"], row);
  return buildUsername(nombre, apellido);
}

export function buildOutputRow(row: DataRow, mapping: MappingConfig): Record<string, string> {
  const nombre = resolveBaseValue(mapping["Nombre"], row);
  const apellido = resolveBaseValue(mapping["Apellido"], row);
  const username = buildUsername(nombre, apellido);
  const displayName = [apellido, nombre].filter(Boolean).join(" ").trim();

  const output: Record<string, string> = {};

  OUTLOOK_HEADERS.forEach((header) => {
    const rule = mapping[header];
    let value = "";

    if (rule?.type === "generated") {
      if (rule.value === GENERATED_USERNAME) {
        value = username;
      } else if (rule.value === GENERATED_DISPLAY) {
        value = displayName;
      }
    } else {
      value = resolveBaseValue(rule, row);
    }

    if (header === HEADER_MOBILE) {
      value = digitsOnly(value);
    }
    if (header === "Fax") {
      value = digitsOnly(value);
    }

    output[header] = value;
  });

  return output;
}

export function buildOutputRows(rows: DataRow[], mapping: MappingConfig): Record<string, string>[] {
  return rows.map((row) => buildOutputRow(row, mapping));
}

export function generateCsvText(rows: DataRow[], mapping: MappingConfig): string {
  const data = buildOutputRows(rows, mapping).map((row) =>
    OUTLOOK_HEADERS.map((header) => row[header] ?? "")
  );

  const csvBody = Papa.unparse(
    {
      fields: OUTLOOK_HEADERS,
      data
    },
    {
      delimiter: ",",
      newline: "\r\n"
    }
  );

  return `\ufeff${csvBody}`;
}

export const GENERATED_OPTIONS = {
  username: GENERATED_USERNAME,
  displayName: GENERATED_DISPLAY
};
