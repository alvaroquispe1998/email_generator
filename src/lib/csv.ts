import Papa from "papaparse";
import {
  buildUsername,
  buildUsernameWithSecondName,
  digitsOnly,
  getLeadingWord,
  toCleanString
} from "./normalization";
import { normalizeLookupKey } from "./matching";
import type { DataRow } from "./xlsx";

export const OUTLOOK_HEADERS = [
  "Nombre de usuario",
  "Nombre",
  "Apellido",
  "Nombre para mostrar",
  "Puesto",
  "Departamento",
  "N\u00famero del trabajo",
  "Tel\u00e9fono de la oficina",
  "Tel\u00e9fono m\u00f3vil",
  "Fax",
  "Direcci\u00f3n de correo electr\u00f3nico alternativa",
  "Direcci\u00f3n",
  "Ciudad",
  "Estado o provincia",
  "C\u00f3digo postal",
  "Pa\u00eds o regi\u00f3n"
];

export const PATERNAL_SURNAME_HEADER = "Apellido paterno";
export const MATERNAL_SURNAME_HEADER = "Apellido materno";
export const EXTRA_MAPPING_HEADERS = [
  PATERNAL_SURNAME_HEADER,
  MATERNAL_SURNAME_HEADER
];
export const MAPPING_HEADERS = OUTLOOK_HEADERS.flatMap((header) =>
  header === "Nombre" ? [header, ...EXTRA_MAPPING_HEADERS] : [header]
);

export type MappingRule = {
  type: "column" | "fixed" | "generated";
  value: string;
};

export type MappingConfig = Record<string, MappingRule>;

const GENERATED_USERNAME = "username";
const GENERATED_DISPLAY = "displayName";
const GENERATED_FULL_SURNAME = "fullSurname";
const HEADER_MOBILE = "Tel\u00e9fono m\u00f3vil";
const STRUCTURED_PATERNAL_COLUMNS = [
  "a_paterno",
  "apellido_paterno",
  "apellido paterno",
  "apaterno"
];
const STRUCTURED_MATERNAL_COLUMNS = [
  "a_materno",
  "apellido_materno",
  "apellido materno",
  "amaterno"
];

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

function normalizeSourceKey(value: string): string {
  return normalizeLookupKey(value);
}

function findStructuredValue(row: DataRow, candidates: string[]): string {
  for (const candidate of candidates) {
    const direct = row[candidate];
    if (direct !== undefined) {
      const value = toCleanString(direct);
      if (value) {
        return value;
      }
    }
  }

  const candidateKeys = candidates.map(normalizeSourceKey);
  for (const [key, rawValue] of Object.entries(row)) {
    if (key === "__rowNumber") {
      continue;
    }
    if (!candidateKeys.includes(normalizeSourceKey(key))) {
      continue;
    }
    const value = toCleanString(rawValue);
    if (value) {
      return value;
    }
  }

  return "";
}

function resolveStructuredValue(
  row: DataRow,
  mapping: MappingConfig,
  mappingHeader: string,
  candidates: string[]
): string {
  return resolveBaseValue(mapping[mappingHeader], row) || findStructuredValue(row, candidates);
}

function getStructuredSurnames(row: DataRow, mapping: MappingConfig): {
  paterno: string;
  materno: string;
  completo: string;
} {
  const paterno = resolveStructuredValue(
    row,
    mapping,
    PATERNAL_SURNAME_HEADER,
    STRUCTURED_PATERNAL_COLUMNS
  );
  const materno = resolveStructuredValue(
    row,
    mapping,
    MATERNAL_SURNAME_HEADER,
    STRUCTURED_MATERNAL_COLUMNS
  );
  return {
    paterno,
    materno,
    completo: [paterno, materno].filter(Boolean).join(" ").trim()
  };
}

function resolveSurnameValue(row: DataRow, mapping: MappingConfig): string {
  const rule = mapping["Apellido"];
  const structured = getStructuredSurnames(row, mapping);

  if (structured.completo) {
    return structured.completo;
  }

  if (rule?.type === "generated" && rule.value === GENERATED_FULL_SURNAME) {
    return structured.completo;
  }

  return resolveBaseValue(rule, row);
}

function resolveUsernameSurname(row: DataRow, mapping: MappingConfig): string {
  const structured = getStructuredSurnames(row, mapping);
  if (structured.paterno) {
    return structured.paterno;
  }

  const apellido = resolveSurnameValue(row, mapping);
  return getLeadingWord(apellido);
}

export function buildGeneratedUsername(row: DataRow, mapping: MappingConfig): string {
  const nombre = resolveBaseValue(mapping["Nombre"], row);
  const apellidoParaUsuario = resolveUsernameSurname(row, mapping);
  return buildUsername(nombre, apellidoParaUsuario);
}

export function buildAlternateGeneratedUsername(row: DataRow, mapping: MappingConfig): string {
  const nombre = resolveBaseValue(mapping["Nombre"], row);
  const apellidoParaUsuario = resolveUsernameSurname(row, mapping);
  return buildUsernameWithSecondName(nombre, apellidoParaUsuario);
}

export function buildOutputRow(row: DataRow, mapping: MappingConfig): Record<string, string> {
  const nombre = resolveBaseValue(mapping["Nombre"], row);
  const apellido = resolveSurnameValue(row, mapping);
  const apellidoParaUsuario = resolveUsernameSurname(row, mapping);
  const username = buildUsername(nombre, apellidoParaUsuario);
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
      } else if (rule.value === GENERATED_FULL_SURNAME) {
        value = apellido;
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
  displayName: GENERATED_DISPLAY,
  fullSurname: GENERATED_FULL_SURNAME
};
