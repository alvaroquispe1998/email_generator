import { buildOutputRow, type MappingConfig } from "./csv";
import { toCleanString } from "./normalization";
import {
  normalizeEmail,
  normalizeStudentCode
} from "./outlook";
import type { DataRow } from "./xlsx";

export type RequiredConfig = {
  dni: boolean;
  celular: boolean;
  codigo: boolean;
};

export type ValidationIssue = {
  rowNumber: number;
  issues: string[];
};

export type ExistingStudentResult = {
  existingEmail: string;
  matchReason: string;
};

export type InternalDuplicateRow = {
  rowNumber: number;
  nombre: string;
  apellido: string;
  generatedEmail: string;
  finalEmail: string;
};

export function rowHasIngresoCondition(row: DataRow, conditionColumn: string): boolean {
  if (!conditionColumn) {
    return true;
  }
  const value = toCleanString(row[conditionColumn]);
  return normalizeStudentCode(value) === "INGRESO";
}

export function findExistingStudent(
  output: Record<string, string>,
  postalHeader: string,
  emailByDni: Map<string, string>,
  emailByCode: Map<string, string>
): ExistingStudentResult | null {
  const dni = toCleanString(output["Fax"]).replace(/\D/g, "");
  const codigo = normalizeStudentCode(output[postalHeader] ?? "");
  const matchByDni = dni ? emailByDni.get(dni) ?? "" : "";
  const matchByCode = codigo ? emailByCode.get(codigo) ?? "" : "";

  if (!matchByDni && !matchByCode) {
    return null;
  }

  return {
    existingEmail: matchByDni || matchByCode,
    matchReason:
      matchByDni && matchByCode ? "DNI y código" : matchByDni ? "DNI" : "Código"
  };
}

function collectValidationIssues(
  output: Record<string, string>,
  required: RequiredConfig,
  mobileHeader: string,
  postalHeader: string,
  generatedEmail: string
): string[] {
  const issues: string[] = [];

  if (required.dni && !toCleanString(output["Fax"])) {
    issues.push("DNI");
  }
  if (required.celular && !toCleanString(output[mobileHeader])) {
    issues.push("Celular");
  }
  if (required.codigo && !toCleanString(output[postalHeader])) {
    issues.push("Código estudiante");
  }
  if (!generatedEmail) {
    issues.push("Correo institucional");
  }

  return issues;
}

export function validateRows(
  rows: DataRow[],
  mapping: MappingConfig,
  required: RequiredConfig,
  conditionColumn: string,
  mobileHeader: string,
  postalHeader: string,
  emailByDni: Map<string, string>,
  emailByCode: Map<string, string>,
  getGeneratedEmail: (row: DataRow, mapping: MappingConfig) => string
): ValidationIssue[] {
  return rows
    .map((row) => {
      if (!rowHasIngresoCondition(row, conditionColumn)) {
        return null;
      }
      const output = buildOutputRow(row, mapping);
      if (findExistingStudent(output, postalHeader, emailByDni, emailByCode)) {
        return null;
      }

      const issues = collectValidationIssues(
        output,
        required,
        mobileHeader,
        postalHeader,
        normalizeEmail(getGeneratedEmail(row, mapping))
      );

      if (issues.length === 0) {
        return null;
      }

      return { rowNumber: row.__rowNumber, issues };
    })
    .filter(Boolean) as ValidationIssue[];
}

export function rowMeetsRequired(
  row: DataRow,
  mapping: MappingConfig,
  required: RequiredConfig,
  conditionColumn: string,
  mobileHeader: string,
  postalHeader: string,
  emailByDni: Map<string, string>,
  emailByCode: Map<string, string>,
  getGeneratedEmail: (row: DataRow, mapping: MappingConfig) => string
): boolean {
  if (!rowHasIngresoCondition(row, conditionColumn)) {
    return false;
  }
  const output = buildOutputRow(row, mapping);
  if (findExistingStudent(output, postalHeader, emailByDni, emailByCode)) {
    return false;
  }

  const issues = collectValidationIssues(
    output,
    required,
    mobileHeader,
    postalHeader,
    normalizeEmail(getGeneratedEmail(row, mapping))
  );

  return issues.length === 0;
}

export function filterRowsForExport(
  rows: DataRow[],
  mapping: MappingConfig,
  required: RequiredConfig,
  conditionColumn: string,
  mobileHeader: string,
  postalHeader: string,
  emailByDni: Map<string, string>,
  emailByCode: Map<string, string>,
  getGeneratedEmail: (row: DataRow, mapping: MappingConfig) => string
): DataRow[] {
  return rows.filter((row) =>
    rowMeetsRequired(
      row,
      mapping,
      required,
      conditionColumn,
      mobileHeader,
      postalHeader,
      emailByDni,
      emailByCode,
      getGeneratedEmail
    )
  );
}

export function buildPlannedEmailCounts(
  rows: DataRow[],
  mapping: MappingConfig,
  emailOverrides: Record<string, string>,
  getGeneratedEmail: (row: DataRow, mapping: MappingConfig) => string
): Map<string, number> {
  const counts = new Map<string, number>();
  rows.forEach((row) => {
    const key = String(row.__rowNumber);
    const override = emailOverrides[key];
    const email = normalizeEmail(override || getGeneratedEmail(row, mapping));
    if (!email) {
      return;
    }
    counts.set(email, (counts.get(email) ?? 0) + 1);
  });
  return counts;
}

export function filterRowsByEmail(
  rows: DataRow[],
  mapping: MappingConfig,
  emailOverrides: Record<string, string>,
  existingEmails: Set<string>,
  plannedEmailCounts: Map<string, number>,
  getGeneratedEmail: (row: DataRow, mapping: MappingConfig) => string
): DataRow[] {
  const seen = new Set<string>();
  return rows.filter((row) => {
    const key = String(row.__rowNumber);
    const override = emailOverrides[key];
    const email = normalizeEmail(override || getGeneratedEmail(row, mapping));
    if (!email) {
      return false;
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

export function findInternalDuplicateRows(
  rows: DataRow[],
  mapping: MappingConfig,
  emailOverrides: Record<string, string>,
  existingEmails: Set<string>,
  plannedEmailCounts: Map<string, number>,
  getGeneratedEmail: (row: DataRow, mapping: MappingConfig) => string
): InternalDuplicateRow[] {
  return rows
    .map((row) => {
      const key = String(row.__rowNumber);
      const output = buildOutputRow(row, mapping);
      const generatedEmail = normalizeEmail(getGeneratedEmail(row, mapping));
      const finalEmail = normalizeEmail(emailOverrides[key] || generatedEmail);
      if (!finalEmail || existingEmails.has(finalEmail)) {
        return null;
      }
      if ((plannedEmailCounts.get(finalEmail) ?? 0) <= 1) {
        return null;
      }
      return {
        rowNumber: row.__rowNumber,
        nombre: toCleanString(output["Nombre"]),
        apellido: toCleanString(output["Apellido"]),
        generatedEmail,
        finalEmail
      };
    })
    .filter(Boolean) as InternalDuplicateRow[];
}
