import { digitsOnly, stripAccents, toCleanString } from "./normalization";
import { findBestColumn } from "./matching";

export type OutlookIdentifierConflict = {
  kind: "dni" | "codigo";
  value: string;
  emails: string[];
};

export type OutlookDirectoryData = {
  rowCount: number;
  emails: Set<string>;
  emailByDni: Map<string, string>;
  emailByCode: Map<string, string>;
  inconsistencies: OutlookIdentifierConflict[];
  error: string | null;
};

export function normalizeEmail(value: string): string {
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

export function normalizeStudentCode(value: string): string {
  return stripAccents(toCleanString(value)).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

export function buildCorrelativeEmail(baseEmail: string, usedEmails: Set<string>): string {
  const normalized = normalizeEmail(baseEmail);
  if (!normalized) {
    return "";
  }
  const [local, domain] = normalized.split("@");
  if (!local || !domain) {
    return "";
  }
  for (let index = 2; index < 1000; index += 1) {
    const candidate = `${local}${index}@${domain}`;
    if (!usedEmails.has(candidate)) {
      return candidate;
    }
  }
  return "";
}

function collectIdentifierConflicts(entries: Map<string, Set<string>>, kind: "dni" | "codigo") {
  return Array.from(entries.entries())
    .filter(([, emails]) => emails.size > 1)
    .map(([value, emails]) => ({
      kind,
      value,
      emails: Array.from(emails).sort()
    }));
}

export function buildOutlookDirectoryData(
  rows: Record<string, string>[],
  fields: string[]
): OutlookDirectoryData {
  const faxField = findBestColumn(fields, ["Fax", "FAX"]);
  const postalField = findBestColumn(fields, [
    "Postal code",
    "PostalCode",
    "Codigo postal",
    "Codigo estudiante"
  ]);
  const upnField = findBestColumn(fields, [
    "User principal name",
    "UserPrincipalName",
    "User principalname"
  ]);

  if (!upnField || (!faxField && !postalField)) {
    return {
      rowCount: rows.length,
      emails: new Set(),
      emailByDni: new Map(),
      emailByCode: new Map(),
      inconsistencies: [],
      error:
        "No se encontraron User principal name y al menos una columna de cruce (Fax o Postal code) en el CSV."
    };
  }

  const emails = new Set<string>();
  const emailByDni = new Map<string, string>();
  const emailByCode = new Map<string, string>();
  const dniCandidates = new Map<string, Set<string>>();
  const codeCandidates = new Map<string, Set<string>>();

  rows.forEach((row) => {
    const email = normalizeEmail(row[upnField] ?? "");
    if (email) {
      emails.add(email);
    }
    if (!email) {
      return;
    }
    if (faxField) {
      const dni = digitsOnly(toCleanString(row[faxField] ?? ""));
      if (dni) {
        if (!dniCandidates.has(dni)) {
          dniCandidates.set(dni, new Set());
        }
        dniCandidates.get(dni)?.add(email);
        if (!emailByDni.has(dni)) {
          emailByDni.set(dni, email);
        }
      }
    }
    if (postalField) {
      const codigo = normalizeStudentCode(row[postalField] ?? "");
      if (codigo) {
        if (!codeCandidates.has(codigo)) {
          codeCandidates.set(codigo, new Set());
        }
        codeCandidates.get(codigo)?.add(email);
        if (!emailByCode.has(codigo)) {
          emailByCode.set(codigo, email);
        }
      }
    }
  });

  return {
    rowCount: rows.length,
    emails,
    emailByDni,
    emailByCode,
    inconsistencies: [
      ...collectIdentifierConflicts(dniCandidates, "dni"),
      ...collectIdentifierConflicts(codeCandidates, "codigo")
    ],
    error: null
  };
}
