export function stripAccents(value: string): string {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizedWords(value: string): string[] {
  const cleaned = stripAccents(value)
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .trim();
  if (!cleaned) {
    return [];
  }
  return cleaned.split(/\s+/);
}

function firstWord(value: string): string {
  return normalizedWords(value)[0] ?? "";
}

function secondWord(value: string): string {
  return normalizedWords(value)[1] ?? "";
}

export function buildUsername(nombres: string, apellidos: string): string {
  const firstName = firstWord(nombres);
  const firstSurname = firstWord(apellidos);
  if (!firstName || !firstSurname) {
    return "";
  }
  const local = `${firstName}.${firstSurname}`.replace(/[^a-z0-9.]/g, "");
  return `${local}@autonomadeica.edu.pe`;
}

export function buildUsernameWithSecondName(nombres: string, apellidos: string): string {
  const secondName = secondWord(nombres);
  const firstSurname = firstWord(apellidos);
  if (!secondName || !firstSurname) {
    return "";
  }
  const local = `${secondName}.${firstSurname}`.replace(/[^a-z0-9.]/g, "");
  return `${local}@autonomadeica.edu.pe`;
}

export function digitsOnly(value: string): string {
  return value.replace(/\D/g, "");
}

export function toCleanString(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}
