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

function compactWords(value: string): string {
  return normalizedWords(value).join("");
}

export function getLeadingWord(value: string): string {
  return firstWord(value);
}

export function buildUsername(nombres: string, apellidoParaUsuario: string): string {
  const firstName = firstWord(nombres);
  const surname = compactWords(apellidoParaUsuario);
  if (!firstName || !surname) {
    return "";
  }
  const local = `${firstName}.${surname}`.replace(/[^a-z0-9.]/g, "");
  return `${local}@autonomadeica.edu.pe`;
}

export function buildUsernameWithSecondName(
  nombres: string,
  apellidoParaUsuario: string
): string {
  const secondName = secondWord(nombres);
  const surname = compactWords(apellidoParaUsuario);
  if (!secondName || !surname) {
    return "";
  }
  const local = `${secondName}.${surname}`.replace(/[^a-z0-9.]/g, "");
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
