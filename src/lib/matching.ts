import { stripAccents } from "./normalization";

export function normalizeLookupKey(value: string): string {
  return stripAccents(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function tokenizeLookupValue(value: string): string[] {
  const cleaned = stripAccents(value)
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .trim();
  if (!cleaned) {
    return [];
  }
  return cleaned.split(/\s+/);
}

function scoreColumnMatch(column: string, candidate: string): number {
  const normalizedColumn = normalizeLookupKey(column);
  const normalizedCandidate = normalizeLookupKey(candidate);
  if (!normalizedColumn || !normalizedCandidate) {
    return -1;
  }
  if (normalizedColumn === normalizedCandidate) {
    return 1000;
  }

  const columnTokens = tokenizeLookupValue(column);
  const candidateTokens = tokenizeLookupValue(candidate);
  const matchedTokens = candidateTokens.filter((token) =>
    columnTokens.includes(token)
  ).length;

  if (matchedTokens === candidateTokens.length && matchedTokens > 0) {
    return 700 + matchedTokens * 20 - Math.max(0, columnTokens.length - matchedTokens);
  }

  if (normalizedColumn.startsWith(normalizedCandidate)) {
    return 500;
  }
  if (normalizedColumn.endsWith(normalizedCandidate)) {
    return 450;
  }
  if (normalizedColumn.includes(normalizedCandidate)) {
    return 300;
  }

  return matchedTokens > 0 ? matchedTokens * 50 : -1;
}

export function findBestColumn(columns: string[], candidates: string[]): string {
  let bestColumn = "";
  let bestScore = -1;

  columns.forEach((column) => {
    candidates.forEach((candidate) => {
      const score = scoreColumnMatch(column, candidate);
      if (score > bestScore) {
        bestScore = score;
        bestColumn = column;
      }
    });
  });

  return bestColumn;
}
