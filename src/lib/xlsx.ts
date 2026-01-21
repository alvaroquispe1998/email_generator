import * as XLSX from "xlsx";
import { toCleanString } from "./normalization";

export type DataRow = {
  __rowNumber: number;
  [key: string]: string | number;
};

export type SheetData = {
  name: string;
  columns: string[];
  rows: DataRow[];
};

export type WorkbookData = {
  sheets: SheetData[];
};

function ensureUnique(columns: string[]): string[] {
  const seen = new Map<string, number>();
  return columns.map((col) => {
    const base = col || "Columna";
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    if (count === 0) {
      return base;
    }
    return `${base} (${count + 1})`;
  });
}

export function parseWorkbook(arrayBuffer: ArrayBuffer): WorkbookData {
  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: false
  });

  const sheets = workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      raw: false
    });
    const headerRow = rows[0] ?? [];
    const columns = ensureUnique(
      headerRow.map((cell, index) => {
        const value = toCleanString(cell);
        return value || `Columna ${index + 1}`;
      })
    );

    const dataRows: DataRow[] = [];
    rows.slice(1).forEach((row, index) => {
      const record: DataRow = { __rowNumber: index + 2 };
      let hasValues = false;
      columns.forEach((column, colIndex) => {
        const value = toCleanString(row[colIndex]);
        if (value) {
          hasValues = true;
        }
        record[column] = value;
      });
      if (hasValues) {
        dataRows.push(record);
      }
    });

    return { name, columns, rows: dataRows };
  });

  return { sheets };
}
