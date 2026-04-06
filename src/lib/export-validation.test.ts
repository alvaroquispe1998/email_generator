import { describe, expect, it } from "vitest";
import {
  buildPlannedEmailCounts,
  filterRowsByEmail,
  findInternalDuplicateRows,
  validateRows,
  type RequiredConfig
} from "./export-validation";
import { buildOutlookDirectoryData } from "./outlook";
import type { MappingConfig } from "./csv";
import type { DataRow } from "./xlsx";

const mapping: MappingConfig = {
  "Nombre de usuario": { type: "generated", value: "username" },
  Nombre: { type: "column", value: "NOMBRES" },
  "Apellido paterno": { type: "column", value: "A_PATERNO" },
  "Apellido materno": { type: "column", value: "A_MATERNO" },
  Apellido: { type: "generated", value: "fullSurname" },
  "Nombre para mostrar": { type: "generated", value: "displayName" },
  Puesto: { type: "fixed", value: "Estudiante" },
  Departamento: { type: "fixed", value: "" },
  "Número del trabajo": { type: "fixed", value: "" },
  "Teléfono de la oficina": { type: "fixed", value: "" },
  "Teléfono móvil": { type: "column", value: "CELULAR" },
  Fax: { type: "column", value: "DNI" },
  "Dirección de correo electrónico alternativa": { type: "fixed", value: "" },
  Dirección: { type: "fixed", value: "" },
  Ciudad: { type: "fixed", value: "" },
  "Estado o provincia": { type: "fixed", value: "" },
  "Código postal": { type: "column", value: "CODIGO" },
  "País o región": { type: "fixed", value: "Perú" }
};

const required: RequiredConfig = {
  dni: true,
  celular: true,
  codigo: true
};

describe("outlook directory parsing", () => {
  it("detects conflicting emails for the same identifier", () => {
    const directory = buildOutlookDirectoryData(
      [
        { Fax: "12345678", "Postal code": "A1", "User principal name": "uno@demo.pe" },
        { Fax: "12345678", "Postal code": "A1", "User principal name": "dos@demo.pe" }
      ],
      ["Fax", "Postal code", "User principal name"]
    );

    expect(directory.inconsistencies).toHaveLength(2);
    expect(directory.inconsistencies[0]?.emails).toContain("uno@demo.pe");
    expect(directory.inconsistencies[0]?.emails).toContain("dos@demo.pe");
  });
});

describe("export validations", () => {
  it("flags rows without institutional email", () => {
    const row: DataRow = {
      __rowNumber: 2,
      NOMBRES: "",
      A_PATERNO: "",
      A_MATERNO: "",
      CELULAR: "999999999",
      DNI: "12345678",
      CODIGO: "A1",
      CONDICION: "INGRESO"
    };

    const issues = validateRows(
      [row],
      mapping,
      required,
      "CONDICION",
      "Teléfono móvil",
      "Código postal",
      new Map(),
      new Map(),
      () => ""
    );

    expect(issues).toHaveLength(1);
    expect(issues[0]?.issues).toContain("Correo institucional");
  });

  it("removes rows with empty email and reports internal duplicates", () => {
    const rows: DataRow[] = [
      {
        __rowNumber: 2,
        NOMBRES: "Juan Luis",
        A_PATERNO: "Pérez",
        A_MATERNO: "Quispe",
        CELULAR: "999999999",
        DNI: "12345678",
        CODIGO: "A1",
        CONDICION: "INGRESO"
      },
      {
        __rowNumber: 3,
        NOMBRES: "Juan Luis",
        A_PATERNO: "Pérez",
        A_MATERNO: "Rojas",
        CELULAR: "988888888",
        DNI: "87654321",
        CODIGO: "A2",
        CONDICION: "INGRESO"
      }
    ];

    const getEmail = () => "juan.perez@autonomadeica.edu.pe";
    const counts = buildPlannedEmailCounts(rows, mapping, {}, getEmail);
    const duplicates = findInternalDuplicateRows(rows, mapping, {}, new Set(), counts, getEmail);
    const exportable = filterRowsByEmail(rows, mapping, {}, new Set(), counts, getEmail);

    expect(duplicates).toHaveLength(2);
    expect(exportable).toHaveLength(1);
  });
});
