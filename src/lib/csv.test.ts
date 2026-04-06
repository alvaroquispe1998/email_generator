import { describe, expect, it } from "vitest";
import {
  buildAlternateGeneratedUsername,
  buildGeneratedUsername,
  buildOutputRow,
  type MappingConfig
} from "./csv";
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
  "Teléfono móvil": { type: "fixed", value: "" },
  Fax: { type: "fixed", value: "" },
  "Dirección de correo electrónico alternativa": { type: "fixed", value: "" },
  Dirección: { type: "fixed", value: "" },
  Ciudad: { type: "fixed", value: "" },
  "Estado o provincia": { type: "fixed", value: "" },
  "Código postal": { type: "fixed", value: "" },
  "País o región": { type: "fixed", value: "Perú" }
};

describe("csv generation", () => {
  it("uses full paternal surname for the institutional email", () => {
    const row: DataRow = {
      __rowNumber: 2,
      NOMBRES: "Juan Carlos",
      A_PATERNO: "De la Cruz",
      A_MATERNO: "Quispe"
    };

    expect(buildGeneratedUsername(row, mapping)).toBe("juan.delacruz@autonomadeica.edu.pe");
    expect(buildAlternateGeneratedUsername(row, mapping)).toBe(
      "carlos.delacruz@autonomadeica.edu.pe"
    );
  });

  it("builds the full surname and display name from paternal and maternal surnames", () => {
    const row: DataRow = {
      __rowNumber: 2,
      NOMBRES: "Juan Carlos",
      A_PATERNO: "De la Cruz",
      A_MATERNO: "Quispe"
    };

    const output = buildOutputRow(row, mapping);
    expect(output.Apellido).toBe("De la Cruz Quispe");
    expect(output["Nombre para mostrar"]).toBe("De la Cruz Quispe Juan Carlos");
  });
});
