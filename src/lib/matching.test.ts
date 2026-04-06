import { describe, expect, it } from "vitest";
import { findBestColumn, normalizeLookupKey } from "./matching";

describe("matching", () => {
  it("normalizes accents and separators", () => {
    expect(normalizeLookupKey("Código postal")).toBe("CODIGOPOSTAL");
    expect(normalizeLookupKey("A_PATERNO")).toBe("APATERNO");
  });

  it("prefers exact-ish matches over broader includes", () => {
    const column = findBestColumn(
      ["CODIGO MODULAR", "CODIGO DE ESTUDIANTE", "CODIGO"],
      ["CODIGO DE ESTUDIANTE", "CODIGO ESTUDIANTE", "CODIGO"]
    );
    expect(column).toBe("CODIGO DE ESTUDIANTE");
  });

  it("detects uppercase surname columns", () => {
    const column = findBestColumn(
      ["NOMBRES", "A_PATERNO", "A_MATERNO"],
      ["APELLIDO PATERNO", "A_PATERNO"]
    );
    expect(column).toBe("A_PATERNO");
  });
});
