# Generador de CSV Outlook (UAI)

Aplicacion web para convertir un XLSX con datos de estudiantes en un CSV compatible con importacion de contactos en Outlook.

## Instalacion

```bash
npm install
```

## Desarrollo

```bash
npm run dev
```

## Uso rapido

1. Abre `http://localhost:3000`.
2. Sube el XLSX con los datos de estudiantes.
3. Selecciona la hoja (si aplica) y revisa el preview.
4. Configura el mapeo de columnas y los campos obligatorios.
5. (Opcional) Sube el CSV exportado de Outlook para validar DNI y correos existentes.
6. Pulsa "Generar CSV" y revisa el preview.
7. Descarga `contactos_outlook.csv` (si supera 249 filas se divide en varios archivos).

Nota: el CSV solo incluye filas con `CONDICIÓN = INGRESÓ`, respeta los campos obligatorios activados y excluye DNIs, correos ya registrados en el CSV de Outlook y duplicados internos.
Para la validación se usan las columnas `Fax` (DNI) y `User principal name` (correo).

## Cambiar defaults de mapeo

Los defaults se definen en `src/app/page.tsx` dentro de la funcion `buildDefaultMapping`. Puedes ajustar:

- Columnas sugeridas para Nombre, Apellido, Celular, DNI, Codigo y Correo personal.
- Valores fijos como `Puesto` o `Pais o region`.

El orden y header del CSV final esta definido en `src/lib/csv.ts` (`OUTLOOK_HEADERS`), basado en `plantilla.csv`.
El CSV se genera con BOM UTF-8 y delimitador por comas para que Outlook y Excel lo lean correctamente.
Si el CSV de Outlook contiene un correo ya usado (columna `User principal name`), se mostrara una tabla para editar el correo final y validar disponibilidad.
