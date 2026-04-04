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
5. (Opcional) Sube el CSV exportado de Outlook para validar alumnos con correo existente y correos ya usados.
6. Pulsa "Generar CSV" y revisa el preview.
7. Descarga `contactos_outlook.csv` (si supera 249 filas se divide en varios archivos).

Nota: el CSV solo incluye filas con `CONDICION = INGRESO`, respeta los campos obligatorios activados y excluye alumnos que ya tengan correo en Outlook por `Fax` (DNI) o `Postal code` (codigo de estudiante), correos ya registrados en el CSV de Outlook y duplicados internos.
Para esta validacion, el registro de Outlook solo bloquea la exportacion si tiene `User principal name`. En esos casos se muestra una tabla con el alumno detectado, el tipo de coincidencia y el correo existente.
Si el Excel trae `a_paterno` y `a_materno`, el campo `Apellido` se arma con ambos y el correo se genera usando solo `a_paterno` completo. Asi, `De la Cruz` pasa a `delacruz` en el usuario.

## Cambiar defaults de mapeo

Los defaults se definen en `src/app/page.tsx` dentro de la funcion `buildDefaultMapping`. Puedes ajustar:

- Columnas sugeridas para Nombre, Apellido, `a_paterno`, `a_materno`, Celular, DNI, Codigo y Correo personal.
- Valores fijos como `Puesto` o `Pais o region`.

El orden y header del CSV final esta definido en `src/lib/csv.ts` (`OUTLOOK_HEADERS`), basado en `plantilla.csv`.
El CSV se genera con BOM UTF-8 y delimitador por comas para que Outlook y Excel lo lean correctamente.
Si el CSV de Outlook contiene un correo ya usado (columna `User principal name`), se mostrara una tabla para editar el correo final y validar disponibilidad.
El cruce con Outlook acepta encabezados `Fax`, `User principal name`, `Postal code` y `Codigo postal`.
