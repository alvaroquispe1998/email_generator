# Funcionamiento del sistema

## Objetivo

La aplicacion toma un archivo Excel con datos de estudiantes y genera un CSV compatible con la importacion de contactos en Outlook.

El flujo tiene 4 etapas principales:

1. Leer y normalizar el XLSX.
2. Mapear columnas del Excel hacia los campos de Outlook.
3. Validar que cada fila sea exportable.
4. Generar el CSV final y, si hace falta, dividirlo en partes.

## Archivos de entrada

### 1. XLSX principal

Se usa como fuente de alumnos.

Comportamiento:

- Lee todas las hojas del libro.
- Usa la primera fila como encabezados.
- Si hay encabezados repetidos, les agrega sufijos como `(2)` para hacerlos unicos.
- Descarta filas totalmente vacias.
- Guarda el numero de fila original del Excel en `__rowNumber` para mostrar errores y conflictos.

### 2. CSV de Outlook

Es opcional y sirve para detectar alumnos que ya tienen correo o correos que ya existen.

Busca estas columnas:

- `User principal name`: para conocer los correos ya usados.
- `Fax`: para cruzar por DNI.
- `Postal code` o variantes: para cruzar por codigo de estudiante.

Si no encuentra `User principal name` y al menos una columna de cruce (`Fax` o `Postal code`), muestra error y no usa ese CSV para validacion.

## Mapeo

El sistema trabaja con dos tipos de campos:

- Campos reales del CSV final de Outlook.
- Campos auxiliares de mapeo que existen solo dentro de la UI.

### Campos auxiliares

En el editor de mapeo existen:

- `Apellido paterno`
- `Apellido materno`

Estos no se exportan al CSV final. Se usan para construir:

- `Apellido`
- `Nombre de usuario`
- sugerencias de correo alternativas

### Defaults de mapeo

Al cargar una hoja, el sistema intenta detectar automaticamente columnas conocidas por nombre aproximado, por ejemplo:

- `NOMBRES`, `NOMBRE`
- `A_PATERNO`, `APELLIDO PATERNO`
- `A_MATERNO`, `APELLIDO MATERNO`
- `APELLIDOS`, `APELLIDOS COMPLETOS`
- `DNI`
- `CODIGO`
- `CORREO PERSONAL`

La deteccion no exige coincidencia exacta. Normaliza acentos, mayusculas y simbolos, prioriza coincidencias exactas o por tokens y solo luego cae a coincidencias mas amplias.

## Reglas de generacion

### Apellido

Prioridad:

1. Si hay `Apellido paterno` o `Apellido materno`, construye `Apellido = paterno + materno`.
2. Si no existen esos campos, usa el campo mapeado a `Apellido`.

### Nombre de usuario institucional

Regla base:

- usa `primer nombre + "." + apellido paterno`

Detalles:

- Si existe `Apellido paterno`, usa ese valor completo.
- Si el paterno tiene espacios, los compacta.
- Ejemplo: `De la Cruz` se convierte en `delacruz`.
- Si no existe `Apellido paterno`, cae al primer bloque del campo `Apellido`.

Ejemplo:

- `Juan Carlos`
- `a_paterno = De la Cruz`

Resultado:

- `juan.delacruz@autonomadeica.edu.pe`

### Correo alternativo sugerido

Cuando el correo generado ya existe, el sistema puede sugerir otra variante usando:

- `segundo nombre + "." + apellido paterno`

Si eso tampoco ayuda, puede proponer un correlativo:

- `juan.delacruz2@...`
- `juan.delacruz3@...`

### Nombre para mostrar

Se genera como:

- `Apellido + Nombre`

### Limpieza de datos

Antes de exportar:

- `Telefono movil` queda solo con digitos.
- `Fax` queda solo con digitos.
- Los correos se normalizan a minusculas y se limpian caracteres no validos.

## Validaciones

Una fila se considera exportable solo si pasa todas estas reglas.

### 1. Condicion de ingreso

Si existe una columna parecida a `CONDICION`, solo se exportan filas cuyo valor normalizado sea:

- `INGRESO`

Si no existe esa columna, no se filtra por condicion.

### 2. Campos obligatorios

Los obligatorios configurables en la UI son:

- DNI
- Celular
- Codigo de estudiante

Si alguno esta activado y falta, la fila no se exporta y aparece en la tabla de problemas.

### 2.1 Correo institucional obligatorio

Aunque no se configure como checkbox, el sistema exige que la fila pueda generar `Nombre de usuario`.

Si no puede construir el correo institucional:

- la fila no se exporta
- aparece con el problema `Correo institucional`

### 3. Cruce con Outlook por alumno existente

Si se cargo el CSV de Outlook, el sistema intenta encontrar al alumno por:

- DNI, usando `Fax`
- Codigo de estudiante, usando `Postal code`

Si encuentra coincidencia por cualquiera de los dos, la fila no se exporta.

La UI muestra:

- numero de fila
- DNI
- codigo
- nombre
- apellido
- motivo de coincidencia
- correo existente

### 4. Conflictos de correo ya existente

Si el correo generado ya aparece en el `User principal name` del CSV de Outlook:

- la fila entra en la tabla de conflictos
- puedes editar el correo manualmente
- puedes pedir sugerencias automaticas

### 4.1 Inconsistencias del CSV de Outlook

Si el CSV de Outlook tiene el mismo DNI o codigo asociado a mas de un correo:

- no bloquea por si solo la exportacion
- pero se muestra en una tabla de inconsistencias para revision
- esas inconsistencias pueden exportarse a un CSV aparte

### 5. Duplicados internos

Si dos filas terminan con el mismo correo final dentro de la misma exportacion:

- el sistema conserva una
- excluye las demas

Ahora tambien se muestra una tabla dedicada de duplicados internos para que puedas identificarlos antes de exportar.

## Como hace match

### Match de columnas

Para reconocer encabezados del Excel o del CSV de Outlook:

- quita acentos
- pasa a mayusculas
- elimina simbolos
- compara por prioridad de score:
- coincidencia exacta normalizada
- coincidencia por tokens
- prefijos o sufijos
- inclusion amplia como ultimo recurso

Eso permite detectar variantes como:

- `A_PATERNO`
- `a_paterno`
- `Apellido Paterno`

### Match de alumnos con Outlook

Se hace sobre los datos ya transformados por el mapeo:

- `Fax` para DNI
- `Postal code` para codigo

No cruza por nombre ni por correo personal.

### Match de correos

Los correos se normalizan antes de comparar:

- minusculas
- sin acentos
- solo letras, numeros y punto en local/domain

## Generacion del CSV final

Cuando pulsas `Generar CSV`:

- se toman solo las filas exportables
- se aplican overrides manuales de correo si existen
- se arma el CSV con el orden exacto definido por `OUTLOOK_HEADERS`
- se agrega BOM UTF-8
- se usa delimitador coma

Cuando pulsas `Descargar`:

- si hay hasta 249 filas, descarga un solo archivo
- si hay mas, lo divide en partes de 249 filas

## Persistencia local

En el navegador se guarda:

- el ultimo mapeo usado
- los campos obligatorios activados

Esto permite que la configuracion se mantenga entre recargas.

## Limitaciones actuales

- La deteccion automatica de columnas es mas robusta que antes, pero todavia puede equivocarse si el archivo tiene encabezados muy ambiguos o casi identicos.
- Si el CSV de Outlook tiene inconsistencias por DNI o codigo, el sistema las reporta, pero sigue usando la primera coincidencia encontrada para el cruce operativo.
- No hay todavia una suite de pruebas de interfaz; las pruebas actuales cubren la logica principal de matching, generacion y validacion.
