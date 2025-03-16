# Documentación del Proyecto "Forced Narratives to STL" (FN2STL)

## Descripción General

"Forced Narratives to STL" (FN2STL) es una herramienta desarrollada en Google Apps Script que permite convertir subtítulos almacenados en hojas de cálculo de Google (Google Sheets) al formato EBU STL (European Broadcasting Union Subtitling data exchange format), siguiendo la especificación técnica EBU Tech 3264-1991.

Esta herramienta está diseñada para facilitar el proceso de exportación de subtítulos forzados (forced narratives) a un formato estándar de la industria audiovisual, permitiendo su integración en flujos de trabajo profesionales de subtitulado.

## Especificación Funcional

### Función Principal: `convertSheetToSTL`

#### Descripción
Convierte los datos de subtítulos de una hoja de cálculo de Google al formato EBU STL y guarda el archivo resultante en Google Drive.

#### Sintaxis
```javascript
function convertSheetToSTL(sheetId, country, languageCode, tempFolderId) {
  // Implementación de la función
}
```

#### Parámetros
- **sheetId** (String): ID de la hoja de cálculo de Google que contiene los subtítulos.
- **country** (String): Código del país de origen. Valores aceptados:
  - "ARG": Argentina
  - "BRA": Brasil
  - "MEX": México
- **languageCode** (String): Código del idioma. Valores aceptados:
  - "0A": Español
  - "21": Portugués
- **tempFolderId** (String): ID de la carpeta en Google Drive donde se guardarán los archivos temporales.

#### Valor de Retorno
- **Object**: Objeto que contiene información sobre el archivo STL generado:
  - `fileId`: ID del archivo en Google Drive
  - `fileName`: Nombre del archivo
  - `downloadUrl`: URL para descargar el archivo

## Estructura del Formato EBU STL

El formato EBU STL consta de dos bloques principales:

### 1. Bloque GSI (General Subtitle Information)
Contiene metadatos sobre el archivo de subtítulos, incluyendo:
- Código de país de origen (CO)
- Código de idioma (LC)
- Título del programa
- Información de tiempo y formato

### 2. Bloque TTI (Text and Timing Information)
Contiene la información de cada subtítulo individual, incluyendo:
- Número de subtítulo
- Tiempo de entrada (Time In)
- Tiempo de salida (Time Out)
- Texto del subtítulo
- Posición y formato

## Mapeo de Datos

### Especificación Detallada del Archivo de Entrada (Google Sheets)

#### Información General del Documento
- **Tipo de documento**: Lista de subtítulos para traducción/localización
- **Propósito**: Contiene subtítulos forzados (forced narratives) para su conversión a formato EBU STL
- **Nomenclatura**: Típicamente incluye información del programa (ejemplo: "Descendants_ShuffleOfLove_EP01_LAS_SL")

#### Estructura del Documento

##### Sección de Encabezado (Filas 1-10)
- **Fila 2**: Contiene "SUBTITLE LIST" en la columna B
- **Fila 4**: Contiene "ENGLISH" en columna A y el título original en columna C
- **Fila 5**: Contiene el idioma destino (ej. "SPANISH") en columna A y el título traducido en columna C
- **Fila 9**: Contiene los encabezados de la tabla de subtítulos:
  - Columna A: "Time code in"
  - Columna B: Idioma de destino (ej. "Spanish")
  - Columna C: "Time code out"
  - Columna D: "English"

##### Sección de Subtítulos (A partir de la Fila 11)
Cada fila representa un subtítulo con el siguiente formato:
- **Columna A**: Código de tiempo de entrada en formato HH:MM:SS:FF (horas:minutos:segundos:frames)
- **Columna B**: Texto del subtítulo en el idioma de destino (español o portugués)
- **Columna C**: Código de tiempo de salida en formato HH:MM:SS:FF
- **Columna D**: Texto del subtítulo en inglés (referencia, no se utiliza en la conversión)

#### Ejemplos de Subtítulos
Basado en el documento de ejemplo "Descendants_ShuffleOfLove_EP01_LAS_SL":

1. **Fila 11**:
   - Tiempo entrada: 01:00:00:16
   - Texto español: "El mazo del amor: un corto de Descendientes"
   - Tiempo salida: 01:00:04:16
   - Texto inglés: "Disney Shuffle of Love: A Descendants Short Story"

2. **Fila 12**:
   - Tiempo entrada: 01:02:45:03
   - Texto español: "Yo amo"
   - Tiempo salida: 01:02:46:13
   - Texto inglés: "I LOVE"

3. **Fila 13**:
   - Tiempo entrada: 01:02:49:00
   - Texto español: "Academia Merlín"
   - Tiempo salida: 01:02:49:20
   - Texto inglés: "THE MERLIN ACADEMY"

#### Notas sobre el Formato
- Los códigos de tiempo utilizan el estándar de la industria audiovisual (HH:MM:SS:FF)
- El último componente (FF) representa frames, considerando 24/25/30 frames por segundo dependiendo del estándar de video
- Los textos en el idioma destino (columna B) son los que se convierten al formato EBU STL
- Los textos en inglés (columna D) sirven como referencia pero no se incluyen en el archivo STL final

### Proceso de Conversión

1. La función lee la hoja de cálculo especificada
2. Extrae el título del programa de la sección de encabezado (fila 5, columna C)
3. Ignora las primeras 10 filas (sección de encabezado)
4. Para cada fila a partir de la 11:
   - Extrae el código de tiempo de entrada (columna A)
   - Extrae el texto en español o portugués (columna B)
   - Extrae el código de tiempo de salida (columna C)
   - Ignora el texto en inglés (columna D)
5. Crea el bloque GSI con los datos proporcionados (país, idioma, título)
6. Procesa cada fila de subtítulos para crear los bloques TTI
7. Combina los bloques en un archivo binario según la especificación EBU STL
8. Guarda el archivo en Google Drive con el mismo nombre que la hoja de cálculo pero con extensión .STL

## Consideraciones Técnicas

### Limitaciones
- El formato EBU STL tiene restricciones de caracteres por línea (normalmente 35-40)
- Soporta un conjunto limitado de caracteres especiales
- Los tiempos deben estar en el formato correcto y ser consistentes

### Manejo de Errores
- Validación de parámetros de entrada
- Verificación de formato de tiempos
- Manejo de caracteres no soportados

## Ejemplo de Uso

```javascript
// Ejemplo de cómo llamar a la función
function ejemploDeUso() {
  const sheetId = "1YlaOCinPTChLLxDF-Ce7mIyu2UAMHqGSKcf4t0yLCM0";
  const country = "ARG";
  const languageCode = "0A"; // Español
  const tempFolderId = "1aBcDeFgHiJkLmNoPqRsTuVwXyZ";
  
  const result = convertSheetToSTL(sheetId, country, languageCode, tempFolderId);
  
  Logger.log("Archivo generado: " + result.fileName);
  Logger.log("URL de descarga: " + result.downloadUrl);
}
```

## Especificación Técnica del Formato EBU STL

El formato EBU STL (EBU Tech 3264-1991) es un formato binario que sigue una estructura específica:

### Bloque GSI (1024 bytes)
- CPN (Code Page Number): 3 bytes
- DFC (Disk Format Code): 8 bytes
- DSC (Display Standard Code): 1 byte
- CCT (Character Code Table): 2 bytes
- LC (Language Code): 2 bytes
- OPT (Original Programme Title): 32 bytes
- OET (Original Episode Title): 32 bytes
- TPT (Translated Programme Title): 32 bytes
- TET (Translated Episode Title): 32 bytes
- TN (Translator Name): 32 bytes
- TCD (Translator Contact Details): 32 bytes
- SLR (Subtitle List Reference): 16 bytes
- CD (Creation Date): 6 bytes
- RD (Revision Date): 6 bytes
- RN (Revision Number): 2 bytes
- TNB (Total Number of TTI Blocks): 5 bytes
- TNS (Total Number of Subtitles): 5 bytes
- TNG (Total Number of Subtitle Groups): 3 bytes
- MNC (Maximum Number of Characters): 2 bytes
- MNR (Maximum Number of Rows): 2 bytes
- TCS (Time Code Status): 1 byte
- TCP (Time Code Start): 8 bytes
- TCF (Time Code End): 8 bytes
- CO (Country of Origin): 3 bytes
- Relleno: 75 bytes

### Bloque TTI (128 bytes por subtítulo)
- SGN (Subtitle Group Number): 1 byte
- SN (Subtitle Number): 2 bytes
- EBN (Extension Block Number): 1 byte
- CS (Cumulative Status): 1 byte
- TCI (Time Code In): 4 bytes
- TCO (Time Code Out): 4 bytes
- VP (Vertical Position): 1 byte
- JC (Justification Code): 1 byte
- CF (Comment Flag): 1 byte
- TF (Text Field): 112 bytes

## Recursos Adicionales

Para más información sobre el formato EBU STL, consulte:
- [Especificación EBU Tech 3264-1991](https://tech.ebu.ch/docs/tech/tech3264.pdf) 