# Árbol de Funciones - Conversor Excel a STL

Este documento presenta la estructura jerárquica de las funciones principales del conversor de Excel a STL, mostrando el flujo de ejecución y la relación entre los diferentes componentes del sistema.

## Punto de Entrada Principal

- **doGet(e)** - Función principal que maneja todas las solicitudes HTTP
  - Crea la interfaz web según los parámetros recibidos (página principal, ayuda o diagnóstico)
  - Detecta el idioma del usuario y configura la interfaz

## Flujo Principal de Conversión

```
processExcelBase64
├── importExcel
│   ├── convertExcelToGoogleSheet - Convierte Excel a Google Sheets
│   ├── processGoogleSheet - Procesa la hoja de cálculo convertida
│   │   ├── extractMetadata - Extrae títulos y números de episodio
│   │   ├── extractSubtitles - Extrae timecodes y textos de subtítulos
│   │   ├── preprocessSubtitleText - Limpia y formatea el texto de subtítulos
│   │   │   └── removeHTMLTags - Elimina etiquetas HTML preservando contenido
│   │   └── validateTimecodes - Verifica que los timecodes sean correctos
│   └── cleanupTempFiles - Elimina archivos temporales
└── generateExactReferenceSTL - Genera el archivo STL final
    ├── createExactGSIHeader - Crea el encabezado GSI (1024 bytes)
    │   └── writeASCIIToBuffer - Escribe texto ASCII en posiciones específicas
    ├── createExactTTIBlock - Crea bloques TTI para cada subtítulo (128 bytes)
    │   ├── parseTimecode - Convierte texto de timecode a valores numéricos
    │   ├── preprocessSubtitleText - Formatea texto de subtítulos
    │   └── debugBuffer - Registra contenido hexadecimal para depuración
    └── updateSTLHeaderCounters - Actualiza contadores en el encabezado
```

## Diagnóstico y Herramientas de Análisis

```
showSTLDiagnosticTool
├── analyzeSTLFile - Analiza un archivo STL existente
│   ├── analyzeGSIHeader - Analiza el encabezado GSI
│   │   └── extractGSIFields - Extrae campos específicos del encabezado
│   ├── analyzeTTIBlocks - Analiza los bloques TTI
│   │   └── analyzeBlock - Analiza un bloque TTI individual
│   └── generateCompatibilityReport - Analiza compatibilidad con editores
└── displayDiagnosticResults - Muestra resultados del análisis
```

## Configuración y Utilidades

```
getConfig - Obtiene configuración específica del sistema
├── CONFIG.stl - Configuración relacionada con el formato STL
├── CONFIG.excel - Configuración relacionada con procesamiento de Excel
└── CONFIG.limits - Límites de la aplicación

getTranslation - Obtiene traducciones según el idioma
cleanupMyTempFiles - Limpia archivos temporales antiguos
```

## Componentes Críticos

### 1. Procesamiento de Archivo Excel

- **importExcel**: Convierte archivos Excel a Google Sheets para su procesamiento
- **processGoogleSheet**: Extrae datos de subtítulos y metadatos de la hoja de cálculo

### 2. Limpieza y Formateo de Texto

- **preprocessSubtitleText**: Función crucial que limpia texto HTML y prepara subtítulos
  - Elimina etiquetas HTML (`<font>`, `<br>`, etc.)
  - Convierte `<br>` en saltos de línea
  - Formatea correctamente para el formato STL

### 3. Generación de Archivo STL

- **createExactGSIHeader**: Crea el encabezado GSI siguiendo el estándar EBU
  - Configura parámetros como página de código, formato, idioma, etc.
  - Establece metadatos como título, episodio, creador, etc.

- **createExactTTIBlock**: Crea bloques TTI para cada subtítulo
  - Establece número de subtítulo, timecodes, posición vertical
  - Procesa el texto y utiliza códigos EBU para caracteres especiales
  - Termina el bloque con ETX (0x03) y rellena con 0x00

### 4. Diagnóstico y Compatibilidad

- **showSTLDiagnosticTool**: Herramienta para analizar archivos STL
  - Verifica integridad del formato
  - Valida compatiblidad con diferentes editores de subtítulos
  - Muestra detalles técnicos y posibles problemas

## Flujo Detallado del Procesamiento

1. El usuario sube un archivo Excel a través de la interfaz web
2. El archivo se convierte a base64 en el navegador y se envía al servidor
3. Se convierte a Google Sheets para procesar su contenido
4. Se extraen metadatos (título, episodio) y subtítulos (timecodes y textos)
5. Se genera el archivo STL con encabezado GSI y bloques TTI
6. Se almacena temporalmente y se proporciona un enlace de descarga
7. Opcionalmente se realiza diagnóstico para verificar compatibilidad 