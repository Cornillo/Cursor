/**
 * Archivo de configuración para el Conversor Excel a STL
 */

// Configuración general de la aplicación
const CONFIG = {
  // Versión de la aplicación
  version: '1.1.0',
  
  // Configuración STL
  stl: {
    // Formato de código STL (25fps para PAL, 30fps para NTSC)
    format: 'STL25.01',
    
    // Código de idioma (ISO 639-2)
    languageCode: {
      'es': 'ESP',
      'pt-BR': 'POR'
    },
    
    // Código de idioma EBU 
    ebuLanguageCode: {
      'es': '0A',
      'pt-BR': '0F'
    },
    
    // Códigos de página
    codePage: '850',
    
    // Código de país (ISO 3166)
    countryCode: {
      'es': 'ES',
      'pt-BR': 'BR'
    },
    
    // Tabla de caracteres
    charTable: '00', // Latin
    
    // Tamaño del GSI header en bytes
    gsiSize: 1024,
    
    // Tamaño de cada bloque TTI en bytes
    ttiSize: 128,
    
    // Justificación de texto
    justification: {
      left: 1,
      center: 2,
      right: 3
    },
    
    // Posición vertical para subtítulos
    verticalPosition: {
      singleLine: 16,    // Una línea
      twoLinesFirst: 14, // Primera línea de dos
      twoLinesSecond: 15 // Segunda línea de dos
    }
  },
  
  // Configuración de Excel
  excel: {
    // Fila donde comienzan los datos de subtítulos (1-indexado)
    startRow: 11,
    
    // Índices de columnas (0-indexado)
    columns: {
      timeIn: 1,   // Columna B
      text: 2,     // Columna C
      timeOut: 3   // Columna D
    },
    
    // Filas de metadatos (1-indexado)
    metadata: {
      titleEnglish: 3,   // Título en inglés
      titleSpanish: 4,   // Título en español
      episodeNumber: 5   // Número de episodio
    }
  },
  
  // Límites de la aplicación
  limits: {
    // Tamaño máximo de archivo en bytes (5MB)
    maxFileSize: 5 * 1024 * 1024,
    
    // Número máximo de subtítulos por archivo
    maxSubtitles: 5000
  },
  
  // Configuración de carpetas
  folders: {
    // Retención en días para archivos temporales (después de esto se limpian)
    tempRetentionDays: 7
  }
};

/**
 * Obtiene toda la configuración o una sección específica
 * @param {String} section - Sección opcional de configuración
 * @return {Object} Configuración solicitada
 */
function getConfig(section) {
  if (section && CONFIG[section]) {
    return CONFIG[section];
  }
  return CONFIG;
} 