/**
 * @fileoverview Script para detectar y gestionar registros duplicados en hojas de cálculo.
 * 
 * Este script busca registros duplicados basándose en una columna clave definida en la estructura
 * de la base de datos. Puede operar en dos modos:
 * - Modo visualización: Solo muestra los duplicados encontrados
 * - Modo limpieza: Limpia el contenido de los registros duplicados más antiguos
 * 
 */

/**
 * Árbol de llamadas del script:
 * 
 * iniciarProceso()
 *   └── checkDuplicates() 
 *         ├── getKeyColumnForSheet() // Determina qué columna usar como clave para cada hoja
 *         └── procesarHoja()
 *               ├── checkTimeLeft()  // Verifica tiempo restante de ejecución
 *               └── guardarEstado()  // Guarda el estado para continuar después
 * 
 * Descripción de funciones:
 * 
 * - iniciarProceso: Punto de entrada principal. Inicia el proceso de detección/eliminación de duplicados
 *   con un ID de hoja específico.
 * 
 * - checkDuplicates: Función principal que controla el flujo del proceso. Maneja la apertura del documento,
 *   recuperación de estado y coordinación del procesamiento de cada hoja.
 * 
 * - procesarHoja: Procesa una hoja individual, detectando duplicados y eliminándolos si está habilitado
 *   el modo borrado. Implementa procesamiento por lotes para manejar timeouts.
 * 
 * - getKeyColumnForSheet: Determina qué columna usar como clave para buscar duplicados según el nombre
 *   de la hoja.
 * 
 * - checkTimeLeft: Calcula el tiempo restante de ejecución para prevenir timeouts.
 * 
 * - guardarEstado: Guarda el estado actual del proceso para poder continuar después de un timeout.
 */

const START_TIME = Date.now();
const MAX_RUNTIME = 8 * 60 * 1000; // 5 minutos en milisegundos
const BATCH_SIZE = 2000; // Número de filas a borrar por lote

// Función de entrada principal
function iniciarProceso() {
  const worksheetId = '175X3hrifGLN_UwPCFHF04mJ7YiVsRbCRZhudkGMrc0c';
  const modoBorrado = true;
  checkDuplicates(worksheetId, modoBorrado);
}

// Función principal de control
function checkDuplicates(worksheetId, modoBorrado = false) {
  try {
    // Abrir documento
    const spreadsheet = SpreadsheetApp.openById(worksheetId);
    if (!spreadsheet) {
      Logger.log('No se pudo abrir el documento');
      return false;
    }

    // Obtener todas las hojas y filtrar las excluidas
    const EXCLUDED_SHEETS = ['Index', 'DWO-LogLabels'];
    const sheets = spreadsheet.getSheets()
      .filter(sheet => !EXCLUDED_SHEETS.includes(sheet.getName()));
    
    // Recuperar estado si existe
    const estado = PropertiesService.getScriptProperties().getProperty('estado');
    let hojaInicial = 0;
    
    if (estado) {
      const estadoObj = JSON.parse(estado);
      // Encontrar índice de la hoja donde nos quedamos
      hojaInicial = sheets.findIndex(s => s.getName() === estadoObj.hoja);
      if (hojaInicial === -1) hojaInicial = 0;
    }

    // Procesar cada hoja
    for (let i = hojaInicial; i < sheets.length; i++) {
      const sheet = sheets[i];
      Logger.log(`Procesando hoja: ${sheet.getName()}`);
      
      // Determinar columna clave según la hoja
      const columnaIndice = getKeyColumnForSheet(sheet.getName());
      
      // Procesar la hoja
      const completado = procesarHoja(sheet, columnaIndice, modoBorrado);
      
      if (!completado) {
        return false; // Se interrumpió por timeout
      }
    }
    
    // Limpiar estado al completar
    PropertiesService.getScriptProperties().deleteProperty('estado');
    return true;
    
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    return false;
  }
}

// Función principal de procesamiento
function procesarHoja(sheet, columnaIndice, modoBorrado) {
  // Recuperar estado si existe
  const estado = PropertiesService.getScriptProperties().getProperty('estado');
  let filasABorrar = [];
  
  if (estado) {
    const estadoObj = JSON.parse(estado);
    if (estadoObj.hoja === sheet.getName() && estadoObj.filasABorrar) {
      // Recuperar filas pendientes de borrar
      filasABorrar = estadoObj.filasABorrar;
      Logger.log(`Recuperando ${filasABorrar.length} filas pendientes de borrar`);
    } else {
      // Si es una hoja nueva, detectar duplicados
      filasABorrar = detectarDuplicados(sheet, columnaIndice);
    }
  } else {
    // Si no hay estado, detectar duplicados
    filasABorrar = detectarDuplicados(sheet, columnaIndice);
  }

  // Si no hay duplicados o no estamos en modo borrado, terminar
  if (filasABorrar.length === 0 || !modoBorrado) {
    return true;
  }

  // Borrar por lotes
  for (let i = 0; i < filasABorrar.length; i += BATCH_SIZE) {
    // Verificar tiempo restante antes de procesar el lote
    if (checkTimeLeft() < 30) { // Si quedan menos de 30 segundos
      Logger.log(`Timeout durante limpieza - Guardando estado: Hoja ${sheet.getName()}, Fila ${i}`);
      
      // Guardar estado y filas restantes
      const filasRestantes = filasABorrar.slice(i);
      guardarEstado(sheet.getName(), i, filasRestantes);
      return false; // Indicar que no se completó
    }
    
    // Obtener el lote actual
    const lote = filasABorrar.slice(i, i + BATCH_SIZE);
    
    // Limpiar filas del lote actual
    lote.forEach(fila => {
      const numColumnas = sheet.getLastColumn();
      const rango = sheet.getRange(fila, 1, 1, numColumnas);
      rango.clearContent(); // Deja las celdas en blanco pero mantiene el formato
    });
  }
  
  return true; // Indicar que se completó
}

function detectarDuplicados(sheet, columnaIndice) {
  const datos = sheet.getDataRange().getValues();
  const duplicados = new Map();
  const filasABorrar = [];
  
  // Detectar duplicados
  for (let i = 1; i < datos.length; i++) {
    const valor = datos[i][columnaIndice];
    // Ignorar filas con valores vacíos o nulos
    if (!valor || valor.toString().trim() === '') {
      continue;
    }
    
    if (!duplicados.has(valor)) {
      duplicados.set(valor, [i + 1]); // Guardar número de fila (1-based)
    } else {
      duplicados.get(valor).push(i + 1);
    }
  }

  // Procesar duplicados encontrados
  for (const [valor, filas] of duplicados.entries()) {
    if (filas.length > 1) {
      Logger.log(`Valor duplicado "${valor}" encontrado en las filas: ${filas.join(', ')}`);
      
      // Mantener la primera ocurrencia, marcar el resto para borrar
      for (let i = 1; i < filas.length; i++) {
        filasABorrar.push(filas[i]);
      }
    }
  }

  // Ordenar filas de mayor a menor para no afectar índices al borrar
  filasABorrar.sort((a, b) => b - a);
  
  Logger.log(`Se encontraron ${filasABorrar.length} registros duplicados en ${sheet.getName()}`);
  
  return filasABorrar;
}

/**
 * Determina la columna clave para cada hoja.
 * 
 * Hojas que usan Columna B (índice 1):
 * - DWO: ID de Proyecto
 * - DWO-ChannelEventType: ID de Tipo de Evento
 * - DWO-Series: ID de Serie
 * 
 * Todas las demás hojas usan Columna A (índice 0):
 * - DWO-SynopsisProject: ID de Synopsis
 * - DWO-Production: ID de Producción
 * - DWO-SynopsisProduction: ID de Synopsis
 * - DWO-Event: ID de Evento
 * - etc.
 * 
 * @param {string} sheetName - Nombre de la hoja
 * @returns {number} Índice de la columna clave (0-based)
 */
function getKeyColumnForSheet(sheetName) {
  // Hojas que usan columna B como clave
  const usesColumnB = [
    'DWO',
    'DWO-ChannelEventType',
    'DWO-Series'
  ];
  
  // Retorna índice 1 (columna B) para las hojas especificadas, 0 (columna A) para el resto
  return usesColumnB.includes(sheetName) ? 1 : 0;
}

function checkTimeLeft() {
  const runtime = Date.now() - START_TIME;
  return (MAX_RUNTIME - runtime) / 1000; // Tiempo restante en segundos
}

function guardarEstado(nombreHoja, ultimaFila, filasRestantes) {
  // Guardar estado actual
  PropertiesService.getScriptProperties().setProperty('estado', JSON.stringify({
    hoja: nombreHoja,
    fila: ultimaFila,
    filasABorrar: filasRestantes,
    duplicadosDetectados: true
  }));

  // Eliminar triggers anteriores si existen
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Crear nuevo trigger para continuar en 1 minuto
  ScriptApp.newTrigger('iniciarProceso')
    .timeBased()
    .after(1 * 60 * 1000) // 1 minuto
    .create();
    
  Logger.log('Nuevo trigger creado para continuar el proceso en 1 minuto');
}

/**
 * Reinicia todas las variables de control del script.
 * Esto incluye:
 * - Eliminar la propiedad 'estado' del almacenamiento
 * - Eliminar todos los triggers existentes
 * 
 * @returns {boolean} true si se completó el reinicio correctamente
 */
function reiniciarControl() {
  try {
    // Eliminar estado guardado
    PropertiesService.getScriptProperties().deleteProperty('estado');
    
    // Eliminar todos los triggers existentes
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    
    Logger.log('Variables de control reiniciadas correctamente');
    return true;
  } catch (error) {
    Logger.log(`Error al reiniciar variables de control: ${error.message}`);
    return false;
  }
}