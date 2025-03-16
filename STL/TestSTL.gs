/**
 * Script de prueba para la generación de archivos STL
 * Este archivo contiene funciones de prueba para verificar la correcta implementación
 * del formato STL según la especificación EBU Tech 3264
 */

/**
 * Ejecuta pruebas para verificar la generación de STL
 */
function testGenerateSTL() {
  try {
    Logger.log('Iniciando pruebas de generación STL...');
    
    // Crear datos de ejemplo
    var sampleData = createSampleSubtitles();
    
    // Crear metadatos de ejemplo
    var metadata = {
      englishTitle: "Big City Greens",
      spanishTitle: "Los Green en la Ciudad",
      episodeNumber: "F100",
      language: "es"
    };
    
    // Generar el STL
    var stlBlob = generateSTL(sampleData, metadata);
    
    // Verificar el resultado
    if (stlBlob && stlBlob.getBytes().length > 1024) {
      Logger.log('✅ Prueba exitosa: Archivo STL generado correctamente');
      Logger.log('  - Tamaño: ' + stlBlob.getBytes().length + ' bytes');
      return true;
    } else {
      Logger.log('❌ Prueba fallida: El archivo STL no tiene el tamaño esperado');
      return false;
    }
  } catch (error) {
    Logger.log('❌ Prueba fallida: ' + error.toString());
    return false;
  }
}

/**
 * Crea subtítulos de muestra para probar la generación de STL
 * @return {Array} Array con datos de muestra
 */
function createSampleSubtitles() {
  var sampleData = [];
  
  // Añadir algunos subtítulos de muestra
  sampleData.push(["00:00:05:00", "Este es un subtítulo de prueba", "00:00:08:00"]);
  sampleData.push(["00:00:09:00", "Subtítulo con acentos: á, é, í, ó, ú", "00:00:12:00"]);
  sampleData.push(["00:00:13:00", "Caracteres especiales: ñ, ¿, ¡", "00:00:16:00"]);
  sampleData.push(["00:00:17:00", "Subtítulo con\ndos líneas", "00:00:20:00"]);
  sampleData.push(["00:00:21:00", "Subtítulo muy largo que debería ser recortado si excede el límite de caracteres permitido", "00:00:25:00"]);
  
  return sampleData;
}

/**
 * Función para ejecutar manualmente la prueba desde el editor
 */
function runSTLTest() {
  testGenerateSTL();
}

/**
 * Prueba la generación de archivos STL en portugués
 */
function testGenerateSTLPortuguese() {
  try {
    Logger.log('Iniciando pruebas de generación STL en portugués...');
    
    // Crear datos de ejemplo
    var sampleData = [];
    
    // Añadir algunos subtítulos de muestra en portugués
    sampleData.push(["00:00:05:00", "Este é um exemplo de legenda", "00:00:08:00"]);
    sampleData.push(["00:00:09:00", "Legenda com acentos: ã, õ, ç", "00:00:12:00"]);
    sampleData.push(["00:00:13:00", "Caracteres especiais do português", "00:00:16:00"]);
    sampleData.push(["00:00:17:00", "Legenda com\nduas linhas", "00:00:20:00"]);
    
    // Crear metadatos de ejemplo
    var metadata = {
      englishTitle: "Big City Greens",
      spanishTitle: "Os Green na Cidade", // Título en portugués
      episodeNumber: "F100",
      language: "pt-BR"
    };
    
    // Generar el STL
    var stlBlob = generateSTL(sampleData, metadata);
    
    // Verificar el resultado
    if (stlBlob && stlBlob.getBytes().length > 1024) {
      Logger.log('✅ Prueba exitosa: Archivo STL en portugués generado correctamente');
      Logger.log('  - Tamaño: ' + stlBlob.getBytes().length + ' bytes');
      return true;
    } else {
      Logger.log('❌ Prueba fallida: El archivo STL en portugués no tiene el tamaño esperado');
      return false;
    }
  } catch (error) {
    Logger.log('❌ Prueba fallida: ' + error.toString());
    return false;
  }
} 