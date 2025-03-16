// Variables globales para caché
let cacheSheetsList = null;
// Variable para almacenar el HTML del Sidebar
let cacheSidebarHTML = null;

function mostrarSidebar() {
  // Si no hay caché de hojas, obtenerlas primero
  if (!verificarCacheHojas()) {
    actualizarCacheHojas();
  }
  
  // Crear el HTML del sidebar con los botones ya generados
  const htmlCompleto = crearHTMLSidebarConBotones();
  
  const html = HtmlService.createHtmlOutput(htmlCompleto)
    .setTitle('Navegador de Hojas');
  
  SpreadsheetApp.getUi().showSidebar(html);
}

function obtenerHojas() {
  // Verificar si tenemos una caché válida
  if (cacheSheetsList && cacheSheetsList.length > 0) {
    console.log("Usando caché de hojas");
    return cacheSheetsList;
  }
  
  // Si no hay caché, obtener datos frescos
  console.log("Obteniendo lista fresca de hojas");
  return actualizarCacheHojas();
}

// Función para actualizar la caché de hojas
function actualizarCacheHojas() {
  const hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // Obtener nombres y ordenarlos alfabéticamente
  const nombresHojas = hojas.map(hoja => hoja.getName());
  const listaOrdenada = nombresHojas.sort(); // Ordena alfabéticamente
  
  // Actualizar la caché
  cacheSheetsList = listaOrdenada;
  
  return listaOrdenada;
}

// Función para limpiar la caché (útil si se agregan/eliminan hojas)
function limpiarCacheHojas() {
  cacheSheetsList = null;
  // También limpiamos la caché del HTML para asegurarnos de que se regenere
  cacheSidebarHTML = null;
  return actualizarCacheHojas(); // Actualiza inmediatamente la caché
}

// Función para limpiar solo la caché del HTML
function limpiarCacheHTML() {
  cacheSidebarHTML = null;
  return true;
}

function navegarAHoja(nombreHoja) {
  // Intentar activar la hoja directamente
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (hoja) {
    hoja.activate();
    return true;
  } else {
    // Si no se encuentra la hoja, podría ser que la caché esté desactualizada
    // Actualizamos la caché y verificamos si la hoja existe en la lista actualizada
    actualizarCacheHojas();
    
    // Verificamos si después de actualizar la caché, la hoja aparece en la lista
    if (cacheSheetsList && cacheSheetsList.includes(nombreHoja)) {
      // Intentamos nuevamente activar la hoja
      const hojaActualizada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
      if (hojaActualizada) {
        hojaActualizada.activate();
        return true;
      }
    }
    
    // Si aún no se encuentra, mostramos el error
    SpreadsheetApp.getUi().alert('No se pudo encontrar la hoja: ' + nombreHoja);
    return false;
  }
}

// Función para crear el HTML completo con los botones ya generados
function crearHTMLSidebarConBotones() {
  // Obtener la lista de hojas (ya sea de la caché o del servidor)
  const hojas = obtenerHojas();
  
  // Generar los botones HTML directamente
  let botonesHTML = '';
  if (hojas && hojas.length > 0) {
    hojas.forEach(function(nombreHoja) {
      botonesHTML += `<button onclick="google.script.run.navegarAHoja('${nombreHoja}');">${nombreHoja}</button>`;
    });
  } else {
    botonesHTML = '<p style="text-align:center">No se encontraron hojas</p>';
  }
  
  // Crear el HTML completo
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 10px; }
    button { 
      margin: 5px 0; 
      width: 100%; 
      padding: 8px;
      background-color: #4285f4;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
    }
    button:hover {
      background-color: #3367d6;
    }
    .refresh-btn {
      background-color: #34a853;
      margin-top: 15px;
      font-size: 0.9em;
    }
    .refresh-btn:hover {
      background-color: #2e8b57;
    }
  </style>
</head>
<body>
  <div id="contenedor">
    ${botonesHTML}
  </div>
  <button id="refreshBtn" class="refresh-btn">🔄 Actualizar Lista de Hojas</button>
  <script>
    document.addEventListener('DOMContentLoaded', function() {
      document.getElementById('refreshBtn').addEventListener('click', function() {
        this.textContent = '⏳ Actualizando...';
        this.disabled = true;
        
        // Mostrar mensaje de actualización en el contenedor
        document.getElementById('contenedor').innerHTML = '<p style="text-align:center">Actualizando lista de hojas...</p>';
        
        // Limpiar la caché y mostrar el sidebar nuevamente
        google.script.run
          .withSuccessHandler(function(result) {
            // Llamar a mostrarSidebar nuevamente para regenerar todo
            google.script.run.mostrarSidebar();
          })
          .withFailureHandler(function(error) {
            document.getElementById('refreshBtn').textContent = '❌ Error al actualizar';
            document.getElementById('refreshBtn').disabled = false;
            document.getElementById('contenedor').innerHTML = '<p style="text-align:center; color:red;">Error al actualizar. Intente nuevamente.</p>';
            setTimeout(function() {
              document.getElementById('refreshBtn').textContent = '🔄 Actualizar Lista de Hojas';
            }, 3000);
          })
          .limpiarCacheHojas();
      });
    });
  </script>
</body>
</html>`;
}

// Función para verificar si ya tenemos hojas en caché
function verificarCacheHojas() {
  return cacheSheetsList !== null && cacheSheetsList.length > 0;
}