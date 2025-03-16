/**
 * Archivo con las traducciones para internacionalización
 */

// Objeto con las traducciones para cada idioma soportado
const TRANSLATIONS = {
  // Español (idioma predeterminado)
  'es': {
    // Textos de la interfaz principal
    'app_title': 'Conversor Excel a STL',
    'instructions_title': 'Instrucciones',
    'instructions_item1': 'Suba un archivo Excel (XLS o XLSX) con formato de subtítulos.',
    'instructions_item2': 'El archivo debe tener los timecodes y diálogos a partir de la fila 11.',
    'instructions_item3': 'La columna B debe contener el Timecode de entrada.',
    'instructions_item4': 'La columna C debe contener el texto del subtítulo.',
    'instructions_item5': 'La columna D debe contener el Timecode de salida.',
    'drop_files': 'Arrastre y suelte su archivo Excel aquí',
    'or': 'o',
    'select_file': 'Seleccionar archivo',
    'processing': 'Procesando archivo...',
    'convert_stl': 'Convirtiendo a STL...',
    'conversion_complete': 'Conversión completada',
    'download_btn': 'Descargar archivo STL',
    'format_info': 'Información del formato:',
    'format_description': 'El formato STL (EBU Subtitle Data Exchange format) es un estándar para el intercambio de subtítulos desarrollado por la Unión Europea de Radiodifusión (EBU).',
    'help': 'Ayuda',
    'copyright': 'Conversor Excel a STL © {0} - Media Access Co.',
    'more_details': 'Para más detalles consulte la',
    'help_page': 'página de ayuda',
    
    // Mensajes de estado
    'success_message': 'El archivo fue convertido exitosamente a formato STL.',
    'error_format': 'Por favor, seleccione un archivo Excel válido (.xls o .xlsx).',
    'error_read': 'Error al leer el archivo: ',
    'warning_excel': 'ADVERTENCIA: El archivo Excel no pudo ser interpretado correctamente. Se ha generado un archivo STL con un mensaje de error.',
    'notice_sample': 'AVISO: Debido a limitaciones técnicas, se ha generado un archivo STL de muestra. Este archivo contiene únicamente subtítulos de ejemplo.',
    'no_file': 'No hay archivo STL disponible para descargar.',
    'download_started': 'Descarga iniciada. Puede cargar un nuevo archivo cuando desee.',
    
    // Mensajes de limpieza de archivos temporales
    'cleanup_files': 'Limpiar mis archivos temporales',
    'confirm_cleanup': '¿Estás seguro de que deseas eliminar tus archivos temporales?'
  },
  
  // Portugués brasileño
  'pt-BR': {
    // Textos da interface principal
    'app_title': 'Conversor Excel para STL',
    'instructions_title': 'Instruções',
    'instructions_item1': 'Carregue um arquivo Excel (XLS ou XLSX) com formato de legendas.',
    'instructions_item2': 'O arquivo deve ter os timecodes e diálogos a partir da linha 11.',
    'instructions_item3': 'A coluna B deve conter o Timecode de entrada.',
    'instructions_item4': 'A coluna C deve conter o texto da legenda.',
    'instructions_item5': 'A coluna D deve conter o Timecode de saída.',
    'drop_files': 'Arraste e solte seu arquivo Excel aqui',
    'or': 'ou',
    'select_file': 'Selecionar arquivo',
    'processing': 'Processando arquivo...',
    'convert_stl': 'Convertendo para STL...',
    'conversion_complete': 'Conversão concluída',
    'download_btn': 'Baixar arquivo STL',
    'format_info': 'Informação do formato:',
    'format_description': 'O formato STL (EBU Subtitle Data Exchange format) é um padrão para o intercâmbio de legendas desenvolvido pela União Europeia de Radiodifusão (EBU).',
    'help': 'Ajuda',
    'copyright': 'Conversor Excel para STL © {0} - Media Access Co.',
    'more_details': 'Para mais detalhes consulte a',
    'help_page': 'página de ajuda',
    
    // Mensagens de estado
    'success_message': 'O arquivo foi convertido com sucesso para o formato STL.',
    'error_format': 'Por favor, selecione um arquivo Excel válido (.xls ou .xlsx).',
    'error_read': 'Erro ao ler o arquivo: ',
    'warning_excel': 'AVISO: O arquivo Excel não pôde ser interpretado corretamente. Foi gerado um arquivo STL com uma mensagem de erro.',
    'notice_sample': 'AVISO: Devido a limitações técnicas, foi gerado um arquivo STL de amostra. Este arquivo contém apenas legendas de exemplo.',
    'no_file': 'Não há arquivo STL disponível para download.',
    'download_started': 'Download iniciado. Você pode carregar um novo arquivo quando quiser.',
    
    // Mensagens de limpeza de arquivos temporários
    'cleanup_files': 'Limpar meus arquivos temporários',
    'confirm_cleanup': 'Tem certeza de que deseja excluir seus arquivos temporários?'
  }
};

/**
 * Obtiene la traducción para una clave específica en el idioma especificado
 * @param {String} key - Clave de traducción
 * @param {String} language - Código de idioma ('es' o 'pt-BR')
 * @param {Array} params - Parámetros opcionales para sustituir en el texto
 * @return {String} Texto traducido
 */
function getTranslation(key, language, params) {
  // Si el idioma no está soportado, usar español por defecto
  if (!TRANSLATIONS[language]) {
    language = 'es';
  }
  
  // Obtener el texto traducido
  var text = TRANSLATIONS[language][key] || TRANSLATIONS['es'][key] || key;
  
  // Reemplazar parámetros si existen
  if (params && params.length > 0) {
    for (var i = 0; i < params.length; i++) {
      text = text.replace('{' + i + '}', params[i]);
    }
  }
  
  return text;
}

/**
 * Obtiene todas las traducciones para un idioma específico
 * @param {String} language - Código de idioma ('es' o 'pt-BR')
 * @return {Object} Objeto con todas las traducciones
 */
function getAllTranslations(language) {
  // Si el idioma no está soportado, usar español por defecto
  if (!TRANSLATIONS[language]) {
    language = 'es';
  }
  
  return TRANSLATIONS[language];
} 