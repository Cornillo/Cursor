# Función `debugBuffer` - Análisis Detallado

## Descripción General

La función `debugBuffer` es una herramienta crítica de diagnóstico que permite visualizar el contenido hexadecimal de buffers binarios, especialmente bloques TTI (Text and Timing Information) y encabezados GSI (General Subtitle Information) dentro del formato STL.

## Sintaxis

```javascript
function debugBuffer(buffer, label) {
  // Implementación
}
```

### Parámetros
- **buffer**: Array de bytes que se desea visualizar (típicamente un bloque TTI de 128 bytes o parte del GSI)
- **label**: Etiqueta descriptiva que se mostrará en los logs para identificar el buffer analizado

## Implementación Detallada

La función realiza las siguientes operaciones:

1. **Inicialización**:
   - Crea arrays para almacenar representaciones hexadecimales y ASCII
   - Configura constantes para el formato de visualización

2. **Procesamiento por líneas**:
   - Recorre cada byte del buffer
   - Genera representación hexadecimal de cada byte
   - Crea representación ASCII usando caracteres imprimibles
   - Utiliza caracteres especiales para representar códigos de control de EBU

3. **Registro en el log**:
   - Registra cada línea formateada en el Logger
   - Muestra offset hexadecimal + bytes en hexadecimal + representación ASCII
   - Incluye un resumen final con la etiqueta proporcionada

## Códigos Especiales Representados

| Byte | Significado en STL | Representación en debugBuffer |
|------|--------------------|-----------------------------|
| 0x8A | Salto de línea EBU | ¶ (símbolo de párrafo)     |
| 0x03 | ETX (Fin de texto) | . (punto)                   |
| 0x00 | Byte nulo          | . (punto)                   |

## Formato de Salida

El formato de salida típico en los logs es:

```
0000: xx xx xx xx xx xx xx xx xx xx xx xx xx xx xx xx | ................
0010: xx xx xx xx xx xx xx xx xx xx xx xx xx xx xx xx | ................
...
[label] - Final: xxxx: xx xx xx xx ... | ....
```

Donde:
- `xxxx` representa el offset en hexadecimal
- `xx` son los valores de bytes en hexadecimal
- Los caracteres después de `|` son la representación ASCII

## Ejemplo de Uso

```javascript
// Dentro de createExactTTIBlock después de procesar un bloque TTI
debugBuffer(buffer, "TTI Block " + subtitleNumber);

// Dentro de funciones de diagnóstico al analizar GSI
debugBuffer(gsiBytes.slice(0, 128), "GSI Header (primeros 128 bytes)");
```

## Importancia en el Flujo de Trabajo

Esta función es fundamental para:

1. **Depuración de bloques TTI incorrectos**:
   - Permite identificar problemas en la codificación de caracteres
   - Ayuda a verificar que los timecodes están correctamente formateados
   - Confirma la correcta terminación del texto con ETX (0x03)

2. **Verificación de compatibilidad con editores**:
   - Los editores de subtítulos pueden interpretar de manera diferente ciertos bytes
   - `debugBuffer` permite inspeccionar exactamente qué bytes están presentes
   - Facilita comparar con archivos de referencia que funcionan correctamente

3. **Validación del formato STL**:
   - Confirma que los bytes del encabezado GSI siguen el estándar EBU
   - Verifica la correcta implementación de campos obligatorios

## Integración con el Diagnóstico

La herramienta de diagnóstico STL utiliza esta función para:

1. Mostrar detalles técnicos de los bloques TTI
2. Visualizar partes críticas del encabezado GSI
3. Generar informes de compatibilidad con editores

## Ejemplo de Salida Real

```
0000: 00 00 01 00 ff 00 01 00 00 00 00 01 00 00 00 0e | ................
0010: 02 00 48 6f 6c 61 20 6d 75 6e 64 6f 8a 45 73 74 | ..Hola mundo.Est
0020: 6f 20 65 73 20 75 6e 61 20 70 72 75 65 62 61 03 | o es una prueba.
0030: 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 | ................
0040: 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 | ................
0050: 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 | ................
0060: 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 | ................
0070: 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 | ................
TTI Block 1 - Final: 0070: 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 | ................
```

Esta salida muestra un bloque TTI con:
- Número de subtítulo: 1 (bytes 2-3: 01 00)
- Texto: "Hola mundo" en primera línea, "Esto es una prueba" en segunda línea
- Salto de línea EBU (0x8A) representado con ¶
- Terminador ETX (0x03) después del texto
- Relleno con bytes 0x00 hasta completar los 128 bytes 