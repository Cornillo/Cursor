# FN2STL - Notas técnicas

## Estructura de archivos

- **Main.gs**: Punto de entrada principal con las funciones expuestas al usuario.
- **GSIBlock.gs**: Implementación del bloque GSI (General Subtitle Information).
- **TTIBlock.gs**: Implementación del bloque TTI (Text and Timing Information).
- **DataAccess.gs**: Funciones para acceso a datos y procesamiento de timecodes (**ACTUAL Y PRINCIPAL**).
- **Utils.gs**: Versión anterior de DataAccess.gs (**DEPRECADO - NO USAR**).

## Notas importantes

1. Las funciones en `DataAccess.gs` son las que deben utilizarse. El archivo `Utils.gs` es una versión desactualizada que se mantiene por compatibilidad pero no debe usarse para nuevas implementaciones.

2. Para la conversión de timecodes, se utiliza el formato 24fps (STL24.01) que es más compatible con Subtitle Edit.

3. El manejo de caracteres especiales se basa en CP437 (DOS Latin US) con mapeos específicos para caracteres acentuados en español.

## Solución a problemas conocidos

### Funciones duplicadas

El sistema tiene duplicada la función `validateAndFormatTimecode` en dos archivos:
1. **DataAccess.gs**: Versión actualizada y correcta, optimizada para 24fps.
2. **Utils.gs**: Versión desactualizada con una advertencia.

Esto puede causar problemas porque Google Apps Script no utiliza espacios de nombres y las funciones se cargan en un orden no predecible. Para resolver este problema:

1. La función en `Utils.gs` ahora imprime una advertencia en los logs.
2. Si detecta problemas de formato de timecode, verifique que no esté importando o utilizando `Utils.gs` inadvertidamente.
3. La mejor solución es utilizar **explícitamente** la función en DataAccess.gs antes de cualquier otra importación.

### Pruebas

Si ejecuta la función `probarCorrecciones()` y observa errores en el manejo de timecodes, puede deberse a que se está utilizando la versión incorrecta de `validateAndFormatTimecode`. Ejecute la función directamente desde el archivo `DataAccess.gs` para asegurarse de usar la versión correcta. 