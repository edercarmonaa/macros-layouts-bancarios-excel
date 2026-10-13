# Codigo VBA exportado

Esta carpeta esta pensada para publicar el codigo de las macros como archivos de texto (`.bas`, `.cls` y `.frm`).

Los libros `.xlsm` son binarios y pueden ocultar datos personales, bancarios o metadatos. Por eso el flujo recomendado es:

1. Sanitizar el libro en Excel.
2. Revisar que no contenga CLABEs, cuentas, nombres reales, importes reales ni metadatos personales.
3. Exportar los modulos VBA con:

   ```powershell
   .\scripts\export-vba.ps1
   ```

4. Revisar los archivos generados en esta carpeta.
5. Hacer commit del codigo exportado, no de los libros privados.

Si Excel muestra un error al exportar, habilita en Excel:

`Archivo > Opciones > Centro de confianza > Configuracion del Centro de confianza > Configuracion de macros > Confiar en el acceso al modelo de objetos de proyectos de VBA`

## Formato recomendado

Versiona cada modulo VBA como archivo de texto (`.bas`, `.cls` o `.frm`) dentro de una carpeta que identifique el libro o proceso correspondiente. Ese formato es preferible a guardar el codigo solo dentro de un archivo `.xlsm`, porque permite revisar cambios y detectar datos sensibles antes de publicar.
