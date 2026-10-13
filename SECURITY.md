# Seguridad

Este repositorio no debe incluir libros de Excel reales con datos bancarios, personales o de nomina.

## Reglas para publicar

- No subir archivos `.xlsm`, `.xls` o `.xlsx` que contengan CLABEs, cuentas, importes reales, nombres de empleados, datos de clientes o metadatos personales.
- Usar datos ficticios en ejemplos, pruebas y documentacion.
- Revisar los metadatos de los libros antes de compartirlos publicamente.
- Generar plantillas sanitizadas antes de publicar cualquier macro.
- No publicar layouts generados por banca en linea si fueron creados con informacion real.

## Si se expuso informacion sensible

Si alguna cuenta, CLABE, credencial, layout bancario o dato personal fue publicado por error:

1. Retira el archivo del repositorio actual.
2. Limpia el historial Git si el archivo llego a un commit.
3. Notifica internamente a la persona responsable de los datos.
4. Rota o sustituye credenciales si existieran.
5. Vuelve a ejecutar una auditoria antes de publicar.
