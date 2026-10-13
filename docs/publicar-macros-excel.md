# Como publicar macros de Excel en GitHub

Hay dos formas razonables de publicar este proyecto.

## Opcion recomendada: codigo VBA exportado

Cuando Excel permite acceder al proyecto VBA, exporta los modulos como texto:

```powershell
.\scripts\export-vba.ps1
```

Esto genera archivos `.bas`, `.cls` y `.frm` dentro de `vba/`. Es la mejor opcion porque GitHub puede mostrar diferencias entre versiones y revisar el codigo sin abrir Excel.

En esta revision, Excel detecto que los libros tienen proyecto VBA, pero no expuso componentes exportables mediante automatizacion. Por eso no se generaron archivos `.bas`, `.cls` o `.frm`.

## Opcion practica: plantillas sanitizadas

Si el codigo VBA no puede exportarse, publica solo libros `.xlsm` completamente sanitizados dentro de:

```text
templates/
```

La regla de `.gitignore` permite versionar archivos Excel unicamente en esa carpeta. Los libros fuera de `templates/` siguen ignorados para evitar publicar datos privados por accidente.

## Opcion manual: copiar modulos a archivos `.bas`

Si ya tienes el codigo visible en el editor VBA:

1. Abre Excel y presiona `Alt + F11`.
2. Abre el modulo que contiene la macro.
3. Copia el contenido del modulo.
4. Crea un archivo `.bas` dentro de `vba/`.
5. Agrega al inicio una linea como:

   ```vb
   Attribute VB_Name = "NombreDelModulo"
   Option Explicit
   ```

6. Pega el codigo revisado y guarda el archivo.

Este metodo es valido cuando Excel no permite exportar automaticamente los componentes del proyecto VBA.

## Checklist antes de hacer commit

1. Abre cada libro en Excel.
2. Revisa hojas visibles y ocultas.
3. Revisa nombres definidos, conexiones, consultas, tablas dinamicas y comentarios.
4. Revisa propiedades del documento y elimina autores o nombres personales innecesarios.
5. Ejecuta una busqueda de CLABEs, cuentas, RFC, CURP, NSS, correos y telefonos.
6. Ejecuta las macros con datos ficticios.
7. Confirma que los layouts generados tambien usan datos ficticios.
8. Agrega al commit solo archivos dentro de `templates/`, documentacion y codigo exportado.

## Comandos utiles

Ver que Excel esta ignorado fuera de `templates/`:

```powershell
git status --short --ignored
```

Agregar una plantilla revisada:

```powershell
git add templates\NOMBRE-DE-PLANTILLA.xlsm
```

Confirmar que no hay libros privados staged:

```powershell
git diff --cached --name-only
```
