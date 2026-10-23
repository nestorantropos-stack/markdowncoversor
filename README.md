# Conversor a Markdown

Aplicación de escritorio portable para convertir documentos PDF y Word `.docx` a Markdown. Funciona localmente en Windows 10/11 de 64 bits, sin instalar Python ni Word y sin conexión a Internet.

## Descargar para Windows

Descarga [`ConversorMarkdown-Portable-Windows.zip`](ConversorMarkdown-Portable-Windows.zip) desde este repositorio, extrae el archivo y abre `ConversorMarkdown.exe`. Selecciona tus documentos y una carpeta de salida; admite lotes, informa el progreso y permite cancelar.

El paquete incluye instrucciones, el código fuente y las licencias. También puedes consultar la [guía completa](LEEME.md).

## Formatos y límites

- PDF con texto: extrae el texto página por página y puede intentar reconocer títulos y tablas.
- Word `.docx`: conserva títulos, formato básico, listas comunes, enlaces, tablas y notas.
- PDF escaneados: necesitan OCR antes de la conversión.
- Las imágenes no se exportan. Revisa tablas y documentos con columnas o diseños complejos.
- Los archivos `.doc` antiguos deben guardarse primero en formato `.docx`.

La conversión procesa documentos extensos, incluidos los de 300 páginas. Los archivos Markdown existentes se conservan; si el nombre ya existe, se genera otro nombre.

## Ejecutar desde el código fuente

Con Python 3.12 de 64 bits y Tk instalado:

```powershell
python -m pip install --target .build-deps -r requirements.txt
$env:PYTHONPATH = "$PWD\.build-deps"
python app.py
```

Para compilar el ejecutable y el paquete portable en Windows:

```powershell
.\build.ps1 -Python python
```

Para reconstruir solamente el ZIP con un ejecutable ya generado:

```powershell
.\build.ps1 -Python python -PackageOnly
```

## Licencia

El proyecto se distribuye bajo la GNU AGPL, versión 3 o posterior. Revisa `LICENSE` y `portable/licencias/` para los términos de PyMuPDF/MuPDF y las licencias de sus dependencias.
