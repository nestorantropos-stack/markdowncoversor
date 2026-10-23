# Conversor a Markdown · Portable

Herramienta local para Windows 10/11 de 64 bits. Convierte PDF con texto y Word `.docx` a Markdown UTF-8. No necesitas instalar Python, Word, servicios en la nube ni una cuenta. No requiere conexión a Internet al usarla.

## Uso

1. Extrae el ZIP en una carpeta local o en una memoria USB.
2. Abre **ConversorMarkdown.exe** con doble clic.
3. Pulsa **Agregar documentos** y selecciona uno o varios PDF / DOCX.
4. Elige la carpeta de destino y pulsa **Convertir a Markdown**.
5. Abre la carpeta con **Abrir destino**. Revisa los avisos de la ventana.

El ejecutable también funciona por sí solo. El ZIP incluye esta guía, las licencias y el código fuente. No solicita permisos de administrador. Al iniciarse, el ejecutable extrae sus componentes a una carpeta temporal de Windows y los elimina al cerrar normalmente; necesita espacio temporal y permiso de ejecución.

## Documentos grandes

No hay un límite programado de páginas. PDF se procesa página por página y se escribe progresivamente en un archivo temporal. Word se carga como documento XML en memoria; el consumo depende del tamaño y contenido. La interfaz se mantiene disponible durante la conversión. Se verificó con pruebas automatizadas un PDF de 300 páginas con texto extenso y un Word con 299 saltos de página. También se probó el ejecutable independiente con un PDF de 300 páginas y un DOCX, sin las rutas del entorno Python.

La cancelación se comprueba entre páginas/bloques. Abrir un documento o analizar una página muy compleja puede tardar antes de atender la cancelación. Se eliminan los temporales al cancelar o fallar. Los archivos ya terminados de un lote se conservan.

Los Markdown existentes nunca se sobrescriben: se generan nombres como `libro (2).md`. Un PDF y un DOCX con el mismo nombre también producen archivos separados. Cada documento genera un único `.md`.

## Qué conserva y qué revisar

- **Word DOCX:** párrafos en orden, títulos por estilo, negrita/cursiva explícitas, listas comunes, enlaces, tablas y notas al pie/finales. La primera fila de cada tabla se representa como cabecera Markdown. Las listas numeradas especiales se normalizan a números decimales.
- **PDF:** texto digital, separadores de página opcionales, títulos aproximados por tamaño de letra y tablas con estructura reconocible. Puedes desactivar la detección de tablas o títulos para documentos donde no funcione bien.
- **PDF escaneados:** no incluye OCR. Las páginas sin texto extraíble se señalan en el resultado y en los avisos. Usa OCR antes de convertir; tener algo de texto en una página no garantiza que todo su contenido sea texto digital.
- **PDF con contraseña:** primero guarda una copia desbloqueada con una herramienta apropiada.
- **Word `.doc` antiguo:** guarda una copia como `.docx` antes de convertir.
- **Imágenes, diagramas, ecuaciones, cuadros de texto y diseño de página:** no se reproducen. Encabezados/pies de Word no se exportan. Las tablas combinadas/anidadas y documentos con columnas requieren revisión. El formato PDF no almacena necesariamente el orden semántico: la extracción no garantiza reproducirlo.

Los archivos se leen localmente y no se suben a ningún servidor. La aplicación no utiliza IA ni API de pago.

## Línea de comandos (opcional)

En PowerShell, espera al proceso con `Start-Process -Wait` porque es una aplicación gráfica:

```powershell
Start-Process -FilePath '.\ConversorMarkdown.exe' -ArgumentList '--convert "C:\Documentos\libro.pdf" --output "C:\Documentos\Markdown" --report "C:\Documentos\informe.json"' -Wait
```

Admite varios archivos después de `--convert`, y las opciones `--no-tables`, `--no-headings`, `--no-page-markers`. El informe JSON incluye rutas, páginas sin texto y avisos. No lo guardes sobre un archivo que quieras conservar: `--report` escribe en la ruta indicada.

## Código fuente y compilación

Código disponible en `codigo/`. Licencia AGPL-3.0-o-posterior; al redistribuir, conserva las licencias y proporciona el código fuente correspondiente. PyMuPDF/MuPDF también utiliza AGPL. Python y las demás dependencias conservan sus propias licencias, incluidas en `licencias/`.

Fuentes de los componentes principales: [PyMuPDF 1.26.7](https://github.com/pymupdf/PyMuPDF/tree/1.26.7), [MuPDF](https://mupdf.com/releases/), [python-docx 1.2.0](https://github.com/python-openxml/python-docx/tree/v1.2.0), [lxml](https://github.com/lxml/lxml), [Pillow](https://github.com/python-pillow/Pillow), [Python](https://www.python.org/downloads/source/), [Tcl/Tk](https://www.tcl.tk/software/tcltk/).

Para desarrollar o recompilar en Windows con Python 3.12 de 64 bits y Tk:

```powershell
python -m pip install --target .build-deps -r requirements.txt
.\build.ps1 -Python python
```

Para ejecutar desde código:

```powershell
$env:PYTHONPATH = "$PWD\.build-deps"
python app.py
```

En el proyecto completo, las pruebas se ejecutan con `python -m unittest discover -s tests -v` después de configurar `PYTHONPATH`. La compilación necesita Internet para obtener las dependencias; la herramienta terminada funciona sin conexión.
