param([string]$Python = "python", [switch]$PackageOnly)
$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
Set-Location -LiteralPath $projectRoot
$env:PYTHONPATH = Join-Path $projectRoot '.build-deps'
$env:PYINSTALLER_CONFIG_DIR = Join-Path $projectRoot 'build/config'
New-Item -ItemType Directory -Path 'build/tmp' -Force | Out-Null
$env:TEMP = Join-Path $projectRoot 'build/tmp'
$env:TMP = $env:TEMP
if (-not $PackageOnly) {
    & $Python -m PyInstaller --noconfirm --clean --onefile --windowed --name ConversorMarkdown --distpath portable --workpath build/pyinstaller --specpath build --exclude-module matplotlib --exclude-module numpy --exclude-module pandas --exclude-module onnxruntime --exclude-module networkx --exclude-module pytest --exclude-module IPython --collect-all pymupdf app.py
    if ($LASTEXITCODE -ne 0) { throw 'No se pudo generar el ejecutable.' }
}
if (-not (Test-Path -LiteralPath 'portable/ConversorMarkdown.exe')) { throw 'Falta el ejecutable portable.' }
Copy-Item -LiteralPath 'LEEME.md' -Destination 'portable/LEEME.md' -Force
Copy-Item -LiteralPath 'LICENSE' -Destination 'portable/LICENSE' -Force
New-Item -ItemType Directory -Path 'portable/codigo' -Force | Out-Null
foreach ($sourceFile in @('app.py', 'converter.py', 'requirements.txt', 'build.ps1')) {
    Copy-Item -LiteralPath $sourceFile -Destination 'portable/codigo' -Force
}
New-Item -ItemType Directory -Path 'portable/codigo/tests' -Force | Out-Null
Copy-Item -Path 'tests/*.py' -Destination 'portable/codigo/tests' -Force
Copy-Item -LiteralPath 'LEEME.md' -Destination 'portable/codigo/LEEME.md' -Force
Copy-Item -LiteralPath 'LICENSE' -Destination 'portable/codigo/LICENSE' -Force
New-Item -ItemType Directory -Path 'portable/licencias' -Force | Out-Null
Copy-Item -LiteralPath '.build-deps/pymupdf-1.26.7.dist-info/COPYING' -Destination 'portable/licencias/PyMuPDF-COPYING.txt' -Force
Copy-Item -LiteralPath 'LICENSE' -Destination 'portable/licencias/AGPL-3.0.txt' -Force
$runtimeRoot = & $Python -c 'import sys; print(sys.base_prefix)'
if (Test-Path -LiteralPath (Join-Path $runtimeRoot 'LICENSE.txt')) {
    Copy-Item -LiteralPath (Join-Path $runtimeRoot 'LICENSE.txt') -Destination 'portable/licencias/Python-LICENSE.txt' -Force
}
foreach ($tclPackage in @('tcl8.6', 'tk8.6')) {
    $tclLicense = Join-Path $runtimeRoot "tcl/$tclPackage/license.terms"
    if (Test-Path -LiteralPath $tclLicense) {
        Copy-Item -LiteralPath $tclLicense -Destination "portable/licencias/$tclPackage-license.terms" -Force
    }
}
$pillowLicenses = & $Python -c 'import importlib.metadata as m; print(m.distribution("Pillow").locate_file("pillow-" + m.version("Pillow") + ".dist-info/licenses"))'
if (Test-Path -LiteralPath $pillowLicenses) {
    Copy-Item -LiteralPath $pillowLicenses -Destination 'portable/licencias/Pillow' -Recurse -Force
}
$tclCoreLicense = Join-Path $runtimeRoot 'tcl/tcl8/license.terms'
if (Test-Path -LiteralPath $tclCoreLicense) {
    Copy-Item -LiteralPath $tclCoreLicense -Destination 'portable/licencias/tcl-license.terms' -Force
}
foreach ($depDir in Get-ChildItem -LiteralPath '.build-deps' -Directory -Filter '*.dist-info') {
    if ($depDir.Name -match '^(pymupdf-1|python_docx-|lxml-)') {
        $licenseDir = Join-Path $depDir.FullName 'licenses'
        if (Test-Path -LiteralPath $licenseDir) {
            Copy-Item -LiteralPath $licenseDir -Destination (Join-Path 'portable/licencias' $depDir.Name) -Recurse -Force
        }
    }
}
Compress-Archive -Path 'portable/*' -DestinationPath 'ConversorMarkdown-Portable-Windows.zip' -Force
Write-Output 'Portable generado en portable/ y en ConversorMarkdown-Portable-Windows.zip'
