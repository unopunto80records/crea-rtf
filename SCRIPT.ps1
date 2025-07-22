# Define la carpeta raíz de destino como la carpeta donde se encuentra este script
$destinationRoot = $PSScriptRoot

# Define el nombre del archivo Excel esperado
$excelFileName = "PONER-NOMBRE-DEL-ARCHIVOt.xlsx"

# Construye la ruta completa al archivo Excel combinando la ruta raíz y el nombre del archivo
$excelPath = Join-Path -Path $destinationRoot -ChildPath $excelFileName

# Define el nombre de la hoja del Excel desde donde se importarán los datos
$worksheetName = "in"

# Verifica si el módulo ImportExcel está instalado en el sistema
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    # Si no está instalado, muestra un mensaje de error
    Write-Host "ERROR: El módulo 'ImportExcel' no está instalado." -ForegroundColor Red
    Write-Host "Por favor, ejecuta este comando en PowerShell y vuelve a intentarlo:" -ForegroundColor Yellow
    Write-Host "Install-Module -Name ImportExcel -Scope CurrentUser" -ForegroundColor Cyan
    # Espera que el usuario presione Enter antes de salir
    Read-Host "Presiona Enter para salir"
    return
}

# Verifica si el archivo Excel existe en la ruta especificada
if (-not (Test-Path -Path $excelPath)) {
    # Si no se encuentra, muestra un mensaje de error y detiene la ejecución
    Write-Host "ERROR: No se encuentra el fichero '$excelFileName' en esta carpeta." -ForegroundColor Red
    Read-Host "Presiona Enter para salir"
    return
}

# Intenta importar los datos desde el archivo Excel
try {
    Write-Host "Importando datos desde el fichero Excel..." -ForegroundColor Cyan
    $data = Import-Excel -Path $excelPath -WorksheetName $worksheetName
    Write-Host "Excel importado correctamente. Procesando $($data.Count) registros..." -ForegroundColor Green
}
catch {
    # Si ocurre un error al importar, muestra un mensaje y termina la ejecución
    Write-Error "No se pudo leer el fichero Excel. Verifica que no esté corrupto y que el nombre de la hoja ('$worksheetName') sea correcto."
    Read-Host "Presiona Enter para salir"
    return
}

# Itera sobre cada fila del Excel importado
foreach ($row in $data) {
    # Verifica si la propiedad CODARTICULO existe y no está vacía
    if (-not $row.PSObject.Properties['CODARTICULO'] -or $null -eq $row.CODARTICULO) {
        Write-Warning "Fila saltada (CODARTICULO no encontrado o vacío)."
        continue
    }

    # Extrae y limpia los valores de las columnas relevantes
    $articleFolder = $row.CODARTICULO.ToString().Trim()
    $fileName = $row.NOMBRE.ToString().Trim()
    $rtfContent = $row.TEXTO.ToString().Trim()

    # Verifica que los valores necesarios no estén vacíos
    if ([string]::IsNullOrEmpty($articleFolder) -or [string]::IsNullOrEmpty($fileName)) {
        Write-Warning "Registro saltado (CODARTICULO o NOMBRE vacíos)."
        continue
    }

    # Construye la ruta de la carpeta donde se guardará el archivo
    $destinationFolderPath = Join-Path -Path $destinationRoot -ChildPath $articleFolder

    # Si la carpeta no existe, la crea
    if (-not (Test-Path -Path $destinationFolderPath)) {
        New-Item -Path $destinationFolderPath -ItemType Directory | Out-Null
    }

    # Construye la ruta completa del archivo RTF que se va a guardar
    $finalRtfPath = Join-Path -Path $destinationFolderPath -ChildPath "$fileName.rtf"

    try {
        # Crea o sobrescribe el archivo RTF con el contenido proporcionado
        Set-Content -Path $finalRtfPath -Value $rtfContent -Encoding Default -NoNewline -Force
        Write-Host "Generado fichero: '$finalRtfPath'" -ForegroundColor Cyan
    }
    catch {
        # Si hay un error al guardar el archivo, muestra el error
        Write-Error "No se pudo crear el fichero '$finalRtfPath'. Error: $_"
    }
}

# Indica que el proceso ha finalizado correctamente
Write-Host "-------------------------------------------" -ForegroundColor Green
Write-Host "Proceso completado." -ForegroundColor Green

# Espera que el usuario presione Enter antes de cerrar la consola
Read-Host "Presiona Enter para finalizar"
