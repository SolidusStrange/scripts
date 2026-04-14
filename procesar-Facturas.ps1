param(
    [Parameter(Mandatory=$true)]
    [string]$RutaBase,

    [Parameter(Mandatory=$false)]
    [string]$Salida = "resultado.csv"
)

# Validación
if (-not (Test-Path $RutaBase)) {
    Write-Error "La ruta no existe: $RutaBase"
    exit 1
}

$resultados = @()

Get-ChildItem -Path $RutaBase -Directory | ForEach-Object {

    $numeroFactura = $_.Name

    $notaCredito = $null
    $anulada = $null
    $refact = $null

    Get-ChildItem -Path $_.FullName -Filter *.pdf | ForEach-Object {

        $nombre = $_.Name.ToLower()

        if ($nombre -match "^(\d+)\s*-") {
            $numero = $matches[1]
        } else {
            return
        }

        switch -Regex ($nombre) {
            "credito"   { $notaCredito = $numero }
            "anulada"   { $anulada = $numero }
            "refact"    { $refact = $numero }
        }
    }

    $resultados += [PSCustomObject]@{
        FACTURA         = $numeroFactura
        NOTA_CREDITO    = $notaCredito
        ANULADA         = $anulada
        REFACTURACION   = $refact
    }
}

# Exportar
$salidaCompleta = Join-Path $RutaBase $Salida
$resultados | Export-Csv -Path $salidaCompleta -NoTypeInformation -Encoding UTF8

Write-Host "Proceso completado. Archivo generado en: $salidaCompleta"