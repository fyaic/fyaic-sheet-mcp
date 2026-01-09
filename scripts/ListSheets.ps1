param(
    [string]$excelFile = ""
)

$OutputEncoding = [System.Text.UTF8Encoding]::new()
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

if ([string]::IsNullOrWhiteSpace($excelFile)) {
    $excelPath = Get-ChildItem -LiteralPath $PSScriptRoot -File -Filter "*.xlsx" | Select-Object -First 1 | ForEach-Object FullName
} else {
    if (Test-Path -Path $excelFile -PathType Leaf) {
        $excelPath = (Resolve-Path -LiteralPath $excelFile).Path
    } else {
        # Fallback: try relative to PSScriptRoot just in case, but usually not needed
        $tryPath = Join-Path -Path $PSScriptRoot -ChildPath $excelFile
        if (Test-Path -Path $tryPath -PathType Leaf) {
            $excelPath = $tryPath
        } else {
            $excelPath = $excelFile # Let the error handling below catch it
        }
    }
}

if ([string]::IsNullOrWhiteSpace($excelPath) -or -not (Test-Path -LiteralPath $excelPath)) {
    throw "Excel file not found: $excelPath"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::OpenRead($excelPath)

try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/workbook.xml' } | Select-Object -First 1
    if (-not $entry) {
        throw "Workbook manifest not found in: $excelPath"
    }

    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream)
    $content = $reader.ReadToEnd()
    $reader.Dispose()
    $stream.Dispose()

    $xml = [xml]$content
    $result = @()
    $index = 1

    foreach ($sheet in $xml.workbook.sheets.sheet) {
        $result += [PSCustomObject]@{
            index = $index
            name = $sheet.name
        }
        $index++
    }

    $json = $result | ConvertTo-Json -Depth 3
    Write-Output $json
} finally {
    $zip.Dispose()
}

