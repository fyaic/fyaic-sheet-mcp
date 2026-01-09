param(
    [string]$excelFile = "",
    [string[]]$Sheet,
    [int]$MaxRows = 20,
    [switch]$AsJson
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

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($excelPath)
    $sheetCount = $workbook.Worksheets.Count
    if ($AsJson) {
        $worksheetsToProcess = @()
        if ($Sheet -and $Sheet.Count -gt 0) {
            foreach ($s in $Sheet) {
                if ($s -match "^\d+$") {
                    $i = [int]$s
                    if ($i -ge 1 -and $i -le $sheetCount) {
                        $worksheetsToProcess += $workbook.Worksheets.Item($i)
                    }
                } else {
                    foreach ($ws in $workbook.Worksheets) {
                        if ($ws.Name -eq $s) {
                            $worksheetsToProcess += $ws
                            break
                        }
                    }
                }
            }
        }
        if (-not $worksheetsToProcess) {
            $worksheetsToProcess += $workbook.Worksheets.Item(1)
        }
        $result = @()
        foreach ($worksheet in $worksheetsToProcess) {
            $usedRange = $worksheet.UsedRange
            $rows = $usedRange.Rows.Count
            $cols = $usedRange.Columns.Count
            $header = @()
            for ($col = 1; $col -le $cols; $col++) {
                $header += $worksheet.Cells.Item(1, $col).Value2
            }
            $data = @()
            $maxDataRows = 0
            if ($rows -gt 1) {
                $maxDataRows = [Math]::Min($MaxRows, [Math]::Max(0, $rows - 1))
                for ($row = 2; $row -le 1 + $maxDataRows; $row++) {
                    $rowData = @()
                    for ($col = 1; $col -le $cols; $col++) {
                        $rowData += $worksheet.Cells.Item($row, $col).Value2
                    }
                    $data += ,$rowData
                }
            }
            $truncated = $false
            if ($rows -gt (1 + $maxDataRows)) {
                $truncated = $true
            }
            $result += [PSCustomObject]@{
                sheetName = $worksheet.Name
                rowCount = $rows
                colCount = $cols
                header = $header
                rows = $data
                truncated = $truncated
            }
        }
        $json = $result | ConvertTo-Json -Depth 5
        Write-Output $json
    } else {
        $worksheetsToProcess = @()
        if ($sheetCount -ge 5) {
            Write-Host "`n检测到该工作簿包含 $sheetCount 个工作表:"
            for ($i = 1; $i -le $sheetCount; $i++) {
                $sheet = $workbook.Worksheets.Item($i)
                Write-Host "$i. $($sheet.Name)"
            }
            $selection = Read-Host "请输入要读取的工作表编号，例如 1 或 1,3,5"
            if ([string]::IsNullOrWhiteSpace($selection)) {
                $indices = 1
            } else {
                $indices = $selection -split "[,，;；\s]+" | Where-Object { $_ -match "^\d+$" } | ForEach-Object { [int]$_ } | Where-Object { $_ -ge 1 -and $_ -le $sheetCount } | Select-Object -Unique
                if (-not $indices) {
                    $indices = 1
                }
            }
            foreach ($i in $indices) {
                $worksheetsToProcess += $workbook.Worksheets.Item($i)
            }
        } else {
            $worksheetsToProcess = $workbook.Worksheets
        }
        foreach ($worksheet in $worksheetsToProcess) {
            $usedRange = $worksheet.UsedRange
            $rows = $usedRange.Rows.Count
            $cols = $usedRange.Columns.Count
            Write-Host "`n--- 工作表: $($worksheet.Name) ---"
            Write-Host "工作表行数: $rows"
            Write-Host "工作表列数: $cols"
            Write-Host "表头:"
            $header = @()
            for ($col = 1; $col -le $cols; $col++) {
                $header += $worksheet.Cells.Item(1, $col).Value2
            }
            Write-Host $header
            Write-Host "`n前20行数据:"
            $maxRows = [Math]::Min(20, $rows)
            for ($row = 2; $row -le $maxRows; $row++) {
                $rowData = @()
                for ($col = 1; $col -le $cols; $col++) {
                    $rowData += $worksheet.Cells.Item($row, $col).Value2
                }
                Write-Host $rowData
            }
            if ($rows -le 50) {
                Write-Host "`n完整数据:"
                for ($row = 1; $row -le $rows; $row++) {
                    $rowData = @()
                    for ($col = 1; $col -le $cols; $col++) {
                        $rowData += $worksheet.Cells.Item($row, $col).Value2
                    }
                    Write-Host $rowData
                }
            } else {
                Write-Host "`n数据量较大($rows行)，仅显示前20行"
            }
        }
    }
} finally {
    if ($workbook) {
        $workbook.Close($false) | Out-Null
    }
    $excel.Quit() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel -ErrorAction SilentlyContinue
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

