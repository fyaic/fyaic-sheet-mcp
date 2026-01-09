param(
    [string]$excelFile = "",
    [int]$MaxRows = 20,
    [switch]$AsJson,
    [string]$SheetIndices = "",
    [switch]$NonInteractive
)

$OutputEncoding = [System.Text.UTF8Encoding]::new()
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

$rootPath = $PSScriptRoot
$guidelinePath = Join-Path -Path $rootPath -ChildPath "prompts\guideline.md"
if (-not (Test-Path -LiteralPath $guidelinePath)) {
    throw "Guideline file not found: $guidelinePath"
}
$guideline = Get-Content -LiteralPath $guidelinePath -Raw
Write-Output $guideline

if ([string]::IsNullOrWhiteSpace($excelFile)) {
    $excelPath = Get-ChildItem -LiteralPath $rootPath -File -Filter "*.xlsx" | Select-Object -First 1 | ForEach-Object FullName
} else {
    if (Test-Path -Path $excelFile -PathType Leaf) {
        $excelPath = (Resolve-Path -LiteralPath $excelFile).Path
    } else {
        $tryPath = Join-Path -Path $rootPath -ChildPath $excelFile
        if (Test-Path -Path $tryPath -PathType Leaf) {
            $excelPath = $tryPath
        } else {
            $excelPath = $excelFile
        }
    }
}

if ([string]::IsNullOrWhiteSpace($excelPath) -or -not (Test-Path -LiteralPath $excelPath)) {
    throw "Excel file not found: $excelPath"
}

$listScript = Join-Path -Path $rootPath -ChildPath "scripts\ListSheets.ps1"
if (-not (Test-Path -LiteralPath $listScript)) {
    throw "ListSheets script not found: $listScript"
}

$json = & $listScript -excelFile $excelPath
if ([string]::IsNullOrWhiteSpace($json)) {
    throw "ListSheets returned empty result"
}
$sheets = $json | ConvertFrom-Json
if (-not $sheets) {
    throw "Failed to parse sheets from ListSheets result"
}

if ($sheets -isnot [System.Array]) {
    $sheets = @($sheets)
}

$sheetCount = $sheets.Count

if (-not [string]::IsNullOrWhiteSpace($SheetIndices)) {
    $indices = $SheetIndices -split "[,;\s]+" | Where-Object { $_ -match "^\d+$" } | ForEach-Object { [int]$_ } | Where-Object { $_ -ge 1 -and $_ -le $sheetCount } | Select-Object -Unique
    if (-not $indices) {
        $indices = 1
    }
    $sheetArgs = @()
    foreach ($i in $indices) {
        $match = $sheets | Where-Object { $_.index -eq $i } | Select-Object -First 1
        if ($match) {
            $sheetArgs += $match.name
        }
    }
} elseif ($NonInteractive) {
    Write-Output ""
    Write-Output "This workbook contains $sheetCount sheets:"
    foreach ($s in $sheets) {
        Write-Output "$($s.index). $($s.name)"
    }
    throw "Sheet selection required. Call this script again with -SheetIndices, for example: -SheetIndices 1 or -SheetIndices 1,3,5."
} else {
    if ($sheetCount -ge 5) {
        Write-Host ""
        Write-Host "This workbook contains $sheetCount sheets:"
        foreach ($s in $sheets) {
            Write-Host "$($s.index). $($s.name)"
        }
        $selection = Read-Host "Enter sheet indices to read, for example 1 or 1,3,5"
    } else {
        Write-Host ""
        Write-Host "This workbook contains $sheetCount sheets:"
        foreach ($s in $sheets) {
            Write-Host "$($s.index). $($s.name)"
        }
        $selection = Read-Host "Enter sheet indices to read, leave empty to read the first"
    }
    if ([string]::IsNullOrWhiteSpace($selection)) {
        $indices = 1
    } else {
        $indices = $selection -split "[,;\s]+" | Where-Object { $_ -match "^\d+$" } | ForEach-Object { [int]$_ } | Where-Object { $_ -ge 1 -and $_ -le $sheetCount } | Select-Object -Unique
        if (-not $indices) {
            $indices = 1
        }
    }
    $sheetArgs = @()
    foreach ($i in $indices) {
        $match = $sheets | Where-Object { $_.index -eq $i } | Select-Object -First 1
        if ($match) {
            $sheetArgs += $match.name
        }
    }
}

if (-not $sheetArgs -or $sheetArgs.Count -eq 0) {
    throw "No valid sheets selected"
}

# 智能选择读取方式：
# 1. 优先检查是否存在 _xml 文件夹（已导出），使用 ParseExcelXml.py（快速）
# 2. 如果不存在，使用 ReadExcelAsXml.ps1（自动解压）

$excelFileName = [System.IO.Path]::GetFileNameWithoutExtension($excelPath)
$xmlDir = Join-Path -Path (Split-Path -Path $excelPath -Parent) -ChildPath "$excelFileName\_xml"

$usePythonParser = $false
if (Test-Path -LiteralPath $xmlDir -PathType Container) {
    $sharedStringsPath = Join-Path -Path $xmlDir -ChildPath "sharedStrings.xml"
    $sheetPath = Join-Path -Path $xmlDir -ChildPath "sheet$($indices[0]).xml"
    if ((Test-Path -LiteralPath $sharedStringsPath) -and (Test-Path -LiteralPath $sheetPath)) {
        $usePythonParser = $true
        Write-Host "`n检测到已导出的 XML 文件夹，使用快速解析器 ParseExcelXml.py`n" -ForegroundColor Green
    }
}

if ($usePythonParser) {
    # 使用 Python 解析器（快速）
    $pythonScript = Join-Path -Path $rootPath -ChildPath "ParseExcelXml.py"
    if (-not (Test-Path -LiteralPath $pythonScript)) {
        throw "ParseExcelXml.py not found: $pythonScript"
    }

    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if (-not $pythonCmd) {
        throw "Python not found. Please install Python or use ReadExcelAsXml.ps1 instead."
    }

    # 调用 Python 脚本
    foreach ($i in $indices) {
        $output = & python $pythonScript $excelPath $i $MaxRows
        Write-Output $output
    }
} else {
    # 使用 PowerShell 解析器（完整功能，自动解压）
    Write-Host "`n未检测到已导出的 XML 文件夹，使用完整解析器 ReadExcelAsXml.ps1（将自动解压）`n" -ForegroundColor Yellow

    $readScript = Join-Path -Path $rootPath -ChildPath "scripts\ReadExcelAsXml.ps1"
    if (-not (Test-Path -LiteralPath $readScript)) {
        throw "ReadExcelAsXml.ps1 not found: $readScript"
    }

    $readArgs = @("-excelFile", $excelPath, "-MaxRows", $MaxRows, "-AsJson")
    foreach ($i in $indices) {
        $readArgs += @("-Sheet", $i)
    }

    & $readScript @readArgs
}

