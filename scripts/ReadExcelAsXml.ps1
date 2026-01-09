param(
    # Path to the Excel .xlsx file (can be relative or absolute)
    [string]$excelFile = "",
    # Optional list of sheet indices or names to read; default is first sheet
    [string[]]$Sheet,
    # Maximum row index to scan per sheet (hard cutoff for performance)
    [int]$MaxRows = 100,
    # When set, output is converted to JSON; otherwise native objects are returned
    [switch]$AsJson
)

# Ensure script and console output use UTF-8 encoding to avoid truncation/garbling
$OutputEncoding = [System.Text.UTF8Encoding]::new()
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

# Convert column letters (e.g. "A", "BC") to 1-based numeric index
function Get-ColumnIndex {
    param([string]$colLetters)
    $sum = 0
    foreach ($ch in $colLetters.ToCharArray()) {
        $c = [char]::ToUpper($ch)
        if ($c -lt 'A' -or $c -gt 'Z') {
            continue
        }
        $sum = $sum * 26 + ([int]$c - [int][char]'A' + 1)
    }
    return $sum
}

# Split a cell reference like "B12" into row/column indexes and column letters
function Split-CellRef {
    param([string]$ref)
    if ($ref -match '^([A-Z]+)(\d+)$') {
        $colLetters = $matches[1]
        $rowIndex = [int]$matches[2]
        $colIndex = Get-ColumnIndex -colLetters $colLetters
        return [PSCustomObject]@{
            row = $rowIndex
            col = $colIndex
            colLetters = $colLetters
        }
    } else {
        return $null
    }
}

# Parse a merge range reference like "A2:A10" or "B3:D5" into numeric bounds
function Parse-MergeRef {
    param([string]$ref)
    $parts = $ref -split ":", 2
    if ($parts.Count -eq 1) {
        # Single cell merge; start and end are the same
        $start = Split-CellRef -ref $parts[0]
        if (-not $start) { return $null }
        return [PSCustomObject]@{
            ref = $ref
            startRow = $start.row
            endRow = $start.row
            startCol = $start.col
            endCol = $start.col
        }
    } else {
        # Range merge; parse both ends
        $start = Split-CellRef -ref $parts[0]
        $end = Split-CellRef -ref $parts[1]
        if (-not $start -or -not $end) { return $null }
        return [PSCustomObject]@{
            ref = $ref
            startRow = $start.row
            endRow = $end.row
            startCol = $start.col
            endCol = $end.col
        }
    }
}

# Resolve the actual Excel file path to open
if ([string]::IsNullOrWhiteSpace($excelFile)) {
    # If no file passed in, default to the first .xlsx in the script directory
    $excelPath = Get-ChildItem -LiteralPath $PSScriptRoot -File -Filter "*.xlsx" | Select-Object -First 1 | ForEach-Object FullName
} else {
    # Try direct path first
    if (Test-Path -Path $excelFile -PathType Leaf) {
        $excelPath = (Resolve-Path -LiteralPath $excelFile).Path
    } else {
        # Fallback: treat as a file under the script directory
        $tryPath = Join-Path -Path $PSScriptRoot -ChildPath $excelFile
        if (Test-Path -Path $tryPath -PathType Leaf) {
            $excelPath = $tryPath
        } else {
            # Last resort: use the raw string as path
            $excelPath = $excelFile
        }
    }
}

# Validate that the target Excel file actually exists
if ([string]::IsNullOrWhiteSpace($excelPath) -or -not (Test-Path -LiteralPath $excelPath)) {
    throw "Excel file not found: $excelPath"
}

# Load .NET ZIP support used to read the xlsx package
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Will hold the open ZipArchive; initialized to null so we can safely dispose it
$zip = $null

try {
    # Open the .xlsx file as a read-only ZIP archive
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($excelPath)
    } catch {
        $msg = $_.Exception.Message
        throw "Failed to open Excel file as zip archive: $excelPath`n$msg"
    }

    # Locate the workbook manifest (sheet list and relationships)
    $workbookEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/workbook.xml' } | Select-Object -First 1
    if (-not $workbookEntry) {
        throw "Workbook manifest not found in: $excelPath"
    }

    # Read workbook.xml into memory
    $stream = $workbookEntry.Open()
    $reader = New-Object System.IO.StreamReader($stream)
    $workbookContent = $reader.ReadToEnd()
    $reader.Dispose()
    $stream.Dispose()

    # Parse sheet metadata (names and indices)
    $workbookXml = [xml]$workbookContent
    $sheetNodes = $workbookXml.workbook.sheets.sheet
    if (-not $sheetNodes) {
        throw "No sheets found in workbook: $excelPath"
    }

    # Build a simple array of sheet descriptors: index, name, path
    $sheetMeta = @()
    $index = 1
    foreach ($sheet in $sheetNodes) {
        $sheetMeta += [PSCustomObject]@{
            index = $index
            name = $sheet.name
            path = "xl/worksheets/sheet$index.xml"
        }
        $index++
    }

    # Load shared strings table (all cell string literals are stored here)
    $sharedStringsEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' } | Select-Object -First 1
    $sharedStrings = @()
    if ($sharedStringsEntry) {
        $ssStream = $sharedStringsEntry.Open()
        $ssReader = New-Object System.IO.StreamReader($ssStream)
        $ssContent = $ssReader.ReadToEnd()
        $ssReader.Dispose()
        $ssStream.Dispose()
        $ssXml = [xml]$ssContent
        foreach ($si in $ssXml.sst.si) {
            $sharedStrings += $si.InnerText
        }
    }

    # Decide which sheets to read based on -Sheet parameter (indices or names)
    $targetSheets = @()
    if ($Sheet -and $Sheet.Count -gt 0) {
        foreach ($s in $Sheet) {
            if ($s -match "^\d+$") {
                # Treat numeric value as 1-based sheet index
                $i = [int]$s
                $match = $sheetMeta | Where-Object { $_.index -eq $i } | Select-Object -First 1
                if ($match) {
                    $targetSheets += $match
                }
            } else {
                # Treat non-numeric value as sheet name
                $match = $sheetMeta | Where-Object { $_.name -eq $s } | Select-Object -First 1
                if ($match) {
                    $targetSheets += $match
                }
            }
        }
    }

    # Default to the first sheet if no valid selection is provided
    if (-not $targetSheets -or $targetSheets.Count -eq 0) {
        $targetSheets += $sheetMeta | Select-Object -First 1
    }

    # Aggregate results from all selected sheets
    $result = @()

    foreach ($meta in $targetSheets) {
        # Locate the XML file for the current worksheet
        $entry = $zip.Entries | Where-Object { $_.FullName -eq $meta.path } | Select-Object -First 1
        if (-not $entry) {
            continue
        }

        # Read worksheet XML into memory
        $sStream = $entry.Open()
        $sReader = New-Object System.IO.StreamReader($sStream)
        $sheetContent = $sReader.ReadToEnd()
        $sReader.Dispose()
        $sStream.Dispose()

        # Parse sheet XML
        $sheetXml = [xml]$sheetContent

        # Collect all cell values and basic sheet statistics
        $cells = @()
        $maxRowSeen = 0
        $maxColSeen = 0

        $rows = $sheetXml.worksheet.sheetData.row
        foreach ($row in $rows) {
            $rowIndex = [int]$row.r
            # Hard stop once we exceed MaxRows
            if ($rowIndex -gt $MaxRows) {
                break
            }
            if ($rowIndex -gt $maxRowSeen) {
                $maxRowSeen = $rowIndex
            }
            foreach ($c in $row.c) {
                $addr = $c.r
                $pos = Split-CellRef -ref $addr
                if (-not $pos) {
                    continue
                }
                if ($pos.col -gt $maxColSeen) {
                    $maxColSeen = $pos.col
                }
                $t = $c.t
                $vNode = $c.v
                $raw = $null
                $val = $null

                if ($t -eq "s") {
                    if ($vNode) {
                        $raw = $vNode.InnerText
                    }
                    if ($raw -ne $null -and $raw -ne "") {
                        $idx = [int]$raw
                        if ($idx -ge 0 -and $idx -lt $sharedStrings.Count) {
                            $val = $sharedStrings[$idx]
                        } else {
                            $val = $raw
                        }
                    }
                } elseif ($t -eq "inlineStr") {
                    if ($c.is -and $c.is.t) {
                        $val = $c.is.t.InnerText
                        $raw = $val
                    }
                } else {
                    if ($vNode) {
                        $raw = $vNode.InnerText
                        $val = $raw
                    }
                }
                $cells += [PSCustomObject]@{
                    address = $addr
                    row = $pos.row
                    col = $pos.col
                    type = $t
                    value = $val
                    raw = $raw
                }
            }
        }

        # Parse merged cell regions for this sheet
        $merges = @()
        $mergeNodes = $sheetXml.worksheet.mergeCells.mergeCell
        if ($mergeNodes) {
            foreach ($m in $mergeNodes) {
                $ref = $m.ref
                if (-not $ref) { continue }
                $parsed = Parse-MergeRef -ref $ref
                if ($parsed) {
                    $merges += $parsed
                }
            }
        }

        # Build the final object for this sheet
        $sheetResult = [PSCustomObject]@{
            sheetName = $meta.name
            sheetIndex = $meta.index
            maxRow = $maxRowSeen
            maxCol = $maxColSeen
            maxRowsScanned = $MaxRows
            cells = $cells
            merges = $merges
        }
        $result += $sheetResult
    }

    # Output either JSON or native PowerShell objects
    if ($AsJson) {
        $json = $result | ConvertTo-Json -Depth 10
        Write-Output $json
    } else {
        Write-Output $result
    }
} finally {
    # Always dispose the ZipArchive if it was successfully opened
    if ($zip -ne $null) {
        $zip.Dispose()
    }
}
