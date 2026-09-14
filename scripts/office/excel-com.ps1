# Drives Excel via COM to perform one action. Invoked by lib/office/excel.js.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File excel-com.ps1 `
#     -Action <Name> -PayloadPath <json in> -ResultPath <json out>
#
# Contract: never throws to stderr for expected failures - always writes
# { ok: bool, error?: string, ...data } to ResultPath so the Node caller can
# read a structured result regardless of outcome.

param(
    [Parameter(Mandatory = $true)][string]$Action,
    [Parameter(Mandatory = $true)][string]$PayloadPath,
    [Parameter(Mandatory = $true)][string]$ResultPath
)

$ErrorActionPreference = 'Stop'

function Write-Result($obj) {
    $obj | ConvertTo-Json -Depth 10 -Compress | Set-Content -Path $ResultPath -Encoding UTF8
}

$excel = $null
$workbook = $null

try {
    $payload = Get-Content -Path $PayloadPath -Raw | ConvertFrom-Json

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $filePath = $payload.filePath

    switch ($Action) {

        'CreateWorkbook' {
            $workbook = $excel.Workbooks.Add()

            $sheetNames = @($payload.sheets)
            if ($sheetNames.Count -gt 0) {
                while ($workbook.Sheets.Count -lt $sheetNames.Count) {
                    $workbook.Sheets.Add() | Out-Null
                }
                for ($i = 0; $i -lt $sheetNames.Count; $i++) {
                    $workbook.Sheets.Item($i + 1).Name = $sheetNames[$i]
                }
            }

            $workbook.SaveAs($filePath, 51) # xlOpenXMLWorkbook (.xlsx)
            Write-Result @{ ok = $true; filePath = $filePath; sheets = @($workbook.Sheets | ForEach-Object { $_.Name }) }
        }

        'ListSheets' {
            $workbook = $excel.Workbooks.Open($filePath)
            Write-Result @{ ok = $true; sheets = @($workbook.Sheets | ForEach-Object { $_.Name }) }
        }

        'AddSheet' {
            $workbook = $excel.Workbooks.Open($filePath)
            $sheet = $workbook.Sheets.Add([Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
            $sheet.Name = $payload.sheetName
            $workbook.Save()
            Write-Result @{ ok = $true; sheets = @($workbook.Sheets | ForEach-Object { $_.Name }) }
        }

        'WriteData' {
            $workbook = $excel.Workbooks.Open($filePath)
            $sheet = $workbook.Sheets.Item($payload.sheetName)

            $rows = @($payload.rows)
            $startRow = [int]$payload.startRow
            $startCol = [int]$payload.startCol

            for ($r = 0; $r -lt $rows.Count; $r++) {
                $rowData = @($rows[$r])
                for ($c = 0; $c -lt $rowData.Count; $c++) {
                    $sheet.Cells.Item($startRow + $r, $startCol + $c) = $rowData[$c]
                }
            }

            $workbook.Save()
            Write-Result @{ ok = $true; rowsWritten = $rows.Count }
        }

        'ReadData' {
            $workbook = $excel.Workbooks.Open($filePath)
            $sheet = $workbook.Sheets.Item($payload.sheetName)
            $range = $sheet.Range($payload.range)
            $values = $range.Value2

            $out = New-Object System.Collections.Generic.List[object]
            if ($values -is [System.Array]) {
                for ($r = 1; $r -le $values.GetLength(0); $r++) {
                    $rowOut = New-Object System.Collections.Generic.List[object]
                    for ($c = 1; $c -le $values.GetLength(1); $c++) {
                        $cellValue = $values[$r, $c]
                        $rowOut.Add($cellValue)
                    }
                    $out.Add($rowOut)
                }
            } else {
                $out.Add(@($values))
            }

            Write-Result @{ ok = $true; values = $out }
        }

        'ApplyFormula' {
            $workbook = $excel.Workbooks.Open($filePath)
            $sheet = $workbook.Sheets.Item($payload.sheetName)
            $sheet.Range($payload.cell).Formula = $payload.formula
            $workbook.Save()
            $result = $sheet.Range($payload.cell).Value2
            Write-Result @{ ok = $true; cell = $payload.cell; result = $result }
        }

        default {
            Write-Result @{ ok = $false; error = "Unknown action: $Action" }
        }
    }
}
catch {
    Write-Result @{ ok = $false; error = $_.Exception.Message }
}
finally {
    if ($workbook) {
        try { $workbook.Close($true) } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
