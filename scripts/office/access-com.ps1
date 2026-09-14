# Drives Access via COM/DAO to perform one action. Invoked by lib/office/access.js.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File access-com.ps1 `
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

# DAO RecordsetTypeEnum
$dbOpenTable = 1
$dbOpenSnapshot = 4

$access = $null
$dbOpened = $false

try {
    $payload = Get-Content -Path $PayloadPath -Raw | ConvertFrom-Json

    $access = New-Object -ComObject Access.Application
    $filePath = $payload.filePath

    switch ($Action) {

        'CreateDatabase' {
            $access.NewCurrentDatabase($filePath)
            $dbOpened = $true
            Write-Result @{ ok = $true; filePath = $filePath }
        }

        'CreateTable' {
            $access.OpenCurrentDatabase($filePath)
            $dbOpened = $true
            $db = $access.CurrentDb()

            $columns = @($payload.columns)
            $colDefs = @('ID COUNTER PRIMARY KEY')
            foreach ($col in $columns) {
                $colDefs += "[$($col.name)] $($col.type)"
            }
            $sql = "CREATE TABLE [$($payload.tableName)] (" + ($colDefs -join ', ') + ')'
            $db.Execute($sql)
            Write-Result @{ ok = $true; tableName = $payload.tableName }
        }

        'InsertRows' {
            $access.OpenCurrentDatabase($filePath)
            $dbOpened = $true
            $db = $access.CurrentDb()
            $rs = $db.OpenRecordset($payload.tableName, $dbOpenTable)

            $rows = @($payload.rows)
            $inserted = 0
            foreach ($row in $rows) {
                $rs.AddNew()
                foreach ($prop in $row.PSObject.Properties) {
                    $field = $rs.Fields.Item($prop.Name)
                    # Late-bound property set via reflection: PowerShell's early-bound
                    # dispatch to DAO's Field2.Value setter throws spurious InvalidCastException
                    # depending on the .NET type of the value being assigned.
                    [System.__ComObject].InvokeMember('Value', [System.Reflection.BindingFlags]::SetProperty, $null, $field, @($prop.Value)) | Out-Null
                }
                $rs.Update()
                $inserted++
            }
            $rs.Close()
            Write-Result @{ ok = $true; rowsInserted = $inserted }
        }

        'QueryTable' {
            $access.OpenCurrentDatabase($filePath)
            $dbOpened = $true
            $db = $access.CurrentDb()
            $rs = $db.OpenRecordset($payload.sql, $dbOpenSnapshot)

            $fieldNames = @()
            foreach ($f in $rs.Fields) { $fieldNames += $f.Name }

            $rows = New-Object System.Collections.Generic.List[object]
            if (-not $rs.EOF) {
                $rs.MoveFirst()
                while (-not $rs.EOF) {
                    $row = [ordered]@{}
                    foreach ($name in $fieldNames) {
                        $row[$name] = $rs.Fields.Item($name).Value
                    }
                    $rows.Add($row)
                    $rs.MoveNext()
                }
            }
            $rs.Close()
            Write-Result @{ ok = $true; rows = $rows }
        }

        'ListTables' {
            $access.OpenCurrentDatabase($filePath)
            $dbOpened = $true
            $db = $access.CurrentDb()

            $names = New-Object System.Collections.Generic.List[object]
            foreach ($t in $db.TableDefs) {
                if ($t.Name -notlike 'MSys*' -and $t.Name -notlike '~*') {
                    $names.Add($t.Name)
                }
            }
            Write-Result @{ ok = $true; tables = $names }
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
    if ($access) {
        if ($dbOpened) {
            try { $access.CloseCurrentDatabase() } catch {}
        }
        try { $access.Quit() } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
