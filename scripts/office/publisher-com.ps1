# Drives Publisher via COM to perform one action. Invoked by lib/office/publisher.js.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File publisher-com.ps1 `
#     -Action <Name> -PayloadPath <json in> -ResultPath <json out>
#
# Note: Publisher.Application has no Visible property (unlike Word/Excel) -
# its window may appear briefly during automation. This is expected.
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

$pub = $null
$doc = $null

# Publisher.Quit() does not reliably terminate MSPUB.EXE (unlike Word/Excel/
# PowerPoint), leaking a process per call. Snapshot existing PIDs so the
# finally block can force-kill only the instance(s) this run spawned - never
# one the user already had open.
$pidsBefore = @(Get-Process -Name MSPUB -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
    $payload = Get-Content -Path $PayloadPath -Raw | ConvertFrom-Json

    $pub = New-Object -ComObject Publisher.Application
    $filePath = $payload.filePath

    switch ($Action) {

        'CreatePublication' {
            $doc = $pub.Documents.Add()
            $page = $doc.Pages.Item(1)

            if ($payload.text) {
                $shape = $page.Shapes.AddTextbox(1, 72, 72, 400, 200) # msoTextOrientationHorizontal
                $shape.TextFrame.TextRange.Text = [string]$payload.text
            }

            $doc.SaveAs($filePath)
            Write-Result @{ ok = $true; filePath = $filePath }
        }

        'AddTextBox' {
            $doc = $pub.Open($filePath)
            $pageIndex = if ($payload.pageIndex) { [int]$payload.pageIndex } else { 1 }
            $page = $doc.Pages.Item($pageIndex)

            $left = if ($null -ne $payload.left) { [double]$payload.left } else { 72 }
            $top = if ($null -ne $payload.top) { [double]$payload.top } else { 72 }
            $width = if ($payload.width) { [double]$payload.width } else { 300 }
            $height = if ($payload.height) { [double]$payload.height } else { 100 }

            $shape = $page.Shapes.AddTextbox(1, $left, $top, $width, $height)
            $shape.TextFrame.TextRange.Text = [string]$payload.text

            $doc.Save()
            Write-Result @{ ok = $true }
        }

        'ReadText' {
            $doc = $pub.Open($filePath)
            $texts = New-Object System.Collections.Generic.List[object]
            foreach ($page in $doc.Pages) {
                foreach ($shape in $page.Shapes) {
                    if ($shape.HasTextFrame) {
                        if ($shape.TextFrame.HasText) {
                            $texts.Add($shape.TextFrame.TextRange.Text)
                        }
                    }
                }
            }
            Write-Result @{ ok = $true; texts = $texts }
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
    if ($doc) {
        try { $doc.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
    }
    if ($pub) {
        try { $pub.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pub) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()

    Start-Sleep -Milliseconds 500
    $pidsAfter = @(Get-Process -Name MSPUB -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPids = $pidsAfter | Where-Object { $pidsBefore -notcontains $_ }
    foreach ($newPid in $newPids) {
        try { Stop-Process -Id $newPid -Force -ErrorAction SilentlyContinue } catch {}
    }
}
