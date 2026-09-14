# Drives PowerPoint via COM to perform one action. Invoked by lib/office/pptx.js.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File powerpoint-com.ps1 `
#     -Action <Name> -PayloadPath <json in> -ResultPath <json out>
#
# Note: unlike Excel/Word, PowerPoint's COM object model does not reliably
# support a fully hidden Application window on all Office versions - a
# presentation window may briefly flash on screen during automation. This is
# a known limitation of the PowerPoint COM API, not a bug in this script.
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

$ppt = $null
$presentation = $null

# PowerPoint layout constants (PpSlideLayout)
$ppLayoutTitle = 1
$ppLayoutText = 2

try {
    $payload = Get-Content -Path $PayloadPath -Raw | ConvertFrom-Json

    $ppt = New-Object -ComObject PowerPoint.Application
    try { $ppt.DisplayAlerts = 1 } catch {} # ppAlertsNone

    $filePath = $payload.filePath

    switch ($Action) {

        'CreatePresentation' {
            $presentation = $ppt.Presentations.Add()

            $slide = $presentation.Slides.Add(1, $ppLayoutTitle)
            if ($payload.title) {
                $slide.Shapes.Title.TextFrame.TextRange.Text = [string]$payload.title
            }
            if ($payload.subtitle) {
                $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = [string]$payload.subtitle
            }

            $presentation.SaveAs($filePath, 11) # ppSaveAsOpenXMLPresentation (.pptx)
            Write-Result @{ ok = $true; filePath = $filePath; slideCount = $presentation.Slides.Count }
        }

        'AddSlide' {
            $presentation = $ppt.Presentations.Open($filePath, $false, $false, $false)
            $index = $presentation.Slides.Count + 1
            $slide = $presentation.Slides.Add($index, $ppLayoutText)

            if ($payload.title) {
                $slide.Shapes.Title.TextFrame.TextRange.Text = [string]$payload.title
            }

            $bullets = @($payload.bullets)
            if ($bullets.Count -gt 0) {
                $body = ($bullets -join "`r")
                $slide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $body
            }

            $presentation.Save()
            Write-Result @{ ok = $true; slideIndex = $index; slideCount = $presentation.Slides.Count }
        }

        'ListSlides' {
            $presentation = $ppt.Presentations.Open($filePath, $false, $false, $false)
            $titles = New-Object System.Collections.Generic.List[object]
            foreach ($slide in $presentation.Slides) {
                $t = ''
                try {
                    if ($slide.Shapes.HasTitle) {
                        $t = $slide.Shapes.Title.TextFrame.TextRange.Text
                    }
                } catch {}
                $titles.Add($t)
            }
            Write-Result @{ ok = $true; slideCount = $presentation.Slides.Count; titles = $titles }
        }

        'ReadSlide' {
            $presentation = $ppt.Presentations.Open($filePath, $false, $false, $false)
            $slide = $presentation.Slides.Item([int]$payload.slideIndex)
            $texts = New-Object System.Collections.Generic.List[object]
            foreach ($shape in $slide.Shapes) {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $texts.Add($shape.TextFrame.TextRange.Text)
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
    if ($presentation) {
        try { $presentation.Close() } catch {}
    }
    if ($ppt) {
        try { $ppt.Quit() } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
