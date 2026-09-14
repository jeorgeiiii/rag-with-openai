# Drives Word via COM to perform one action. Invoked by lib/office/word.js.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File word-com.ps1 `
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

# NOTE: Document.SaveAs/SaveAs2 has been observed to hang indefinitely on some
# Word installations regardless of content/format/target path, while
# Documents.Open() + Document.Save() on an already-path-bound document is
# reliable. To avoid SaveAs entirely, new documents are created as a minimal
# OOXML package on disk directly (no Word needed for that step), then opened
# with Word to set content via the working Open+Save path.
function New-BlankDocx([string]$Path) {
    Add-Type -AssemblyName WindowsBase
    if (Test-Path $Path) { Remove-Item $Path -Force }

    $package = [System.IO.Packaging.Package]::Open($Path, [System.IO.FileMode]::Create)
    try {
        $mainUri = New-Object System.Uri('/word/document.xml', [System.UriKind]::Relative)
        $mainPart = $package.CreatePart($mainUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')

        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr/></w:body></w:document>'

        $stream = $mainPart.GetStream()
        $writer = New-Object System.IO.StreamWriter($stream)
        try {
            $writer.Write($xml)
            $writer.Flush()
        } finally {
            $writer.Close()
        }

        $package.CreateRelationship($mainUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'rId1') | Out-Null
    } finally {
        $package.Close()
    }
}

$word = $null
$doc = $null

try {
    $payload = Get-Content -Path $PayloadPath -Raw | ConvertFrom-Json

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0 # wdAlertsNone

    $filePath = $payload.filePath

    switch ($Action) {

        'CreateDocument' {
            New-BlankDocx -Path $filePath
            $doc = $word.Documents.Open($filePath)

            if ($payload.title) {
                $range = $doc.Content
                $range.Text = [string]$payload.title
                try { $range.Style = $doc.Styles.Item('Title') } catch {}
                $range.Collapse(0) # wdCollapseEnd
                $range.InsertParagraphAfter()
            }

            $doc.Save()
            Write-Result @{ ok = $true; filePath = $filePath }
        }

        'AddParagraph' {
            $doc = $word.Documents.Open($filePath)
            $range = $doc.Content
            $range.Collapse(0) # wdCollapseEnd

            if ($range.Text -and $range.Text.Trim().Length -gt 0) {
                $range.InsertParagraphAfter()
                $range.Collapse(0)
            }

            $range.Text = [string]$payload.text

            $style = [string]$payload.style
            if ($style) {
                try { $range.Style = $doc.Styles.Item($style) } catch {}
            }

            $doc.Save()
            Write-Result @{ ok = $true }
        }

        'ReadText' {
            $doc = $word.Documents.Open($filePath)
            $text = $doc.Content.Text
            Write-Result @{ ok = $true; text = $text }
        }

        'ReplaceText' {
            $doc = $word.Documents.Open($filePath)
            $find = $doc.Content.Find
            $replaced = $find.Execute(
                [string]$payload.findText, $false, $false, $false, $false, $false,
                $true, 1, $false, [string]$payload.replaceText, 2
            )
            $doc.Save()
            Write-Result @{ ok = $true; replaced = [bool]$replaced }
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
        try { $doc.Close($true) } catch {}
    }
    if ($word) {
        try { $word.Quit() } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
