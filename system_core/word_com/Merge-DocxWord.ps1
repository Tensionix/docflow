param(
    [Parameter(Mandatory=$true)]
    [string]$InputDir,

    [string]$Files = "",

    [Parameter(Mandatory=$true)]
    [string]$Out,

    [Parameter(Mandatory=$true)]
    [string]$Report
)

$ErrorActionPreference = "Stop"

function Resolve-InputFiles {
    param(
        [string]$InputDir,
        [string]$Files
    )

    if (-not (Test-Path -LiteralPath $InputDir)) {
        throw "Input folder does not exist: $InputDir"
    }

    if ($Files.Trim()) {
        $result = @()
        foreach ($item in ($Files -split '[;|]')) {
            $name = $item.Trim().Trim('"')
            if (-not $name) { continue }
            $path = Join-Path $InputDir $name
            if (-not (Test-Path -LiteralPath $path)) {
                throw "DOCX file not found: $path"
            }
            $result += (Resolve-Path -LiteralPath $path).Path
        }
        return $result
    }

    return @(Get-ChildItem -LiteralPath $InputDir -Filter "*.docx" -File -Recurse |
        Where-Object { -not $_.Name.StartsWith("~$") } |
        Sort-Object FullName |
        ForEach-Object { $_.FullName })
}

function New-ParentFolder {
    param([string]$Path)
    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent | Out-Null
    }
}

$inputFiles = @(Resolve-InputFiles -InputDir $InputDir -Files $Files)
if ($inputFiles.Count -lt 2) {
    throw "At least two DOCX files are required for Word COM merge."
}

$outPath = [System.IO.Path]::GetFullPath($Out)
$reportPath = [System.IO.Path]::GetFullPath($Report)
New-ParentFolder -Path $outPath
New-ParentFolder -Path $reportPath

$wdCollapseEnd = 0
$wdSectionBreakNextPage = 2
$wdStatisticPages = 2
$wdStatisticWords = 0
$wdFormatXMLDocument = 16

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false

    Copy-Item -LiteralPath $inputFiles[0] -Destination $outPath -Force
    $doc = $word.Documents.Open($outPath, $false, $false)

    for ($i = 1; $i -lt $inputFiles.Count; $i++) {
        $range = $doc.Content
        $range.Collapse($wdCollapseEnd)
        $range.InsertBreak($wdSectionBreakNextPage)
        $range.Collapse($wdCollapseEnd)
        $range.InsertFile($inputFiles[$i])
    }

    $doc.SaveAs2($outPath, $wdFormatXMLDocument)
    $pages = $doc.ComputeStatistics($wdStatisticPages)
    $words = $doc.ComputeStatistics($wdStatisticWords)
    $paragraphs = $doc.Paragraphs.Count
    $tables = $doc.Tables.Count
    $sections = $doc.Sections.Count
    $inlineShapes = $doc.InlineShapes.Count
    $shapes = $doc.Shapes.Count
    $doc.Close($true)
    $doc = $null

    $lines = @()
    $lines += "# Отчёт склейки DOCX через Word COM"
    $lines += ""
    $lines += "- Выходной файл: ``$outPath``"
    $lines += "- Документов склеено: **$($inputFiles.Count)**"
    $lines += "- Метод: Microsoft Word COM, разрыв секции + InsertFile."
    $lines += "- Логика: близко к ручной вставке следующего документа в Word после предыдущего."
    $lines += ""
    $lines += "## Статистика результата"
    $lines += ""
    $lines += "| Показатель | Значение |"
    $lines += "|---|---:|"
    $lines += "| Страницы | $pages |"
    $lines += "| Слова | $words |"
    $lines += "| Абзацы | $paragraphs |"
    $lines += "| Таблицы | $tables |"
    $lines += "| Секции | $sections |"
    $lines += "| Встроенные рисунки | $inlineShapes |"
    $lines += "| Плавающие фигуры | $shapes |"
    $lines += ""
    $lines += "## Входные документы"
    $lines += ""
    $lines += "| Порядок | Файл |"
    $lines += "|---:|---|"
    for ($i = 0; $i -lt $inputFiles.Count; $i++) {
        $safe = $inputFiles[$i].Replace("|", "\|")
        $lines += "| $($i + 1) | ``$safe`` |"
    }
    Set-Content -LiteralPath $reportPath -Value ($lines -join [Environment]::NewLine) -Encoding UTF8

    Write-Host "[OK] Merged DOCX: $outPath"
    Write-Host "[OK] Documents merged: $($inputFiles.Count)"
    Write-Host "[OK] Pages: $pages"
    Write-Host "[OK] Sections: $sections"
    Write-Host "[OK] Tables: $tables"
    Write-Host "[OK] Inline shapes: $inlineShapes"
    Write-Host "[OK] Report: $reportPath"
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close($false) | Out-Null } catch {}
    }
    if ($word -ne $null) {
        try { $word.Quit() | Out-Null } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
