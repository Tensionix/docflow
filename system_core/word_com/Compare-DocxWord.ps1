param(
    [Parameter(Mandatory=$true)]
    [string]$Original,

    [Parameter(Mandatory=$true)]
    [string]$Revised,

    [Parameter(Mandatory=$true)]
    [string]$Out,

    [Parameter(Mandatory=$true)]
    [string]$Report,

    [string]$Author = "Audion DocFlow"
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
$ErrorActionPreference = "Stop"

function New-ParentFolder {
    param([string]$Path)
    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent | Out-Null
    }
}

function Get-HResultText {
    param([System.Exception]$Exception)
    if ($null -eq $Exception) {
        return ""
    }
    $value = [int64]$Exception.HResult
    if ($value -lt 0) {
        $value += 4294967296
    }
    return ("0x{0:X8}" -f $value)
}

function Write-CompareReport {
    param(
        [string]$Status,
        [string]$ErrorMessage = "",
        [string]$HResultText = "",
        [hashtable]$Stats = @{}
    )

    $lines = @()
    $lines += "# Отчёт сравнения DOCX через Microsoft Word COM"
    $lines += ""
    $lines += "- Статус: **$Status**"
    $lines += '- Выходной файл: `' + $outPath + '`'
    $lines += '- Старый документ: `' + $originalPath + '`'
    $lines += '- Новый документ: `' + $revisedPath + '`'
    $lines += "- Метод: Microsoft Word COM CompareDocuments."
    $lines += "- Назначение: получить нативный DOCX с правками Word, когда Word способен обработать пару документов."
    $lines += ""

    if ($Status -ne "OK") {
        $lines += "## Ошибка"
        $lines += ""
        if ($HResultText) {
            $lines += '- HResult: `' + $HResultText + '`'
        }
        $lines += ""
        $lines += '```text'
        $lines += $ErrorMessage
        $lines += '```'
        $lines += ""
        $lines += "Если Word COM не может обработать эту пару документов, используйте сравнение DocFlow по структуре разделов."
    }
    else {
        $lines += "## Статистика результата"
        $lines += ""
        $lines += "| Показатель | Значение |"
        $lines += "|---|---:|"
        $lines += "| Страницы | $($Stats.Pages) |"
        $lines += "| Слова | $($Stats.Words) |"
        $lines += "| Абзацы | $($Stats.Paragraphs) |"
        $lines += "| Таблицы | $($Stats.Tables) |"
        $lines += "| Секции | $($Stats.Sections) |"
        $lines += "| Правки Word | $($Stats.Revisions) |"
        $lines += "| Комментарии | $($Stats.Comments) |"
        $lines += "| Встроенные рисунки | $($Stats.InlineShapes) |"
        $lines += "| Плавающие фигуры | $($Stats.Shapes) |"
    }

    Set-Content -LiteralPath $reportPath -Value ($lines -join [Environment]::NewLine) -Encoding UTF8
}

if (-not (Test-Path -LiteralPath $Original)) {
    throw "Original DOCX does not exist: $Original"
}
if (-not (Test-Path -LiteralPath $Revised)) {
    throw "Revised DOCX does not exist: $Revised"
}

$originalPath = (Resolve-Path -LiteralPath $Original).Path
$revisedPath = (Resolve-Path -LiteralPath $Revised).Path
$outPath = [System.IO.Path]::GetFullPath($Out)
$reportPath = [System.IO.Path]::GetFullPath($Report)

if ([System.String]::Equals($originalPath, $revisedPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Original and revised DOCX must be different files."
}

New-ParentFolder -Path $outPath
New-ParentFolder -Path $reportPath

$wdCompareDestinationNew = 2
$wdGranularityWordLevel = 1
$wdStatisticPages = 2
$wdStatisticWords = 0
$wdFormatDocumentDefault = 16

$word = $null
$originalDoc = $null
$revisedDoc = $null
$resultDoc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false

    $originalDoc = $word.Documents.Open($originalPath, $false, $true)
    $revisedDoc = $word.Documents.Open($revisedPath, $false, $true)

    $resultDoc = $word.CompareDocuments(
        $originalDoc,
        $revisedDoc,
        $wdCompareDestinationNew,
        $wdGranularityWordLevel,
        $true,  # CompareFormatting
        $true,  # CompareCaseChanges
        $true,  # CompareWhitespace
        $true,  # CompareTables
        $true,  # CompareHeaders
        $true,  # CompareFootnotes
        $true,  # CompareTextboxes
        $true,  # CompareFields
        $true,  # CompareComments
        $true,  # CompareMoves
        $Author,
        $true   # IgnoreAllComparisonWarnings
    )

    $resultDoc.SaveAs2($outPath, $wdFormatDocumentDefault)
    $pages = $resultDoc.ComputeStatistics($wdStatisticPages)
    $words = $resultDoc.ComputeStatistics($wdStatisticWords)
    $paragraphs = $resultDoc.Paragraphs.Count
    $tables = $resultDoc.Tables.Count
    $sections = $resultDoc.Sections.Count
    $revisions = $resultDoc.Revisions.Count
    $comments = $resultDoc.Comments.Count
    $inlineShapes = $resultDoc.InlineShapes.Count
    $shapes = $resultDoc.Shapes.Count

    $resultDoc.Close($true)
    $resultDoc = $null
    $revisedDoc.Close($false)
    $revisedDoc = $null
    $originalDoc.Close($false)
    $originalDoc = $null

    Write-CompareReport -Status "OK" -Stats @{
        Pages = $pages
        Words = $words
        Paragraphs = $paragraphs
        Tables = $tables
        Sections = $sections
        Revisions = $revisions
        Comments = $comments
        InlineShapes = $inlineShapes
        Shapes = $shapes
    }

    Write-Host "[OK] Word compare DOCX: $outPath"
    Write-Host "[OK] Revisions: $revisions"
    Write-Host "[OK] Pages: $pages"
    Write-Host "[OK] Tables: $tables"
    Write-Host "[OK] Report: $reportPath"
}
catch {
    $message = $_.Exception.Message
    $hresult = Get-HResultText -Exception $_.Exception
    try {
        Write-CompareReport -Status "FAILED" -ErrorMessage $message -HResultText $hresult
    } catch {}
    Write-Output "[ERROR] Word COM CompareDocuments failed."
    if ($hresult) {
        Write-Output "[ERROR] HResult: $hresult"
    }
    Write-Output "[ERROR] $message"
    exit 1
}
finally {
    if ($resultDoc -ne $null) {
        try { $resultDoc.Close($false) | Out-Null } catch {}
    }
    if ($revisedDoc -ne $null) {
        try { $revisedDoc.Close($false) | Out-Null } catch {}
    }
    if ($originalDoc -ne $null) {
        try { $originalDoc.Close($false) | Out-Null } catch {}
    }
    if ($word -ne $null) {
        try { $word.Quit() | Out-Null } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
