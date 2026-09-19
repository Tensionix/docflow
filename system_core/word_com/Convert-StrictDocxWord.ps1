param(
    [Parameter(Mandatory=$true)]
    [string]$ListFile
)

# Resave DOCX files through Microsoft Word as ordinary (transitional) DOCX.
# Each line of the list file: source path|target path. One Word instance for all files;
# a file Word cannot open is reported and the rest go on.

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$wdFormatXMLDocument = 12
$word = $null
$failed = 0

try {
    try {
        $word = New-Object -ComObject Word.Application
    }
    catch {
        Write-Output "[WORD-UNAVAILABLE] $($_.Exception.Message)"
        exit 3
    }
    $word.Visible = $false
    $word.DisplayAlerts = 0

    foreach ($line in (Get-Content -LiteralPath $ListFile -Encoding UTF8)) {
        if (-not $line.Trim()) { continue }
        $parts = $line -split '\|', 2
        $source = $parts[0]
        $target = $parts[1]
        $doc = $null
        try {
            $parent = Split-Path -Parent $target
            if ($parent -and -not (Test-Path -LiteralPath $parent)) {
                New-Item -ItemType Directory -Path $parent | Out-Null
            }
            $doc = $word.Documents.Open($source, $false, $true)
            $doc.SaveAs2($target, $wdFormatXMLDocument)
            Write-Output "[OK] $target"
        }
        catch {
            $failed++
            Write-Output "[FAILED] ${source}: $($_.Exception.Message)"
        }
        finally {
            if ($doc -ne $null) {
                try { $doc.Close($false) | Out-Null } catch {}
            }
        }
    }
}
finally {
    if ($word -ne $null) {
        try { $word.Quit() | Out-Null } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    }
}

if ($failed -gt 0) { exit 2 }
exit 0
