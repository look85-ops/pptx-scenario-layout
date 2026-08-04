<#
.SYNOPSIS
  Reorder slides of a PowerPoint deck according to a target map using PowerPoint COM.

.DESCRIPTION
  Map file format (target_order.txt):
    - one line per destination position, 1-based;
    - line N contains the ORIGINAL slide index that must end up at position N;
    - empty lines and lines starting with '#' are ignored.

  The reorder runs from the END to the START of the deck, so slides that are
  already placed in their final tail position never get shifted again.

.PARAMETER PptxPath
  Path to the .pptx file (may contain Cyrillic characters).

.PARAMETER MapFile
  Path to the target order map. Default: target_order.txt

.PARAMETER Backup
  If set, copies the source file to <name>_original.pptx before reordering,
  so validate.ps1 can compare content afterwards.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File layout.ps1 -PptxPath ".\deck.pptx" -MapFile "target_order.txt" -Backup
#>
param(
    [Parameter(Mandatory = $true)][string]$PptxPath,
    [string]$MapFile = "target_order.txt",
    [switch]$Backup
)

$ErrorActionPreference = "Stop"

function Read-TargetMap {
    param([string]$Path)
    $target = [System.Collections.Generic.List[int]]::new()
    Get-Content -LiteralPath $Path -Encoding UTF8 | ForEach-Object {
        $t = $_.Trim()
        if ($t -eq "" -or $t.StartsWith("#")) { return }
        $target.Add([int]$t)
    }
    return , $target
}

$target = Read-TargetMap -Path $MapFile
$n = $target.Count
Write-Host "Map entries: $n"

# --- Validate the map BEFORE touching the deck ---
$sorted = @($target.ToArray() | Sort-Object)
for ($i = 0; $i -lt $n; $i++) {
    if ($sorted[$i] -ne ($i + 1)) {
        throw "Map is not a valid permutation: expected $($i+1) at index $i, got $($sorted[$i])"
    }
}
Write-Host "Map validated: full permutation of 1..$n"

if ($Backup) {
    $copy = [System.IO.Path]::ChangeExtension($PptxPath, $null) + "_original.pptx"
    Copy-Item -LiteralPath $PptxPath -Destination $copy -Force
    Write-Host "Backup saved: $copy"
}

# --- Open deck via COM ---
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $null
try {
    $pres = $app.Presentations.Open(
        $PptxPath,
        [Microsoft.Office.Core.MsoTriState]::msoFalse, # ReadOnly
        [Microsoft.Office.Core.MsoTriState]::msoFalse, # Untitled
        [Microsoft.Office.Core.MsoTriState]::msoFalse) # WithWindow -> window hidden
    $slides = $pres.Slides
    $count = $slides.Count
    Write-Host "Opened deck, slides: $count"
    if ($count -ne $n) {
        throw "Slide count ($count) does not match map size ($n)"
    }

    # Snapshot of original SlideIDs so a slide can be located after previous moves.
    $ids = New-Object 'System.Collections.Generic.List[int]'
    for ($i = 1; $i -le $count; $i++) {
        $ids.Add([int]$slides.Item($i).SlideID)
    }

    # target[d-1] = original slide that must land on position d.
    # Process from the end to the start.
    $moves = 0
    for ($d = $n; $d -ge 1; $d--) {
        $srcId = $ids[$target[$d - 1] - 1]
        $cur = -1
        for ($i = 1; $i -le $slides.Count; $i++) {
            if ([int]$slides.Item($i).SlideID -eq $srcId) { $cur = $i; break }
        }
        if ($cur -lt 0) { throw "Slide with ID $srcId not found" }
        if ($cur -ne $d) {
            $slides.Item($cur).MoveTo($d)
            $moves++
        }
    }
    Write-Host "Slides moved: $moves of $count"

    $pres.Save()
    Write-Host "Saved."
}
finally {
    if ($pres) { $pres.Close() }
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
Write-Host "DONE"
