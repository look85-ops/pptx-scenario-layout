<#
.SYNOPSIS
  Update a PowerPoint deck from a scenario: clone template slides, replace
  text, tag new/changed slides, hide obsolete ones.

.DESCRIPTION
  The core of course-deck actualization. Never builds a slide from scratch:
  every new slide is a CLONE of an existing template slide (this keeps the
  corporate design) with its text replaced and a tag added.

  Operations are described in a JSON file (see operations.example.json):
    { "operations": [ ... ] }

  Operation types:
    edit  : find a slide by text, replace substrings in all shapes,
            optionally add a tag "Изменено".
    clone : duplicate template slide (from), replace its text, add tag,
            move to final position (to).
    hide  : hide a slide (by slide number or by text) - show="0",
            keeps it in the file but out of the show.

.PARAMETER PptxPath
  Path to the .pptx file (may contain Cyrillic characters).

.PARAMETER Operations
  Path to the JSON file with the operations table.

.PARAMETER Backup
  If set, copies the source file to <name>_backup.pptx before updating.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File update.ps1 -PptxPath ".\deck.pptx" -Operations "operations.json" -Backup
#>
param(
    [Parameter(Mandatory = $true)][string]$PptxPath,
    [Parameter(Mandatory = $true)][string]$Operations,
    [switch]$Backup
)

$ErrorActionPreference = "Stop"

# --- Read operations table ---
$json = Get-Content -LiteralPath $Operations -Encoding UTF8 -Raw | ConvertFrom-Json
if (-not $json.operations) { throw "No 'operations' array in $Operations" }

if ($Backup) {
    $copy = [System.IO.Path]::ChangeExtension($PptxPath, $null) + "_backup.pptx"
    Copy-Item -LiteralPath $PptxPath -Destination $copy -Force
    Write-Host "Backup saved: $copy"
}

# --- COM helpers ---
function Find-SlideByText {
    param($Slides, [string]$Find)
    for ($i = 1; $i -le $Slides.Count; $i++) {
        $s = $Slides.Item($i)
        foreach ($shape in $s.Shapes) {
            if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                if ($shape.TextFrame.TextRange.Text -like "*$Find*") { return $s }
            }
        }
    }
    return $null
}

function Replace-Texts {
    param($Slide, $Map)
    foreach ($shape in $Slide.Shapes) {
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
            foreach ($prop in $Map.PSObject.Properties) {
                $shape.TextFrame.TextRange.Replace([string]$prop.Name, [string]$prop.Value) | Out-Null
            }
        }
    }
}

function Add-Tag {
    param($Pres, $Slide, [string]$Text)
    if (-not $Text) { return }
    $w = $Pres.PageSetup.SlideWidth
    $tb = $Slide.Shapes.AddTextbox(1, [int]$w - 280, 15, 260, 40)  # MSOAutoShapeType 1 = msoTextOrientationHorizontal
    $range = $tb.TextFrame.TextRange
    $range.Text = $Text
    $range.Font.Color.RGB = 255        # red
    $range.Font.Italic = -1            # msoTrue
    $range.Font.Size = 24
}

function Hide-Slide {
    param($Slide)
    $Slide.SlideShowTransition.Hidden = [Microsoft.Office.Core.MsoTriState]::msoTrue
}

# --- Open deck ---
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
    Write-Host "Opened deck, slides: $($slides.Count)"

    $clones = 0; $edits = 0; $hides = 0
    foreach ($op in $json.operations) {
        switch ($op.op) {
            "clone" {
                $src = [int]$op.from
                $to = [int]$op.to
                $dupe = $slides.Item($src).Duplicate()
                $newSlide = $dupe.Item(1)
                if ($op.replace) { Replace-Texts -Slide $newSlide -Map $op.replace }
                Add-Tag -Pres $pres -Slide $newSlide -Text ([string]$op.tag)
                $newSlide.MoveTo($to)
                Write-Host "  clone: slide $src -> position $to (tag: $($op.tag))"
                $clones++
            }
            "edit" {
                $target = $null
                if ($op.slide) { $target = $slides.Item([int]$op.slide) }
                elseif ($op.find) { $target = Find-SlideByText -Slides $slides -Find ([string]$op.find) }
                if (-not $target) { throw "edit: slide not found ($($op.find))" }
                if ($op.replace) { Replace-Texts -Slide $target -Map $op.replace }
                Add-Tag -Pres $pres -Slide $target -Text ([string]$op.tag)
                Write-Host "  edit: slide $($target.SlideIndex) (tag: $($op.tag))"
                $edits++
            }
            "hide" {
                $target = $null
                if ($op.slide) { $target = $slides.Item([int]$op.slide) }
                elseif ($op.find) { $target = Find-SlideByText -Slides $slides -Find ([string]$op.find) }
                if (-not $target) { throw "hide: slide not found ($($op.find))" }
                Hide-Slide -Slide $target
                Write-Host "  hide: slide $($target.SlideIndex)"
                $hides++
            }
            default { throw "Unknown operation: $($op.op)" }
        }
    }

    $pres.Save()
    Write-Host "Saved. clones: $clones, edits: $edits, hides: $hides"
}
finally {
    if ($pres) { $pres.Close() }
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
Write-Host "DONE"
