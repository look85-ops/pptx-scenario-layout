<#
.SYNOPSIS
  Validate a reordered deck: slide count, hidden slides, sldIdLst order,
  and (optional) content permutation against an original copy.

.DESCRIPTION
  Use AFTER layout.ps1 -Backup. Checks:
    1. slide count in the deck == map size;
    2. every destination position in sldIdLst resolves to slide{N}.xml
       (PowerPoint renumbers slide files sequentially after save);
    3. number of hidden slides (show="0");
    4. if -OriginalPath is given: for each position d the normalized text of
       slide d must equal the normalized text of original slide map[d].
       Pure-digit runs (auto slide numbers) are ignored.

.PARAMETER PptxPath
  Path to the reordered .pptx.

.PARAMETER MapFile
  Path to the target order map. Default: target_order.txt

.PARAMETER OriginalPath
  Path to the <name>_original.pptx created by layout.ps1 -Backup.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File validate.ps1 -PptxPath ".\deck.pptx" -MapFile "target_order.txt" -OriginalPath ".\deck_original.pptx"
#>
param(
    [Parameter(Mandatory = $true)][string]$PptxPath,
    [string]$MapFile = "target_order.txt",
    [string]$OriginalPath
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

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

function Get-SlideXmlMap {
    param($Zip, [string]$Prefix)
    # returns ordered dict: position -> slide entry name, and count
    $presEntry = $Zip.Entries | Where-Object { $_.FullName -eq "$Prefix/presentation.xml" }
    if (-not $presEntry) { throw "$Prefix/presentation.xml not found" }
    $doc = New-Object System.Xml.XmlDocument
    $reader = New-Object System.IO.StreamReader($presEntry.Open(), [System.Text.Encoding]::UTF8)
    $doc.LoadXml($reader.ReadToEnd())
    $reader.Close()
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("a", "http://schemas.openxmlformats.org/presentationml/2006/main")
    $ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $ids = $doc.SelectNodes("//a:sldIdLst/a:sldId", $ns)

    $relsEntry = $Zip.Entries | Where-Object { $_.FullName -eq "$Prefix/_rels/presentation.xml.rels" }
    $relsDoc = New-Object System.Xml.XmlDocument
    $relReader = New-Object System.IO.StreamReader($relsEntry.Open(), [System.Text.Encoding]::UTF8)
    $relsDoc.LoadXml($relReader.ReadToEnd())
    $relReader.Close()
    $nsR = New-Object System.Xml.XmlNamespaceManager($relsDoc.NameTable)
    $nsR.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")
    $rels = @{}
    foreach ($rel in $relsDoc.SelectNodes("//r:Relationship", $nsR)) {
        $rels[$rel.GetAttribute("Id")] = $rel.GetAttribute("Target")
    }

    $order = New-Object System.Collections.Generic.List[string]
    foreach ($id in $ids) {
        $rid = $id.Attributes["r:id"].Value
        $order.Add($rels[$rid])
    }
    return $order
}

function Get-EntryName {
    param([string]$Target)
    # rel Target is relative to ppt/ ("slides/slide1.xml") unless absolute
    if ($Target.StartsWith("/")) { return $Target.TrimStart("/") }
    return "ppt/" + $Target
}

function Get-SlideText {
    param($Zip, [string]$EntryName)
    $entry = $Zip.Entries | Where-Object { $_.FullName -eq $EntryName }
    if (-not $entry) { return $null }
    $doc = New-Object System.Xml.XmlDocument
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $doc.LoadXml($reader.ReadToEnd())
    $reader.Close()
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $doc.SelectNodes("//*[local-name()='t']")) {
        [void]$sb.Append($t.InnerText).Append(" ")
    }
    return $sb.ToString()
}

function Normalize-Text {
    param([string]$Text)
    if (-not $Text) { return "" }
    # drop pure-digit runs (auto slide numbers), collapse whitespace, lowercase
    $cleaned = [regex]::Replace($Text, "\b\d+\b", " ")
    $cleaned = [regex]::Replace($cleaned, "\s+", " ").Trim()
    return $cleaned.ToLowerInvariant()
}

$target = Read-TargetMap -Path $MapFile
$n = $target.Count

$zip = [System.IO.Compression.ZipFile]::OpenRead($PptxPath)
try {
    $order = Get-SlideXmlMap -Zip $zip -Prefix "ppt"

    Write-Host "== Order check =="
    $ok = $true
    if ($order.Count -ne $n) {
        Write-Host "FAIL: sldIdLst has $($order.Count) slides, map expects $n"
        $ok = $false
    } else {
        for ($d = 1; $d -le $n; $d++) {
            $resolved = [System.IO.Path]::GetFileName($order[$d - 1])
            $expected = "slide$d.xml"
            if ($resolved -ne $expected) {
                Write-Host "  position $d -> $resolved (expected $expected)"
                $ok = $false
            }
        }
        if ($ok) { Write-Host "OK: sequential sldIdLst, $n slides" }
    }

    Write-Host "== Hidden slides =="
    $hidden = 0
    foreach ($e in $zip.Entries) {
        if ($e.FullName -match "^ppt/slides/slide\d+\.xml$") {
            $doc = New-Object System.Xml.XmlDocument
            $reader = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
            $doc.LoadXml($reader.ReadToEnd())
            $reader.Close()
            $show = $doc.DocumentElement.GetAttribute("show")
            if ($show -eq "0") { $hidden++ }
        }
    }
    Write-Host "Hidden slides: $hidden"

    if ($OriginalPath) {
        Write-Host "== Content permutation check =="
        $origZip = [System.IO.Compression.ZipFile]::OpenRead($OriginalPath)
        try {
            $origOrder = Get-SlideXmlMap -Zip $origZip -Prefix "ppt"
            $mismatch = 0
            for ($d = 1; $d -le $n; $d++) {
                $src = $target[$d - 1]
                # new slide d must match original slide src (by file name)
                $newEntry = "ppt/slides/slide$d.xml"
                $origEntry = Get-EntryName ([string]$origOrder[$src - 1])
                $newText = Normalize-Text (Get-SlideText -Zip $zip -EntryName $newEntry)
                $origText = Normalize-Text (Get-SlideText -Zip $origZip -EntryName $origEntry)
                if ($newText -ne $origText) {
                    $mismatch++
                    Write-Host "  position ${d}: content differs from original slide $src"
                }
            }
            if ($mismatch -eq 0) {
                Write-Host "OK: all $n slides match their mapped original (content preserved)"
            } else {
                Write-Host "MISMATCH: $mismatch slide(s) differ. Check manually."
                $ok = $false
            }
        }
        finally { $origZip.Dispose() }
    }

    Write-Host ""
    if ($ok) { Write-Host "VALIDATION PASSED" } else { Write-Host "VALIDATION FAILED" }
}
finally { $zip.Dispose() }
