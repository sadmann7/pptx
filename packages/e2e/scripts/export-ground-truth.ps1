# Exports every fixture deck's slides to PNG using real PowerPoint (COM
# automation), producing the ground truth the oracle specs compare against.
#
# Run locally on a machine with desktop PowerPoint:
#   pnpm oracle:export          (from packages/e2e)
#
# Output: fixtures/ground-truth/<deck>/slide-<n>.png (1280x720, committed).
# Rerun only when fixtures change; CI never runs this, it just compares
# against the committed PNGs.
#
# Note: Office COM automation is not supported unattended and can flake
# (dialogs, licensing). Keep this a local, on-demand step.

$ErrorActionPreference = "Stop"

$e2eDir = Split-Path -Parent $PSScriptRoot
$fixturesDir = Join-Path $e2eDir "fixtures"
$groundTruthDir = Join-Path $fixturesDir "ground-truth"

# Native pixel size of the fixture decks (12192000 x 6858000 EMU at 96dpi).
$exportWidth = 1280
$exportHeight = 720

$decks = Get-ChildItem -Path $fixturesDir -Filter "*.pptx" -File
if ($decks.Count -eq 0) {
    Write-Error "No .pptx fixtures found in $fixturesDir. Run 'pnpm fixtures' first."
}

Write-Host "Starting PowerPoint..."
$powerpoint = New-Object -ComObject PowerPoint.Application

$failed = @()
try {
    foreach ($deck in $decks) {
        $deckName = [System.IO.Path]::GetFileNameWithoutExtension($deck.Name)
        $outDir = Join-Path $groundTruthDir $deckName

        try {
            # WithWindow:=false keeps the export headless-ish; ReadOnly avoids
            # PowerPoint writing repair changes back into the fixture.
            # MsoTriState raw values: msoTrue = -1, msoFalse = 0.
            $presentation = $powerpoint.Presentations.Open($deck.FullName, -1, 0, 0)
        }
        catch {
            $failed += "$($deck.Name): failed to open ($($_.Exception.Message))"
            continue
        }

        try {
            if (Test-Path $outDir) { Remove-Item -Recurse -Force $outDir }
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null

            $slideCount = $presentation.Slides.Count
            for ($i = 1; $i -le $slideCount; $i++) {
                $outPath = Join-Path $outDir ("slide-{0}.png" -f ($i - 1))
                $presentation.Slides.Item($i).Export($outPath, "PNG", $exportWidth, $exportHeight)
            }
            Write-Host "Exported $deckName ($slideCount slides)"
        }
        catch {
            $failed += "$($deck.Name): export failed ($($_.Exception.Message))"
        }
        finally {
            $presentation.Close()
        }
    }
}
finally {
    $powerpoint.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
}

if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "Failures:" -ForegroundColor Red
    $failed | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}

Write-Host ""
Write-Host "Ground truth written to $groundTruthDir"
