param([string]$Path, [int]$Dpr, [double]$TableTop, [double[]]$Dividers)
Add-Type -AssemblyName System.Drawing
$img = [System.Drawing.Bitmap]::FromFile($Path)
$top = [int][Math]::Floor($TableTop * $Dpr)
foreach ($d in $Dividers) {
  $dx = [int][Math]::Round($d * $Dpr)
  $hits = @()
  for ($y = $top - (12 * $Dpr); $y -lt $top; $y++) {
    for ($x = $dx - (3 * $Dpr); $x -le $dx + (3 * $Dpr); $x++) {
      $c = $img.GetPixel($x, $y)
      # cream background is 239,233,217; report anything that deviates
      if ([Math]::Abs($c.R - 239) -gt 1 -or [Math]::Abs($c.G - 233) -gt 1 -or [Math]::Abs($c.B - 217) -gt 1) {
        $hits += "y=$y x=$x $($c.R),$($c.G),$($c.B)"
      }
    }
  }
  "divider @$d (px $dx): $($hits.Count) deviating px above table top ($top)"
  $hits | Select-Object -First 12 | ForEach-Object { "    $_" }
}
$img.Dispose()
