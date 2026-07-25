param([double]$Dpr, [double]$Top, [double[]]$Dividers, [string]$Variant)
Add-Type -AssemblyName System.Drawing
$path = "D:\Code\web\pptx\scratch\visual\corner-$Dpr-$Variant.png"
$img = [System.Drawing.Bitmap]::FromFile($path)
$top = [int][Math]::Floor($Top * $Dpr)
"=== v=$Variant dpr=$Dpr  table top dev-row $top"
foreach ($d in $Dividers) {
  $dx = [int][Math]::Round($d * $Dpr)
  $lines = @()
  for ($y = $top - 3; $y -le $top + 3; $y++) {
    $row = @()
    for ($x = $dx - 3; $x -le $dx + 3; $x++) {
      $c = $img.GetPixel($x, $y)
      $row += "{0,3}" -f $c.R
    }
    # reference pixel far from any divider on the same row
    $ref = $img.GetPixel($dx - 40, $y).R
    $lines += ("  y={0,4} ref={1,3} : {2}" -f $y, $ref, ($row -join ' '))
  }
  "  divider dev-x=$dx"
  $lines
}
$img.Dispose()
