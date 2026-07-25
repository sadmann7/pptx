param(
  [string]$Path,
  [int]$X, [int]$Y, [int]$W, [int]$H,
  [int]$Zoom = 6,
  [string]$Out
)
Add-Type -AssemblyName System.Drawing
$src = [System.Drawing.Bitmap]::FromFile($Path)
$dst = New-Object System.Drawing.Bitmap ($W * $Zoom), ($H * $Zoom)
$g = [System.Drawing.Graphics]::FromImage($dst)
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::Half
$g.DrawImage($src, (New-Object System.Drawing.Rectangle 0, 0, ($W * $Zoom), ($H * $Zoom)), (New-Object System.Drawing.Rectangle $X, $Y, $W, $H), [System.Drawing.GraphicsUnit]::Pixel)
$g.Dispose()
$dst.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$dst.Dispose()
$src.Dispose()
"wrote $Out"
