# Generates the StockData logo set into ../logos using System.Drawing.
# Theme: green up-trend area chart on dark navy, matching oxt/icons/icon.png.
Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = "Stop"

$root = Split-Path $PSScriptRoot -Parent
$out  = Join-Path $root "logos"
New-Item -ItemType Directory -Force -Path $out | Out-Null

$NAVY_TOP = [System.Drawing.Color]::FromArgb(255, 32, 50, 78)
$NAVY_BOT = [System.Drawing.Color]::FromArgb(255, 16, 26, 44)
$GREEN    = [System.Drawing.Color]::FromArgb(255, 46, 204, 113)
$GREEN_FILL = [System.Drawing.Color]::FromArgb(70, 46, 204, 113)
$WHITE    = [System.Drawing.Color]::FromArgb(255, 245, 247, 250)
$SUBTLE   = [System.Drawing.Color]::FromArgb(255, 150, 200, 175)

function New-RoundedPath($x, $y, $w, $h, $r) {
  $d = $r * 2
  $p = New-Object System.Drawing.Drawing2D.GraphicsPath
  $p.AddArc($x, $y, $d, $d, 180, 90)
  $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
  $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
  $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
  $p.CloseFigure()
  return $p
}

# Draws the up-trend area chart inside the box (bx,by,bw,bh) onto $g.
function Draw-Chart($g, $bx, $by, $bw, $bh, $lineWidth) {
  # normalized trend points (x fraction, y fraction; y=0 top)
  $norm = @(
    @(0.04, 0.74), @(0.20, 0.55), @(0.34, 0.64),
    @(0.52, 0.30), @(0.68, 0.40), @(0.96, 0.10)
  )
  $pts = foreach ($n in $norm) {
    [System.Drawing.PointF]::new($bx + $n[0] * $bw, $by + $n[1] * $bh)
  }
  [System.Drawing.PointF[]]$line = $pts

  # filled area under the line
  $poly = New-Object System.Collections.Generic.List[System.Drawing.PointF]
  foreach ($p in $line) { $poly.Add($p) }
  $poly.Add([System.Drawing.PointF]::new($bx + $bw, $by + $bh))
  $poly.Add([System.Drawing.PointF]::new($bx, $by + $bh))
  $fill = New-Object System.Drawing.SolidBrush $GREEN_FILL
  $g.FillPolygon($fill, [System.Drawing.PointF[]]$poly)

  # the trend line
  $pen = New-Object System.Drawing.Pen $GREEN, $lineWidth
  $pen.StartCap = 'Round'; $pen.EndCap = 'Round'; $pen.LineJoin = 'Round'
  $g.DrawLines($pen, $line)

  # arrow head at the last point
  $last = $line[$line.Length - 1]
  $ah = $lineWidth * 2.2
  $g.DrawLine($pen, $last.X, $last.Y, $last.X - $ah, $last.Y + $ah * 0.35)
  $g.DrawLine($pen, $last.X, $last.Y, $last.X - $ah * 0.35, $last.Y + $ah)
}

function New-Graphics($bmp) {
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'
  $g.InterpolationMode = 'HighQualityBicubic'
  $g.TextRenderingHint = 'ClearTypeGridFit'
  return $g
}

# --- 1) square app logo, 512x512, dark, transparent corners -------------------
$sz = 512
$bmp = New-Object System.Drawing.Bitmap $sz, $sz
$g = New-Graphics $bmp
$path = New-RoundedPath 0 0 $sz $sz 96
$rect = New-Object System.Drawing.Rectangle 0, 0, $sz, $sz
$grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, $NAVY_TOP, $NAVY_BOT, 90
$g.FillPath($grad, $path)
$g.SetClip($path)
Draw-Chart $g 70 70 372 372 18
$g.ResetClip()
$bmp.Save((Join-Path $out "logo.png"), [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()

# --- 2) square logo on white (for light surfaces) -----------------------------
$bmp = New-Object System.Drawing.Bitmap $sz, $sz
$g = New-Graphics $bmp
$g.Clear([System.Drawing.Color]::White)
$path = New-RoundedPath 0 0 $sz $sz 96
$grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, $NAVY_TOP, $NAVY_BOT, 90
$g.FillPath($grad, $path)
$g.SetClip($path)
Draw-Chart $g 70 70 372 372 18
$g.ResetClip()
$bmp.Save((Join-Path $out "logo-light.png"), [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()

# --- 3) horizontal wordmark banner, 1000x300 ----------------------------------
$bw = 1000; $bh = 300
$bmp = New-Object System.Drawing.Bitmap $bw, $bh
$g = New-Graphics $bmp
$brect = New-Object System.Drawing.Rectangle 0, 0, $bw, $bh
$bpath = New-RoundedPath 0 0 $bw $bh 40
$bgrad = New-Object System.Drawing.Drawing2D.LinearGradientBrush $brect, $NAVY_TOP, $NAVY_BOT, 90
$g.FillPath($bgrad, $bpath)
# left mark tile
$tile = New-RoundedPath 36 36 228 228 44
$tileFill = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 22, 36, 60))
$g.FillPath($tileFill, $tile)
$g.SetClip($tile)
Draw-Chart $g 70 70 160 160 12
$g.ResetClip()
# wordmark
$f1 = New-Object System.Drawing.Font "Segoe UI", 70, ([System.Drawing.FontStyle]::Bold), 3
$f2 = New-Object System.Drawing.Font "Segoe UI", 28, ([System.Drawing.FontStyle]::Regular), 3
$wb = New-Object System.Drawing.SolidBrush $WHITE
$sb = New-Object System.Drawing.SolidBrush $SUBTLE
$gb = New-Object System.Drawing.SolidBrush $GREEN
$g.DrawString("Stock", $f1, $wb, 300, 92)
$stockW = $g.MeasureString("Stock", $f1).Width
$g.DrawString("Picker", $f1, $gb, (300 + $stockW), 92)
$g.DrawString("Yahoo Finance for LibreOffice Calc", $f2, $sb, 306, 196)
$bmp.Save((Join-Path $out "logo-banner.png"), [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()

Get-ChildItem $out | ForEach-Object {
  $i = [System.Drawing.Image]::FromFile($_.FullName)
  "{0,-18} {1}x{2}" -f $_.Name, $i.Width, $i.Height
  $i.Dispose()
}
