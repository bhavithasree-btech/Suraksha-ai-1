$ErrorActionPreference = 'Stop'

$source = 'C:\Users\bhavi\Documents\Codex\2026-07-30\m\work\source.pptx'
$outputDir = 'C:\Users\bhavi\Documents\Codex\2026-07-30\m\outputs'
$output = Join-Path $outputDir 'Comfort Helmet - Redesigned.pptx'
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
Copy-Item -LiteralPath $source -Destination $output -Force

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($output, [System.IO.Compression.ZipArchiveMode]::Update)

function Shape([int]$id, [string]$name, [int]$x, [int]$y, [int]$w, [int]$h, [string]$fill, [string]$geom = 'rect', [string]$alpha = '') {
  $opacity = if ($alpha) { '<a:alpha val="' + $alpha + '"/>' } else { '' }
  return @"
<p:sp><p:nvSpPr><p:cNvPr id="$id" name="$name"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$w" cy="$h"/></a:xfrm><a:prstGeom prst="$geom"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="$fill">$opacity</a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>
"@
}

function TextBox([int]$id, [int]$x, [int]$y, [int]$w, [int]$h, [string]$text, [int]$size, [string]$color, [bool]$bold = $false, [string]$align = 'l') {
  $b = if ($bold) { ' b="1"' } else { '' }
  $safe = [Security.SecurityElement]::Escape($text)
  return @"
<p:sp><p:nvSpPr><p:cNvPr id="$id" name="Text $id"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$w" cy="$h"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/><a:lstStyle/><a:p><a:pPr algn="$align"><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="$size"$b><a:solidFill><a:srgbClr val="$color"/></a:solidFill><a:latin typeface="Montserrat"/></a:rPr><a:t>$safe</a:t></a:r><a:endParaRPr lang="en-US" sz="$size"/></a:p></p:txBody></p:sp>
"@
}

$shapes = @()
$shapes += Shape 2 'Background' 0 0 9144000 5143500 '0B1220'
$shapes += Shape 3 'Teal glow' 6500000 -700000 3300000 3300000 '14B8A6' 'ellipse' '13000'
$shapes += Shape 4 'Orange glow' -900000 4000000 2600000 1600000 'F97316' 'ellipse' '9000'
$shapes += Shape 5 'Top accent' 571500 500000 650000 60000 '2DD4BF'
$shapes += TextBox 6 571500 640000 7600000 500000 'DESIGN OVERVIEW' 2700 'F8FAFC' $true
$shapes += TextBox 7 571500 1190000 7400000 300000 'A layered comfort system engineered around the rider.' 1250 'CBD5E1'

$cards = @(
  @{x=571500; color='2DD4BF'; label='01  COOL'; title='Cooling gel packs'; body='Chilled inserts help absorb heat before it builds up inside the helmet.'},
  @{x=3225800; color='38BDF8'; label='02  BREATHE'; title='3D air mesh'; body='A breathable layer encourages airflow and keeps the interior feeling light.'},
  @{x=5880100; color='FB923C'; label='03  SOFTEN'; title='Satin contact layer'; body='A smooth, low-friction finish adds comfort without changing protection.'}
)
$id = 10
foreach ($card in $cards) {
  $shapes += Shape $id "Card $($card.label)" $card.x 1850000 2260000 1980000 '172033' 'roundRect'; $id++
  $shapes += Shape $id 'Card accent' ($card.x + 180000) 2070000 300000 300000 $card.color 'ellipse'; $id++
  $shapes += TextBox $id ($card.x + 600000) 2100000 1400000 220000 $card.label 900 $card.color $true; $id++
  $shapes += TextBox $id ($card.x + 180000) 2580000 1900000 350000 $card.title 1700 'F8FAFC' $true; $id++
  $shapes += TextBox $id ($card.x + 180000) 3100000 1850000 550000 $card.body 950 'CBD5E1'; $id++
}
$shapes += Shape $id 'Bottom line' 571500 4290000 8000000 15000 '334155'; $id++
$shapes += TextBox $id 571500 4500000 8000000 210000 'COOL  |  SMOOTH  |  BREATHABLE' 900 '94A3B8' $true 'c'

$slideXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>$($shapes -join '')</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>
"@

$entry = $zip.GetEntry('ppt/slides/slide8.xml')
$entry.Delete()
$entry = $zip.CreateEntry('ppt/slides/slide8.xml')
$writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
$writer.Write($slideXml)
$writer.Dispose()
$zip.Dispose()
Write-Output $output
