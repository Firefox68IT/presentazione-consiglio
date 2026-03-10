param(
    [string]$ProjectRoot = (Get-Location).Path,
    [string]$OutputName = 'Viaggio - Presentazione Consiglio.pptx',
    [string]$ClipFolderName = 'clip_video_15s'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$imgExt = '.jpg','.jpeg','.png','.heic'
$vidExt = '.mp4','.mov','.avi','.m4v'

$projectRoot = (Resolve-Path $ProjectRoot).Path
$outputPath = Join-Path $projectRoot $OutputName
$reportPath = Join-Path $projectRoot 'presentazione_generata.txt'
$clipDir = Join-Path $projectRoot $ClipFolderName

function Find-Ffmpeg {
    $cmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $fallback = Join-Path $env:USERPROFILE 'tools\ffmpeg\bin\ffmpeg.exe'
    if (Test-Path $fallback) { return $fallback }
    throw 'ffmpeg non trovato. Installalo o aggiungilo al PATH.'
}

$ffmpeg = Find-Ffmpeg

function Get-RelativeTimeLabel([datetime]$dt) { $dt.ToString('dd MMM yyyy - HH:mm') }
function Get-DateLabelFromFolder([string]$folderName) {
    if ($folderName -match '^(\d{4})(\d{2})(\d{2})') {
        return ([datetime]::ParseExact($matches[1]+$matches[2]+$matches[3], 'yyyyMMdd', $null)).ToString('dd MMMM yyyy')
    }
    return $folderName
}
function Get-PlaceName([string]$folderName) {
    if ($folderName -match '^\d{8}\s+(.+)$') { return $matches[1] }
    return $folderName
}
function Read-SectionText([string]$folderPath, [string]$folderName, [int]$photoCount, [int]$videoCount) {
    $textPath = Join-Path $folderPath 'Testo.txt'
    if (Test-Path $textPath) {
        $txt = (Get-Content $textPath -Raw -Encoding UTF8).Trim()
        if ($txt) { return $txt }
    }
    $place = Get-PlaceName $folderName
    if ($photoCount -gt 0 -or $videoCount -gt 0) {
        return "Tappa del viaggio a $place. La sezione raccoglie la documentazione fotografica e video disponibile per questa visita."
    }
    return "Tappa del viaggio a $place. In questa cartella non risultano contenuti multimediali locali, quindi la sezione resta come introduzione narrativa del percorso."
}
function Add-TextBox($slide, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height, [int]$fontSize, [bool]$bold, [int]$rgb, [int]$fillRgb, [double]$fillTransparency) {
    $shape = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    $shape.Fill.ForeColor.RGB = $fillRgb
    $shape.Fill.Transparency = $fillTransparency
    $shape.Line.Visible = 0
    $shape.TextFrame.MarginLeft = 8
    $shape.TextFrame.MarginRight = 8
    $shape.TextFrame.MarginTop = 4
    $shape.TextFrame.MarginBottom = 4
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.TextRange.Font.Name = 'Aptos'
    $shape.TextFrame.TextRange.Font.Size = $fontSize
    $shape.TextFrame.TextRange.Font.Bold = [int]$bold
    $shape.TextFrame.TextRange.Font.Color.RGB = $rgb
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.AutoSize = 0
    return $shape
}
function Fit-ShapeInBox($shape, [double]$left, [double]$top, [double]$width, [double]$height) {
    $shape.LockAspectRatio = -1
    $origW = [double]$shape.Width
    $origH = [double]$shape.Height
    if ($origW -le 0 -or $origH -le 0) { return }
    $scale = [Math]::Min($width / $origW, $height / $origH)
    $newW = $origW * $scale
    $newH = $origH * $scale
    $shape.Width = $newW
    $shape.Height = $newH
    $shape.Left = $left + (($width - $newW) / 2)
    $shape.Top = $top + (($height - $newH) / 2)
}
function Ensure-Clip([string]$sectionName, $videoFile) {
    if (-not (Test-Path $clipDir)) { New-Item -ItemType Directory -Path $clipDir | Out-Null }
    $datePrefix = if ($sectionName -match '^(\d{8})') { $matches[1] } else { $videoFile.LastWriteTime.ToString('yyyyMMdd') }
    $clipName = "$datePrefix - $($videoFile.BaseName) - clip15s.mp4"
    $clipPath = Join-Path $clipDir $clipName
    if (-not (Test-Path $clipPath)) {
        & $ffmpeg -y -i $videoFile.FullName -t 15 -c:v libx264 -preset veryfast -crf 23 -c:a aac -movflags +faststart $clipPath | Out-Null
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $clipPath)) {
            throw "Impossibile creare la clip per $($videoFile.Name)"
        }
    }
    return $clipPath
}

$ppt = $null
$pres = $null
$window = $null
try {
    $sections = Get-ChildItem $projectRoot -Directory | Where-Object { $_.Name -match '^2026|^\d{8}\s' } | Sort-Object Name
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $pres = $ppt.Presentations.Add()
    $pres.PageSetup.SlideSize = 3
    $window = $pres.NewWindow()

    $slideW = [double]$pres.PageSetup.SlideWidth
    $slideH = [double]$pres.PageSetup.SlideHeight
    $marginX = 20
    $titleY = 12
    $titleH = 30
    $mediaY = 54
    $mediaH = $slideH - 110
    $footerY = $slideH - 44
    $footerH = 22
    $contentW = $slideW - (2 * $marginX)
    $mediaX = 30
    $mediaW = $slideW - 60

    $cover = $pres.Slides.Add(1, 12)
    $cover.FollowMasterBackground = 0
    $cover.Background.Fill.ForeColor.RGB = 15393258
    Add-TextBox $cover 'Viaggio' 50 120 ($slideW - 100) 60 24 $true 3027998 15393258 0 | Out-Null
    Add-TextBox $cover 'Presentazione Consiglio' 50 195 ($slideW - 100) 36 16 $true 13134786 12495754 0 | Out-Null
    Add-TextBox $cover ('Generata da Codex | ' + (Get-Date)) 50 245 ($slideW - 100) 28 12 $false 4207920 15393258 0 | Out-Null

    $summary = New-Object System.Collections.Generic.List[string]
    $summary.Add('Presentazione generata con la skill travel-council-presentation.') | Out-Null
    $summary.Add('') | Out-Null

    foreach ($section in $sections) {
        $files = @(Get-ChildItem $section.FullName -File -ErrorAction SilentlyContinue)
        $photos = @($files | Where-Object { $imgExt -contains $_.Extension.ToLower() } | Sort-Object LastWriteTime, Name)
        $videos = @($files | Where-Object { $vidExt -contains $_.Extension.ToLower() } | Sort-Object LastWriteTime, Name)
        $sectionText = Read-SectionText $section.FullName $section.Name $photos.Count $videos.Count
        $place = Get-PlaceName $section.Name

        $sectionSlide = $pres.Slides.Add($pres.Slides.Count + 1, 12)
        $sectionSlide.FollowMasterBackground = 0
        $sectionSlide.Background.Fill.ForeColor.RGB = 15393258
        Add-TextBox $sectionSlide $place 40 28 ($slideW - 80) 46 22 $true 2952225 15393258 0 | Out-Null
        Add-TextBox $sectionSlide ((Get-DateLabelFromFolder $section.Name) + ' | ' + $photos.Count + ' foto | ' + $videos.Count + ' video') 40 82 ($slideW - 80) 28 11 $true 16119285 13134786 0 | Out-Null
        Add-TextBox $sectionSlide $sectionText 40 130 ($slideW - 80) 310 14 $false 4207920 16776440 0 | Out-Null

        $media = @(($photos + $videos) | Sort-Object LastWriteTime, Name)
        foreach ($file in $media) {
            if ($imgExt -contains $file.Extension.ToLower()) {
                $slide = $pres.Slides.Add($pres.Slides.Count + 1, 12)
                $slide.FollowMasterBackground = 0
                $slide.Background.Fill.ForeColor.RGB = 1382197
                $pic = $slide.Shapes.AddPicture($file.FullName, 0, -1, 0, 0)
                Fit-ShapeInBox $pic $mediaX $mediaY $mediaW $mediaH
                Add-TextBox $slide $place $marginX $titleY $contentW $titleH 16 $true 16777215 0 0.45 | Out-Null
                Add-TextBox $slide (Get-RelativeTimeLabel $file.LastWriteTime) $marginX $footerY $contentW $footerH 10 $false 16119285 0 0.35 | Out-Null
            } else {
                $clipPath = Ensure-Clip $section.Name $file
                $slide = $pres.Slides.Add($pres.Slides.Count + 1, 12)
                $slide.FollowMasterBackground = 0
                $slide.Background.Fill.ForeColor.RGB = 1382197
                $videoShape = $slide.Shapes.AddMediaObject2($clipPath, $false, $true, $mediaX, $mediaY, $mediaW, $mediaH)
                Fit-ShapeInBox $videoShape $mediaX $mediaY $mediaW $mediaH
                Add-TextBox $slide ($place + ' | Clip video') $marginX $titleY $contentW $titleH 16 $true 16777215 0 0.45 | Out-Null
                Add-TextBox $slide ('Avvio automatico dopo 3 secondi | ' + (Get-RelativeTimeLabel $file.LastWriteTime)) $marginX $footerY $contentW $footerH 10 $false 16119285 0 0.35 | Out-Null
                $videoShape.AnimationSettings.PlaySettings.PlayOnEntry = $true
                $videoShape.AnimationSettings.PlaySettings.RewindMovie = $true
                $window.Activate()
                $window.View.GotoSlide($slide.SlideIndex)
                $videoShape.Select()
                $ppt.CommandBars.ExecuteMso('MoviePlayFullScreen')
                $slide.TimeLine.MainSequence.Item(1).Timing.TriggerDelayTime = 3
            }
        }

        $summary.Add(($section.Name + ' -> ' + $photos.Count + ' foto, ' + $videos.Count + ' video')) | Out-Null
    }

    if (Test-Path $outputPath) { Remove-Item -Force $outputPath }
    $pres.SaveAs($outputPath)
    Set-Content -Path $reportPath -Value ($summary -join [Environment]::NewLine) -Encoding UTF8
}
finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt) { try { $ppt.Quit() } catch {} }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
