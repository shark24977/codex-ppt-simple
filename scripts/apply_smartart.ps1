param(
    [Parameter(Mandatory = $true)][string]$PptxPath,
    [Parameter(Mandatory = $true)][string]$SpecPath,
    [int]$SlideOffset = 0
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

$layoutIds = @{
    flow = "urn:microsoft.com/office/officeart/2005/8/layout/process1"
    hierarchy = "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1"
    cycle = "urn:microsoft.com/office/officeart/2005/8/layout/cycle3"
}

function Find-SmartArtLayout($app, [string]$layoutId) {
    for ($index = 1; $index -le $app.SmartArtLayouts.Count; $index++) {
        $layout = $app.SmartArtLayouts.Item($index)
        if ($layout.Id -eq $layoutId) {
            return $layout
        }
    }
    return $null
}

$spec = Get-Content -LiteralPath $SpecPath -Raw -Encoding UTF8 | ConvertFrom-Json
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$bootstrap = $null
$presentation = $null
$converted = 0
$skipped = 0
$failures = @()

try {
    # PowerPoint populates SmartArtLayouts lazily after a presentation window exists.
    $bootstrap = $app.Presentations.Add()
    $presentation = $app.Presentations.Open($PptxPath, $false, $false, $true)
    for ($specIndex = 0; $specIndex -lt $spec.slides.Count; $specIndex++) {
        $slideSpec = $spec.slides[$specIndex]
        if ($slideSpec.type -ne "diagram" -or $slideSpec.diagram.render -ne "smartart") {
            continue
        }
        $kind = [string]$slideSpec.diagram.kind
        if (-not $layoutIds.ContainsKey($kind)) {
            $skipped++
            continue
        }

        $targetIndex = $SlideOffset + $specIndex + 1
        $slide = $presentation.Slides.Item($targetIndex)
        $prefix = "CPTF_SMARTART_${targetIndex}_"
        $fallbackShapes = @()
        $left = [double]::PositiveInfinity
        $top = [double]::PositiveInfinity
        $right = 0.0
        $bottom = 0.0
        for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
            $shape = $slide.Shapes.Item($shapeIndex)
            if ($shape.Name.StartsWith($prefix)) {
                $fallbackShapes += $shape
                $left = [Math]::Min($left, $shape.Left)
                $top = [Math]::Min($top, $shape.Top)
                $right = [Math]::Max($right, $shape.Left + $shape.Width)
                $bottom = [Math]::Max($bottom, $shape.Top + $shape.Height)
            }
        }
        if ($fallbackShapes.Count -eq 0) {
            $skipped++
            continue
        }

        $layout = Find-SmartArtLayout $app $layoutIds[$kind]
        if ($null -eq $layout) {
            $failures += "slide ${targetIndex}: SmartArt layout not found for ${kind}"
            continue
        }

        $smartShape = $null
        try {
            $smartShape = $slide.Shapes.AddSmartArt(
                $layout,
                [single]$left,
                [single]$top,
                [single]($right - $left),
                [single]($bottom - $top)
            )
            $smartShape.Name = "CPTF_NATIVE_SMARTART_${targetIndex}"
            $nodeSpecs = @($slideSpec.diagram.nodes)
            while ($smartShape.SmartArt.AllNodes.Count -lt $nodeSpecs.Count) {
                $null = $smartShape.SmartArt.Nodes.Add()
            }
            while ($smartShape.SmartArt.AllNodes.Count -gt $nodeSpecs.Count) {
                $smartShape.SmartArt.AllNodes.Item($smartShape.SmartArt.AllNodes.Count).Delete()
            }
            for ($nodeIndex = 1; $nodeIndex -le $nodeSpecs.Count; $nodeIndex++) {
                $nodeSpec = $nodeSpecs[$nodeIndex - 1]
                $text = [string]$nodeSpec.label
                if ($nodeSpec.body) {
                    $text += "`n" + [string]$nodeSpec.body
                }
                $smartNode = $smartShape.SmartArt.AllNodes.Item($nodeIndex)
                $smartNode.TextFrame2.TextRange.Text = $text
                if ($kind -eq "hierarchy" -and $nodeIndex -gt 1 -and $nodeSpec.level) {
                    for ($levelIndex = 0; $levelIndex -lt [int]$nodeSpec.level; $levelIndex++) {
                        try { $smartNode.Demote() } catch { break }
                    }
                }
            }
            for ($deleteIndex = $fallbackShapes.Count - 1; $deleteIndex -ge 0; $deleteIndex--) {
                $fallbackShapes[$deleteIndex].Delete()
            }
            $converted++
        }
        catch {
            if ($null -ne $smartShape) {
                try { $smartShape.Delete() } catch {}
            }
            $failures += "slide ${targetIndex}: " + $_.Exception.Message
        }
    }
    $presentation.Save()
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch {}
    }
    if ($null -ne $bootstrap) {
        try { $bootstrap.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($bootstrap) | Out-Null } catch {}
    }
    try { $app.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
}

[pscustomobject]@{
    converted = $converted
    skipped = $skipped
    failures = $failures
} | ConvertTo-Json -Depth 4
