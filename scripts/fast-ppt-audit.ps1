[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PptPath,

    [int]$ExpectedMathCount = -1,

    [double]$AxisTolerancePt = 0.25,

    [string]$RouteNameRegex = '^ROUTE_'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2

function Add-Issue {
    param(
        [System.Collections.Generic.List[object]]$List,
        [string]$Severity,
        [string]$Code,
        [int]$Slide,
        [string]$Shape,
        [string]$Message
    )
    $List.Add([pscustomobject]@{
        severity = $Severity
        code = $Code
        slide = $Slide
        shape = $Shape
        message = $Message
    })
}

function Get-FreeformPoints {
    param($Shape)

    $points = New-Object System.Collections.Generic.List[object]
    try {
        $count = [int]$Shape.Nodes.Count
        for ($i = 1; $i -le $count; $i++) {
            $raw = $Shape.Nodes.Item($i).Points
            $x = $null
            $y = $null
            try {
                $x = [double]$raw.GetValue(0, 0)
                $y = [double]$raw.GetValue(0, 1)
            } catch {
                try {
                    $x = [double]$raw[0, 0]
                    $y = [double]$raw[0, 1]
                } catch {}
            }
            if ($null -ne $x -and $null -ne $y) {
                $points.Add(@($x, $y))
            }
        }
    } catch {}
    return $points
}

$resolved = (Resolve-Path -LiteralPath $PptPath).Path
$issues = New-Object System.Collections.Generic.List[object]
$ppt = $null
$pres = $null
$shapeCount = 0
$routeCount = 0
$mathCount = 0

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($resolved, $true, $false, $false)
    $slideWidth = [double]$pres.PageSetup.SlideWidth
    $slideHeight = [double]$pres.PageSetup.SlideHeight

    for ($s = 1; $s -le $pres.Slides.Count; $s++) {
        $slide = $pres.Slides.Item($s)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $shapeCount++
            $name = [string]$shape.Name
            $left = [double]$shape.Left
            $top = [double]$shape.Top
            $width = [double]$shape.Width
            $height = [double]$shape.Height

            if ($width -lt 0 -or $height -lt 0 -or ($width -eq 0 -and $height -eq 0)) {
                Add-Issue $issues 'error' 'NONPOSITIVE_SIZE' $s $name 'Shape has invalid dimensions.'
            }
            if ($left -lt -0.5 -or $top -lt -0.5 -or ($left + $width) -gt ($slideWidth + 0.5) -or ($top + $height) -gt ($slideHeight + 0.5)) {
                Add-Issue $issues 'error' 'OUT_OF_BOUNDS' $s $name 'Shape extends beyond the slide boundary.'
            }

            try {
                if ([int]$shape.HasTextFrame -eq -1 -and [int]$shape.TextFrame2.HasText -eq -1) {
                    $availableWidth = [double]$shape.Width - [double]$shape.TextFrame2.MarginLeft - [double]$shape.TextFrame2.MarginRight
                    $availableHeight = [double]$shape.Height - [double]$shape.TextFrame2.MarginTop - [double]$shape.TextFrame2.MarginBottom
                    $boundWidth = [double]$shape.TextFrame2.TextRange.BoundWidth
                    $boundHeight = [double]$shape.TextFrame2.TextRange.BoundHeight
                    if ($boundWidth -gt ($availableWidth + 0.75) -or $boundHeight -gt ($availableHeight + 0.75)) {
                        Add-Issue $issues 'error' 'TEXT_OVERFLOW' $s $name ("Rendered text {0:N1}x{1:N1} pt exceeds available box {2:N1}x{3:N1} pt." -f $boundWidth,$boundHeight,$availableWidth,$availableHeight)
                    }
                }
            } catch {
                Add-Issue $issues 'warning' 'TEXT_METRICS_UNAVAILABLE' $s $name 'PowerPoint did not expose rendered text metrics.'
            }

            $progId = $null
            try { $progId = [string]$shape.OLEFormat.ProgID } catch {}
            if ($progId -eq 'Equation.DSMT4') {
                $mathCount++
            }

            if ($name -match $RouteNameRegex) {
                $routeCount++
                $points = Get-FreeformPoints $shape
                if ($points.Count -ge 2) {
                    for ($p = 1; $p -lt $points.Count; $p++) {
                        $x1 = [double]$points[$p - 1][0]
                        $y1 = [double]$points[$p - 1][1]
                        $x2 = [double]$points[$p][0]
                        $y2 = [double]$points[$p][1]
                        $dx = [math]::Abs($x2 - $x1)
                        $dy = [math]::Abs($y2 - $y1)
                        if ($dx -gt $AxisTolerancePt -and $dy -gt $AxisTolerancePt) {
                            Add-Issue $issues 'error' 'NON_ORTHOGONAL_ROUTE' $s $name ("Segment {0} is diagonal: dx={1:N2}, dy={2:N2} pt." -f $p,$dx,$dy)
                        }
                    }
                } else {
                    try {
                        if ([int]$shape.Connector -eq -1 -and [int]$shape.ConnectorFormat.Type -eq 1) {
                            if ($width -gt $AxisTolerancePt -and $height -gt $AxisTolerancePt) {
                                Add-Issue $issues 'error' 'DIAGONAL_STRAIGHT_ROUTE' $s $name 'Straight connector has both nonzero width and height.'
                            }
                        } else {
                            Add-Issue $issues 'warning' 'ROUTE_NODES_UNAVAILABLE' $s $name 'Route nodes were unavailable; validate this route from the route manifest.'
                        }
                    } catch {
                        Add-Issue $issues 'warning' 'ROUTE_NODES_UNAVAILABLE' $s $name 'Route nodes were unavailable; validate this route from the route manifest.'
                    }
                }
            }
        }
    }

    if ($ExpectedMathCount -ge 0 -and $mathCount -ne $ExpectedMathCount) {
        Add-Issue $issues 'error' 'MATHTYPE_COUNT_MISMATCH' 0 '' ("Expected {0} Equation.DSMT4 objects, found {1}." -f $ExpectedMathCount,$mathCount)
    }

    $errors = @($issues | Where-Object { $_.severity -eq 'error' })
    $warnings = @($issues | Where-Object { $_.severity -eq 'warning' })
    [pscustomobject]@{
        ppt_path = $resolved
        slides = [int]$pres.Slides.Count
        shapes = $shapeCount
        routes = $routeCount
        mathtype_equations = $mathCount
        errors = $errors.Count
        warnings = $warnings.Count
        passed = ($errors.Count -eq 0)
        issues = $issues
    } | ConvertTo-Json -Depth 6

    if ($errors.Count -gt 0) { exit 2 }
} finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($ppt) { try { $ppt.Quit() } catch {} }
    if ($pres) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($pres) } catch {} }
    if ($ppt) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
