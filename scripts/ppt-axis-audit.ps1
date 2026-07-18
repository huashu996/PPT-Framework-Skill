param(
    [Parameter(Mandatory = $true)]
    [string]$PptPath,

    [string[]]$HorizontalNames = @(),

    [string[]]$VerticalNames = @(),

    [double]$TolerancePt = 0.25,

    [switch]$RequireConnected,

    [switch]$RequireEndArrow
)

$ErrorActionPreference = 'Stop'

$resolvedPpt = [System.IO.Path]::GetFullPath($PptPath)
if (-not (Test-Path -LiteralPath $resolvedPpt)) {
    throw "PowerPoint file not found: $resolvedPpt"
}

$createdApp = $false
$openedPresentation = $false
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
} catch {
    $app = New-Object -ComObject PowerPoint.Application
    $createdApp = $true
}

$presentation = $null
foreach ($candidate in $app.Presentations) {
    if ($candidate.FullName -eq $resolvedPpt) {
        $presentation = $candidate
        break
    }
}
if ($null -eq $presentation) {
    $presentation = $app.Presentations.Open($resolvedPpt, $false, $true, $false)
    $openedPresentation = $true
}

try {
    $slide = $presentation.Slides.Item(1)
    $results = @()
    $horizontalExpanded = @($HorizontalNames | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
    $verticalExpanded = @($VerticalNames | ForEach-Object { $_ -split ',' } | Where-Object { $_ })

    foreach ($spec in @(
        @{ Orientation = 'horizontal'; Names = $horizontalExpanded },
        @{ Orientation = 'vertical'; Names = $verticalExpanded }
    )) {
        foreach ($name in $spec.Names) {
            $shape = $slide.Shapes.Item($name)
            $isConnector = $false
            try { $isConnector = ($shape.Connector -eq -1) } catch {}

            $axisDelta = if ($spec.Orientation -eq 'horizontal') {
                [double]$shape.Height
            } else {
                [double]$shape.Width
            }

            $beginConnected = $false
            $endConnected = $false
            if ($isConnector) {
                $beginConnected = [bool]$shape.ConnectorFormat.BeginConnected
                $endConnected = [bool]$shape.ConnectorFormat.EndConnected
            }

            $endArrow = [int]$shape.Line.EndArrowheadStyle
            $pass = ($isConnector -and $axisDelta -le $TolerancePt)
            if ($RequireConnected) {
                $pass = ($pass -and $beginConnected -and $endConnected)
            }
            if ($RequireEndArrow) {
                $pass = ($pass -and $endArrow -ne 1)
            }

            $results += [pscustomobject]@{
                Name = $name
                Orientation = $spec.Orientation
                IsConnector = $isConnector
                AxisDeltaPt = [math]::Round($axisDelta, 3)
                BeginConnected = $beginConnected
                EndConnected = $endConnected
                EndArrowheadStyle = $endArrow
                Pass = $pass
            }
        }
    }

    [pscustomobject]@{
        Action = 'axis-audit'
        Success = (@($results | Where-Object { -not $_.Pass }).Count -eq 0)
        PptPath = $resolvedPpt
        TolerancePt = $TolerancePt
        Count = $results.Count
        Failures = @($results | Where-Object { -not $_.Pass })
        Results = $results
    } | ConvertTo-Json -Depth 5
} finally {
    if ($openedPresentation) {
        $presentation.Close()
    }
    if ($createdApp) {
        $app.Quit()
    }
}
