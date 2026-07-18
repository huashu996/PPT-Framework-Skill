[CmdletBinding()]
param(
    [ValidateSet('detect','preflight','insert','edit','fit','inspect','validate','inspect-validate')]
    [string]$Action = 'detect',

    [string]$PptPath,
    [int]$SlideNumber = 1,
    [string]$ShapeName,
    [double]$Left = 72,
    [double]$Top = 72,
    [double]$Width = 120,
    [double]$Height = 36,
    [int]$ExpectedCount = -1,
    [switch]$Replace
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2

$MathTypeProgId = 'Equation.DSMT4'

function Get-RegistryDefaultValue {
    param([string]$LiteralPath)

    if (-not (Test-Path -LiteralPath $LiteralPath)) {
        return $null
    }
    return (Get-ItemProperty -LiteralPath $LiteralPath).'(default)'
}

function Get-MathTypeRegistration {
    $progIdPath = "Registry::HKEY_CLASSES_ROOT\$MathTypeProgId"
    $clsid = Get-RegistryDefaultValue "$progIdPath\CLSID"
    $server = $null
    if ($clsid) {
        $server = Get-RegistryDefaultValue "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\LocalServer32"
        if (-not $server) {
            $server = Get-RegistryDefaultValue "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\LocalServer"
        }
    }

    $exePath = $null
    if ($server) {
        if ($server -match '^"([^"]+\.exe)"') {
            $exePath = $Matches[1]
        } elseif ($server -match '^(.+?\.exe)(?:\s|$)') {
            $exePath = $Matches[1].Trim()
        }
    }

    if (-not $exePath) {
        foreach ($candidate in @(
            'C:\Program Files (x86)\MathType\MathType.exe',
            'C:\Program Files\MathType\MathType.exe'
        )) {
            if (Test-Path -LiteralPath $candidate) {
                $exePath = $candidate
                break
            }
        }
    }

    [pscustomobject]@{
        Available = [bool]($clsid -and $exePath -and (Test-Path -LiteralPath $exePath))
        ProgId = $MathTypeProgId
        Description = Get-RegistryDefaultValue $progIdPath
        Clsid = $clsid
        Executable = $exePath
    }
}

function Write-JsonResult {
    param($Value)
    $Value | ConvertTo-Json -Depth 8
}

function Resolve-PresentationPath {
    if ([string]::IsNullOrWhiteSpace($PptPath)) {
        throw '-PptPath is required for this action.'
    }
    if (-not (Test-Path -LiteralPath $PptPath -PathType Leaf)) {
        throw "PowerPoint file not found: $PptPath"
    }
    return (Resolve-Path -LiteralPath $PptPath).Path
}

function Get-MathTypeShapeInfo {
    param($Presentation)

    $items = @()
    foreach ($slide in $Presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            $progId = $null
            try {
                $progId = [string]$shape.OLEFormat.ProgID
            } catch {}

            if ($progId -eq $MathTypeProgId) {
                $items += [pscustomobject]@{
                    Slide = [int]$slide.SlideIndex
                    Name = [string]$shape.Name
                    ProgId = $progId
                    Left = [math]::Round([double]$shape.Left, 3)
                    Top = [math]::Round([double]$shape.Top, 3)
                    Width = [math]::Round([double]$shape.Width, 3)
                    Height = [math]::Round([double]$shape.Height, 3)
                }
            }
        }
    }
    return @($items)
}

function Get-UniqueMathShapeName {
    param($Slide)

    $index = 1
    while ($true) {
        $candidate = "MATH_S$($Slide.SlideIndex)_$index"
        try {
            $null = $Slide.Shapes.Item($candidate)
            $index++
        } catch {
            return $candidate
        }
    }
}

function Get-UsablePowerPointApplication {
    $application = $null
    try {
        $application = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    } catch {}

    if ($application -and $application.Presentations) {
        return [pscustomobject]@{
            Application = $application
            OwnsApplication = $false
            Source = 'active'
        }
    }

    if ($application) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) } catch {}
        $application = $null
    }

    $application = New-Object -ComObject PowerPoint.Application
    if (-not $application -or -not $application.Presentations) {
        if ($application) {
            try { $application.Quit() } catch {}
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) } catch {}
        }
        throw 'PowerPoint COM returned an unusable application. Open PowerPoint once and rerun; the tool will attach to the active instance without closing unrelated presentations.'
    }

    return [pscustomobject]@{
        Application = $application
        OwnsApplication = $true
        Source = 'new'
    }
}

function Open-TargetPresentation {
    param(
        $Application,
        [string]$Path,
        [bool]$ReadOnly,
        [bool]$WithWindow
    )

    foreach ($candidate in $Application.Presentations) {
        if ([string]::Equals([string]$candidate.FullName, $Path, [StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{
                Presentation = $candidate
                OpenedHere = $false
            }
        }
    }

    $opened = $Application.Presentations.Open($Path, $ReadOnly, $false, $WithWindow)
    return [pscustomobject]@{
        Presentation = $opened
        OpenedHere = $true
    }
}

$registration = Get-MathTypeRegistration
if ($Action -eq 'detect') {
    Write-JsonResult ([pscustomobject]@{
        Action = 'detect'
        Success = $registration.Available
        MathType = $registration
    })
    if (-not $registration.Available) {
        throw "MathType OLE server '$MathTypeProgId' is not available."
    }
    return
}

if (-not $registration.Available) {
    throw "MathType OLE server '$MathTypeProgId' is not available. Run -Action detect first."
}

$pptContext = $null
if ($Action -eq 'preflight') {
    try {
        $pptContext = Get-UsablePowerPointApplication
        Write-JsonResult ([pscustomobject]@{
            Action = 'preflight'
            Success = $true
            MathType = $registration
            PowerPoint = [pscustomobject]@{
                Source = $pptContext.Source
                Version = [string]$pptContext.Application.Version
                OpenPresentations = [int]$pptContext.Application.Presentations.Count
            }
        })
    } finally {
        if ($pptContext) {
            if ($pptContext.OwnsApplication) {
                try { $pptContext.Application.Quit() } catch {}
            }
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($pptContext.Application) } catch {}
        }
    }
    return
}

$resolvedPptPath = Resolve-PresentationPath
$effectiveAction = $Action
if ($Action -eq 'inspect-validate') {
    $effectiveAction = 'validate'
}
$ppt = $null
$presentation = $null
$leaveOpen = $false
$ownsPpt = $false
$openedPresentation = $false

try {
    $pptContext = Get-UsablePowerPointApplication
    $ppt = $pptContext.Application
    $ownsPpt = [bool]$pptContext.OwnsApplication

    switch ($effectiveAction) {
        'insert' {
            $openContext = Open-TargetPresentation $ppt $resolvedPptPath $false $false
            $presentation = $openContext.Presentation
            $openedPresentation = [bool]$openContext.OpenedHere
            if ($SlideNumber -lt 1 -or $SlideNumber -gt $presentation.Slides.Count) {
                throw "SlideNumber $SlideNumber is outside 1..$($presentation.Slides.Count)."
            }
            if ($Width -le 0 -or $Height -le 0) {
                throw 'Width and Height must be positive PowerPoint point values.'
            }

            $slide = $presentation.Slides.Item($SlideNumber)
            if ([string]::IsNullOrWhiteSpace($ShapeName)) {
                $ShapeName = Get-UniqueMathShapeName $slide
            } else {
                try {
                    $existing = $slide.Shapes.Item($ShapeName)
                    if ($Replace) {
                        $existing.Delete()
                    } else {
                        throw "A shape named '$ShapeName' already exists on slide $SlideNumber. Use -Replace to replace it."
                    }
                } catch {
                    if ($_.Exception.Message -like "A shape named*") {
                        throw
                    }
                }
            }

            $shape = $slide.Shapes.AddOLEObject(
                $Left,
                $Top,
                $Width,
                $Height,
                $MathTypeProgId
            )
            $shape.Name = $ShapeName
            $actualProgId = [string]$shape.OLEFormat.ProgID
            if ($actualProgId -ne $MathTypeProgId) {
                throw "Inserted OLE object has unexpected ProgID '$actualProgId'."
            }

            $presentation.Save()
            Write-JsonResult ([pscustomobject]@{
                Action = 'insert'
                Success = $true
                PptPath = $resolvedPptPath
                Shape = [pscustomobject]@{
                    Slide = $SlideNumber
                    Name = $shape.Name
                    ProgId = $actualProgId
                    Left = [double]$shape.Left
                    Top = [double]$shape.Top
                    Width = [double]$shape.Width
                    Height = [double]$shape.Height
                }
            })
        }

        'edit' {
            if ([string]::IsNullOrWhiteSpace($ShapeName)) {
                throw '-ShapeName is required for -Action edit.'
            }
            $ppt.Visible = -1
            $openContext = Open-TargetPresentation $ppt $resolvedPptPath $false $true
            $presentation = $openContext.Presentation
            $openedPresentation = [bool]$openContext.OpenedHere
            if ($SlideNumber -lt 1 -or $SlideNumber -gt $presentation.Slides.Count) {
                throw "SlideNumber $SlideNumber is outside 1..$($presentation.Slides.Count)."
            }
            $shape = $presentation.Slides.Item($SlideNumber).Shapes.Item($ShapeName)
            $actualProgId = [string]$shape.OLEFormat.ProgID
            if ($actualProgId -ne $MathTypeProgId) {
                throw "Shape '$ShapeName' is '$actualProgId', not '$MathTypeProgId'."
            }

            $ppt.ActiveWindow.View.GotoSlide($SlideNumber)
            $shape.Select()
            $shape.OLEFormat.DoVerb(2)
            $leaveOpen = $true

            Write-JsonResult ([pscustomobject]@{
                Action = 'edit'
                Success = $true
                PptPath = $resolvedPptPath
                Slide = $SlideNumber
                ShapeName = $ShapeName
                Message = 'PowerPoint and the MathType editor were opened. Save the presentation after editing.'
            })
        }

        'fit' {
            if ([string]::IsNullOrWhiteSpace($ShapeName)) {
                throw '-ShapeName is required for -Action fit.'
            }
            if ($Width -le 0 -or $Height -le 0) {
                throw 'Width and Height must be positive reserved-rectangle dimensions.'
            }
            $openContext = Open-TargetPresentation $ppt $resolvedPptPath $false $false
            $presentation = $openContext.Presentation
            $openedPresentation = [bool]$openContext.OpenedHere
            if ($SlideNumber -lt 1 -or $SlideNumber -gt $presentation.Slides.Count) {
                throw "SlideNumber $SlideNumber is outside 1..$($presentation.Slides.Count)."
            }
            $shape = $presentation.Slides.Item($SlideNumber).Shapes.Item($ShapeName)
            $actualProgId = [string]$shape.OLEFormat.ProgID
            if ($actualProgId -ne $MathTypeProgId) {
                throw "Shape '$ShapeName' is '$actualProgId', not '$MathTypeProgId'."
            }

            $currentWidth = [double]$shape.Width
            $currentHeight = [double]$shape.Height
            if ($currentWidth -le 0 -or $currentHeight -le 0) {
                throw "Shape '$ShapeName' has a non-positive size."
            }
            $scale = [math]::Min($Width / $currentWidth, $Height / $currentHeight)
            $newWidth = $currentWidth * $scale
            $newHeight = $currentHeight * $scale
            $shape.LockAspectRatio = -1
            $shape.Width = $newWidth
            $shape.Left = $Left + (($Width - [double]$shape.Width) / 2)
            $shape.Top = $Top + (($Height - [double]$shape.Height) / 2)
            $presentation.Save()

            Write-JsonResult ([pscustomobject]@{
                Action = 'fit'
                Success = $true
                PptPath = $resolvedPptPath
                Shape = [pscustomobject]@{
                    Slide = $SlideNumber
                    Name = [string]$shape.Name
                    ProgId = $actualProgId
                    Left = [math]::Round([double]$shape.Left, 3)
                    Top = [math]::Round([double]$shape.Top, 3)
                    Width = [math]::Round([double]$shape.Width, 3)
                    Height = [math]::Round([double]$shape.Height, 3)
                }
            })
        }

        'inspect' {
            $openContext = Open-TargetPresentation $ppt $resolvedPptPath $true $false
            $presentation = $openContext.Presentation
            $openedPresentation = [bool]$openContext.OpenedHere
            $items = @(Get-MathTypeShapeInfo $presentation)
            Write-JsonResult ([pscustomobject]@{
                Action = 'inspect'
                Success = $true
                PptPath = $resolvedPptPath
                Count = $items.Count
                Shapes = $items
            })
        }

        'validate' {
            $openContext = Open-TargetPresentation $ppt $resolvedPptPath $true $false
            $presentation = $openContext.Presentation
            $openedPresentation = [bool]$openContext.OpenedHere
            $items = @(Get-MathTypeShapeInfo $presentation)
            $errors = @()

            if ($ExpectedCount -ge 0 -and $items.Count -ne $ExpectedCount) {
                $errors += "Expected $ExpectedCount MathType objects but found $($items.Count)."
            }
            if (-not [string]::IsNullOrWhiteSpace($ShapeName)) {
                if (-not ($items | Where-Object { $_.Slide -eq $SlideNumber -and $_.Name -eq $ShapeName })) {
                    $errors += "MathType shape '$ShapeName' was not found on slide $SlideNumber."
                }
            }

            $slideWidth = [double]$presentation.PageSetup.SlideWidth
            $slideHeight = [double]$presentation.PageSetup.SlideHeight
            foreach ($item in $items) {
                if ($item.Width -le 0 -or $item.Height -le 0) {
                    $errors += "MathType shape '$($item.Name)' has a non-positive size."
                }
                if (
                    $item.Left -lt 0 -or
                    $item.Top -lt 0 -or
                    ($item.Left + $item.Width) -gt $slideWidth -or
                    ($item.Top + $item.Height) -gt $slideHeight
                ) {
                    $errors += "MathType shape '$($item.Name)' is outside the slide bounds."
                }
            }

            $valid = $errors.Count -eq 0
            Write-JsonResult ([pscustomobject]@{
                Action = $Action
                Success = $valid
                PptPath = $resolvedPptPath
                Count = $items.Count
                Shapes = $items
                Errors = $errors
            })
            if (-not $valid) {
                throw ($errors -join ' ')
            }
        }
    }
} finally {
    if (-not $leaveOpen) {
        if ($presentation -and $openedPresentation) {
            try { $presentation.Close() } catch {}
        }
        if ($ppt -and $ownsPpt) {
            try { $ppt.Quit() } catch {}
        }
    }

    if ($presentation) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) } catch {}
    }
    if ($ppt) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
