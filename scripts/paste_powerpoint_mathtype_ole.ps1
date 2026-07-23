param(
    [Parameter(Mandatory = $true)][int]$SlideIndex,
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][string]$PresentationPath,
    [string]$TargetName,
    [double[]]$Rect,
    [double]$Padding = 4,
    [switch]$Save
)

$ErrorActionPreference = 'Stop'
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
if ($Rect -and $Rect.Count -ne 4) { throw '-Rect must contain left, top, width, height.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'MathType clipboard paste requires STA. Run this script with powershell -Sta.'
}

Add-Type -AssemblyName System.Windows.Forms
$formats = @([Windows.Forms.Clipboard]::GetDataObject().GetFormats())
if (($formats -notcontains 'MathType EF') -and ($formats -notcontains 'Embedded Object')) {
    throw 'Clipboard does not contain a MathType OLE object. Copy from MathType and run the clipboard inspector first.'
}

$powerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$wantedPath = [IO.Path]::GetFullPath($PresentationPath)
$presentation = $powerPoint.ActivePresentation
if ($null -eq $presentation) { throw 'Prepared PowerPoint session has no active presentation.' }
if (-not [string]::Equals([IO.Path]::GetFullPath($presentation.FullName), $wantedPath, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Prepared PowerPoint session is no longer active for: $wantedPath"
}
$powerPoint.ActiveWindow.View.GotoSlide($SlideIndex)
$currentSlide = $powerPoint.ActiveWindow.View.Slide.SlideIndex
if ($currentSlide -ne $SlideIndex) { throw "Failed to activate slide $SlideIndex." }
$slide = $presentation.Slides.Item($SlideIndex)

$before = $slide.Shapes.Count
[void]$powerPoint.ActiveWindow.View.PasteSpecial(10)
$after = $slide.Shapes.Count
if ($after -le $before) { throw 'PowerPoint PasteSpecial(10) did not add a shape.' }

$shape = $slide.Shapes.Item($after)
$progId = $null
try { $progId = $shape.OLEFormat.ProgID } catch {}
if ($progId -ne 'Equation.DSMT4') {
    for ($index = $after; $index -gt $before; $index--) {
        try { $slide.Shapes.Item($index).Delete() } catch {}
    }
    throw "Paste produced '$progId' instead of Equation.DSMT4; inserted shapes were removed."
}

$shape.Name = $Name
$shape.LockAspectRatio = -1

$left = $null; $top = $null; $width = $null; $height = $null
if ($Rect) {
    $left = $Rect[0]; $top = $Rect[1]; $width = $Rect[2]; $height = $Rect[3]
}
elseif ($TargetName) {
    $target = $slide.Shapes.Item($TargetName)
    $left = $target.Left; $top = $target.Top; $width = $target.Width; $height = $target.Height
}

if ($null -ne $width) {
    $maxWidth = [Math]::Max(1, $width - 2 * $Padding)
    $maxHeight = [Math]::Max(1, $height - 2 * $Padding)
    $scale = [Math]::Min($maxWidth / $shape.Width, $maxHeight / $shape.Height)
    if ($scale -lt 1) { $shape.Width = $shape.Width * $scale }
    $shape.Left = $left + ($width - $shape.Width) / 2
    $shape.Top = $top + ($height - $shape.Height) / 2
    $insideTarget = $shape.Left -ge $left -and $shape.Top -ge $top -and ($shape.Left + $shape.Width) -le ($left + $width) -and ($shape.Top + $shape.Height) -le ($top + $height)
    if (-not $insideTarget) {
        try { $shape.Delete() } catch {}
        throw "Inserted equation $Name is outside its target bounds; failed object was removed."
    }
    if ($null -ne $target) {
        try { $target.TextFrame2.TextRange.Text = '' } catch {}
        try { $target.Fill.Visible = 0 } catch {}
        try { $target.Line.Visible = 0 } catch {}
        try { $target.ZOrder(1) } catch {}
    }
}

if ($shape.Parent.SlideIndex -ne $SlideIndex -or [string]$shape.Name -ne $Name) {
    try { $shape.Delete() } catch {}
    throw "Inserted equation failed slide/name verification and was removed."
}

if ($Save) { $presentation.Save() }
$stopwatch.Stop()

[pscustomobject]@{
    slide = $SlideIndex
    name = $shape.Name
    prog_id = $progId
    left = [Math]::Round($shape.Left, 2)
    top = [Math]::Round($shape.Top, 2)
    width = [Math]::Round($shape.Width, 2)
    height = [Math]::Round($shape.Height, 2)
    saved = [bool]$Save
    prepared_session = $true
    inside_target = $insideTarget
    elapsed_ms = $stopwatch.ElapsedMilliseconds
} | ConvertTo-Json -Depth 4
