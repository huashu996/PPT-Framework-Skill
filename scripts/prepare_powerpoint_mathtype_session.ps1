param(
    [Parameter(Mandatory = $true)][string]$PresentationPath,
    [int]$SlideIndex = 1,
    [string]$ManifestPath
)

$ErrorActionPreference = 'Stop'
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$path = [IO.Path]::GetFullPath($PresentationPath)
if (-not (Test-Path -LiteralPath $path)) { throw "Presentation does not exist: $path" }

. (Join-Path $PSScriptRoot 'powerpoint_window_binding.ps1')
$binding = Get-FormulaSkillExactVisiblePresentation -PresentationPath $path
$alreadyOpen = $null -ne $binding
if ($null -eq $binding) { $binding = Get-FormulaSkillExactVisiblePresentation -PresentationPath $path -LaunchIfMissing }
$powerPoint = $binding.Application
$presentation = $binding.Presentation
$reusedPowerPoint = $alreadyOpen
if (-not [string]::Equals([IO.Path]::GetFullPath([string]$presentation.FullName), $path, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Active presentation does not match requested output.'
}
if ($SlideIndex -lt 1 -or $SlideIndex -gt $presentation.Slides.Count) {
    throw "Slide index $SlideIndex is outside 1..$($presentation.Slides.Count)."
}

try { $presentation.Windows.Item(1).Activate() } catch {}
$powerPoint.ActiveWindow.ViewType = 9
$powerPoint.ActiveWindow.View.GotoSlide($SlideIndex)
try { $powerPoint.ActiveWindow.Panes.Item(2).Activate() } catch {}
try { $powerPoint.ActiveWindow.Selection.Unselect() } catch {}

$targetCount = 0
if ($ManifestPath) {
    $manifestFile = [IO.Path]::GetFullPath($ManifestPath)
    $manifest = [IO.File]::ReadAllText($manifestFile, [Text.Encoding]::UTF8) | ConvertFrom-Json
    foreach ($equation in @($manifest.equations)) {
        $target = $equation.target
        if ($null -eq $target -or [string]$target.type -notin @('powerpoint_shape','powerpoint_rect')) {
            throw "Equation $($equation.id) does not have a PowerPoint target."
        }
        $slideNumber = [int]$target.slide_index
        if ($slideNumber -lt 1 -or $slideNumber -gt $presentation.Slides.Count) {
            throw "Invalid target slide for $($equation.id): $slideNumber"
        }
        if ([string]$target.type -eq 'powerpoint_shape') {
            $targetName = [string]$target.shape_name
            try { [void]$presentation.Slides.Item($slideNumber).Shapes.Item($targetName); $targetCount++ }
            catch { throw "Missing PowerPoint target shape '$targetName' on slide $slideNumber." }
        }
        else { $targetCount++ }
    }
}

$stopwatch.Stop()
[pscustomobject]@{
    presentation = $presentation.FullName
    slide = $SlideIndex
    reused_powerpoint = $reusedPowerPoint
    target_count = $targetCount
    elapsed_ms = $stopwatch.ElapsedMilliseconds
} | ConvertTo-Json -Depth 4
