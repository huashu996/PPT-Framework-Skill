$ErrorActionPreference = 'Stop'

function Get-FormulaSkillExactVisiblePresentation {
    param(
        [Parameter(Mandatory = $true)][string]$PresentationPath,
        [switch]$LaunchIfMissing
    )
    $expected = [IO.Path]::GetFullPath($PresentationPath)
    function Find-ExactPresentation {
        try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
        catch { return $null }
        $matches = @()
        foreach ($candidate in @($application.Presentations)) {
            try {
                if ([string]::Equals([IO.Path]::GetFullPath([string]$candidate.FullName), $expected, [StringComparison]::OrdinalIgnoreCase)) {
                    $matches += $candidate
                }
            } catch {}
        }
        if ($matches.Count -gt 1) { throw "Multiple PowerPoint presentations have the same target path: $expected" }
        if ($matches.Count -eq 1) {
            return [pscustomobject]@{ Application = $application; Presentation = $matches[0] }
        }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) } catch {}
        return $null
    }

    $binding = Find-ExactPresentation
    if ($null -ne $binding -or -not $LaunchIfMissing) { return $binding }
    Start-Process -FilePath $expected | Out-Null
    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    while ($null -eq $binding -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 200
        $binding = Find-ExactPresentation
    }
    if ($null -eq $binding) { throw "PowerPoint did not expose the exact presentation within 20 seconds: $expected" }
    return $binding
}
