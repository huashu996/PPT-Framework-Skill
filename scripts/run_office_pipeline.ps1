param(
    [Parameter(Mandatory = $true)][string]$ManifestPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$output = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $output)) { throw "Office target does not exist: $output" }
if ([IO.Path]::GetExtension($output).ToLowerInvariant() -ne '.pptx') {
    throw 'Paper-Fig-Skill Office output must be PPTX.'
}
& (Join-Path $PSScriptRoot 'insert_powerpoint_office.ps1') -PresentationPath $output -ManifestPath $ManifestPath
