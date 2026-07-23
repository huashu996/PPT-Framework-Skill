param([Parameter(Mandatory=$true)][string]$PresentationPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'powerpoint_window_binding.ps1')
$binding=Get-FormulaSkillExactVisiblePresentation -PresentationPath $PresentationPath
if($null-eq$binding){throw "The exact PowerPoint presentation is not open: $PresentationPath"}
$presentation=$binding.Presentation;$powerPoint=$binding.Application
try{$presentation.Save();$presentation.Close()}finally{try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)}catch{};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint)}catch{}}
[pscustomobject]@{presentation=[IO.Path]::GetFullPath($PresentationPath);saved=$true;closed=$true}|ConvertTo-Json
