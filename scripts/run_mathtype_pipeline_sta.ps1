param(
    [Parameter(Mandatory=$true)][string]$ManifestPath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [Parameter(Mandatory=$true)][string]$BuildDir,
    [string]$PostProcessScript
)
$ErrorActionPreference='Stop'
if([Threading.Thread]::CurrentThread.ApartmentState-ne'STA'){throw 'run_mathtype_pipeline_sta.ps1 must run once in STA.'}
$manifestFile=[IO.Path]::GetFullPath($ManifestPath);$output=[IO.Path]::GetFullPath($OutputPath);$build=[IO.Path]::GetFullPath($BuildDir)
$stagePath=Join-Path $build 'mathtype-stage.txt';function Set-Stage([string]$Value){[IO.File]::WriteAllText($stagePath,$Value,[Text.Encoding]::UTF8)}
Set-Stage 'loading_manifest'
$manifest=[IO.File]::ReadAllText($manifestFile,[Text.Encoding]::UTF8)|ConvertFrom-Json
if([string]$manifest.mode-ne'MathType'){throw 'STA pipeline accepts MathType manifests only.'}
if(-not(Test-Path -LiteralPath $output)){throw "Prepared output does not exist: $output"}
$environmentPath=Join-Path $build 'environment.json'
& (Join-Path $PSScriptRoot 'detect_environment.ps1') -OutputPath $environmentPath|Out-Null
Set-Stage 'environment_detected'
$environment=[IO.File]::ReadAllText($environmentPath,[Text.Encoding]::UTF8)|ConvertFrom-Json
if(-not[bool]$environment.mathtype.available){throw 'Equation.DSMT4 is not registered.'}
function Get-MathTypeEditorProcess {
    return @(Get-Process -Name 'MathType' -ErrorAction SilentlyContinue|Where-Object{$_.MainWindowHandle-ne0}|Sort-Object StartTime|Select-Object -First 1)
}
function Invoke-OfficePasteWithRetry([scriptblock]$Operation) {
    for($attempt=1;$attempt-le10;$attempt++){
        try{return & $Operation}catch{
            $rejected=$_.Exception.HResult-eq-2147418111 -or $_.Exception.Message -match 'RPC_E_CALL_REJECTED|rejected by callee'
            if(-not$rejected -or $attempt-eq10){throw}
            Start-Sleep -Milliseconds (150*$attempt)
        }
    }
}
$mathTypeProcess=Get-MathTypeEditorProcess
if($mathTypeProcess.Count-eq0){
    $exe=[string]$environment.mathtype.executable
    if([string]::IsNullOrWhiteSpace($exe)){throw 'MathType is registered but no executable path was found.'}
    $exe=$exe.Trim().Trim('"');Start-Process -FilePath $exe|Out-Null
    $deadline=[DateTime]::UtcNow.AddSeconds(15)
    do{Start-Sleep -Milliseconds 250;$mathTypeProcess=Get-MathTypeEditorProcess}while($mathTypeProcess.Count-eq0 -and [DateTime]::UtcNow-lt$deadline)
    if($mathTypeProcess.Count-eq0){throw 'MathType did not expose an editor window.'}
}
$mathTypePid=[int]$mathTypeProcess[0].Id;$equations=@($manifest.equations);$extension=[IO.Path]::GetExtension($output).ToLowerInvariant();$records=New-Object Collections.Generic.List[object];$pending=New-Object Collections.Generic.List[object]
Set-Stage "mathtype_ready_pid_$mathTypePid"
if($extension-ne'.pptx'){throw 'Paper-Fig-Skill MathType output must be PPTX.'}
$firstSlide=[int]$equations[0].target.slide_index
& (Join-Path $PSScriptRoot 'prepare_powerpoint_mathtype_session.ps1') -PresentationPath $output -SlideIndex $firstSlide -ManifestPath $manifestFile|Out-Null
Set-Stage 'office_target_prepared'
for($index=0;$index-lt$equations.Count;$index++){
    $item=$equations[$index];$result=$null;$editor=$null
    try{
        $powerPoint=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        $presentation=$powerPoint.ActivePresentation
        $existing=$presentation.Slides.Item([int]$item.target.slide_index).Shapes.Item([string]$item.id)
        $existingProgId=$null
        try{$existingProgId=[string]$existing.OLEFormat.ProgID}catch{}
        if($existingProgId-eq'Equation.DSMT4'){
            $existingRecord=[pscustomobject]@{
                slide=[int]$item.target.slide_index
                name=[string]$item.id
                prog_id=$existingProgId
                left=[Math]::Round($existing.Left,2)
                top=[Math]::Round($existing.Top,2)
                width=[Math]::Round($existing.Width,2)
                height=[Math]::Round($existing.Height,2)
                saved=$false
                prepared_session=$true
                inside_target=$true
                existing=$true
            }
            $records.Add([pscustomobject]@{id=[string]$item.id;editor=$null;insertion=$existingRecord})
            Set-Stage "formula_$($index+1)_existing"
            continue
        }
    }catch{}
    $targetName=$null;$rect=$null;$padding=4
    if([string]$item.target.type-eq'powerpoint_shape'){$targetName=[string]$item.target.shape_name}else{$rect=[double[]]@($item.target.rect)}
    if($null-ne$item.target.padding){$padding=[double]$item.target.padding}
    $lastClipboardError=$null
    for($clipboardAttempt=1;$clipboardAttempt-le3;$clipboardAttempt++){
        try{
            Set-Stage "formula_$($index+1)_activating_mathtype_attempt_$clipboardAttempt"
            $mathTypeProcess=Get-MathTypeEditorProcess
            if($mathTypeProcess.Count-eq0){throw 'MathType editor window disappeared during the batch.'}
            $mathTypePid=[int]$mathTypeProcess[0].Id
            $editorJson=& (Join-Path $PSScriptRoot 'update_mathtype_editor.ps1') -MathMlPath ([string]$item.mathml_file) -MathTypeProcessId $mathTypePid
            $editor=$editorJson|ConvertFrom-Json
            if(-not[bool]$editor.ole_clipboard_ready){throw "MathType clipboard gate failed for $($item.id)."}
            # The system clipboard now contains the OLE object. Paste it on the
            # very next statement so no foreground switch, stage-file write, or
            # layout preparation can expose a clipboard race.
            $result=Invoke-OfficePasteWithRetry { & (Join-Path $PSScriptRoot 'paste_powerpoint_mathtype_ole.ps1') -PresentationPath $output -SlideIndex ([int]$item.target.slide_index) -Name ([string]$item.id) -TargetName $targetName -Rect $rect -Padding $padding }
            break
        }catch{
            $lastClipboardError=$_.Exception.Message
            Set-Stage "formula_$($index+1)_retrying_after_attempt_$clipboardAttempt"
            if($lastClipboardError -match 'MathType rejected the imported clipboard equation'){
                # The modal has already received Enter. Do not paste the same
                # unmodified MathML again; leave this formula pending and exit
                # the current formula retry circle.
                break
            }
            if($clipboardAttempt-lt3){Start-Sleep -Milliseconds 180}
        }
    }
    if($null-eq$result){
        $pending.Add([pscustomobject]@{
            id=[string]$item.id
            formula_index=$index+1
            stage=[string](Get-Content -Raw -Encoding UTF8 -LiteralPath $stagePath)
            verification=$lastClipboardError
        })
        Set-Stage "formula_$($index+1)_pending_continuing"
        continue
    }
    Set-Stage "formula_$($index+1)_pasted"
    $records.Add([pscustomobject]@{id=[string]$item.id;editor=$editor;insertion=($result|ConvertFrom-Json)})
    Start-Sleep -Milliseconds 180
}
$postProcessResult=$null
if(-not[string]::IsNullOrWhiteSpace($PostProcessScript)){
    $postProcessPath=[IO.Path]::GetFullPath($PostProcessScript)
    if(-not(Test-Path -LiteralPath $postProcessPath)){throw "Post-process script does not exist: $postProcessPath"}
    Set-Stage 'running_postprocess_in_prepared_powerpoint_session'
    $postProcessResult=& $postProcessPath -PresentationPath $output
}
$comValidation=[pscustomobject]@{method='insertion_time_powerpoint_com';blocker_count=$pending.Count;equations=@($records.ToArray()|ForEach-Object{$_.insertion})}
Set-Stage 'closing_saved_powerpoint_presentation'
& (Join-Path $PSScriptRoot 'close_powerpoint_exact.ps1') -PresentationPath $output|Out-Null
Set-Stage 'complete'
[pscustomobject]@{mode='MathType';output=$output;mathtype_process_id=$mathTypePid;inserted_count=$records.Count;pending_count=$pending.Count;pending=$pending.ToArray();records=$records.ToArray();postprocess=$postProcessResult;com_validation=$comValidation}|ConvertTo-Json -Depth 12
