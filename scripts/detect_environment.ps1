param(
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Test-ComApplication([string]$ProgId) {
    $registrationKey = "Registry::HKEY_CLASSES_ROOT\$ProgId"
    try {
        $active = [Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        $version = $active.Version
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($active)
        return [pscustomobject]@{ available = $true; version = $version; method = 'active_latebound'; type_library_error = $false; latebound_fallback_recommended = $false; error = $null }
    }
    catch {
        $message = $_.Exception.ToString()
        $shortMessage = $_.Exception.Message
        $typeLibraryError = $message -match 'TYPE_E_CANTLOADLIBRARY|0x80029C4A|80029C4A'
        if (Test-Path -LiteralPath $registrationKey) {
            return [pscustomobject]@{ available = $true; version = $null; method = 'registered_not_launched'; type_library_error = $typeLibraryError; latebound_fallback_recommended = $typeLibraryError; error = $(if ($typeLibraryError) { $shortMessage } else { $null }) }
        }
        return [pscustomobject]@{ available = $false; version = $null; method = $null; type_library_error = $typeLibraryError; latebound_fallback_recommended = $typeLibraryError; error = $message }
    }
}

$powerPoint = Test-ComApplication 'PowerPoint.Application'

$mathType = [ordered]@{
    available = $false
    prog_id = 'Equation.DSMT4'
    description = $null
    clsid = $null
    executable = $null
}

$progIdKey = 'Registry::HKEY_CLASSES_ROOT\Equation.DSMT4'
if (Test-Path -LiteralPath $progIdKey) {
    $mathType.available = $true
    $mathType.description = (Get-Item -LiteralPath $progIdKey).GetValue('')
    $mathType.clsid = (Get-Item -LiteralPath (Join-Path $progIdKey 'CLSID')).GetValue('')
    if ($mathType.clsid) {
        $serverKey = "Registry::HKEY_CLASSES_ROOT\CLSID\$($mathType.clsid)\LocalServer32"
        if (Test-Path -LiteralPath $serverKey) {
            $mathType.executable = (Get-Item -LiteralPath $serverKey).GetValue('')
        }
    }
}

$addIns = @()
if ($powerPoint.available) {
    try {
        $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        foreach ($addIn in $app.AddIns) {
            if ($addIn.Name -match 'MathType|WIRIS|Design Science') {
                $addIns += [pscustomobject]@{ name = $addIn.Name; path = $addIn.FullName; loaded = [bool]$addIn.Loaded }
            }
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
    catch {}
}

$report = [pscustomobject]@{
    timestamp = (Get-Date).ToString('o')
    operating_system = [Environment]::OSVersion.VersionString
    powershell = $PSVersionTable.PSVersion.ToString()
    powerpoint = $powerPoint
    mathtype = $mathType
    powerpoint_mathtype_addins = $addIns
}

$json = $report | ConvertTo-Json -Depth 8
if ($OutputPath) {
    $parent = Split-Path -Parent ([IO.Path]::GetFullPath($OutputPath))
    if ($parent) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $json | Set-Content -LiteralPath $OutputPath -Encoding utf8
}
$json
