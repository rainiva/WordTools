# Layer 3: smoke test - discovery, COM add-in load check, optional image insert.

[CmdletBinding()]
param(
    [string]$RepoRoot,
    [ValidateSet("discovery", "com_load", "insert_image")]
    [string]$Mode = "discovery",
    [ValidateSet("Word", "WPS", "Both")]
    [string]$HostTarget = "Both",
    [string]$ImagePath = "",
    [string]$ScreenshotDir = "",
    [string]$ProgId = "WordTools.ThisAddIn"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "Matrix.Common.ps1")

function Test-ComAddinLoaded {
    param(
        [Parameter(Mandatory = $true)]
        $Application,
        [Parameter(Mandatory = $true)]
        [string]$ProgId
    )

    foreach ($addin in @($Application.COMAddIns)) {
        if ([string]$addin.ProgId -eq $ProgId -and $addin.Connect) {
            return $true
        }
    }
    return $false
}

function Invoke-WordSmoke {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModeName,
        [Parameter(Mandatory = $true)]
        [string]$AddinProgId,
        [string]$ImageFile = ""
    )

    $result = [ordered]@{
        host = "Word"
        launched = $false
        addin_loaded = $false
        image_inserted = $false
        image_persisted = $false
        error = $null
    }

    $word = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $result.launched = $true

        if ($ModeName -eq "discovery") {
            return $result
        }

        $result.addin_loaded = Test-ComAddinLoaded -Application $word -ProgId $AddinProgId
        if ($ModeName -eq "com_load") {
            $result.image_persisted = $result.addin_loaded
            return $result
        }

        if ($ModeName -eq "insert_image" -and $result.addin_loaded -and -not [string]::IsNullOrWhiteSpace($ImageFile) -and (Test-Path -LiteralPath $ImageFile)) {
            $doc = $word.Documents.Add()
            $selection = $word.Selection
            $inlineShape = $selection.InlineShapes.AddPicture($ImageFile)
            $result.image_inserted = ($null -ne $inlineShape)
            $tempDoc = Join-Path $env:TEMP ("wordtools-smoke-" + [Guid]::NewGuid().ToString("N") + ".docx")
            $doc.SaveAs([ref]$tempDoc)
            $shapeCountBeforeClose = $doc.InlineShapes.Count
            $doc.Close($false)

            $verifyDoc = $word.Documents.Open($tempDoc, $false, $true)
            $result.image_persisted = ($verifyDoc.InlineShapes.Count -ge 1) -and ($shapeCountBeforeClose -ge 1)
            $verifyDoc.Close($false)
            Remove-Item -LiteralPath $tempDoc -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        $result.error = $_.Exception.Message
    }
    finally {
        if ($null -ne $word) {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }

    return $result
}

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Get-MatrixRepoRoot
}

$wordResult = $null
$wpsResult = [ordered]@{
    host = "WPS"
    launched = $false
    addin_loaded = $false
    image_inserted = $false
    image_persisted = $false
    error = "WPS COM smoke is not implemented in this phase."
}

if ($HostTarget -in @("Word", "Both")) {
    if ($Mode -eq "discovery") {
        $wordPath = Get-RegistryExePath -AppName "WINWORD.EXE"
        $wordResult = [ordered]@{
            host = "Word"
            launched = ($null -ne $wordPath)
            addin_loaded = $false
            image_inserted = $false
            image_persisted = $false
            path = if ($null -eq $wordPath) { $null } else { [string]$wordPath.path }
        }
    }
    else {
        $wordResult = Invoke-WordSmoke -ModeName $Mode -AddinProgId $ProgId -ImageFile $ImagePath
    }
}

if ($Mode -eq "discovery") {
    $wpsList = @(Find-WpsExecutable)
    $wpsResult = [ordered]@{
        host = "WPS"
        launched = ($wpsList.Count -gt 0)
        addin_loaded = $false
        image_inserted = $false
        image_persisted = $false
        path = if ($wpsList.Count -gt 0) { $wpsList[0] } else { $null }
    }
}

$pass = $true
if ($HostTarget -in @("Word", "Both") -and $null -ne $wordResult) {
    if ($Mode -eq "com_load") {
        $pass = [bool]$wordResult.addin_loaded
    }
    elseif ($Mode -eq "insert_image") {
        $pass = [bool]$wordResult.addin_loaded -and [bool]$wordResult.image_persisted
    }
    else {
        $pass = [bool]$wordResult.launched
    }
}

$payload = [ordered]@{
    layer = "smoke"
    mode = $Mode
    host_target = $HostTarget
    word = $wordResult
    wps = $wpsResult
    image_path = $ImagePath
    screenshot_dir = $ScreenshotDir
    pass = $pass
}

Write-MatrixJsonResult -Payload $payload
