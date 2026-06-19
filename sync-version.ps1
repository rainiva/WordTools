param(
    [string]$RepoRoot = $PSScriptRoot
)

$ErrorActionPreference = "Stop"

$versionFile = Join-Path $RepoRoot "version.json"
if (-not (Test-Path -LiteralPath $versionFile)) {
    throw "version.json not found at $versionFile"
}

$meta = Get-Content -LiteralPath $versionFile -Raw -Encoding UTF8 | ConvertFrom-Json
$version = [string]$meta.version
$productName = [string]$meta.productName
$company = [string]$meta.company
$copyrightYear = [string]$meta.copyrightYear
$copyrightNotice = "Copyright $([char]0x00A9) $copyrightYear $company"

if ($version -notmatch '^\d+\.\d+\.\d+\.\d+$') {
    throw "version.json: version must use four numeric segments, e.g. 1.2.0.0"
}

function Write-Utf8TextFile {
    param(
        [string]$Path,
        [string]$Content,
        [string]$ExpectedFragment
    )

    $encoding = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($Path, $Content, $encoding)

    # 写入后校验：文件必须存在、可读取、且包含预期片段
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Write verification failed: file not found after write: $Path"
    }

    $readBack = [System.IO.File]::ReadAllText($Path, $encoding)
    if ($readBack -ne $Content) {
        throw "Write verification failed: content mismatch for $Path"
    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedFragment) -and $readBack -notlike "*$ExpectedFragment*") {
        throw "Write verification failed: expected fragment not found in $Path"
    }
}

function Update-FileContent {
    param(
        [string]$Path,
        [scriptblock]$Transform,
        [string]$ExpectedFragment
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "File not found: $Path"
    }

    $content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $updated = & $Transform $content
    if ($null -eq $updated) {
        throw "Failed to transform $Path"
    }

    Write-Utf8TextFile -Path $Path -Content $updated -ExpectedFragment $ExpectedFragment
}

$assemblyInfoPath = Join-Path $RepoRoot "WordTools\Properties\AssemblyInfo.cs"
Update-FileContent -Path $assemblyInfoPath -ExpectedFragment $version -Transform {
    param($content)

    $content = [regex]::Replace($content, '\[assembly: AssemblyProduct\("[^"]*"\)\]', "[assembly: AssemblyProduct(`"$productName`")]")
    $content = [regex]::Replace($content, '\[assembly: AssemblyCopyright\("[^"]*"\)\]', "[assembly: AssemblyCopyright(`"$copyrightNotice`")]")
    $content = [regex]::Replace($content, '\[assembly: AssemblyVersion\("[^"]*"\)\]', "[assembly: AssemblyVersion(`"$version`")]")
    $content = [regex]::Replace($content, '\[assembly: AssemblyFileVersion\("[^"]*"\)\]', "[assembly: AssemblyFileVersion(`"$version`")]")

    if ($content -match 'AssemblyInformationalVersion') {
        return [regex]::Replace($content, '\[assembly: AssemblyInformationalVersion\("[^"]*"\)\]', "[assembly: AssemblyInformationalVersion(`"$version`")]")
    }

    return [regex]::Replace(
        $content,
        '(\[assembly: AssemblyFileVersion\("[^"]*"\)\])',
        "`$1`r`n[assembly: AssemblyInformationalVersion(`"$version`")]"
    )
}

$setupPath = Join-Path $RepoRoot "Setup.iss"
Update-FileContent -Path $setupPath -ExpectedFragment $version -Transform {
    param($content)
    return [regex]::Replace($content, '(#define MyAppVersion ")[^"]*(")', "`${1}$version`${2}")
}

$csprojPath = Join-Path $RepoRoot "WordTools\WordTools.csproj"
$versionParts = $version.Split('.')
$applicationVersion = "$($versionParts[0]).$($versionParts[1]).$($versionParts[2]).*"
Update-FileContent -Path $csprojPath -ExpectedFragment $applicationVersion -Transform {
    param($content)
    return [regex]::Replace($content, '(<ApplicationVersion>)[^<]*(</ApplicationVersion>)', "`${1}$applicationVersion`${2}")
}

Write-Host "Synced version $version from version.json" -ForegroundColor Green

return @{
    Version     = $version
    ProductName = $productName
}
