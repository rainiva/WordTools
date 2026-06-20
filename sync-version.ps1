param(
    [string]$RepoRoot = $PSScriptRoot,
    [ValidateSet('None', 'Patch', 'Minor', 'Major')]
    [string]$Bump = 'None'
)

$ErrorActionPreference = 'Stop'

$versionFile = Join-Path $RepoRoot 'version.json'
if (-not (Test-Path -LiteralPath $versionFile)) {
    throw "version.json not found at $versionFile"
}

function Normalize-SemVer {
    param([string]$Value)

    if ($Value -match '^(\d+)\.(\d+)\.(\d+)\.(\d+)$') {
        return "$($Matches[1]).$($Matches[2]).$($Matches[3])"
    }

    if ($Value -match '^\d+\.\d+\.\d+$') {
        return $Value
    }

    throw 'version.json: version must use semver x.x.x, e.g. 1.3.0'
}

function Bump-SemVer {
    param(
        [string]$Value,
        [ValidateSet('Patch', 'Minor', 'Major')]
        [string]$Kind
    )

    $parts = Normalize-SemVer -Value $Value
    $segments = $parts.Split('.') | ForEach-Object { [int]$_ }

    switch ($Kind) {
        'Major' { return '{0}.0.0' -f ($segments[0] + 1) }
        'Minor' { return '{0}.{1}.0' -f $segments[0], ($segments[1] + 1) }
        'Patch' { return '{0}.{1}.{2}' -f $segments[0], $segments[1], ($segments[2] + 1) }
    }
}

function Write-Utf8TextFile {
    param(
        [string]$Path,
        [string]$Content,
        [string]$ExpectedFragment
    )

    $encoding = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($Path, $Content, $encoding)

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

$meta = Get-Content -LiteralPath $versionFile -Raw -Encoding UTF8 | ConvertFrom-Json
$version = Normalize-SemVer -Value ([string]$meta.version)

if ($Bump -ne 'None') {
    $version = Bump-SemVer -Value $version -Kind $Bump
    $meta.version = $version
    $encoding = [System.Text.UTF8Encoding]::new($false)
    $json = ($meta | ConvertTo-Json -Depth 3) + [Environment]::NewLine
    Write-Utf8TextFile -Path $versionFile -Content $json -ExpectedFragment $version
    Write-Host "Bumped version to $version ($Bump)" -ForegroundColor Yellow
}

$productName = [string]$meta.productName
$company = [string]$meta.company
$copyrightYear = [string]$meta.copyrightYear
$copyrightNotice = "Copyright $([char]0x00A9) $copyrightYear $company"
$assemblyVersion = "$version.0"

$assemblyInfoPath = Join-Path $RepoRoot 'WordTools\Properties\AssemblyInfo.cs'
Update-FileContent -Path $assemblyInfoPath -ExpectedFragment $assemblyVersion -Transform {
    param($content)

    $content = [regex]::Replace($content, '\[assembly: AssemblyProduct\("[^"]*"\)\]', "[assembly: AssemblyProduct(`"$productName`")]")
    $content = [regex]::Replace($content, '\[assembly: AssemblyCopyright\("[^"]*"\)\]', "[assembly: AssemblyCopyright(`"$copyrightNotice`")]")
    $content = [regex]::Replace($content, '\[assembly: AssemblyVersion\("[^"]*"\)\]', "[assembly: AssemblyVersion(`"$assemblyVersion`")]")
    $content = [regex]::Replace($content, '\[assembly: AssemblyFileVersion\("[^"]*"\)\]', "[assembly: AssemblyFileVersion(`"$assemblyVersion`")]")

    if ($content -match 'AssemblyInformationalVersion') {
        return [regex]::Replace($content, '\[assembly: AssemblyInformationalVersion\("[^"]*"\)\]', "[assembly: AssemblyInformationalVersion(`"$version`")]")
    }

    return [regex]::Replace(
        $content,
        '(\[assembly: AssemblyFileVersion\("[^"]*"\)\])',
        "`$1`r`n[assembly: AssemblyInformationalVersion(`"$version`")]"
    )
}

$setupPath = Join-Path $RepoRoot 'Setup.iss'
Update-FileContent -Path $setupPath -ExpectedFragment $version -Transform {
    param($content)
    return [regex]::Replace($content, '(#define MyAppVersion ")[^"]*(")', "`${1}$version`${2}")
}

$csprojPath = Join-Path $RepoRoot 'WordTools\WordTools.csproj'
Update-FileContent -Path $csprojPath -ExpectedFragment $assemblyVersion -Transform {
    param($content)
    return [regex]::Replace($content, '(<ApplicationVersion>)[^<]*(</ApplicationVersion>)', "`${1}$assemblyVersion`${2}")
}

Write-Host "Synced version $version from version.json (assembly $assemblyVersion)" -ForegroundColor Green

return @{
    Version          = $version
    AssemblyVersion  = $assemblyVersion
    ProductName      = $productName
}
