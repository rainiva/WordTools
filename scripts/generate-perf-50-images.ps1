# Generates 50 small JPEGs for PERF-03 batch insert benchmark.
# Binary-safe: uses .NET System.Drawing only.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir
$dest = Join-Path $repoRoot 'automation\assets\images\selected-50'

function New-SmallJpeg {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [int]$Width = 64,
        [int]$Height = 48,
        [int]$ColorArgb = 0xFF336699
    )

    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object System.Drawing.Bitmap $Width, $Height
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($ColorArgb))
            try {
                $graphics.FillRectangle($brush, 0, 0, $Width, $Height)
            }
            finally {
                $brush.Dispose()
            }
        }
        finally {
            $graphics.Dispose()
        }

        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    }
    finally {
        $bitmap.Dispose()
    }
}

New-Item -ItemType Directory -Path $dest -Force | Out-Null

1..50 | ForEach-Object {
    $path = Join-Path $dest ('{0:D2}.jpg' -f $_)
    $color = 0xFF000000 -bor (($_ * 12345) -band 0xFFFFFF)
    New-SmallJpeg -Path $path -ColorArgb $color
}

Write-Host "Created 50 JPEGs under $dest"
