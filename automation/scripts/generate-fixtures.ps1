# Generates automation/assets fixtures for batch-insert E2E.
# Safe for binary output: uses .NET APIs only, does not rewrite source files.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$assetsRoot = Join-Path (Split-Path -Parent $scriptDir) "assets"

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

function New-TableTemplateDocx {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$Rows = 12,
        [int]$Columns = 2
    )

    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Add()

        $range = $doc.Range()
        $range.Collapse(0) # wdCollapseStart
        # 与 Word「插入 → 表格」对话框默认一致：
        # DefaultTableBehavior=1 (wdWord9TableBehavior), AutoFitBehavior=0 (wdAutoFitFixed)
        # 实测：列宽约均分正文区（~415pt），自动套用「网格型/Table Grid」实线边框
        $table = $doc.Tables.Add($range, $Rows, $Columns, 1, 0)

        $table.Cell(1, 1).Select() | Out-Null

        $tempPath = [string](Join-Path $env:TEMP ("wordtools-table-template-" + [Guid]::NewGuid().ToString("N") + ".docx"))
        $saveTarget = $tempPath
        $doc.SaveAs([ref]$saveTarget)
        Copy-Item -LiteralPath $tempPath -Destination $Path -Force
        Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue

        return [ordered]@{
            row_count = $Rows
            column_count = $Columns
            style = "word_insert_table_autofit_fixed"
        }
    }
    finally {
        if ($null -ne $doc) {
            $doc.Close($false) | Out-Null
        }
        if ($null -ne $word) {
            $word.Quit() | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }
}

New-Item -ItemType Directory -Path $assetsRoot -Force | Out-Null

$templateMeta = New-TableTemplateDocx -Path (Join-Path $assetsRoot "table-template.docx") -Rows 12 -Columns 2
$null = New-TableTemplateDocx -Path (Join-Path $assetsRoot "blank.docx") -Rows 12 -Columns 2
$templateMeta | ConvertTo-Json | ForEach-Object {
    $manifestPath = Join-Path $assetsRoot "table-template.manifest.json"
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($manifestPath, $_, $utf8NoBom)
}
New-SmallJpeg -Path (Join-Path $assetsRoot "test-small.jpg") -ColorArgb 0xFFCC6633

$selectedDir = Join-Path $assetsRoot "images\selected-4"
$folderRoot = Join-Path $assetsRoot "images\folder-root"
$folderSub = Join-Path $folderRoot "sub-a"
$singleDir = Join-Path $assetsRoot "images\single"

$colors = @(0xFF336699, 0xFF993366, 0xFF669933, 0xFF996633)
for ($i = 0; $i -lt 4; $i++) {
    $name = "{0:D2}.jpg" -f ($i + 1)
    New-SmallJpeg -Path (Join-Path $selectedDir $name) -ColorArgb $colors[$i]
}

for ($i = 0; $i -lt 3; $i++) {
    $name = "{0:D2}.jpg" -f ($i + 1)
    New-SmallJpeg -Path (Join-Path $folderRoot $name) -ColorArgb (0xFF224466 + ($i * 0x1111))
}

for ($i = 0; $i -lt 2; $i++) {
    $name = "{0:D2}.jpg" -f ($i + 1)
    New-SmallJpeg -Path (Join-Path $folderSub $name) -ColorArgb (0xFF662244 + ($i * 0x1111))
}

New-SmallJpeg -Path (Join-Path $singleDir "01.jpg") -ColorArgb 0xFF445566

Write-Output ("Generated fixtures under " + $assetsRoot)
