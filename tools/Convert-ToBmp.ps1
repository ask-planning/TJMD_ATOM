#Requires -Version 5.1
<#
.SYNOPSIS
    画像を SVF の様式に貼れる 24bit BMP へ変換します。

.DESCRIPTION
    SVF の様式に貼る画像は BMP が無難です（既存様式でも BMP が使われています）。
    透過 PNG をそのまま使うと透明部分が黒くなることがあるため、
    **白を敷いてから** 24bit（アルファ無し）の BMP として保存します。

    変換した BMP は、様式と同じフォルダ（D:\SVFDATA\form\ATOM\）へ置いてください。
    SVFX-Designer は手元の PC、PDF を作るのは SVF サーバーのため、
    ローカルのパスを参照していると出力されません。

.PARAMETER Path
    変換元の画像ファイル。PNG / JPEG / GIF など。

.PARAMETER OutputPath
    出力先。省略すると変換元と同じ場所に拡張子 .bmp で保存します。

.PARAMETER BackgroundColor
    透過部分に敷く色。既定は白。

.EXAMPLE
    .\Convert-ToBmp.ps1 -Path .\logo.png

.EXAMPLE
    .\Convert-ToBmp.ps1 -Path .\logo.png -OutputPath D:\work\tajima_logo.bmp
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path,

    [Parameter(Position = 1)]
    [string] $OutputPath,

    [string] $BackgroundColor = 'White'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

$sourcePath = (Resolve-Path -LiteralPath $Path).ProviderPath

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = [System.IO.Path]::ChangeExtension($sourcePath, '.bmp')
}
$outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
    throw "出力先のフォルダがありません。$outputDirectory"
}

$source = $null
$target = $null
$graphics = $null

try {
    $source = [System.Drawing.Image]::FromFile($sourcePath)

    # 24bit（アルファ無し）で作り直す
    $target = New-Object System.Drawing.Bitmap($source.Width, $source.Height, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $target.SetResolution($source.HorizontalResolution, $source.VerticalResolution)

    $graphics = [System.Drawing.Graphics]::FromImage($target)
    $graphics.Clear([System.Drawing.Color]::FromName($BackgroundColor))
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.DrawImage($source, 0, 0, $source.Width, $source.Height)

    $target.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Bmp)

    Write-Host "変換しました。" -ForegroundColor Green
    Write-Host "  元    : $sourcePath"
    Write-Host "  出力  : $OutputPath"
    Write-Host "  大きさ: $($source.Width) x $($source.Height) px"
    Write-Host ""
    Write-Host "様式と同じフォルダ（D:\SVFDATA\form\ATOM\）へ置いてください。"
}
finally {
    if ($null -ne $graphics) { $graphics.Dispose() }
    if ($null -ne $target) { $target.Dispose() }
    if ($null -ne $source) { $source.Dispose() }
}
