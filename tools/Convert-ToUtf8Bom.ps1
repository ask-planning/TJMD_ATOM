<#
.SYNOPSIS
    リポジトリ内のソースを UTF-8（BOM 付き）・CRLF へ変換します。

.DESCRIPTION
    開発基本方針 §4.1「文字コードは UTF-8（BOM 付き）、改行は CRLF」に合わせます。
    BOM が無いと .aspx の解析や VB のコンパイルで日本語が Shift-JIS として読まれ、
    文字化けやパーサーエラーの原因になります。

    対象: src\ と docs\ 配下、およびリポジトリ直下の設定ファイル
    除外: Backup\ bin\ obj\ packages\ .vs\ .git\

.EXAMPLE
    # 変換対象の確認だけ行う（ファイルは変更しない）
    powershell -ExecutionPolicy Bypass -File tools\Convert-ToUtf8Bom.ps1 -WhatIf

.EXAMPLE
    # 実際に変換する
    powershell -ExecutionPolicy Bypass -File tools\Convert-ToUtf8Bom.ps1

.NOTES
    実行前に必ずコミットまたはバックアップを取ってください。
    変換後は Visual Studio でソリューションを再ビルドしてください。
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # 対象リポジトリのルート。既定はこのスクリプトの 1 つ上の階層。
    [string]$RootPath = (Split-Path -Parent $PSScriptRoot)
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 変換対象の拡張子
$targetExtensions = @('.vb', '.aspx', '.ascx', '.master', '.asax', '.config', '.css', '.js', '.json', '.md', '.sln', '.vbproj', '.editorconfig')

# 除外するフォルダ名
$excludedFolders = @('Backup', 'bin', 'obj', 'packages', '.vs', '.git')

function Test-Excluded {
    param([string]$FullPath)

    $relative = $FullPath.Substring($RootPath.Length).TrimStart('\')
    foreach ($segment in $relative.Split('\')) {
        if ($excludedFolders -contains $segment) {
            return $true
        }
    }
    return $false
}

$utf8WithBom = New-Object System.Text.UTF8Encoding($true)

$converted = 0
$skipped = 0

Get-ChildItem -Path $RootPath -Recurse -File | ForEach-Object {
    $file = $_

    if ($targetExtensions -notcontains $file.Extension.ToLowerInvariant()) {
        return
    }
    if (Test-Excluded -FullPath $file.FullName) {
        return
    }

    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

    # すでに BOM 付きかどうか
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)

    # 内容を UTF-8 として読み込む（BOM があれば除去される）
    $text = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)

    # 改行を CRLF へ統一する
    $normalized = $text -replace "`r`n", "`n"
    $normalized = $normalized -replace "`r", "`n"
    $normalized = $normalized -replace "`n", "`r`n"

    $needsBom = -not $hasBom
    $needsNewline = ($normalized -ne $text)

    if (-not $needsBom -and -not $needsNewline) {
        $skipped++
        return
    }

    $reason = @()
    if ($needsBom) { $reason += 'BOM 付与' }
    if ($needsNewline) { $reason += '改行 CRLF 化' }

    if ($PSCmdlet.ShouldProcess($file.FullName, ($reason -join ' / '))) {
        [System.IO.File]::WriteAllText($file.FullName, $normalized, $utf8WithBom)
        Write-Host ('変換: {0}  [{1}]' -f $file.FullName.Substring($RootPath.Length + 1), ($reason -join ' / '))
    }
    $converted++
}

Write-Host ''
Write-Host ('変換対象: {0} 件 / 変更不要: {1} 件' -f $converted, $skipped)
