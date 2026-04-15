<#
.SYNOPSIS
  Word2Png.ps1 — 把 Word 文档 (.doc/.docx) 每一页转换为 PNG.

.DESCRIPTION
  使用 Microsoft Word COM 自动化, 逐页选择范围 -> CopyAsPicture -> 从剪贴板
  拿到图像 -> 以指定 DPI 重新采样 -> 保存为 PNG.
  无需 PDF 库、无需 ImageMagick, 只要本机装了 Word + .NET Framework 即可.

.PARAMETER WordPath
  源 .doc 或 .docx 的完整路径.

.PARAMETER OutputDir
  PNG 输出目录, 不存在会自动创建.

.PARAMETER Dpi
  输出分辨率, 默认 200. 数值越大图片越清晰、文件越大.

.NOTES
  本脚本被 ZwCadWordToPng.lsp 同步调用.
  所有错误信息会写入 <OutputDir>\Word2Png.log 方便排查.
#>

param(
    [Parameter(Mandatory = $true)][string] $WordPath,
    [Parameter(Mandatory = $true)][string] $OutputDir,
    [int] $Dpi = 200
)

# 强制 STA, 才能使用 System.Windows.Forms.Clipboard
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    # 重新以 STA 模式启动自己
    $psi = @(
        '-NoLogo', '-NoProfile', '-STA',
        '-ExecutionPolicy', 'Bypass',
        '-File', "`"$PSCommandPath`"",
        '-WordPath', "`"$WordPath`"",
        '-OutputDir', "`"$OutputDir`"",
        '-Dpi', "$Dpi"
    )
    $p = Start-Process -FilePath powershell.exe -ArgumentList $psi -Wait -PassThru -WindowStyle Hidden
    exit $p.ExitCode
}

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

$logPath = Join-Path $OutputDir 'Word2Png.log'
function Write-Log([string]$msg) {
    $line = "{0:yyyy-MM-dd HH:mm:ss} {1}" -f (Get-Date), $msg
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    Write-Host $line
}

Write-Log "Start: WordPath=$WordPath OutputDir=$OutputDir Dpi=$Dpi"

# Word 常量
$wdStatisticPages   = 2
$wdGoToPage         = 1
$wdGoToAbsolute     = 1
$wdAlertsNone       = 0

$word = $null
$doc  = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible      = $false
    $word.DisplayAlerts = $wdAlertsNone

    # 以只读方式打开, 不写最近文件
    $doc = $word.Documents.Open(
        [ref]$WordPath,              # FileName
        [ref]$false,                 # ConfirmConversions
        [ref]$true,                  # ReadOnly
        [ref]$false,                 # AddToRecentFiles
        [ref][Type]::Missing,        # PasswordDocument
        [ref][Type]::Missing,        # PasswordTemplate
        [ref]$false,                 # Revert
        [ref][Type]::Missing,        # WritePasswordDocument
        [ref][Type]::Missing,        # WritePasswordTemplate
        [ref][Type]::Missing,        # Format
        [ref][Type]::Missing,        # Encoding
        [ref]$false,                 # Visible
        [ref][Type]::Missing,        # OpenAndRepair
        [ref][Type]::Missing,        # DocumentDirection
        [ref]$false,                 # NoEncodingDialog
        [ref][Type]::Missing         # XMLTransform
    )

    # 先重分页一次, 保证分页数正确
    $doc.Repaginate() | Out-Null
    $pageCount = [int]$doc.ComputeStatistics($wdStatisticPages)
    Write-Log "PageCount = $pageCount"

    if ($pageCount -lt 1) {
        throw "Word 统计到的页数为 $pageCount, 无法继续."
    }

    $baseName  = [IO.Path]::GetFileNameWithoutExtension($WordPath)
    $padWidth  = [Math]::Max(3, $pageCount.ToString().Length)

    for ($i = 1; $i -le $pageCount; $i++) {

        # 计算当前页的起止位置
        $sel = $word.Selection
        $sel.GoTo([ref]$wdGoToPage, [ref]$wdGoToAbsolute, [ref]$i, [ref][Type]::Missing) | Out-Null
        $startPos = $sel.Start

        if ($i -lt $pageCount) {
            $next = $i + 1
            $sel.GoTo([ref]$wdGoToPage, [ref]$wdGoToAbsolute, [ref]$next, [ref][Type]::Missing) | Out-Null
            $endPos = [Math]::Max($sel.Start - 1, $startPos + 1)
        } else {
            $endPos = $doc.Content.End - 1
        }

        $range = $doc.Range($startPos, $endPos)
        [System.Windows.Forms.Clipboard]::Clear()
        $range.CopyAsPicture()

        $img = $null
        # Word 粘贴稳定性: 有时第一次 GetImage 返回 null, 短暂重试几次
        for ($try = 0; $try -lt 5 -and $img -eq $null; $try++) {
            Start-Sleep -Milliseconds 120
            if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
                $img = [System.Windows.Forms.Clipboard]::GetImage()
            }
        }

        if ($img -eq $null) {
            Write-Log "page $i : 剪贴板无图像, 已跳过"
            continue
        }

        # 按 DPI 重新采样 (Word CopyAsPicture 默认 96dpi, 放大到 Dpi)
        $scale    = $Dpi / 96.0
        $newW     = [int][Math]::Max(1, [Math]::Round($img.Width  * $scale))
        $newH     = [int][Math]::Max(1, [Math]::Round($img.Height * $scale))

        $bmp = New-Object System.Drawing.Bitmap $newW, $newH
        $bmp.SetResolution($Dpi, $Dpi)

        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $g.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $g.Clear([System.Drawing.Color]::White)
            $g.DrawImage($img, 0, 0, $newW, $newH)
        } finally {
            $g.Dispose()
        }

        $idx     = $i.ToString().PadLeft($padWidth, '0')
        $pngPath = Join-Path $OutputDir ("{0}_page{1}.png" -f $baseName, $idx)
        $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

        $bmp.Dispose()
        $img.Dispose()

        Write-Log "page $i -> $pngPath"
    }

    [System.Windows.Forms.Clipboard]::Clear()
    Write-Log "Done."
    exit 0
}
catch {
    Write-Log ("ERROR: " + $_.Exception.Message)
    Write-Log $_.ScriptStackTrace
    exit 1
}
finally {
    if ($doc  -ne $null) { try { $doc.Close([ref]$false) } catch {} }
    if ($word -ne $null) { try { $word.Quit([ref]$false) } catch {} }

    # 释放 COM 对象
    foreach ($o in @($doc, $word)) {
        if ($o -ne $null) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
