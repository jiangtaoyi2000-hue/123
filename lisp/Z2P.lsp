;;; -----------------------------------------------------------------------------
;;; Z2P.lsp  -  中望CAD / AutoCAD 单文件插件
;;;
;;; 功能: 一键把 Word 文档 (.doc / .docx) 的每一页转成 PNG, 顺序插入到当前图纸.
;;; 命令: Z2P
;;;
;;; 全自包含:
;;;   - 内嵌 PowerShell 脚本 (Word COM + 剪贴板 + System.Drawing 重采样)
;;;   - 运行时落盘到 %TEMP%, 用完即删, 无需额外文件
;;;
;;; 加载方式:
;;;   (load "D:/path/to/Z2P.lsp")
;;;   或追加到 acaddoc.lsp / ZWCAD 启动加载列表实现自动加载.
;;;
;;; 依赖:
;;;   - Windows + Microsoft Word (COM 自动化)
;;;   - Windows PowerShell 5.1+ (Win10/Win11 自带)
;;; -----------------------------------------------------------------------------

(vl-load-com)

;;; --- 通用工具 ---------------------------------------------------------------

(defun z2p:slash (s)
  (vl-string-translate "\\" "/" s)
)

(defun z2p:ensure-dir (dir / fso)
  (setq dir (z2p:slash dir))
  (if (not (vl-file-directory-p dir))
    (progn
      (setq fso (vlax-create-object "Scripting.FileSystemObject"))
      (vl-catch-all-apply
        '(lambda () (vlax-invoke fso 'CreateFolder dir)))
      (vlax-release-object fso)
    )
  )
  dir
)

(defun z2p:temp-dir ( / wsh tmp)
  (setq wsh (vlax-create-object "WScript.Shell"))
  (setq tmp (vlax-invoke wsh 'ExpandEnvironmentStrings "%TEMP%"))
  (vlax-release-object wsh)
  (z2p:slash tmp)
)

(defun z2p:list-pngs (dir / files)
  (setq files (vl-directory-files dir "*.png" 1))
  (if files (vl-sort files '<) nil)
)

;; 读 PNG 原始像素尺寸 (用来按图形单位宽度算出高度)
(defun z2p:png-size (path / img w h)
  (vl-catch-all-apply
    '(lambda ()
       (setq img (vlax-create-object "WIA.ImageFile"))
       (vlax-invoke img 'LoadFile path)
       (setq w (vlax-get img 'Width))
       (setq h (vlax-get img 'Height))
       (vlax-release-object img)))
  (if (and w h (> w 0) (> h 0)) (list w h) nil)
)

;;; --- 内嵌 PowerShell 脚本 (Word -> PNG) -------------------------------------
;;; 注意:
;;;   - 字符串中 ASCII-only, 避免落盘后被错误编码解释.
;;;   - LISP -> PS 调用方加 -STA 参数, 故 PS 内不再需要自重启 STA.

(defun z2p:ps-lines ()
  (list
    "param("
    "    [Parameter(Mandatory = $true)][string] $WordPath,"
    "    [Parameter(Mandatory = $true)][string] $OutputDir,"
    "    [int] $Dpi = 200"
    ")"
    ""
    "$ErrorActionPreference = 'Stop'"
    ""
    "Add-Type -AssemblyName System.Windows.Forms"
    "Add-Type -AssemblyName System.Drawing"
    ""
    "if (-not (Test-Path -LiteralPath $OutputDir)) {"
    "    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null"
    "}"
    ""
    "$logPath = Join-Path $OutputDir 'Z2P.log'"
    "function Write-Log([string]$msg) {"
    "    $line = \"{0:yyyy-MM-dd HH:mm:ss} {1}\" -f (Get-Date), $msg"
    "    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8"
    "    Write-Host $line"
    "}"
    ""
    "Write-Log \"Start WordPath=$WordPath OutputDir=$OutputDir Dpi=$Dpi\""
    ""
    "$wdStatisticPages = 2"
    "$wdGoToPage       = 1"
    "$wdGoToAbsolute   = 1"
    "$wdAlertsNone     = 0"
    ""
    "$word = $null"
    "$doc  = $null"
    ""
    "try {"
    "    $word = New-Object -ComObject Word.Application"
    "    $word.Visible       = $false"
    "    $word.DisplayAlerts = $wdAlertsNone"
    ""
    "    $doc = $word.Documents.Open("
    "        [ref]$WordPath, [ref]$false, [ref]$true, [ref]$false,"
    "        [ref][Type]::Missing, [ref][Type]::Missing, [ref]$false,"
    "        [ref][Type]::Missing, [ref][Type]::Missing, [ref][Type]::Missing,"
    "        [ref][Type]::Missing, [ref]$false, [ref][Type]::Missing,"
    "        [ref][Type]::Missing, [ref]$false, [ref][Type]::Missing"
    "    )"
    ""
    "    $doc.Repaginate() | Out-Null"
    "    $pageCount = [int]$doc.ComputeStatistics($wdStatisticPages)"
    "    Write-Log \"PageCount = $pageCount\""
    ""
    "    if ($pageCount -lt 1) {"
    "        throw \"Page count is $pageCount, cannot continue.\""
    "    }"
    ""
    "    $baseName = [IO.Path]::GetFileNameWithoutExtension($WordPath)"
    "    $padWidth = [Math]::Max(3, $pageCount.ToString().Length)"
    ""
    "    for ($i = 1; $i -le $pageCount; $i++) {"
    "        $sel = $word.Selection"
    "        $sel.GoTo([ref]$wdGoToPage, [ref]$wdGoToAbsolute, [ref]$i, [ref][Type]::Missing) | Out-Null"
    "        $startPos = $sel.Start"
    ""
    "        if ($i -lt $pageCount) {"
    "            $next = $i + 1"
    "            $sel.GoTo([ref]$wdGoToPage, [ref]$wdGoToAbsolute, [ref]$next, [ref][Type]::Missing) | Out-Null"
    "            $endPos = [Math]::Max($sel.Start - 1, $startPos + 1)"
    "        } else {"
    "            $endPos = $doc.Content.End - 1"
    "        }"
    ""
    "        $range = $doc.Range($startPos, $endPos)"
    "        [System.Windows.Forms.Clipboard]::Clear()"
    "        $range.CopyAsPicture()"
    ""
    "        $img = $null"
    "        for ($try = 0; $try -lt 5 -and $img -eq $null; $try++) {"
    "            Start-Sleep -Milliseconds 120"
    "            if ([System.Windows.Forms.Clipboard]::ContainsImage()) {"
    "                $img = [System.Windows.Forms.Clipboard]::GetImage()"
    "            }"
    "        }"
    ""
    "        if ($img -eq $null) {"
    "            Write-Log \"page $i : clipboard empty, skipped\""
    "            continue"
    "        }"
    ""
    "        $scale = $Dpi / 96.0"
    "        $newW  = [int][Math]::Max(1, [Math]::Round($img.Width  * $scale))"
    "        $newH  = [int][Math]::Max(1, [Math]::Round($img.Height * $scale))"
    ""
    "        $bmp = New-Object System.Drawing.Bitmap $newW, $newH"
    "        $bmp.SetResolution($Dpi, $Dpi)"
    ""
    "        $g = [System.Drawing.Graphics]::FromImage($bmp)"
    "        try {"
    "            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic"
    "            $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality"
    "            $g.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality"
    "            $g.Clear([System.Drawing.Color]::White)"
    "            $g.DrawImage($img, 0, 0, $newW, $newH)"
    "        } finally {"
    "            $g.Dispose()"
    "        }"
    ""
    "        $idx     = $i.ToString().PadLeft($padWidth, '0')"
    "        $pngPath = Join-Path $OutputDir (\"{0}_page{1}.png\" -f $baseName, $idx)"
    "        $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)"
    ""
    "        $bmp.Dispose()"
    "        $img.Dispose()"
    "        Write-Log \"page $i -> $pngPath\""
    "    }"
    ""
    "    [System.Windows.Forms.Clipboard]::Clear()"
    "    Write-Log \"Done.\""
    "    exit 0"
    "}"
    "catch {"
    "    Write-Log (\"ERROR: \" + $_.Exception.Message)"
    "    Write-Log $_.ScriptStackTrace"
    "    exit 1"
    "}"
    "finally {"
    "    if ($doc  -ne $null) { try { $doc.Close([ref]$false) } catch {} }"
    "    if ($word -ne $null) { try { $word.Quit([ref]$false) } catch {} }"
    "    foreach ($o in @($doc, $word)) {"
    "        if ($o -ne $null) {"
    "            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o)"
    "        }"
    "    }"
    "    [GC]::Collect()"
    "    [GC]::WaitForPendingFinalizers()"
    "}"
  )
)

;; 把内嵌脚本写到 %TEMP%, 返回脚本路径
(defun z2p:write-ps (path / fh)
  (setq fh (open path "w"))
  (cond
    ((null fh)
     (alert (strcat "Z2P: 无法写入临时脚本 " path))
     nil)
    (t
     (foreach line (z2p:ps-lines)
       (write-line line fh))
     (close fh)
     path))
)

;; 同步调用 PowerShell, 返回退出码 (0=成功)
(defun z2p:run-ps (ps1 wordFile outDir dpi / shell cmd rc)
  (setq shell (vlax-create-object "WScript.Shell"))
  (setq cmd
    (strcat
      "powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass"
      " -File \""      ps1      "\""
      " -WordPath \""  wordFile "\""
      " -OutputDir \"" outDir   "\""
      " -Dpi "         (itoa dpi)
    )
  )
  (setq rc
    (vl-catch-all-apply
      '(lambda ()
         ;; style=0 隐藏窗口, bWaitOnReturn=:vlax-true 阻塞至退出
         (vlax-invoke shell 'Run cmd 0 :vlax-true))))
  (vlax-release-object shell)
  (if (vl-catch-all-error-p rc) -1 rc)
)

;;; --- 主命令: Z2P ------------------------------------------------------------

(defun c:Z2P ( / wordFile outDir dpi width ip ps1 rc pngs
                curTopY gap cnt png fullPng size pxW pxH h ok)

  (setq ok T)

  ;; 1. 选 Word 文件
  (setq wordFile (getfiled "选择要插入的 Word 文档" "" "doc;docx" 8))
  (if (null wordFile)
    (progn (princ "\n[Z2P] 已取消.") (setq ok nil))
    (setq wordFile (z2p:slash wordFile))
  )

  ;; 2. 输出目录 (默认: Word 同目录下 <基本名>_png)
  (if ok
    (progn
      (setq outDir
        (strcat (z2p:slash (vl-filename-directory wordFile))
                "/"
                (vl-filename-base wordFile) "_png"))
      (setq outDir (z2p:ensure-dir outDir))
      (princ (strcat "\n[Z2P] PNG 输出目录: " outDir))
    )
  )

  ;; 3. 取插入点 / 宽度 / DPI
  (if ok
    (progn
      (initget 1)
      (setq ip (getpoint "\n[Z2P] 请指定第一张图片的左上角插入点: "))

      (initget 6)
      (setq width (getdist ip "\n[Z2P] 请输入每张图片的宽度 <200>: "))
      (if (not width) (setq width 200.0))

      (initget 4)
      (setq dpi (getint "\n[Z2P] 请输入渲染 DPI [72-600] <200>: "))
      (if (or (null dpi) (< dpi 72)) (setq dpi 200))
      (if (> dpi 600) (setq dpi 600))
    )
  )

  ;; 4. 落盘内嵌 PS 脚本
  (if ok
    (progn
      (setq ps1 (strcat (z2p:temp-dir) "/Z2P_runtime.ps1"))
      (if (null (z2p:write-ps ps1))
        (setq ok nil)
        (princ (strcat "\n[Z2P] 临时脚本: " ps1)))
    )
  )

  ;; 5. 调 PowerShell 做 Word -> PNG
  (if ok
    (progn
      (princ "\n[Z2P] [1/2] 调用 Word 渲染 PNG, 请稍候...")
      (setq rc (z2p:run-ps ps1 wordFile outDir dpi))
      ;; 删掉临时脚本 (失败也尽量清掉)
      (vl-catch-all-apply '(lambda () (vl-file-delete ps1)))
      (if (/= rc 0)
        (progn
          (alert (strcat "Z2P: Word -> PNG 失败 (退出码 = " (itoa rc) ").\n"
                         "详见 " outDir "/Z2P.log"))
          (setq ok nil))
      )
    )
  )

  ;; 6. 收集生成的 PNG
  (if ok
    (progn
      (setq pngs (z2p:list-pngs outDir))
      (if (or (null pngs) (= (length pngs) 0))
        (progn (alert "Z2P: 未发现生成的 PNG.") (setq ok nil))
        (princ (strcat "\n[Z2P]       已生成 " (itoa (length pngs)) " 张 PNG."))
      )
    )
  )

  ;; 7. 顺序插入到模型空间
  (if ok
    (progn
      (princ "\n[Z2P] [2/2] 插入到模型空间...")
      (setq gap (* width 0.05))
      (setq curTopY (cadr ip))
      (setq cnt 0)
      (foreach png pngs
        (setq fullPng (strcat outDir "/" png))
        (setq size (z2p:png-size fullPng))
        (if size
          (progn
            (setq pxW (car size) pxH (cadr size))
            (setq h (* width (/ (float pxH) (float pxW)))))
          (setq h width))
        ;; -IMAGEATTACH: 路径 / 插入点(左下角) / 缩放(图纸单位的宽度) / 旋转
        (command "._-IMAGEATTACH"
                 fullPng
                 (list (car ip) (- curTopY h) (caddr ip))
                 width
                 0)
        (setq curTopY (- curTopY h gap))
        (setq cnt (1+ cnt))
        (princ (strcat "\n[Z2P]       已插入第 " (itoa cnt) " 页: " png))
      )
      (princ (strcat "\n[Z2P] 完成! 共插入 " (itoa cnt)
                     " 张, 输出目录: " outDir))
    )
  )

  (princ)
)

(princ "\n[Z2P] 已加载, 命令: Z2P")
(princ)
