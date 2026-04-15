;;; -----------------------------------------------------------------------------
;;; ZwCadWordToPng.lsp
;;; 中望CAD / AutoCAD 通用 Visual LISP 插件
;;; 功能: 把 Word 文档 (.doc/.docx) 每一页转换为 PNG, 再批量插入到当前图纸.
;;;
;;; 使用:
;;;   1. 把本文件与同目录下的 Word2Png.ps1 放在同一个文件夹内;
;;;   2. 在 CAD 命令行执行 (load "D:/path/to/ZwCadWordToPng.lsp");
;;;   3. 命令行输入 WORD2PNG, 按提示操作即可.
;;;
;;; 依赖:
;;;   - 本机安装 Microsoft Word (COM 自动化)
;;;   - Windows PowerShell 5.1 及以上 (随 Win10/Win11 自带)
;;; -----------------------------------------------------------------------------

(vl-load-com)

;;; --- 工具函数 ---------------------------------------------------------------

;; 把反斜杠替换为正斜杠, 便于在 CAD 命令行使用
(defun zwp:slash (s)
  (vl-string-translate "\\" "/" s)
)

;; 取本 .lsp 文件所在目录 (用来定位同目录下的 Word2Png.ps1)
(defun zwp:script-dir ( / here)
  (setq here (findfile "ZwCadWordToPng.lsp"))
  (if here
    (zwp:slash (vl-filename-directory here))
    (zwp:slash (getvar "DWGPREFIX"))
  )
)

;; 确保目录存在
(defun zwp:ensure-dir (dir / fso)
  (setq dir (zwp:slash dir))
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

;; 列出指定目录下匹配 pattern 的文件, 并按名称排序 (页码顺序)
(defun zwp:list-pngs (dir / files)
  (setq files (vl-directory-files dir "*.png" 1))
  (if files (vl-sort files '<) nil)
)

;; 把图片原始像素尺寸读出来, 用来按图形单位宽度算出对应高度
;; 返回 (list pixelW pixelH), 读不到时返回 nil
(defun zwp:png-size (path / img w h)
  (vl-catch-all-apply
    '(lambda ()
       (setq img (vlax-create-object "WIA.ImageFile"))
       (vlax-invoke img 'LoadFile path)
       (setq w (vlax-get img 'Width))
       (setq h (vlax-get img 'Height))
       (vlax-release-object img)))
  (if (and w h (> w 0) (> h 0)) (list w h) nil)
)

;; 同步调用 PowerShell 脚本, 返回 PS 进程退出码
(defun zwp:run-ps1 (ps1 wordFile outDir dpi / shell cmd rc)
  (setq shell (vlax-create-object "WScript.Shell"))
  (setq cmd
    (strcat
      "powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass"
      " -File \"" ps1 "\""
      " -WordPath \""   wordFile "\""
      " -OutputDir \""  outDir   "\""
      " -Dpi "          (itoa dpi)
    )
  )
  (setq rc
    (vl-catch-all-apply
      '(lambda ()
         ;; style=0 隐藏窗口, bWaitOnReturn=:vlax-true 阻塞直到返回
         (vlax-invoke shell 'Run cmd 0 :vlax-true))))
  (vlax-release-object shell)
  (if (vl-catch-all-error-p rc) -1 rc)
)

;;; --- 命令主体 ---------------------------------------------------------------

(defun c:WORD2PNG ( / wordFile outDir dpi width ip ps1 rc pngs
                     curTopY gap cnt png fullPng size pxW pxH h)
  ;; 1. 选择 Word 文件
  (setq wordFile (getfiled "选择要转换的 Word 文档" "" "doc;docx" 8))
  (if (not wordFile)
    (progn (princ "\n已取消.") (princ) (exit))
  )
  (setq wordFile (zwp:slash wordFile))

  ;; 2. 输出目录 (默认: Word 同目录下 <基本名>_png)
  (setq outDir
    (strcat (zwp:slash (vl-filename-directory wordFile))
            "/"
            (vl-filename-base wordFile) "_png"))
  (setq outDir (zwp:ensure-dir outDir))
  (princ (strcat "\nPNG 输出目录: " outDir))

  ;; 3. 插入点 (图片左上角)
  (initget 1)
  (setq ip (getpoint "\n请指定第一张图片的左上角插入点: "))

  ;; 4. 图片宽度
  (initget 6) ;; 不允许 0 和负数
  (setq width (getdist ip "\n请输入每张图片的宽度 <200>: "))
  (if (not width) (setq width 200.0))

  ;; 5. DPI
  (initget 4) ;; 不允许负数
  (setq dpi (getint "\n请输入渲染 DPI [72-600] <200>: "))
  (if (or (null dpi) (< dpi 72)) (setq dpi 200))
  (if (> dpi 600) (setq dpi 600))

  ;; 6. 定位 PowerShell 脚本
  (setq ps1 (strcat (zwp:script-dir) "/Word2Png.ps1"))
  (if (not (findfile ps1))
    (progn
      (alert (strcat "找不到配套脚本:\n" ps1
                     "\n请确认 Word2Png.ps1 与 ZwCadWordToPng.lsp 放在同一目录."))
      (exit)
    )
  )

  ;; 7. 调用 PowerShell 做 Word -> PNG
  (princ "\n[1/2] 正在调用 Word 逐页渲染 PNG, 请稍候...")
  (setq rc (zwp:run-ps1 ps1 wordFile outDir dpi))
  (if (/= rc 0)
    (progn
      (alert (strcat "Word 转 PNG 失败, PowerShell 退出码 = " (itoa rc)
                     "\n请打开 " outDir "/Word2Png.log 查看详细信息."))
      (exit)
    )
  )

  ;; 8. 收集生成的 PNG 并插入
  (setq pngs (zwp:list-pngs outDir))
  (if (or (null pngs) (= (length pngs) 0))
    (progn (alert "未发现生成的 PNG, 操作中止.") (exit))
  )
  (princ (strcat "\n      已生成 " (itoa (length pngs)) " 张 PNG."))

  (princ "\n[2/2] 正在把 PNG 插入到模型空间...")
  (setq gap (* width 0.05))
  (setq curTopY (cadr ip))
  (setq cnt 0)
  (foreach png pngs
    (setq fullPng (strcat outDir "/" png))
    (setq size (zwp:png-size fullPng))
    (if size
      (progn
        (setq pxW (car size) pxH (cadr size))
        (setq h (* width (/ (float pxH) (float pxW))))
      )
      ;; 读不到尺寸时退化: 假设宽高相同
      (setq h width)
    )
    ;; -IMAGEATTACH: 文件路径 / 插入点 / 缩放 (相对 1 单位 = 1 像素) / 旋转
    ;; 中望CAD 与 AutoCAD 都支持 _-IMAGEATTACH, 缩放为 width (相当于一个像素 = width/pxW)
    ;; 插入点视为左下角, 所以先把 y 下移 h
    (command "._-IMAGEATTACH"
             fullPng
             (list (car ip) (- curTopY h) (caddr ip))
             width
             0)
    (setq curTopY (- curTopY h gap))
    (setq cnt (1+ cnt))
    (princ (strcat "\n      已插入第 " (itoa cnt) " 页: " png))
  )

  (princ (strcat "\n完成! 共插入 " (itoa cnt)
                 " 张图片. 输出目录: " outDir))
  (princ)
)

(princ "\n[ZwCadWordToPng] 已加载, 输入 WORD2PNG 开始使用.")
(princ)
