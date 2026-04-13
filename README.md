# ZwCadWordToPng —— 中望CAD Word 转 PNG 插件

一个面向 **中望CAD (ZWCAD)** 的 .NET 小插件,一键把 Word 文档（`.docx` / `.doc`）
**每一页**转换为 PNG 图片并**直接插入到当前 CAD 图纸**中,按页码顺序纵向排列。

命令名: `WORD2PNG`

## 功能

- 支持 `.doc` / `.docx`,任意页数
- 每页独立导出为 PNG,文件名形如 `MyDoc_page001.png`,`MyDoc_page002.png` ...
- 允许指定插入基点、图片宽度、渲染 DPI,高度按原比例自动计算
- 插入后的图片以 `RasterImage` 实体存在于模型空间,可在 CAD 中移动/缩放/裁剪
- 自动建立 `RasterImageDef` 反应器关联,避免图像失效

## 工作原理

```
 Word (.doc/.docx)
       │  Microsoft.Office.Interop.Word.ExportAsFixedFormat
       ▼
   临时 PDF
       │  PdfiumViewer.PdfDocument.Render (按 DPI 渲染)
       ▼
   每页一张 PNG
       │  ZWCAD RasterImage / RasterImageDef
       ▼
   当前图纸 (模型空间)
```

## 目录结构

```
ZwCadWordToPng.sln
src/
└── ZwCadWordToPng/
    ├── ZwCadWordToPng.csproj     MSBuild 项目
    ├── packages.config           NuGet 包清单 (PdfiumViewer / Word Interop)
    ├── Properties/AssemblyInfo.cs
    ├── WordToPngCommand.cs       WORD2PNG 命令入口
    ├── WordConverter.cs          Word -> PDF -> PNG
    ├── CadImageInserter.cs       PNG -> RasterImage 插入
    └── ZwCadWordToPng.ldr        LISP 启动加载脚本
```

## 编译要求

| 依赖 | 说明 |
|------|------|
| Visual Studio 2019/2022 | 或等价的 MSBuild |
| .NET Framework 4.8 | 与中望CAD 2022+ 保持一致 |
| 中望CAD | 安装目录下需包含 `ZwCADMgd.dll`、`ZwCADDbMgd.dll` |
| Microsoft Word | 需本机已安装 Word(2013+),用于 Interop 调用 |
| NuGet: `PdfiumViewer` | 渲染 PDF 到位图 |
| NuGet: `PdfiumViewer.Native.x86_64.v8-xfa` | PDFium 原生库(x64) |
| NuGet: `Microsoft.Office.Interop.Word` | Word COM 类型库 |

### 构建步骤

1. 打开 `ZwCadWordToPng.sln`,确保选择 **x64** 平台。
2. 首次构建前在环境变量中设置中望CAD的安装路径:
   ```
   setx ZwCadInstallDir "C:\Program Files\ZWSOFT\ZWCAD 2024"
   ```
   或直接修改 `ZwCadWordToPng.csproj` 内的 `<HintPath>`。
3. 右键项目 → **还原 NuGet 包**。
4. Build → 输出 `bin\x64\Release\ZwCadWordToPng.dll`。
5. 把 `ZwCadWordToPng.dll`、`PdfiumViewer.dll` 以及 PDFium 原生库
   (`pdfium.dll`,由 `PdfiumViewer.Native.*` 包解压得到) 拷贝到同一目录。

> 中望CAD 不同版本 .NET API 的命名空间偶有差异,若编译报错
> 找不到 `ZwSoft.ZwCAD.*`,请将 3 个源文件中的命名空间替换为你所用版本
> 实际提供的命名空间(如 `ZwCAD.ApplicationServices` 等)。

## 安装 / 加载

**临时加载**:在中望CAD 命令行执行:
```
NETLOAD
```
在弹出对话框中选择 `ZwCadWordToPng.dll`。

**永久加载**(推荐):
1. 把整个插件目录放到固定位置,例如 `D:\Plugins\ZwCadWordToPng\`。
2. 编辑 `ZwCadWordToPng.ldr`,把 `dllPath` 改为真实路径。
3. 在中望CAD 执行:
   ```
   (load "D:/Plugins/ZwCadWordToPng/ZwCadWordToPng.ldr")
   ```
   或把这行追加到 `acaddoc.lsp`,实现随 CAD 启动自动加载。

## 使用

1. 在中望CAD 打开一个图纸。
2. 命令行输入 `WORD2PNG`,按提示:
   - 选择要转换的 `.doc` / `.docx` 文件
   - 选择 PNG 保存目录(默认使用 Word 同目录下的 `xxx_png` 子目录)
   - 鼠标点选第一张图片的**左上角**插入点
   - 输入每张图片的宽度(当前图形单位,默认 200)
   - 输入渲染 DPI(72–600,默认 200)
3. 命令执行完毕后,会在命令行打印每页插入进度,所有 PNG 图片
   按页码顺序从上往下排列,页间自动留出 5% 宽度的间隙。

## 常见问题

**Q: 报错 "The message filter indicated that the application is busy"**
A: 关闭所有 Word 进程再重试,或在 Word 文档没有启用宏 / 受保护视图时再运行。

**Q: 插入图片显示为 "???" 或空白框**
A: 说明图像文件路径失效,请确认 PNG 输出目录保留未删除;该插件使用的是
引用式图像插入(RasterImage),不会把图像嵌入 DWG。

**Q: 页数非常多怎么办**
A: 可适当调低 DPI(例如 150),PNG 文件会更小,CAD 打开/再生速度更好。

## License

MIT
