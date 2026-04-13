# WordToCadPlugin

一个 AutoCAD .NET 插件：选择一个 Word 文档，插件自动把每一页渲染成 PNG 图片，并按网格排列嵌入到当前 DWG 中。DWG 不再依赖外部 PNG 文件。

## 功能

- 一键把 `.docx` / `.doc` 的所有页作为图片批量插入 CAD。
- 通过 PDFium 高保真渲染，按原排版呈现文字、表格、图形。
- 按"每行 N 张"自动网格排列，用户仅需指定起点。
- 使用 AutoCAD 2022+ 内置的 `EMBEDIMAGE` 把像素数据嵌入 DWG，发送/归档时不会丢失图片。

## 系统要求

| 组件 | 要求 |
| --- | --- |
| AutoCAD | 2020 – 2024 (x64) |
| .NET | .NET Framework 4.8 |
| Microsoft Word | 2013 或更高（用于 Word → PDF 的 COM 互操作） |
| PDF 渲染 | Docnet.Core NuGet（内含 PDFium 原生库）|

> `EMBEDIMAGE` 命令需要 AutoCAD 2022+。如果你使用 2020/2021，图片会以"外部引用"方式插入，需保留 `%TEMP%\W2CAD_*` 文件夹或改用 v2 的 OLE 嵌入方案。

## 构建

```bat
nuget restore WordToCadPlugin.sln
msbuild WordToCadPlugin.sln /p:Configuration=Release /p:Platform=x64
```

生成物位于 `WordToCadPlugin\bin\x64\Release\WordToCadPlugin.dll`。

如果 `.csproj` 中的 AutoCAD `HintPath` 与本机安装路径不一致，请修改为实际路径，或在系统环境变量里设 `ACAD_INSTALL`。

## 使用

1. 启动 AutoCAD，打开或新建一张图纸。
2. 命令行执行 `NETLOAD`，选择刚生成的 `WordToCadPlugin.dll`。
3. 输入命令 `W2CAD`：
   - 在弹出的对话框中选择 Word 文档。
   - 在图纸上点击网格左上角作为插入起点。
   - 输入每行张数（默认 3）。
4. 等待（进度打印在命令行），完成后所有页面已作为图片嵌入到当前 DWG。

### 命令列表

| 命令 | 说明 |
| --- | --- |
| `W2CAD` | 选择 Word → 渲染 → 网格插入 → 嵌入 |
| `W2CADOPT` | 修改 DPI、单张宽度、行列间距，设置保存到 `%APPDATA%\W2CAD\config.json` |

### 默认选项

| 选项 | 默认值 | 说明 |
| --- | --- | --- |
| `Dpi` | 200 | PDF → PNG 渲染分辨率 |
| `CellWidth` | 150 | 单张图片宽度（图纸单位） |
| `GapX` | 20 | 水平间距 |
| `GapY` | 20 | 垂直间距 |

## 架构

```
Commands.cs            入口命令 W2CAD / W2CADOPT
 └── Services/
     ├── WordToPdfConverter.cs   Word COM → PDF
     ├── PdfToPngRenderer.cs     PDFium → PNG
     └── CadImageInserter.cs     RasterImage 网格插入 + EMBEDIMAGE
 └── Models/
     └── PluginOptions.cs        配置持久化
```

## 已知限制

- Word 必须安装在本机（用 COM 自动化导出 PDF）。若需无 Office 方案，可改接 LibreOffice headless 或 Aspose.Words。
- 大文档（几十页）首次渲染可能较慢；DPI 越高越耗时。
- `EMBEDIMAGE` 为 AutoCAD 2022+ 命令；低版本需保留临时 PNG 或使用 OLE 嵌入分支。

## 验证

详细的端到端测试清单见 `plans/` 下的计划文件：

1. 生成 `.dll`，`NETLOAD` 加载无错误。
2. 对 6 页 `.docx` 运行 `W2CAD` 并选择每行 3 张，应得到 2 行 × 3 列网格。
3. 关闭 AutoCAD、删除 `%TEMP%\W2CAD_*`、重新打开 DWG，图片仍正常显示 → 嵌入成功。
4. 运行后检查任务管理器 `WINWORD.EXE` 应已退出（无残留）。
