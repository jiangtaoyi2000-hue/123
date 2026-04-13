using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using WordToCadPlugin.Models;
using WordToCadPlugin.Services;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace WordToCadPlugin
{
    /// <summary>
    /// Entry-point command class. Registered via [assembly: CommandClass] in AssemblyInfo.cs.
    /// After NETLOAD, the user can invoke `W2CAD` or `W2CADOPT` at the AutoCAD command line.
    /// </summary>
    public class Commands
    {
        /// <summary>
        /// Main command: pick a .docx/.doc, render each page to PNG, then embed as a grid
        /// of raster images starting at a user-picked point.
        /// </summary>
        [CommandMethod("W2CAD", CommandFlags.Modal)]
        public void WordToCad()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;

            // 1. Prompt user for the Word file.
            string wordPath;
            using (var ofd = new OpenFileDialog
            {
                Title = "选择要插入的 Word 文档",
                Filter = "Word 文档 (*.docx;*.doc)|*.docx;*.doc|所有文件 (*.*)|*.*",
                CheckFileExists = true
            })
            {
                if (ofd.ShowDialog() != DialogResult.OK) return;
                wordPath = ofd.FileName;
            }

            // 2. Prompt for the grid origin (upper-left corner of page 1).
            var ptRes = ed.GetPoint("\n指定网格起点（首页左上角）: ");
            if (ptRes.Status != PromptStatus.OK) return;

            // 3. Prompt for the number of columns per row.
            var intOpts = new PromptIntegerOptions("\n每行张数")
            {
                DefaultValue = 3,
                UseDefaultValue = true,
                AllowNegative = false,
                AllowZero = false,
                LowerLimit = 1,
                UpperLimit = 20
            };
            var colRes = ed.GetInteger(intOpts);
            if (colRes.Status != PromptStatus.OK) return;
            int cols = colRes.Value;

            // 4. Run the pipeline inside an isolated temp folder; cleanup on exit.
            PluginOptions opts = PluginOptions.LoadOrDefault();
            string tempDir = Path.Combine(
                Path.GetTempPath(),
                "W2CAD_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);

            try
            {
                ed.WriteMessage("\n[W2CAD] 正在启动 Word 导出为 PDF...");
                string pdfPath = WordToPdfConverter.Convert(wordPath, tempDir);

                ed.WriteMessage("\n[W2CAD] 正在将 PDF 逐页渲染为 PNG (DPI={0})...", opts.Dpi);
                List<string> pngs = PdfToPngRenderer.Render(pdfPath, tempDir, opts.Dpi);
                ed.WriteMessage("\n[W2CAD] 共 {0} 页。正在插入并嵌入到 DWG...", pngs.Count);

                CadImageInserter.InsertGrid(doc, pngs, ptRes.Value, cols, opts);

                ed.WriteMessage("\n[W2CAD] 完成。");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n[W2CAD] 错误: {0}", ex.Message);
            }
            finally
            {
                // After EMBEDIMAGE, DWG holds the pixel data itself; safe to drop temp files.
                try { Directory.Delete(tempDir, true); } catch { /* best effort */ }
            }
        }

        /// <summary>
        /// Open a small WinForms dialog to edit persistent options (DPI, cell width, gaps).
        /// Settings are stored at %APPDATA%\W2CAD\config.json.
        /// </summary>
        [CommandMethod("W2CADOPT", CommandFlags.Modal)]
        public void SetOptions()
        {
            Document doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            PluginOptions opts = PluginOptions.LoadOrDefault();
            using (var dlg = new OptionsDialog(opts))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    opts.Save();
                    doc.Editor.WriteMessage(
                        "\n[W2CADOPT] 已保存: DPI={0}, CellWidth={1}, GapX={2}, GapY={3}",
                        opts.Dpi, opts.CellWidth, opts.GapX, opts.GapY);
                }
            }
        }
    }

    /// <summary>
    /// Minimal WinForms options dialog. Kept in the same file to avoid a designer round-trip.
    /// Four numeric inputs bound to the <see cref="PluginOptions"/> instance passed in.
    /// </summary>
    internal sealed class OptionsDialog : Form
    {
        private readonly PluginOptions _opts;
        private readonly NumericUpDown _dpi;
        private readonly NumericUpDown _cellW;
        private readonly NumericUpDown _gapX;
        private readonly NumericUpDown _gapY;

        public OptionsDialog(PluginOptions opts)
        {
            _opts = opts;
            Text = "W2CAD 选项";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ClientSize = new System.Drawing.Size(300, 200);

            AddLabel("渲染 DPI:", 20);
            _dpi = AddNumeric(opts.Dpi, 72, 600, 0, 20);

            AddLabel("单张宽度 (图纸单位):", 60);
            _cellW = AddNumeric((decimal)opts.CellWidth, 10, 10000, 2, 60);

            AddLabel("水平间距:", 100);
            _gapX = AddNumeric((decimal)opts.GapX, 0, 10000, 2, 100);

            AddLabel("垂直间距:", 140);
            _gapY = AddNumeric((decimal)opts.GapY, 0, 10000, 2, 140);

            var ok = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(130, 165),
                Size = new System.Drawing.Size(70, 26)
            };
            ok.Click += (s, e) =>
            {
                _opts.Dpi = (int)_dpi.Value;
                _opts.CellWidth = (double)_cellW.Value;
                _opts.GapX = (double)_gapX.Value;
                _opts.GapY = (double)_gapY.Value;
            };
            var cancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(210, 165),
                Size = new System.Drawing.Size(70, 26)
            };
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        private void AddLabel(string text, int top)
        {
            Controls.Add(new Label
            {
                Text = text,
                Location = new System.Drawing.Point(20, top),
                Size = new System.Drawing.Size(160, 20)
            });
        }

        private NumericUpDown AddNumeric(decimal value, decimal min, decimal max, int decimals, int top)
        {
            var nud = new NumericUpDown
            {
                Minimum = min,
                Maximum = max,
                Value = value,
                DecimalPlaces = decimals,
                Location = new System.Drawing.Point(180, top - 2),
                Size = new System.Drawing.Size(100, 22)
            };
            Controls.Add(nud);
            return nud;
        }
    }
}
