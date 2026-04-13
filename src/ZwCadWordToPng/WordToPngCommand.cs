using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using ZwSoft.ZwCAD.ApplicationServices;
using ZwSoft.ZwCAD.EditorInput;
using ZwSoft.ZwCAD.Geometry;
using ZwSoft.ZwCAD.Runtime;

namespace ZwCadWordToPng
{
    /// <summary>
    /// 中望CAD 命令入口. 在命令行执行 WORD2PNG 即可调用.
    /// </summary>
    public class WordToPngCommand
    {
        /// <summary>
        /// 将选定的 Word 文档逐页转换为 PNG, 然后批量插入到当前图纸.
        /// </summary>
        [CommandMethod("WORD2PNG")]
        public void Word2Png()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("请先打开一个图纸.", "Word2Png");
                return;
            }
            Editor ed = doc.Editor;

            // 1. 选择 Word 文件
            string wordPath = PickWordFile();
            if (string.IsNullOrEmpty(wordPath))
            {
                ed.WriteMessage("\n已取消.");
                return;
            }

            // 2. 选择 PNG 输出目录 (默认为 Word 文件所在目录下的同名子目录)
            string defaultDir = Path.Combine(
                Path.GetDirectoryName(wordPath) ?? Path.GetTempPath(),
                Path.GetFileNameWithoutExtension(wordPath) + "_png");
            string outputDir = PickOutputDir(defaultDir);
            if (string.IsNullOrEmpty(outputDir))
            {
                ed.WriteMessage("\n已取消.");
                return;
            }

            // 3. 指定插入基点 (图纸的左上角)
            PromptPointOptions ppo = new PromptPointOptions("\n请指定第一张图片的左上角插入点: ")
            {
                AllowNone = false
            };
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return;
            Point3d insertPoint = ppr.Value;

            // 4. 指定图片宽度 (以图纸单位计)
            PromptDoubleOptions pdo = new PromptDoubleOptions("\n请输入每张图片的宽度 (单位: 当前图形单位)")
            {
                AllowNegative = false,
                AllowZero = false,
                DefaultValue = 200.0,
                UseDefaultValue = true
            };
            PromptDoubleResult pdr = ed.GetDouble(pdo);
            if (pdr.Status != PromptStatus.OK) return;
            double imageWidth = pdr.Value;

            // 5. 指定渲染 DPI (影响 PNG 清晰度和文件大小)
            PromptIntegerOptions pio = new PromptIntegerOptions("\n请输入渲染 DPI (数值越大图片越清晰)")
            {
                AllowNegative = false,
                AllowZero = false,
                LowerLimit = 72,
                UpperLimit = 600,
                DefaultValue = 200,
                UseDefaultValue = true
            };
            PromptIntegerResult pir = ed.GetInteger(pio);
            if (pir.Status != PromptStatus.OK) return;
            int dpi = pir.Value;

            // 6. 执行转换 + 插入
            List<string> pngFiles = null;
            try
            {
                ed.WriteMessage("\n[1/2] 正在将 Word 文档逐页转换为 PNG, 请稍候...");
                pngFiles = WordConverter.ConvertToPng(wordPath, outputDir, dpi);
                ed.WriteMessage(string.Format("\n      成功生成 {0} 张 PNG 图片.", pngFiles.Count));
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n[错误] Word 转 PNG 失败: " + ex.Message);
                return;
            }

            if (pngFiles == null || pngFiles.Count == 0)
            {
                ed.WriteMessage("\n没有生成任何图片, 操作结束.");
                return;
            }

            try
            {
                ed.WriteMessage("\n[2/2] 正在将 PNG 图片插入到图纸...");
                double gap = imageWidth * 0.05; // 页与页之间留一点空隙
                double currentTopY = insertPoint.Y;
                int pageIndex = 1;
                foreach (string png in pngFiles)
                {
                    double insertedHeight = CadImageInserter.InsertImage(
                        doc,
                        png,
                        new Point3d(insertPoint.X, currentTopY, insertPoint.Z),
                        imageWidth);
                    ed.WriteMessage(string.Format("\n      已插入第 {0} 页: {1}",
                        pageIndex, Path.GetFileName(png)));
                    currentTopY -= (insertedHeight + gap);
                    pageIndex++;
                }

                ed.WriteMessage(string.Format(
                    "\n完成! 共插入 {0} 张图片, 输出目录: {1}", pngFiles.Count, outputDir));
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n[错误] 插入图片失败: " + ex.Message);
            }
        }

        private static string PickWordFile()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "选择要转换的 Word 文档";
                ofd.Filter = "Word 文档 (*.docx;*.doc)|*.docx;*.doc|所有文件 (*.*)|*.*";
                ofd.CheckFileExists = true;
                ofd.Multiselect = false;
                return ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : null;
            }
        }

        private static string PickOutputDir(string defaultDir)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "选择 PNG 图片的保存目录 (建议使用空目录)";
                if (!string.IsNullOrEmpty(defaultDir))
                {
                    try
                    {
                        if (!Directory.Exists(defaultDir)) Directory.CreateDirectory(defaultDir);
                        fbd.SelectedPath = defaultDir;
                    }
                    catch { /* ignore */ }
                }
                return fbd.ShowDialog() == DialogResult.OK ? fbd.SelectedPath : null;
            }
        }
    }
}
