using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

using PdfiumViewer;
using Word = Microsoft.Office.Interop.Word;

namespace ZwCadWordToPng
{
    /// <summary>
    /// 负责把 doc/docx 文档逐页渲染为 PNG 图片.
    /// 流程: Word.Interop -> PDF -> PDFium -> PNG
    /// </summary>
    internal static class WordConverter
    {
        /// <summary>
        /// 将指定 Word 文档转换为若干张 PNG (一页一张), 返回按页码排序的文件路径列表.
        /// </summary>
        /// <param name="wordPath">源 doc/docx 路径</param>
        /// <param name="outputDir">PNG 输出目录</param>
        /// <param name="dpi">渲染 DPI, 默认 200</param>
        public static List<string> ConvertToPng(string wordPath, string outputDir, int dpi = 200)
        {
            if (string.IsNullOrEmpty(wordPath))
                throw new ArgumentNullException("wordPath");
            if (!File.Exists(wordPath))
                throw new FileNotFoundException("Word 文档不存在: " + wordPath, wordPath);
            if (string.IsNullOrEmpty(outputDir))
                throw new ArgumentNullException("outputDir");
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);
            if (dpi < 72) dpi = 72;

            string baseName = SanitizeFileName(Path.GetFileNameWithoutExtension(wordPath));
            string pdfPath = Path.Combine(outputDir, baseName + ".__tmp__.pdf");

            try
            {
                WordToPdf(wordPath, pdfPath);
                return PdfToPng(pdfPath, outputDir, baseName, dpi);
            }
            finally
            {
                // 清理中间 PDF
                try { if (File.Exists(pdfPath)) File.Delete(pdfPath); } catch { /* ignore */ }
            }
        }

        #region Word -> PDF

        private static void WordToPdf(string wordPath, string pdfPath)
        {
            Word.Application app = null;
            Word.Document doc = null;
            try
            {
                app = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone };

                // 使用只读 + 不允许修订 方式打开, 避免提示
                object oWord = wordPath;
                object oFalse = false;
                object oTrue = true;
                object oMissing = Type.Missing;

                doc = app.Documents.Open(
                    ref oWord,
                    ref oMissing, // ConfirmConversions
                    ref oTrue,    // ReadOnly
                    ref oFalse,   // AddToRecentFiles
                    ref oMissing, // PasswordDocument
                    ref oMissing, // PasswordTemplate
                    ref oFalse,   // Revert
                    ref oMissing, // WritePasswordDocument
                    ref oMissing, // WritePasswordTemplate
                    ref oMissing, // Format
                    ref oMissing, // Encoding
                    ref oFalse,   // Visible
                    ref oMissing, // OpenAndRepair
                    ref oMissing, // DocumentDirection
                    ref oFalse,   // NoEncodingDialog
                    ref oMissing  // XMLTransform
                );

                doc.ExportAsFixedFormat(
                    OutputFileName: pdfPath,
                    ExportFormat: Word.WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false,
                    OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Range: Word.WdExportRange.wdExportAllDocument,
                    From: 0,
                    To: 0,
                    Item: Word.WdExportItem.wdExportDocumentContent,
                    IncludeDocProps: false,
                    KeepIRM: true,
                    CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                    DocStructureTags: true,
                    BitmapMissingFonts: true,
                    UseISO19005_1: false);
            }
            finally
            {
                try { if (doc != null) { doc.Close(SaveChanges: false); } } catch { /* ignore */ }
                try { if (app != null) { app.Quit(SaveChanges: false); } } catch { /* ignore */ }
                if (doc != null) Marshal.ReleaseComObject(doc);
                if (app != null) Marshal.ReleaseComObject(app);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion

        #region PDF -> PNG

        private static List<string> PdfToPng(string pdfPath, string outputDir, string baseName, int dpi)
        {
            var results = new List<string>();
            using (var pdfDoc = PdfDocument.Load(pdfPath))
            {
                int pageCount = pdfDoc.PageCount;
                int width = PageCountDigits(pageCount);
                string indexFormat = "D" + Math.Max(3, width);

                for (int i = 0; i < pageCount; i++)
                {
                    SizeF pageSizeF = pdfDoc.PageSizes[i]; // 单位是 point (1/72 inch)
                    int pixelWidth = Math.Max(1, (int)Math.Round(pageSizeF.Width / 72.0 * dpi));
                    int pixelHeight = Math.Max(1, (int)Math.Round(pageSizeF.Height / 72.0 * dpi));

                    using (var img = pdfDoc.Render(i, pixelWidth, pixelHeight, dpi, dpi, false))
                    {
                        string pngPath = Path.Combine(
                            outputDir,
                            string.Format("{0}_page{1}.png", baseName, (i + 1).ToString(indexFormat)));
                        img.Save(pngPath, ImageFormat.Png);
                        results.Add(pngPath);
                    }
                }
            }
            return results;
        }

        private static int PageCountDigits(int count)
        {
            int d = 1;
            while (count >= 10) { count /= 10; d++; }
            return d;
        }

        #endregion

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "document";
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }
}
