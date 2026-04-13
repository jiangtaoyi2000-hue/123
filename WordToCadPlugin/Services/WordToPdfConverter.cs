using System;
using System.IO;
using System.Runtime.InteropServices;

namespace WordToCadPlugin.Services
{
    /// <summary>
    /// Converts a Word document (.doc/.docx) into a PDF using Microsoft Word COM automation.
    /// Late-bound via <c>Type.GetTypeFromProgID</c> so the plugin compiles without referencing
    /// a specific Office PIA and works across Office 2013–2024.
    /// </summary>
    internal static class WordToPdfConverter
    {
        // Microsoft Word constants (from the WdSaveFormat enum)
        private const int WdFormatPDF = 17;

        /// <summary>
        /// Open <paramref name="wordPath"/> in hidden Word, export to PDF in <paramref name="outDir"/>,
        /// and return the resulting PDF path. Throws if Word is not installed.
        /// </summary>
        public static string Convert(string wordPath, string outDir)
        {
            if (string.IsNullOrEmpty(wordPath)) throw new ArgumentNullException(nameof(wordPath));
            if (!File.Exists(wordPath)) throw new FileNotFoundException("Word 文件不存在", wordPath);
            if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);

            Type wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
                throw new InvalidOperationException("未检测到 Microsoft Word。请先安装 Office。");

            dynamic word = Activator.CreateInstance(wordType);
            dynamic wdoc = null;
            try
            {
                word.Visible = false;
                word.DisplayAlerts = 0; // wdAlertsNone
                word.ScreenUpdating = false;

                // ReadOnly=true: avoid locking the user's file; ConfirmConversions=false: suppress prompts.
                wdoc = word.Documents.Open(
                    FileName: wordPath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false);

                string pdfPath = Path.Combine(
                    outDir,
                    Path.GetFileNameWithoutExtension(wordPath) + ".pdf");

                // SaveAs2 is preferred over SaveAs for Word 2010+; both accept WdFormatPDF.
                wdoc.SaveAs2(pdfPath, WdFormatPDF);
                wdoc.Close(SaveChanges: false);
                wdoc = null;

                return pdfPath;
            }
            finally
            {
                // Best-effort cleanup: ensure the hidden WINWORD.EXE process exits.
                try { if (wdoc != null) wdoc.Close(SaveChanges: false); } catch { }
                try { word.Quit(SaveChanges: false); } catch { }
                try { Marshal.FinalReleaseComObject(word); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
