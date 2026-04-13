using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using Docnet.Core;
using Docnet.Core.Models;

namespace WordToCadPlugin.Services
{
    /// <summary>
    /// Renders each page of a PDF to a PNG file using PDFium (via Docnet.Core).
    /// No Ghostscript dependency; no external process spawned.
    /// </summary>
    internal static class PdfToPngRenderer
    {
        /// <summary>
        /// Render <paramref name="pdfPath"/> into <paramref name="outDir"/> at the given
        /// <paramref name="dpi"/>, producing <c>page_001.png</c>, <c>page_002.png</c>, ...
        /// Returns the list of PNG paths in page order.
        /// </summary>
        public static List<string> Render(string pdfPath, string outDir, int dpi)
        {
            var results = new List<string>();
            if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);

            // Docnet expects a PageDimensions struct describing the target output resolution.
            // We render at 1:1 of the requested DPI by asking for a scaled page (dpi / 72).
            double scale = dpi / 72.0;
            var dims = new PageDimensions(scale);

            using (var library = DocLib.Instance)
            using (var reader = library.GetDocReader(pdfPath, dims))
            {
                int total = reader.GetPageCount();
                for (int i = 0; i < total; i++)
                {
                    using (var page = reader.GetPageReader(i))
                    {
                        int w = page.GetPageWidth();
                        int h = page.GetPageHeight();
                        byte[] raw = page.GetImage(); // BGRA, row-major, 4 bytes per pixel

                        using (var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                        {
                            BitmapData data = bmp.LockBits(
                                new Rectangle(0, 0, w, h),
                                ImageLockMode.WriteOnly,
                                bmp.PixelFormat);
                            try
                            {
                                Marshal.Copy(raw, 0, data.Scan0, raw.Length);
                            }
                            finally
                            {
                                bmp.UnlockBits(data);
                            }

                            string png = Path.Combine(outDir, $"page_{i + 1:D3}.png");
                            bmp.Save(png, ImageFormat.Png);
                            results.Add(png);
                        }
                    }
                }
            }
            return results;
        }
    }
}
