using System.Collections.Generic;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using WordToCadPlugin.Models;

namespace WordToCadPlugin.Services
{
    /// <summary>
    /// Inserts a list of PNG files into the current drawing arranged in an N-column grid,
    /// then embeds the pixel data into the DWG via the EMBEDIMAGE command so the drawing
    /// no longer depends on the external PNG files.
    /// </summary>
    internal static class CadImageInserter
    {
        /// <summary>
        /// Build one <see cref="RasterImage"/> per PNG in a grid starting at
        /// <paramref name="origin"/> (treated as the upper-left corner of page 1).
        /// </summary>
        public static void InsertGrid(
            Document doc,
            List<string> pngs,
            Point3d origin,
            int cols,
            PluginOptions opts)
        {
            if (pngs == null || pngs.Count == 0) return;
            if (cols < 1) cols = 1;

            Database db = doc.Database;
            double cellW = opts.CellWidth;
            double gapX = opts.GapX;
            double gapY = opts.GapY;

            var addedDefNames = new List<string>();

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // The image dictionary holds all RasterImageDef entries; SetAt creates or replaces.
                ObjectId dictId = RasterImageDef.GetImageDictionary(db);
                if (dictId == ObjectId.Null)
                    dictId = RasterImageDef.CreateImageDictionary(db);

                var imgDict = (DBDictionary)tr.GetObject(dictId, OpenMode.ForWrite);
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(
                    bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Track the running bottom of each column so variable-height pages still line up.
                // Simpler first-pass: treat all pages as same aspect (PDF pages usually are).
                double firstCellH = 0;

                for (int i = 0; i < pngs.Count; i++)
                {
                    int row = i / cols;
                    int col = i % cols;

                    var def = new RasterImageDef
                    {
                        SourceFileName = pngs[i],
                        ResolutionUnits = UnitsValue.Millimeters
                    };
                    def.Load();

                    // Dictionary names must be unique per DWG; use filename + short hash fallback.
                    string defName = MakeUniqueDefName(imgDict, Path.GetFileNameWithoutExtension(pngs[i]));
                    ObjectId defId = imgDict.SetAt(defName, def);
                    tr.AddNewlyCreatedDBObject(def, true);
                    addedDefNames.Add(defName);

                    // Keep aspect ratio: RasterImageDef.Size is in pixels.
                    Vector2d pxSize = def.Size;
                    double aspect = pxSize.Y / pxSize.X; // height / width
                    double cellH = cellW * aspect;
                    if (i == 0) firstCellH = cellH;

                    double x = origin.X + col * (cellW + gapX);
                    double y = origin.Y - row * (firstCellH + gapY); // grid top-row y; image bottom-left used below

                    var img = new RasterImage
                    {
                        ImageDefId = defId,
                        // Orientation: (bottom-left origin, +U width vector, +V height vector)
                        Orientation = new CoordinateSystem3d(
                            new Point3d(x, y - cellH, 0),
                            new Vector3d(cellW, 0, 0),
                            new Vector3d(0, cellH, 0))
                    };

                    ms.AppendEntity(img);
                    tr.AddNewlyCreatedDBObject(img, true);

                    // Enable reactors so the image survives def edits and supports EMBEDIMAGE.
                    RasterImage.EnableReactors(true);
                    img.AssociateRasterDef(def);
                }

                tr.Commit();
            }

            // EMBEDIMAGE (AutoCAD 2022+) converts a referenced raster image into embedded bytes.
            // We invoke it once per def name. `_.` prefix = international form; the trailing
            // newline acts as the ENTER that confirms the image-name prompt.
            foreach (string name in addedDefNames)
            {
                doc.SendStringToExecute($"_.EMBEDIMAGE\n{name}\n", true, false, false);
            }
        }

        /// <summary>
        /// Guarantees a unique key in the image dictionary by appending _2, _3, ... if needed.
        /// </summary>
        private static string MakeUniqueDefName(DBDictionary dict, string baseName)
        {
            string safe = baseName.Replace(' ', '_');
            if (!dict.Contains(safe)) return safe;
            for (int n = 2; n < 10000; n++)
            {
                string candidate = safe + "_" + n;
                if (!dict.Contains(candidate)) return candidate;
            }
            return safe + "_" + System.Guid.NewGuid().ToString("N").Substring(0, 6);
        }
    }
}
