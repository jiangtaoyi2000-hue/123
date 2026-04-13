using System.Drawing;
using System.IO;

using ZwSoft.ZwCAD.ApplicationServices;
using ZwSoft.ZwCAD.DatabaseServices;
using ZwSoft.ZwCAD.Geometry;

namespace ZwCadWordToPng
{
    /// <summary>
    /// 将 PNG 图片作为 RasterImage 实体插入到当前图纸模型空间.
    /// </summary>
    internal static class CadImageInserter
    {
        /// <summary>
        /// 按指定宽度(保持纵横比)把 PNG 插入到 <paramref name="topLeft"/> 所在位置,
        /// <paramref name="topLeft"/> 被视为图片的 "左上角".
        /// </summary>
        /// <returns>插入后图片在图纸中实际占用的高度(用于让调用者进行垂直排版)</returns>
        public static double InsertImage(Document doc, string pngPath, Point3d topLeft, double targetWidth)
        {
            if (!File.Exists(pngPath))
                throw new FileNotFoundException("图片文件不存在: " + pngPath, pngPath);
            if (targetWidth <= 0)
                throw new System.ArgumentException("targetWidth 必须为正数", "targetWidth");

            // 读取原始像素尺寸, 得到纵横比
            int pixelWidth, pixelHeight;
            using (var bmp = new Bitmap(pngPath))
            {
                pixelWidth = bmp.Width;
                pixelHeight = bmp.Height;
            }
            double aspect = (double)pixelHeight / pixelWidth;
            double targetHeight = targetWidth * aspect;

            Database db = doc.Database;

            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // 1. 确保图像字典存在
                ObjectId dictId = RasterImageDef.GetImageDictionary(db);
                if (dictId.IsNull)
                {
                    dictId = RasterImageDef.CreateImageDictionary(db);
                }
                DBDictionary imgDict = (DBDictionary)tr.GetObject(dictId, OpenMode.ForWrite);

                // 2. 构造唯一 key
                string baseKey = Path.GetFileNameWithoutExtension(pngPath);
                string key = MakeUniqueKey(imgDict, baseKey);

                // 3. 创建 RasterImageDef 指向磁盘上的 PNG
                RasterImageDef imgDef = new RasterImageDef
                {
                    SourceFileName = pngPath
                };
                imgDef.Load();
                ObjectId imgDefId = imgDict.SetAt(key, imgDef);
                tr.AddNewlyCreatedDBObject(imgDef, true);

                // 4. 创建 RasterImage 实体, 并计算朝向: 以 topLeft 为左上角
                RasterImage ri = new RasterImage { ImageDefId = imgDefId };

                // 注意: RasterImage.Orientation 的 origin 是图片 "左下角".
                Point3d origin = new Point3d(topLeft.X, topLeft.Y - targetHeight, topLeft.Z);
                Vector3d uVec = new Vector3d(targetWidth, 0, 0);
                Vector3d vVec = new Vector3d(0, targetHeight, 0);
                ri.Orientation = new CoordinateSystem3d(origin, uVec, vVec);
                ri.ShowImage = true;

                // 5. 添加到模型空间
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                ms.AppendEntity(ri);
                tr.AddNewlyCreatedDBObject(ri, true);

                // 6. 建立反应器关联, 否则 ImageDef 被清理后图像会失效
                RasterImage.EnableReactors(true);
                ri.AssociateRasterDef(imgDef);

                tr.Commit();
            }

            return targetHeight;
        }

        private static string MakeUniqueKey(DBDictionary dict, string baseKey)
        {
            if (string.IsNullOrEmpty(baseKey)) baseKey = "image";
            // DBDictionary 名称不允许某些字符, 做一次过滤
            baseKey = baseKey.Replace(' ', '_');
            foreach (char bad in new[] { '<', '>', '/', '\\', '"', ':', ';', '?', '*', '|', ',', '=', '`' })
                baseKey = baseKey.Replace(bad.ToString(), "_");

            if (!dict.Contains(baseKey)) return baseKey;

            int i = 1;
            while (true)
            {
                string candidate = baseKey + "_" + i;
                if (!dict.Contains(candidate)) return candidate;
                i++;
            }
        }
    }
}
