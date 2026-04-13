using System;
using System.IO;
using System.Web.Script.Serialization;

namespace WordToCadPlugin.Models
{
    /// <summary>
    /// Persistent user preferences for the W2CAD plugin.
    /// Stored as JSON at %APPDATA%\W2CAD\config.json.
    /// </summary>
    public class PluginOptions
    {
        /// <summary>DPI used when rendering PDF pages to PNG. Higher = sharper but larger.</summary>
        public int Dpi { get; set; } = 200;

        /// <summary>Width of each image in drawing units (mm if the drawing is metric).</summary>
        public double CellWidth { get; set; } = 150;

        /// <summary>Horizontal gap between columns, in drawing units.</summary>
        public double GapX { get; set; } = 20;

        /// <summary>Vertical gap between rows, in drawing units.</summary>
        public double GapY { get; set; } = 20;

        private static string ConfigPath
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "W2CAD");
                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "config.json");
            }
        }

        /// <summary>Load from disk, or return defaults if the file is missing/corrupt.</summary>
        public static PluginOptions LoadOrDefault()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    var serializer = new JavaScriptSerializer();
                    var loaded = serializer.Deserialize<PluginOptions>(json);
                    if (loaded != null) return loaded;
                }
            }
            catch
            {
                // Fall through to defaults if the config is unreadable.
            }
            return new PluginOptions();
        }

        /// <summary>Write current values to disk, creating the folder if needed.</summary>
        public void Save()
        {
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(this);
            File.WriteAllText(ConfigPath, json);
        }
    }
}
