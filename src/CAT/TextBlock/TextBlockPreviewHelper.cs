using System;
using System.IO;
using System.Windows.Media.Imaging;

namespace Cat.TextBlocks
{
    internal static class TextBlockPreviewHelper
    {
        public static BitmapImage TryLoadPreviewImage(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return null;
            byte[] bytes = LoadBytes(fileName.Trim());
            if (bytes == null || bytes.Length == 0) return null;
            try
            {
                var bmp = new BitmapImage();
                using (var ms = new MemoryStream(bytes))
                {
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.StreamSource = ms;
                    bmp.EndInit();
                }
                bmp.Freeze();
                return bmp;
            }
            catch { return null; }
        }

        private static byte[] LoadBytes(string fileName)
        {
            var asm = typeof(TextBlockPreviewHelper).Assembly;
            try
            {
                foreach (string name in asm.GetManifestResourceNames())
                {
                    if (!name.EndsWith(fileName, StringComparison.OrdinalIgnoreCase)) continue;
                    using (var stream = asm.GetManifestResourceStream(name))
                    {
                        if (stream == null) continue;
                        using (var ms = new MemoryStream())
                        {
                            stream.CopyTo(ms);
                            return ms.ToArray();
                        }
                    }
                }
            }
            catch { }

            string path = ResolveFilePath(fileName);
            if (path == null) return null;
            try { return File.ReadAllBytes(path); }
            catch { return null; }
        }

        private static string ResolveFilePath(string fileName)
        {
            try
            {
                string dir = Path.GetDirectoryName(typeof(TextBlockPreviewHelper).Assembly.Location) ?? "";
                string nextToDll = Path.Combine(dir, "text-block", fileName);
                if (File.Exists(nextToDll)) return nextToDll;

                var walk = new DirectoryInfo(dir);
                for (int i = 0; i < 8 && walk != null; i++, walk = walk.Parent)
                {
                    string assets = Path.Combine(walk.FullName, "assets", "text-block", fileName);
                    if (File.Exists(assets)) return assets;
                }
            }
            catch { }
            return null;
        }
    }
}
