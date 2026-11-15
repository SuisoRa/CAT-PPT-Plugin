using System;
using System.Collections.Generic;
using System.IO;

namespace Cat.Core
{
    /// <summary>
    /// SVG icon library folder. On installed machines: {CAT install dir}\icon-library
    /// (drop-in updates without reinstall). Dev: repo assets\icon-library, found by walking up from the DLL.
    /// </summary>
    public static class IconLibraryPaths
    {
        public const string SubfolderName = "icon-library";

        public static string Root
        {
            get
            {
                string fallback = null;
                foreach (string start in AssemblyDirs())
                {
                    string nextToDll = Path.Combine(start, SubfolderName);
                    if (ContainsSvg(nextToDll)) return nextToDll;
                    if (fallback == null) fallback = nextToDll;

                    try
                    {
                        var dir = new DirectoryInfo(start);
                        for (int i = 0; i < 8 && dir != null; i++, dir = dir.Parent)
                        {
                            string assets = Path.Combine(dir.FullName, "assets", SubfolderName);
                            if (ContainsSvg(assets)) return assets;
                        }
                    }
                    catch { }
                }

                return fallback ?? "";
            }
        }

        private static bool ContainsSvg(string dir)
        {
            if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir)) return false;
            try
            {
                foreach (string family in Directory.EnumerateDirectories(dir))
                {
                    foreach (string _ in Directory.EnumerateFiles(family, "*.svg", SearchOption.TopDirectoryOnly))
                        return true;
                }
            }
            catch { }
            return false;
        }

        private static string[] AssemblyDirs()
        {
            var list = new System.Collections.Generic.List<string>();
            void Add(string dir)
            {
                if (string.IsNullOrEmpty(dir)) return;
                try { dir = Path.GetFullPath(dir); } catch { return; }
                if (!list.Contains(dir)) list.Add(dir);
            }

            try { Add(Path.GetDirectoryName(typeof(IconLibraryPaths).Assembly.Location)); } catch { }
            try
            {
                string codeBase = typeof(IconLibraryPaths).Assembly.CodeBase;
                if (!string.IsNullOrEmpty(codeBase))
                    Add(Path.GetDirectoryName(new Uri(codeBase).LocalPath));
            }
            catch { }

            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\PowerPoint\Addins\CAT"))
                {
                    string manifest = key?.GetValue("Manifest") as string;
                    if (!string.IsNullOrEmpty(manifest))
                    {
                        int bar = manifest.IndexOf('|');
                        if (bar >= 0) manifest = manifest.Substring(0, bar);
                        if (manifest.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
                            Add(Path.GetDirectoryName(new Uri(manifest).LocalPath));
                    }
                }
            }
            catch { }

            return list.ToArray();
        }
    }
}
