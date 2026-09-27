using System;
using System.IO;

namespace Cat.Core
{
    /// <summary>
    /// SVG icon library folder. On installed machines: {CAT install dir}\icon-library
    /// (drop-in updates without reinstall). Dev: repo assets\icon-library or walk-up from bin.
    /// </summary>
    public static class IconLibraryPaths
    {
        public const string SubfolderName = "icon-library";

        public static string Root
        {
            get
            {
                string asmDir = Path.GetDirectoryName(typeof(IconLibraryPaths).Assembly.Location) ?? "";
                string nextToDll = Path.Combine(asmDir, SubfolderName);
                if (Directory.Exists(nextToDll)) return nextToDll;

                try
                {
                    var fromBin = Path.GetFullPath(Path.Combine(asmDir, "..", "..", "..", "..", "assets", SubfolderName));
                    if (Directory.Exists(fromBin)) return fromBin;
                }
                catch { }

                return nextToDll;
            }
        }
    }
}
