using System;
using System.IO;

namespace Cat.Core
{
    /// <summary>Fixed icon library location under the CAT project assets folder.</summary>
    public static class IconLibraryPaths
    {
        public const string FixedAssetsPath = @"C:\Users\adith\Desktop\CAT\assets\icon-library";

        public static string Root
        {
            get
            {
                if (Directory.Exists(FixedAssetsPath)) return FixedAssetsPath;

                // F5 debug: …\src\CAT\bin\Debug → …\assets\icon-library
                try
                {
                    var asmDir = Path.GetDirectoryName(typeof(IconLibraryPaths).Assembly.Location) ?? "";
                    var fromBin = Path.GetFullPath(Path.Combine(asmDir, "..", "..", "..", "..", "assets", "icon-library"));
                    if (Directory.Exists(fromBin)) return fromBin;
                }
                catch { }

                return FixedAssetsPath;
            }
        }
    }
}
