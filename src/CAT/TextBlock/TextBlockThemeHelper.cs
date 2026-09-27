using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.TextBlocks
{
    internal static class TextBlockThemeHelper
    {
        public static string TryThemeMajorFont(PowerPoint.Slide slide)
        {
            try
            {
                dynamic fonts = slide.Design.SlideMaster.Theme.ThemeFontScheme.MajorFont;
                return (string)fonts[1].Name;
            }
            catch { return "Calibri Light"; }
        }

        public static string TryThemeMinorFont(PowerPoint.Slide slide)
        {
            try
            {
                dynamic fonts = slide.Design.SlideMaster.Theme.ThemeFontScheme.MinorFont;
                return (string)fonts[1].Name;
            }
            catch { return "Calibri"; }
        }

        public static void ApplyAccent1Color(PowerPoint.TextRange tr, PowerPoint.Slide slide)
        {
            try
            {
                var fc = tr.Font.Color;
                fc.ObjectThemeColor = Office.MsoThemeColorIndex.msoThemeColorAccent1;
                fc.TintAndShade = 0f;
            }
            catch
            {
                try
                {
                    dynamic themeColors = slide.ThemeColorScheme;
                    tr.Font.Color.RGB = themeColors[Office.MsoThemeColorSchemeIndex.msoThemeAccent1].RGB;
                }
                catch { }
            }
        }

        public static void ApplyText1Color(PowerPoint.TextRange tr, PowerPoint.Slide slide)
        {
            try
            {
                var fc = tr.Font.Color;
                fc.ObjectThemeColor = Office.MsoThemeColorIndex.msoThemeColorText1;
                fc.TintAndShade = 0f;
            }
            catch
            {
                try
                {
                    dynamic themeColors = slide.ThemeColorScheme;
                    tr.Font.Color.RGB = themeColors[Office.MsoThemeColorSchemeIndex.msoThemeDark1].RGB;
                }
                catch { }
            }
        }
    }
}
