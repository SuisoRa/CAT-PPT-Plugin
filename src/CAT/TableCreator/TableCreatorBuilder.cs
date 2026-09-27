using System;
using System.Collections.Generic;
using Cat.Commands;
using Cat.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.TableCreator
{
    /// <summary>Builds text-box tables from cell data and fixed layout rules.</summary>
    public static class TableCreatorBuilder
    {
        public static PowerPoint.Shape Build(
            PowerPoint.Slide slide,
            string[,] cellTexts,
            bool hasHeader,
            TableCreatorStyle style = null,
            bool useThemeText1Body = false)
        {
            if (slide == null || cellTexts == null) return null;
            style = style ?? TableCreatorStyle.Default;

            int rows = cellTexts.GetLength(0);
            int cols = cellTexts.GetLength(1);
            if (rows < TableCreatorSpec.MinDimension || cols < TableCreatorSpec.MinDimension) return null;

            if (!TableCreatorLayout.TryCompute(slide, rows, cols, out var layout))
                return null;

            string themeFont = TryThemeMinorFont(slide);
            var parts = new List<PowerPoint.Shape>();

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    float left = layout.CellLeft(c);
                    float top = layout.CellTop(r);
                    var tb = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        left, top, layout.CellWidth, layout.CellHeight);
                    ApplyCell(slide, tb, cellTexts[r, c], r == 0 && hasHeader, themeFont, style, useThemeText1Body);
                    parts.Add(tb);
                }
            }

            for (int r = 0; r < rows - 1; r++)
            {
                float y = layout.RowSeparatorY(r);
                var line = slide.Shapes.AddLine(layout.TableLeft, y, layout.TableLeft + layout.TableWidth, y);
                ApplySeparatorLine(line, style);
                parts.Add(line);
            }

            return GroupShapes(slide, parts);
        }

        private static void ApplyCell(
            PowerPoint.Slide slide,
            PowerPoint.Shape tb,
            string text,
            bool isHeader,
            string themeFont,
            TableCreatorStyle style,
            bool useThemeText1Body)
        {
            try
            {
                tb.Line.Visible = Office.MsoTriState.msoFalse;
                var tf = tb.TextFrame;
                Cmd.ClearTextFrameMargins(tf);
                try { tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop; } catch { }
                ShapeHelpers.SetWordWrap(tb, true);
                ShapeHelpers.SetAutoSize(tb, PowerPoint.PpAutoSize.ppAutoSizeNone);

                var tr = tf.TextRange;
                tr.Text = Cmd.TextOrPlaceholder(text);
                tr.Font.Name = themeFont;
                tr.Font.Size = isHeader ? style.HeaderFontSizePt : style.BodyFontSizePt;
                tr.Font.Bold = isHeader && style.HeaderBold
                    ? Office.MsoTriState.msoTrue
                    : Office.MsoTriState.msoFalse;

                if (useThemeText1Body)
                    ApplyThemeText1Color(tr, slide);
            }
            catch { }
        }

        private static void ApplyThemeText1Color(PowerPoint.TextRange tr, PowerPoint.Slide slide)
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

        private static void ApplySeparatorLine(PowerPoint.Shape line, TableCreatorStyle style)
        {
            try
            {
                line.Line.Visible = Office.MsoTriState.msoTrue;
                line.Line.Weight = style.SeparatorWeightPt;
                line.Line.ForeColor.RGB = style.SeparatorColorRgb;
            }
            catch { }
        }

        private static string TryThemeMinorFont(PowerPoint.Slide slide)
        {
            try
            {
                dynamic fonts = slide.Design.SlideMaster.Theme.ThemeFontScheme.MinorFont;
                return (string)fonts[1].Name;
            }
            catch { return "Calibri"; }
        }

        private static PowerPoint.Shape GroupShapes(PowerPoint.Slide slide, List<PowerPoint.Shape> parts)
        {
            if (parts == null || parts.Count == 0) return null;
            if (parts.Count == 1) return parts[0];

            var names = new string[parts.Count];
            for (int i = 0; i < parts.Count; i++)
                names[i] = parts[i].Name;

            try { return slide.Shapes.Range(names).Group(); }
            catch { return parts[0]; }
        }
    }
}
