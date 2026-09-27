using System;
using System.Collections.Generic;
using System.Linq;
using Cat.Commands;
using Cat.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.AutoFormat
{
    internal static class Af
    {
        public const int LineGrey = 0x7F7F7F;
        public const int PanelGrey = 0xF2F2F2;
        public const float LineWeight = 0.5f;
        public const float BracketThickness = 10f;
        public const float BracketGap = 6f;

        public static List<PowerPoint.Shape> Sel(CommandContext c) => ShapeHelpers.ToList(c.SelectedShapes);

        public static List<PowerPoint.Shape> TopToBottom(CommandContext c) =>
            Sel(c).OrderBy(s => s.Top).ThenBy(s => s.Left).ToList();

        public static float FontSize(PowerPoint.Shape s)
        {
            try
            {
                if (ShapeHelpers.HasTextFrame(s))
                {
                    float sz = s.TextFrame.TextRange.Font.Size;
                    if (sz > 0) return sz;
                }
            }
            catch { }
            return 12f;
        }

        public static int Accent1(PowerPoint.Slide slide)
        {
            try
            {
                dynamic themeColors = slide.ThemeColorScheme;
                return themeColors[Office.MsoThemeColorSchemeIndex.msoThemeAccent1].RGB;
            }
            catch { return LineGrey; }
        }
    }

    internal static class AfGapLines
    {
        public static void StyleLine(PowerPoint.Shape line)
        {
            line.Line.Visible = Office.MsoTriState.msoTrue;
            line.Line.Weight = Af.LineWeight;
            line.Line.ForeColor.RGB = Af.LineGrey;
        }
    }

    public sealed class AfInsertVerticalGapLinesCommand : ICommand
    {
        public string Name => "Insert vertical gap lines";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;

        public void Execute(CommandContext c)
        {
            var shapes = Af.Sel(c).OrderBy(s => s.Left).ToList();
            var slide = c.ActiveSlide;
            int made = 0;
            for (int i = 0; i < shapes.Count - 1; i++)
            {
                var a = shapes[i]; var b = shapes[i + 1];
                float gapMid = ((a.Left + a.Width) + b.Left) / 2f;
                float top = Math.Min(a.Top, b.Top);
                float bottom = Math.Max(a.Top + a.Height, b.Top + b.Height);
                AfGapLines.StyleLine(slide.Shapes.AddLine(gapMid, top, gapMid, bottom));
                made++;
            }
            if (made == 0) Notifier.Info("Select at least two shapes with horizontal gaps between them.");
        }
    }

    public sealed class AfInsertHorizontalGapLinesCommand : ICommand
    {
        public string Name => "Insert horizontal gap lines";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;

        public void Execute(CommandContext c)
        {
            var shapes = Af.Sel(c).OrderBy(s => s.Top).ToList();
            var slide = c.ActiveSlide;
            int made = 0;
            for (int i = 0; i < shapes.Count - 1; i++)
            {
                var a = shapes[i]; var b = shapes[i + 1];
                float gapMid = ((a.Top + a.Height) + b.Top) / 2f;
                float left = Math.Min(a.Left, b.Left);
                float right = Math.Max(a.Left + a.Width, b.Left + b.Width);
                AfGapLines.StyleLine(slide.Shapes.AddLine(left, gapMid, right, gapMid));
                made++;
            }
            if (made == 0) Notifier.Info("Select at least two shapes with vertical gaps between them.");
        }
    }

    public enum BracketSide { Left, Right, Top, Bottom }

    public sealed class AfBracketCommand : ICommand
    {
        private readonly BracketSide _side;
        public AfBracketCommand(BracketSide side) { _side = side; }

        public string Name => $"Insert {_side.ToString().ToLowerInvariant()} bracket";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;

        public void Execute(CommandContext c)
        {
            var shapes = Af.Sel(c);
            var slide = c.ActiveSlide;

            float left = shapes.Min(s => s.Left);
            float top = shapes.Min(s => s.Top);
            float right = shapes.Max(s => s.Left + s.Width);
            float bottom = shapes.Max(s => s.Top + s.Height);
            float w = right - left, h = bottom - top;
            float t = Af.BracketThickness, g = Af.BracketGap;

            PowerPoint.Shape brace;
            switch (_side)
            {
                case BracketSide.Left:
                    brace = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeLeftBrace, left - g - t, top, t, h);
                    break;
                case BracketSide.Right:
                    brace = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightBrace, right + g, top, t, h);
                    break;
                case BracketSide.Top:
                    brace = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeLeftBrace, 0, 0, t, w);
                    brace.Rotation = 90f;
                    CentreRotated(brace, left + w / 2f, top - g - t / 2f);
                    break;
                default:
                    brace = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightBrace, 0, 0, t, w);
                    brace.Rotation = 90f;
                    CentreRotated(brace, left + w / 2f, bottom + g + t / 2f);
                    break;
            }

            brace.Line.Visible = Office.MsoTriState.msoTrue;
            brace.Line.Weight = Af.LineWeight;
            brace.Line.ForeColor.RGB = Af.LineGrey;
            brace.Fill.Visible = Office.MsoTriState.msoFalse;
        }

        private static void CentreRotated(PowerPoint.Shape s, float centreX, float centreY)
        {
            s.Left = centreX - s.Width / 2f;
            s.Top = centreY - s.Height / 2f;
        }
    }

    public enum CalloutDirection { Up, Down, Left, Right }

    public sealed class AfCalloutCommand : ICommand
    {
        private readonly CalloutDirection _dir;
        public AfCalloutCommand(CalloutDirection dir) { _dir = dir; }

        public string Name => $"Callout pointing {_dir.ToString().ToLowerInvariant()}";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;

        public void Execute(CommandContext c)
        {
            foreach (var s in Af.Sel(c))
            {
                if (!ShapeHelpers.HasTextFrame(s)) continue;
                try { s.AutoShapeType = Office.MsoAutoShapeType.msoShapeRectangularCallout; } catch { continue; }

                try
                {
                    float px, py;
                    switch (_dir)
                    {
                        case CalloutDirection.Up:    px = 0.5f;  py = -0.4f; break;
                        case CalloutDirection.Down:  px = 0.5f;  py = 1.4f;  break;
                        case CalloutDirection.Left:  px = -0.4f; py = 0.5f;  break;
                        default:                     px = 1.4f;  py = 0.5f;  break;
                    }
                    s.Adjustments[1] = px;
                    s.Adjustments[2] = py;
                }
                catch { }

                try
                {
                    s.Fill.Visible = Office.MsoTriState.msoTrue;
                    s.Fill.Solid();
                    s.Fill.ForeColor.RGB = Af.PanelGrey;
                    s.Line.Visible = Office.MsoTriState.msoFalse;
                }
                catch { }

                if (_dir == CalloutDirection.Up && ShapeHelpers.HasTextFrame(s))
                {
                    Cmd.SetTextFrameMargins(s.TextFrame, Cmd.Margin02CmPt);
                    Cmd.EnsureTextBoxPlaceholder(s.TextFrame);
                }
            }
        }
    }

    public sealed class AfKeyTakeawayCommand : ICommand
    {
        public string Name => "Key takeaway style";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;

        public void Execute(CommandContext c)
        {
            foreach (var s in Af.Sel(c))
            {
                try
                {
                    s.Fill.Visible = Office.MsoTriState.msoTrue;
                    s.Fill.Solid();
                    s.Fill.ForeColor.RGB = Af.PanelGrey;
                    s.Line.Visible = Office.MsoTriState.msoFalse;
                }
                catch { }

                if (ShapeHelpers.HasTextFrame(s))
                {
                    Cmd.SetTextFrameMargins(s.TextFrame, Cmd.Margin02CmPt);
                    Cmd.EnsureTextBoxPlaceholder(s.TextFrame);
                }
            }
        }
    }

    public sealed class AfNumberCircleCommand : ICommand
    {
        private readonly bool _letters;
        public AfNumberCircleCommand(bool letters) { _letters = letters; }

        public string Name => _letters ? "Insert lettered circles" : "Insert numbered circles";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;

        public void Execute(CommandContext c)
        {
            var ordered = Af.TopToBottom(c);
            var slide = c.ActiveSlide;
            int accent = Af.Accent1(slide);

            for (int i = 0; i < ordered.Count; i++)
            {
                var host = ordered[i];
                float fontSize = Af.FontSize(host);
                float diameter = fontSize * 1.7f;
                float gap = 6f;

                var circle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval,
                    host.Left - diameter - gap, host.Top, diameter, diameter);

                try
                {
                    circle.Fill.Visible = Office.MsoTriState.msoTrue;
                    circle.Fill.Solid();
                    circle.Fill.ForeColor.RGB = accent;
                    circle.Line.Visible = Office.MsoTriState.msoFalse;

                    var tf = circle.TextFrame;
                    tf.TextRange.Text = _letters ? Letter(i) : (i + 1).ToString();
                    tf.TextRange.Font.Size = fontSize;
                    tf.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                    tf.TextRange.Font.Color.RGB = 0xFFFFFF;
                    tf.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                    tf.MarginLeft = 0; tf.MarginRight = 0; tf.MarginTop = 0; tf.MarginBottom = 0;
                    tf.WordWrap = Office.MsoTriState.msoFalse;
                }
                catch { }
            }
        }

        private static string Letter(int index)
        {
            string s = "";
            int n = index;
            do { s = (char)('A' + (n % 26)) + s; n = n / 26 - 1; } while (n >= 0);
            return s;
        }
    }

    public sealed class AfPictureToRectangleCommand : ICommand
    {
        public string Name => "Picture to rectangle";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;

        public void Execute(CommandContext c)
        {
            int changed = 0;
            foreach (var s in Af.Sel(c))
            {
                try { s.AutoShapeType = Office.MsoAutoShapeType.msoShapeRectangle; changed++; }
                catch { }
            }
            if (changed == 0) Notifier.Info("None of the selected items could be reshaped.");
        }
    }

    public sealed class AfMakeSquareCommand : ICommand
    {
        public string Name => "Make 1:1";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;

        public void Execute(CommandContext c)
        {
            foreach (var s in Af.Sel(c))
            {
                try
                {
                    float side = Math.Max(s.Width, s.Height);
                    s.LockAspectRatio = Office.MsoTriState.msoFalse;
                    s.Width = side;
                    s.Height = side;
                }
                catch { }
            }
        }
    }

    public enum TemplateElement { Title, Subtitle, Footnote, Source }

    internal static class TemplateElementHelper
    {
        public static string PlaceholderKey(PowerPoint.Shape ph)
        {
            try
            {
                var pf = ph.PlaceholderFormat;
                return $"{(int)pf.Type}:{pf.ContainedType}";
            }
            catch { return ph.Name; }
        }

        public static bool Matches(PowerPoint.Shape s, TemplateElement element)
        {
            string name = "";
            try { name = (s.Name ?? "").ToLowerInvariant(); } catch { }

            PowerPoint.PpPlaceholderType? type = null;
            try
            {
                if (s.Type == Office.MsoShapeType.msoPlaceholder)
                    type = s.PlaceholderFormat.Type;
            }
            catch { }

            switch (element)
            {
                case TemplateElement.Title:
                    return type == PowerPoint.PpPlaceholderType.ppPlaceholderTitle
                        || type == PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                        || name.Contains("title");
                case TemplateElement.Subtitle:
                    return type == PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle
                        || name.Contains("subtitle");
                case TemplateElement.Footnote:
                    return type == PowerPoint.PpPlaceholderType.ppPlaceholderFooter
                        || name.Contains("footnote")
                        || name.Contains("footer");
                default:
                    return name.Contains("source");
            }
        }

        public static PowerPoint.Shape FindLayoutPlaceholder(PowerPoint.Slide slide, TemplateElement element)
        {
            try
            {
                if (slide.CustomLayout == null) return null;
                foreach (PowerPoint.Shape s in slide.CustomLayout.Shapes)
                    if (Matches(s, element)) return s;
            }
            catch { }
            return null;
        }

        public static PowerPoint.Shape FindSlidePlaceholder(PowerPoint.Slide slide, TemplateElement element, PowerPoint.Shape layoutPh)
        {
            if (layoutPh != null)
            {
                string key = PlaceholderKey(layoutPh);
                try
                {
                    foreach (PowerPoint.Shape s in slide.Shapes)
                    {
                        if (s.Type != Office.MsoShapeType.msoPlaceholder) continue;
                        if (PlaceholderKey(s) == key) return s;
                    }
                }
                catch { }
            }

            try
            {
                foreach (PowerPoint.Shape s in slide.Shapes)
                    if (Matches(s, element)) return s;
            }
            catch { }
            return null;
        }

        public static void ApplyAsFixedElement(PowerPoint.Shape slidePh, PowerPoint.Shape layoutPh, string textFromSelection)
        {
            slidePh.Left = layoutPh.Left;
            slidePh.Top = layoutPh.Top;
            slidePh.Width = layoutPh.Width;
            slidePh.Height = layoutPh.Height;
            try { slidePh.Rotation = layoutPh.Rotation; } catch { }

            try
            {
                slidePh.Fill.Visible = layoutPh.Fill.Visible;
                if (layoutPh.Fill.Visible == Office.MsoTriState.msoTrue)
                {
                    slidePh.Fill.Solid();
                    try { slidePh.Fill.ForeColor.RGB = layoutPh.Fill.ForeColor.RGB; } catch { }
                }
            }
            catch { }

            try
            {
                slidePh.Line.Visible = layoutPh.Line.Visible;
                if (layoutPh.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    try { slidePh.Line.ForeColor.RGB = layoutPh.Line.ForeColor.RGB; } catch { }
                    try { slidePh.Line.Weight = layoutPh.Line.Weight; } catch { }
                }
            }
            catch { }

            if (ShapeHelpers.HasTextFrame(slidePh) && ShapeHelpers.HasTextFrame(layoutPh))
            {
                var sf = layoutPh.TextFrame;
                var df = slidePh.TextFrame;
                try { df.VerticalAnchor = sf.VerticalAnchor; } catch { }
                try { df.WordWrap = sf.WordWrap; } catch { }
                try
                {
                    df.MarginLeft = sf.MarginLeft;
                    df.MarginRight = sf.MarginRight;
                    df.MarginTop = sf.MarginTop;
                    df.MarginBottom = sf.MarginBottom;
                }
                catch { }

                var rFont = sf.TextRange.Font;
                var dFont = df.TextRange.Font;
                try { dFont.Name = rFont.Name; } catch { }
                try { if (rFont.Size > 0) dFont.Size = rFont.Size; } catch { }
                try { dFont.Bold = rFont.Bold; } catch { }
                try { dFont.Italic = rFont.Italic; } catch { }
                try { dFont.Color.RGB = rFont.Color.RGB; } catch { }
                try { df.TextRange.ParagraphFormat.Alignment = sf.TextRange.ParagraphFormat.Alignment; } catch { }

                if (textFromSelection != null)
                    df.TextRange.Text = textFromSelection;
            }
        }
    }

    public sealed class AfAdoptTemplateElementCommand : ICommand
    {
        private readonly TemplateElement _element;
        public AfAdoptTemplateElementCommand(TemplateElement element) { _element = element; }

        public string Name => $"Apply template {_element.ToString().ToLowerInvariant()}";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection && c.ActiveSlide != null;

        public void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            var layoutPh = TemplateElementHelper.FindLayoutPlaceholder(slide, _element);
            if (layoutPh == null)
            {
                Notifier.Info($"No {_element.ToString().ToLowerInvariant()} placeholder in this slide layout.");
                return;
            }

            var slidePh = TemplateElementHelper.FindSlidePlaceholder(slide, _element, layoutPh);
            if (slidePh == null)
            {
                Notifier.Info($"This slide has no {_element.ToString().ToLowerInvariant()} placeholder to attach to.");
                return;
            }

            string mergedText = null;
            var toDelete = new List<PowerPoint.Shape>();
            foreach (var s in Af.Sel(c))
            {
                if (s.Id == slidePh.Id) continue;
                if (ShapeHelpers.HasTextFrame(s))
                {
                    var t = s.TextFrame.TextRange.Text;
                    if (!string.IsNullOrWhiteSpace(t))
                        mergedText = mergedText == null ? t.Trim() : mergedText + "\r" + t.Trim();
                }
                toDelete.Add(s);
            }

            TemplateElementHelper.ApplyAsFixedElement(slidePh, layoutPh, mergedText);
            foreach (var s in toDelete)
            {
                try { s.Delete(); } catch { }
            }
            slidePh.Select();
        }
    }
}
