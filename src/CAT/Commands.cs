using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WinForms = System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Cat.AutoFormat;
using Cat.Core;
using Cat.TableCreator;
using Cat.TextBlocks;
using Cat.UI;

namespace Cat.Commands
{
    public interface ICommand
    {
        string Name { get; }
        bool CanExecute(CommandContext ctx);
        void Execute(CommandContext ctx);
    }

    public sealed class CommandContext
    {
        public PowerPoint.Application App { get; }
        public CommandContext(PowerPoint.Application app) { App = app; }

        public PowerPoint.DocumentWindow ActiveWindow
        { get { try { return App.ActiveWindow; } catch { return null; } } }

        public PowerPoint.Selection Selection
        { get { try { return ActiveWindow?.Selection; } catch { return null; } } }

        public PowerPoint.Slide ActiveSlide
        { get { try { return ActiveWindow?.View?.Slide as PowerPoint.Slide; } catch { return null; } } }

        public bool HasShapeSelection
        {
            get
            {
                try
                {
                    var range = SelectedShapes;
                    if (range != null && range.Count > 0) return true;
                }
                catch { }
                return TextEditShape() != null;
            }
        }

        public PowerPoint.ShapeRange SelectedShapes
        {
            get
            {
                try
                {
                    var s = Selection;
                    if (s == null) return null;
                    var t = s.Type;
                    if (t != PowerPoint.PpSelectionType.ppSelectionShapes
                        && t != PowerPoint.PpSelectionType.ppSelectionText)
                        return null;
                    var range = s.ShapeRange;
                    if (range != null && range.Count > 0) return range;
                }
                catch { }
                return null;
            }
        }

        /// <summary>Shape that contains the caret when PowerPoint reports a text selection.</summary>
        public PowerPoint.Shape TextEditShape()
        {
            try
            {
                if (!HasTextSelection) return null;
                try
                {
                    var range = Selection?.ShapeRange;
                    if (range != null && range.Count > 0) return range[1];
                }
                catch { }
                dynamic frame = SelectedText?.Parent;
                return frame?.Parent as PowerPoint.Shape;
            }
            catch { return null; }
        }

        public bool HasTextSelection
        { get { try { return Selection != null && Selection.Type == PowerPoint.PpSelectionType.ppSelectionText; } catch { return false; } } }

        public PowerPoint.TextRange SelectedText
        { get { try { return Selection?.TextRange; } catch { return null; } } }

        public IList<PowerPoint.Shape> OrderedSelection() =>
            ThisAddIn.Instance.SelectionTracker.GetOrderedSelection(SelectedShapes);

        public static float SlideWidth(PowerPoint.Slide s)  => ((PowerPoint.Presentation)s.Parent).PageSetup.SlideWidth;
        public static float SlideHeight(PowerPoint.Slide s) => ((PowerPoint.Presentation)s.Parent).PageSetup.SlideHeight;
    }

    internal static class Cmd
    {
        public const float Pt = ShapeHelpers.PointsPerInch;

        public static void ForEachSelected(CommandContext ctx, Action<PowerPoint.Shape> action)
        {
            var shapes = ShapeHelpers.ToList(ctx?.SelectedShapes);
            if (shapes.Count == 0) shapes = EditableShapes(ctx);
            foreach (var s in shapes) action(s);
        }

        /// <summary>Shape selection, or the text box that contains the caret (text-edit / soft selection).</summary>
        public static List<PowerPoint.Shape> EditableShapes(CommandContext ctx)
        {
            var list = new List<PowerPoint.Shape>();
            if (ctx == null) return list;
            if (ctx.HasShapeSelection && ctx.SelectedShapes != null)
            {
                list.AddRange(ShapeHelpers.ToList(ctx.SelectedShapes));
                if (list.Count > 0) return list;
            }
            var editing = ctx.TextEditShape();
            if (editing != null)
            {
                list.Add(editing);
                return list;
            }
            if (!ctx.HasTextSelection) return list;

            try
            {
                var range = ctx.Selection?.ShapeRange;
                if (range != null && range.Count > 0)
                {
                    list.AddRange(ShapeHelpers.ToList(range));
                    if (list.Count > 0) return list;
                }
            }
            catch { }

            try
            {
                dynamic frame = ctx.SelectedText?.Parent;
                if (frame != null)
                {
                    PowerPoint.Shape shape = frame.Parent as PowerPoint.Shape;
                    if (shape != null) list.Add(shape);
                }
            }
            catch { }
            return list;
        }

        public static void ForEachEditable(CommandContext ctx, Action<PowerPoint.Shape> action)
        { foreach (var s in EditableShapes(ctx)) action(s); }

        public static PowerPoint.Shape Reference(CommandContext ctx)
        { var o = ctx.OrderedSelection(); return o.Count > 0 ? o[0] : null; }

        public static PowerPoint.TextRange TargetText(CommandContext ctx)
        {
            if (ctx.HasTextSelection && ctx.SelectedText != null) return ctx.SelectedText;
            if (ctx.HasShapeSelection && ShapeHelpers.HasTextFrame(ctx.SelectedShapes[1]))
                return ctx.SelectedShapes[1].TextFrame.TextRange;
            return null;
        }

        public static int Clamp(int v, int lo, int hi) => Math.Max(lo, Math.Min(hi, v));

        public static void Select(IList<PowerPoint.Shape> shapes)
        {
            if (shapes == null || shapes.Count == 0) return;
            for (int i = 0; i < shapes.Count; i++)
                shapes[i].Select(i == 0 ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
        }

        public const string TextBoxPlaceholder = "Text";

        /// <summary>0.2 cm in PowerPoint points.</summary>
        public const float Margin02CmPt = 0.2f * 72f / 2.54f;

        public static string TextOrPlaceholder(string text) =>
            string.IsNullOrWhiteSpace(text) ? TextBoxPlaceholder : text;

        public static void EnsureTextBoxPlaceholder(PowerPoint.TextFrame tf)
        {
            if (tf == null) return;
            try
            {
                var tr = tf.TextRange;
                if (string.IsNullOrWhiteSpace(tr.Text)) tr.Text = TextBoxPlaceholder;
            }
            catch { }
        }

        public static void ClearTextFrameMargins(PowerPoint.TextFrame tf)
        {
            if (tf == null) return;
            try { tf.MarginLeft = 0f; tf.MarginRight = 0f; tf.MarginTop = 0f; tf.MarginBottom = 0f; } catch { }
        }

        public static void SetTextFrameMargins(PowerPoint.TextFrame tf, float pt)
        {
            if (tf == null) return;
            try { tf.MarginLeft = pt; tf.MarginRight = pt; tf.MarginTop = pt; tf.MarginBottom = pt; } catch { }
        }

        public static void SetSingleLineSpacing(PowerPoint.TextRange tr)
        {
            SetMultipleLineSpacing(tr, 1f);
        }

        public static void SetMultipleLineSpacing(PowerPoint.TextRange tr, float multiple)
        {
            if (tr == null) return;
            try
            {
                int n = tr.Paragraphs().Count;
                for (int i = 1; i <= n; i++)
                {
                    var pf = tr.Paragraphs(i).ParagraphFormat;
                    pf.LineRuleWithin = Office.MsoTriState.msoTrue;
                    pf.SpaceWithin = multiple;
                }
            }
            catch { }
        }

        /// <summary>One pass: body font, plain weight, Text 1 (or black when requested), no highlight; keeps each fragment's point size.</summary>
        public static void FixTextRangeFormatting(PowerPoint.TextRange tr, PowerPoint.Slide slide, bool blackFont = false, bool keepFont = false, bool keepColor = false)
        {
            if (tr == null || slide == null) return;
            string bodyFont = keepFont ? null : TextBlockThemeHelper.TryThemeMinorFont(slide);

            int len = 0;
            try { len = tr.Length; } catch { }

            bool anyRun = false;
            for (int i = 1; i <= 500; i++)
            {
                try
                {
                    var run = tr.Runs(i);
                    if (run == null || run.Length < 1) break;
                    NormalizeTextRun(run, bodyFont, slide, blackFont, keepColor);
                    anyRun = true;
                }
                catch { break; }
            }

            if (!anyRun && len > 0)
                NormalizeTextRun(tr, bodyFont, slide, blackFont, keepColor);
        }

        public static void FixTextBoxShape(PowerPoint.Shape shape, PowerPoint.Slide slide)
        {
            if (shape == null || slide == null || !ShapeHelpers.HasTextFrame(shape)) return;
            try
            {
                shape.Fill.Visible = Office.MsoTriState.msoFalse;
                shape.Line.Visible = Office.MsoTriState.msoFalse;
                shape.Shadow.Visible = Office.MsoTriState.msoFalse;
                var tf = shape.TextFrame;
                Cmd.ClearTextFrameMargins(tf);
                try { tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop; } catch { }
                ShapeHelpers.SetAutoSize(shape, PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText);
                var tr = tf.TextRange;
                bool kpmgBold = IsOnlyFont(tr, "KPMG Bold");
                if (kpmgBold) Cmd.SetMultipleLineSpacing(tr, 0.8f);
                else Cmd.SetSingleLineSpacing(tr);
                var bulletParagraphs = BulletParagraphIndexes(tr);
                FixTextRangeFormatting(tr, slide, blackFont: !kpmgBold, keepFont: kpmgBold, keepColor: kpmgBold);
                ResetPlainParagraphLayout(tr, bulletParagraphs);
                ClearTextHighlights(shape);
                FinishResetTypography(shape, keepColor: kpmgBold);
                if (bulletParagraphs.Count > 0) ApplyResetBullets(shape, tr, bulletParagraphs);
            }
            catch { }
        }

        private static bool IsOnlyFont(PowerPoint.TextRange tr, string fontName)
        {
            if (tr == null || string.IsNullOrEmpty(fontName)) return false;
            bool any = false;
            try
            {
                for (int i = 1; i <= 500; i++)
                {
                    PowerPoint.TextRange run;
                    try { run = tr.Runs(i); }
                    catch { break; }
                    if (run == null || run.Length < 1) break;

                    string text;
                    try { text = run.Text ?? ""; }
                    catch { text = ""; }
                    if (text.Trim('\r', '\n', '\v', '\t', ' ').Length == 0) continue;

                    string name;
                    try { name = run.Font.Name ?? ""; }
                    catch { return false; }
                    if (!string.Equals(name.Trim(), fontName, StringComparison.OrdinalIgnoreCase))
                        return false;
                    any = true;
                }
            }
            catch { return false; }
            return any;
        }

        private static void NormalizeTextRun(PowerPoint.TextRange run, string bodyFont, PowerPoint.Slide slide, bool blackFont, bool keepColor = false)
        {
            if (run == null) return;
            float size;
            try { size = run.Font.Size; }
            catch { size = 12f; }
            if (size < 1f) size = 12f;

            var f = run.Font;
            try
            {
                f.Bold = Office.MsoTriState.msoFalse;
                f.Italic = Office.MsoTriState.msoFalse;
                f.Underline = Office.MsoTriState.msoFalse;
                f.Emboss = Office.MsoTriState.msoFalse;
                f.Shadow = Office.MsoTriState.msoFalse;
                f.Superscript = Office.MsoTriState.msoFalse;
                f.Subscript = Office.MsoTriState.msoFalse;
                try { f.BaselineOffset = 0f; } catch { }
                if (!string.IsNullOrEmpty(bodyFont)) f.Name = bodyFont;
                f.Size = size;
            }
            catch { }

            try
            {
                dynamic df = f;
                try { df.StrikeThrough = Office.MsoTriState.msoFalse; } catch { }
                try { df.DoubleStrikeThrough = Office.MsoTriState.msoFalse; } catch { }
                try { df.AllCaps = Office.MsoTriState.msoFalse; } catch { }
                try { df.SmallCaps = Office.MsoTriState.msoFalse; } catch { }
                try { df.Spacing = 0f; } catch { }
                try { df.CharacterSpacing = 0f; } catch { }
                try { df.Kerning = 0f; } catch { }
                try { df.Kernings = Office.MsoTriState.msoFalse; } catch { }
                try { df.Highlight = Office.MsoTriState.msoFalse; } catch { }
            }
            catch { }

            if (!keepColor)
            {
                if (blackFont)
                    ApplyExplicitBlack(run);
                else
                    TextBlockThemeHelper.ApplyText1Color(run, slide);
            }
            ClearRunHighlight(run, keepColor);
        }

        /// <summary>RGB 0 is treated as automatic and stays theme white. 0x010000 is BGR for RGB(0,0,1).</summary>
        private static void ApplyExplicitBlack(PowerPoint.TextRange run)
        {
            try
            {
                dynamic color = run.Font.Color;
                try { color.ObjectThemeColor = Office.MsoThemeColorIndex.msoNotThemeColor; } catch { }
                color.RGB = 0x010000;
            }
            catch
            {
                try { run.Font.Color.RGB = 0x010000; } catch { }
            }
        }

        /// <summary>Font dialog spacing (condense/expand) and color live on TextFrame2. Once per run, not per character.</summary>
        private static void FinishResetTypography(PowerPoint.Shape shape, bool keepColor = false)
        {
            try
            {
                dynamic tr2 = shape.TextFrame2.TextRange;
                ApplyFont2Reset(tr2.Font, keepColor);
                try
                {
                    dynamic runs = tr2.Runs;
                    int count = (int)runs.Count;
                    for (int i = 1; i <= count && i <= 500; i++)
                    {
                        try { ApplyFont2Reset(runs[i].Font, keepColor); }
                        catch { try { ApplyFont2Reset(runs.Item(i).Font, keepColor); } catch { } }
                    }
                }
                catch { }
            }
            catch { }
        }

        private static void ApplyFont2Reset(dynamic font, bool keepColor)
        {
            try { font.Spacing = 0f; } catch { }
            try { font.Kerning = 0f; } catch { }
            if (keepColor) return;
            try
            {
                dynamic fill = font.Fill;
                fill.Visible = Office.MsoTriState.msoTrue;
                fill.Solid();
                dynamic fore = fill.ForeColor;
                try { fore.ObjectThemeColor = Office.MsoThemeColorIndex.msoNotThemeColor; } catch { }
                fore.RGB = 0x010000;
            }
            catch { }
        }

        private static List<int> BulletParagraphIndexes(PowerPoint.TextRange tr)
        {
            var indexes = new List<int>();
            try
            {
                int n = tr.Paragraphs().Count;
                for (int i = 1; i <= n; i++)
                    if (ParagraphHasBullet(tr.Paragraphs(i))) indexes.Add(i);
            }
            catch { }
            return indexes;
        }

        private static bool ParagraphHasBullet(PowerPoint.TextRange paragraph)
        {
            try
            {
                var bullet = paragraph.ParagraphFormat.Bullet;
                if (bullet.Visible == Office.MsoTriState.msoFalse) return false;
                var type = bullet.Type;
                return type != PowerPoint.PpBulletType.ppBulletNone;
            }
            catch { return false; }
        }

        /// <summary>Non-bullet paragraphs: no extra spacing or indent. Box margins are already zero.</summary>
        private static void ResetPlainParagraphLayout(PowerPoint.TextRange tr, List<int> bulletParagraphs)
        {
            var bullets = new HashSet<int>(bulletParagraphs);
            try
            {
                int n = tr.Paragraphs().Count;
                for (int i = 1; i <= n; i++)
                {
                    if (bullets.Contains(i)) continue;
                    try
                    {
                        var pf = tr.Paragraphs(i).ParagraphFormat;
                        try { pf.SpaceBefore = 0f; } catch { }
                        try { pf.SpaceAfter = 0f; } catch { }
                        dynamic format = pf;
                        try { format.LeftIndent = 0f; } catch { }
                        try { format.FirstLineIndent = 0f; } catch { }
                    }
                    catch { }
                }
            }
            catch { }
        }

        private static void ApplyResetBullets(PowerPoint.Shape shape, PowerPoint.TextRange tr, List<int> paragraphIndexes)
        {
            foreach (int i in paragraphIndexes)
            {
                try
                {
                    var p = tr.Paragraphs(i);
                    int level = 1;
                    try { level = Clamp(p.IndentLevel, 1, 5); } catch { }

                    var pf = p.ParagraphFormat;
                    var bullet = pf.Bullet;
                    bullet.Visible = Office.MsoTriState.msoTrue;
                    bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
                    bullet.Character = 0x2022;
                    bullet.RelativeSize = 1f;
                    try { bullet.UseTextColor = Office.MsoTriState.msoTrue; } catch { }
                    try { bullet.UseTextFont = Office.MsoTriState.msoTrue; } catch { }
                    try { p.IndentLevel = level; } catch { }

                    float stepPt = (ParagraphFontSize(p) >= 10.5f ? 0.2f : 0.15f) * 72f;
                    // Before text grows by one step per level. Hanging stays one step, so the bullet sits one step to the left of the text.
                    float beforePt = stepPt * level;
                    float bulletPt = stepPt * (level - 1);
                    SetBulletIndent(shape, pf, level, bulletPt, beforePt);
                }
                catch { }
            }
        }

        /// <summary>
        /// The paragraph dialog reads the text-frame ruler for bulleted text. ParagraphFormat indents alone leave the template ruler (often 0.31") in place.
        /// </summary>
        private static void SetBulletIndent(PowerPoint.Shape shape, PowerPoint.ParagraphFormat pf, int level, float bulletPt, float beforePt)
        {
            try
            {
                dynamic format = pf;
                format.FirstLineIndent = bulletPt - beforePt;
                format.LeftIndent = beforePt;
            }
            catch { }

            try
            {
                var rulerLevel = shape.TextFrame.Ruler.Levels[level];
                rulerLevel.FirstMargin = bulletPt;
                rulerLevel.LeftMargin = beforePt;
            }
            catch { }

            try
            {
                dynamic ruler2 = shape.TextFrame2.Ruler;
                dynamic rulerLevel2 = ruler2.Levels[level];
                rulerLevel2.FirstMargin = bulletPt;
                rulerLevel2.LeftMargin = beforePt;
            }
            catch { }
        }

        private static float ParagraphFontSize(PowerPoint.TextRange paragraph)
        {
            try
            {
                int len = paragraph.Length;
                if (len > 1)
                {
                    float size = paragraph.Characters(1, 1).Font.Size;
                    if (size >= 1f) return size;
                }
                float paraSize = paragraph.Font.Size;
                if (paraSize >= 1f) return paraSize;
            }
            catch { }
            return 12f;
        }

        private static void ClearRunHighlight(PowerPoint.TextRange run, bool keepColor)
        {
            if (run == null) return;
            try
            {
                dynamic f = run.Font;
                try { f.Highlight = Office.MsoTriState.msoFalse; } catch { }
                if (!keepColor)
                    try { f.Fill.Visible = Office.MsoTriState.msoFalse; } catch { }
            }
            catch { }
        }

        private static void ClearTextHighlights(PowerPoint.Shape shape)
        {
            try
            {
                dynamic highlight = shape.TextFrame2.TextRange.Font.Highlight;
                highlight.Transparency = 1f;
            }
            catch { }
        }

        public static void ReduceTextFrameMargins(PowerPoint.TextFrame tf, float deltaPt)
        {
            if (tf == null || deltaPt <= 0f) return;
            try { tf.MarginLeft = Math.Max(0f, tf.MarginLeft - deltaPt); } catch { }
            try { tf.MarginRight = Math.Max(0f, tf.MarginRight - deltaPt); } catch { }
            try { tf.MarginTop = Math.Max(0f, tf.MarginTop - deltaPt); } catch { }
            try { tf.MarginBottom = Math.Max(0f, tf.MarginBottom - deltaPt); } catch { }
        }

        /// <summary>Current within-paragraph spacing as a line multiple, one value per indent level 1–5.</summary>
        public static double[] ReadLineSpacingInLines(PowerPoint.TextRange tr)
        {
            var values = new double[] { 1, 1, 1, 1, 1 };
            if (tr == null) return values;
            var seen = new bool[5];
            try
            {
                int n = tr.Paragraphs().Count;
                for (int i = 1; i <= n; i++)
                {
                    var p = tr.Paragraphs(i);
                    int level = Clamp(p.IndentLevel, 1, 5) - 1;
                    if (seen[level]) continue;
                    var pf = p.ParagraphFormat;
                    float space = 1f;
                    bool inLines = true;
                    try { space = pf.SpaceWithin; } catch { }
                    try { inLines = pf.LineRuleWithin != Office.MsoTriState.msoFalse; } catch { }
                    if (!inLines)
                    {
                        float font = 12f;
                        try { if (p.Font.Size > 0) font = p.Font.Size; } catch { }
                        space = font > 0 ? space / font : 1f;
                    }
                    if (space < 0 || float.IsNaN(space) || float.IsInfinity(space)) space = 1f;
                    values[level] = Math.Round(space, 2);
                    seen[level] = true;
                }
            }
            catch { }
            return values;
        }
    }

    public abstract class ShapeCommandBase : ICommand
    {
        public abstract string Name { get; }
        public virtual bool CanExecute(CommandContext ctx) => ctx.HasShapeSelection;
        public abstract void Execute(CommandContext ctx);
        protected static Office.MsoTriState Rel(CommandContext ctx) =>
            ctx.SelectedShapes.Count > 1 ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
    }

    public sealed class AlignTopCommand : ShapeCommandBase
    {
        public override string Name => "Align Top";
        public override void Execute(CommandContext c)
        {
            if (AddInSettings.AlignByFirst)
            {
                var first = Cmd.Reference(c);
                if (first == null) return;
                float top = first.Top;
                Cmd.ForEachSelected(c, s => s.Top = top);
                return;
            }
            c.SelectedShapes.Align(Office.MsoAlignCmd.msoAlignTops, Rel(c));
        }
    }
    public sealed class AlignBottomCommand : ShapeCommandBase
    {
        public override string Name => "Align Bottom";
        public override void Execute(CommandContext c)
        {
            if (AddInSettings.AlignByFirst)
            {
                var first = Cmd.Reference(c);
                if (first == null) return;
                float bottom = first.Top + first.Height;
                Cmd.ForEachSelected(c, s => s.Top = bottom - s.Height);
                return;
            }
            c.SelectedShapes.Align(Office.MsoAlignCmd.msoAlignBottoms, Rel(c));
        }
    }
    public sealed class AlignLeftCommand : ShapeCommandBase
    {
        public override string Name => "Align Left";
        public override void Execute(CommandContext c)
        {
            if (AddInSettings.AlignByFirst)
            {
                var first = Cmd.Reference(c);
                if (first == null) return;
                float left = first.Left;
                Cmd.ForEachSelected(c, s => s.Left = left);
                return;
            }
            c.SelectedShapes.Align(Office.MsoAlignCmd.msoAlignLefts, Rel(c));
        }
    }
    public sealed class AlignRightCommand : ShapeCommandBase
    {
        public override string Name => "Align Right";
        public override void Execute(CommandContext c)
        {
            if (AddInSettings.AlignByFirst)
            {
                var first = Cmd.Reference(c);
                if (first == null) return;
                float right = first.Left + first.Width;
                Cmd.ForEachSelected(c, s => s.Left = right - s.Width);
                return;
            }
            c.SelectedShapes.Align(Office.MsoAlignCmd.msoAlignRights, Rel(c));
        }
    }
    public sealed class AlignCenterCommand : ShapeCommandBase
    {
        public override string Name => "Align Center";
        public override void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            int n = c.SelectedShapes.Count;
            if (n == 0) return;

            if (n == 1)
            {
                var s = c.SelectedShapes[1];
                float sw = CommandContext.SlideWidth(slide);
                s.Left = (sw - s.Width) / 2f;
                return;
            }

            var shapes = ShapeHelpers.ToList(c.SelectedShapes);
            PowerPoint.Shape anchor;
            if (AddInSettings.AlignByFirst)
            {
                anchor = Cmd.Reference(c);
                if (anchor == null) return;
            }
            else
                anchor = shapes.OrderBy(s => s.Left).First();

            float centerX = anchor.Left + anchor.Width / 2f;
            foreach (var s in shapes) s.Left = centerX - s.Width / 2f;
        }
    }

    public sealed class AlignMiddleCommand : ShapeCommandBase
    {
        public override string Name => "Align Middle";
        public override void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            int n = c.SelectedShapes.Count;
            if (n == 0) return;

            if (n == 1)
            {
                var s = c.SelectedShapes[1];
                float sh = CommandContext.SlideHeight(slide);
                s.Top = (sh - s.Height) / 2f;
                return;
            }

            var shapes = ShapeHelpers.ToList(c.SelectedShapes);
            PowerPoint.Shape anchor;
            if (AddInSettings.AlignByFirst)
            {
                anchor = Cmd.Reference(c);
                if (anchor == null) return;
            }
            else
                anchor = shapes.OrderBy(s => s.Top).First();

            float centerY = anchor.Top + anchor.Height / 2f;
            foreach (var s in shapes) s.Top = centerY - s.Height / 2f;
        }
    }

    public sealed class DistributeHorizontallyCommand : ShapeCommandBase
    { public override string Name => "Distribute Horizontally";
      public override bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 3;
      public override void Execute(CommandContext c) => c.SelectedShapes.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, Office.MsoTriState.msoFalse); }
    public sealed class DistributeVerticallyCommand : ShapeCommandBase
    { public override string Name => "Distribute Vertically";
      public override bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 3;
      public override void Execute(CommandContext c) => c.SelectedShapes.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, Office.MsoTriState.msoFalse); }

    public sealed class AlignAndGroupCommand : ShapeCommandBase
    {
        public override string Name => "Align and Group";
        public override bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;
        public override void Execute(CommandContext c)
        {
            if (AddInSettings.AlignByFirst)
            {
                var first = Cmd.Reference(c);
                if (first != null)
                {
                    float top = first.Top;
                    Cmd.ForEachSelected(c, s => s.Top = top);
                }
            }
            else
                c.SelectedShapes.Align(Office.MsoAlignCmd.msoAlignTops, Office.MsoTriState.msoFalse);
            c.SelectedShapes.Group();
        }
    }

    internal static class AlignClusterHelper
    {
        public static bool OverlapVertical(PowerPoint.Shape a, PowerPoint.Shape b) =>
            a.Top < b.Top + b.Height && b.Top < a.Top + a.Height;

        public static bool OverlapHorizontal(PowerPoint.Shape a, PowerPoint.Shape b) =>
            a.Left < b.Left + b.Width && b.Left < a.Left + a.Width;

        public static List<List<PowerPoint.Shape>> ClusterRows(IList<PowerPoint.Shape> shapes) =>
            Cluster(shapes, OverlapVertical);

        public static List<List<PowerPoint.Shape>> ClusterColumns(IList<PowerPoint.Shape> shapes) =>
            Cluster(shapes, OverlapHorizontal);

        private static List<List<PowerPoint.Shape>> Cluster(
            IList<PowerPoint.Shape> shapes,
            Func<PowerPoint.Shape, PowerPoint.Shape, bool> overlaps)
        {
            var pool = shapes.ToList();
            var groups = new List<List<PowerPoint.Shape>>();
            while (pool.Count > 0)
            {
                var group = new List<PowerPoint.Shape> { pool[0] };
                pool.RemoveAt(0);
                bool grew;
                do
                {
                    grew = false;
                    for (int i = pool.Count - 1; i >= 0; i--)
                    {
                        if (group.Any(g => overlaps(g, pool[i])))
                        {
                            group.Add(pool[i]);
                            pool.RemoveAt(i);
                            grew = true;
                        }
                    }
                } while (grew);
                groups.Add(group);
            }
            return groups;
        }

        public static void GroupOnSlide(PowerPoint.Slide slide, IList<PowerPoint.Shape> shapes)
        {
            if (shapes == null || shapes.Count < 2) return;
            shapes[0].Select();
            for (int i = 1; i < shapes.Count; i++)
                shapes[i].Select(Office.MsoTriState.msoFalse);
            try
            {
                var names = new object[shapes.Count];
                for (int i = 0; i < shapes.Count; i++) names[i] = shapes[i].Name;
                slide.Shapes.Range(names).Group();
            }
            catch { }
        }
    }

    /// <summary>Table-friendly: cluster by vertical overlap into rows, top-align each row to its topmost shape, group each row.</summary>
    public sealed class AlignRowsAndGroupCommand : ShapeCommandBase
    {
        public override string Name => "Align rows and group";
        public override bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;
        public override void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            var shapes = ShapeHelpers.ToList(c.SelectedShapes);
            foreach (var row in AlignClusterHelper.ClusterRows(shapes))
            {
                if (row.Count == 0) continue;
                float anchorTop = row.Min(s => s.Top);
                foreach (var s in row) s.Top = anchorTop;
                AlignClusterHelper.GroupOnSlide(slide, row);
            }
        }
    }

    /// <summary>Table-friendly: cluster by horizontal overlap into columns, left-align each column to its leftmost shape, group each column.</summary>
    public sealed class AlignColumnsAndGroupCommand : ShapeCommandBase
    {
        public override string Name => "Align columns and group";
        public override bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;
        public override void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            var shapes = ShapeHelpers.ToList(c.SelectedShapes);
            foreach (var col in AlignClusterHelper.ClusterColumns(shapes))
            {
                if (col.Count == 0) continue;
                float anchorLeft = col.Min(s => s.Left);
                foreach (var s in col) s.Left = anchorLeft;
                AlignClusterHelper.GroupOnSlide(slide, col);
            }
        }
    }

    public sealed class BringForwardCommand : ShapeCommandBase
    { public override string Name => "Bring Forward";
      public override void Execute(CommandContext c) => Cmd.ForEachSelected(c, s => s.ZOrder(Office.MsoZOrderCmd.msoBringForward)); }
    public sealed class BringToFrontCommand : ShapeCommandBase
    { public override string Name => "Bring to Front";
      public override void Execute(CommandContext c) => Cmd.ForEachSelected(c, s => s.ZOrder(Office.MsoZOrderCmd.msoBringToFront)); }
    public sealed class SendBackwardCommand : ShapeCommandBase
    { public override string Name => "Send Backward";
      public override void Execute(CommandContext c) => Cmd.ForEachSelected(c, s => s.ZOrder(Office.MsoZOrderCmd.msoSendBackward)); }
    public sealed class SendToBackCommand : ShapeCommandBase
    { public override string Name => "Send to Back";
      public override void Execute(CommandContext c) => Cmd.ForEachSelected(c, s => s.ZOrder(Office.MsoZOrderCmd.msoSendToBack)); }

    public sealed class MakeSameHeightCommand : ShapeCommandBase
    {
        public override string Name => "Make Same Height";
        public override bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;
        public override void Execute(CommandContext c)
        { var r = Cmd.Reference(c); if (r == null) return; float h = r.Height;
          Cmd.ForEachSelected(c, s => { if (s.Id != r.Id) { s.LockAspectRatio = Office.MsoTriState.msoFalse; s.Height = h; } }); }
    }
    public sealed class MakeSameWidthCommand : ShapeCommandBase
    {
        public override string Name => "Make Same Width";
        public override bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;
        public override void Execute(CommandContext c)
        { var r = Cmd.Reference(c); if (r == null) return; float w = r.Width;
          Cmd.ForEachSelected(c, s => { if (s.Id != r.Id) { s.LockAspectRatio = Office.MsoTriState.msoFalse; s.Width = w; } }); }
    }
    public sealed class MakeSameSizeCommand : ShapeCommandBase
    {
        public override string Name => "Make Same Size";
        public override bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;
        public override void Execute(CommandContext c)
        { var r = Cmd.Reference(c); if (r == null) return; float w = r.Width, h = r.Height;
          Cmd.ForEachSelected(c, s => { if (s.Id != r.Id) { s.LockAspectRatio = Office.MsoTriState.msoFalse; s.Width = w; s.Height = h; } }); }
    }

    public sealed class DoNotResizeCommand : ICommand
    {
        public string Name => "Do Not Resize";
        public bool CanExecute(CommandContext c) => Cmd.EditableShapes(c).Count > 0;
        public void Execute(CommandContext c) => Cmd.ForEachEditable(c, s => ShapeHelpers.SetAutoSize(s, PowerPoint.PpAutoSize.ppAutoSizeNone));
    }
    public sealed class ResizeCommand : ICommand
    {
        public string Name => "Resize";
        public bool CanExecute(CommandContext c) => Cmd.EditableShapes(c).Count > 0;
        public void Execute(CommandContext c) => Cmd.ForEachEditable(c, s => ShapeHelpers.SetAutoSize(s, PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText));
    }

    public sealed class FitToWindowCommand : ICommand
    {
        public string Name => "Fit slide to window";
        public bool CanExecute(CommandContext c) => c.ActiveWindow != null;
        public void Execute(CommandContext c)
        {
            try { ((dynamic)c.ActiveWindow.View).ZoomToFit = Office.MsoTriState.msoTrue; }
            catch
            {
                try { c.App.CommandBars.ExecuteMso("ZoomFitToWindow"); } catch { }
            }
        }
    }

    public sealed class ResetFixedElementsCommand : ICommand
    {
        public string Name => "Reset Fixed Elements";
        public bool CanExecute(CommandContext c) => c.ActiveSlide != null;
        public void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            var layout = slide.CustomLayout;
            if (layout == null) return;

            var map = new Dictionary<string, PowerPoint.Shape>();
            foreach (PowerPoint.Shape lp in layout.Shapes)
                if (lp.Type == Office.MsoShapeType.msoPlaceholder) map[Key(lp)] = lp;

            foreach (PowerPoint.Shape ph in slide.Shapes)
            {
                if (ph.Type != Office.MsoShapeType.msoPlaceholder) continue;
                if (!map.TryGetValue(Key(ph), out var lp)) continue;
                ph.Left = lp.Left; ph.Top = lp.Top; ph.Width = lp.Width; ph.Height = lp.Height;
                try { ph.Rotation = lp.Rotation; } catch { }
            }

            RestoreFooterAndSource(slide, layout);
        }

        private static void RestoreFooterAndSource(PowerPoint.Slide slide, PowerPoint.CustomLayout layout)
        {
            FooterSource.ApplyReference(slide, layout);
        }

        private static string Key(PowerPoint.Shape ph)
        { try { var pf = ph.PlaceholderFormat; return $"{(int)pf.Type}:{pf.ContainedType}"; } catch { return ph.Name; } }
    }

    public sealed class AddFooterSourceCommand : ICommand
    {
        public string Name => "Footer and source";
        public bool CanExecute(CommandContext c) => c.ActiveSlide != null;
        public void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            var layout = slide.CustomLayout;
            if (layout == null)
            {
                Notifier.Info("This slide has no layout to hold the footer and source.");
                return;
            }

            try
            {
                FooterSource.EnsureOnLayout(layout, slide);
                FooterSource.EnsureOnSlide(slide, layout);
                FooterSource.ApplyReference(slide, layout);
            }
            catch (Exception ex)
            {
                Notifier.Error("Could not add the footer and source: " + ex.Message);
            }
        }
    }

    internal static class FooterSource
    {
        public const string SourceName = "Source";
        public const string FooterName = "Footer";

        // Reference deck "Footer and source reference.pptx", 13.333 x 7.5 in. EMU / 12700 = points.
        const float RefSlideHeight = 540f;
        const float BoxLeft = 995364f / 12700f;
        const float BoxWidth = 7071360f / 12700f;
        const float BoxHeight = 123111f / 12700f;
        const float FooterTopOnRef = 5853192f / 12700f;
        const float SourceTopOnRef = 6037152f / 12700f;
        const float FootnoteIndent = 18f;

        public static void EnsureOnLayout(PowerPoint.CustomLayout layout, PowerPoint.Slide slide)
        {
            float slideHeight = SlideHeight(slide);
            if (FindFooter(layout.Shapes) == null)
            {
                try
                {
                    layout.Shapes.AddPlaceholder(
                        PowerPoint.PpPlaceholderType.ppPlaceholderFooter,
                        BoxLeft, TopFromBottom(slideHeight, FooterTopOnRef), BoxWidth, BoxHeight);
                }
                catch { }
            }

            var source = FindSource(layout.Shapes);
            if (source != null && !IsPlaceholder(source))
            {
                string kept = ReadText(source);
                try { source.Delete(); } catch { }
                source = null;
                var created = AddSourcePlaceholder(layout, TopFromBottom(slideHeight, SourceTopOnRef));
                if (created != null && !string.IsNullOrWhiteSpace(kept))
                    try { created.TextFrame.TextRange.Text = kept; } catch { }
            }
            else if (source == null)
            {
                AddSourcePlaceholder(layout, TopFromBottom(slideHeight, SourceTopOnRef));
            }
            else
            {
                MarkSource(source);
            }
        }

        public static void EnsureOnSlide(PowerPoint.Slide slide, PowerPoint.CustomLayout layout)
        {
            if (FindFooter(slide.Shapes) == null)
            {
                try { slide.HeadersFooters.Footer.Visible = Office.MsoTriState.msoTrue; } catch { }
                if (FindFooter(slide.Shapes) == null)
                    CopyShape(slide, FindFooter(layout.Shapes), FooterName);
            }

            var layoutSource = FindSource(layout.Shapes);
            var slideSource = FindSource(slide.Shapes);
            if (slideSource != null && !IsPlaceholder(slideSource))
            {
                string kept = ReadText(slideSource);
                try { slideSource.Delete(); } catch { }
                slideSource = PasteSourcePlaceholder(slide, layoutSource);
                if (slideSource != null && !string.IsNullOrWhiteSpace(kept))
                    try { slideSource.TextFrame.TextRange.Text = kept; } catch { }
            }
            else if (slideSource == null)
            {
                PasteSourcePlaceholder(slide, layoutSource);
            }
            else
            {
                MarkSource(slideSource);
            }
        }

        public static void ApplyReference(PowerPoint.Slide slide, PowerPoint.CustomLayout layout)
        {
            float slideHeight = SlideHeight(slide);
            string textFont = TextBlockThemeHelper.TryThemeMinorFont(slide);
            string bulletFont = TextBlockThemeHelper.TryThemeMajorFont(slide);
            ApplyFootnote(FindFooter(layout.Shapes), slide, slideHeight, textFont, bulletFont);
            ApplyFootnote(FindFooter(slide.Shapes), slide, slideHeight, textFont, bulletFont);
            ApplySource(FindSource(layout.Shapes), slide, slideHeight);
            ApplySource(FindSource(slide.Shapes), slide, slideHeight);
        }

        private static float SlideHeight(PowerPoint.Slide slide)
        {
            try { return ((PowerPoint.Presentation)slide.Parent).PageSetup.SlideHeight; }
            catch { return RefSlideHeight; }
        }

        private static float TopFromBottom(float slideHeight, float topOnRef)
        {
            return slideHeight - (RefSlideHeight - topOnRef);
        }

        private static void ApplyFootnote(PowerPoint.Shape shape, PowerPoint.Slide slide, float slideHeight, string textFont, string bulletFont)
        {
            if (shape == null || !ShapeHelpers.HasTextFrame(shape)) return;
            Place(shape, TopFromBottom(slideHeight, FooterTopOnRef));
            var tf = shape.TextFrame;
            try { tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom; } catch { }
            ZeroMargins(tf);
            try { tf.WordWrap = Office.MsoTriState.msoTrue; } catch { }
            try { tf.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText; } catch { }

            var tr = tf.TextRange;
            try { if (string.IsNullOrWhiteSpace(tr.Text)) tr.Text = "XX"; } catch { }
            try
            {
                tr.Font.Name = textFont;
                tr.Font.Size = 8f;
                tr.Font.Bold = Office.MsoTriState.msoFalse;
                TextBlockThemeHelper.ApplyText1Color(tr, slide);
            }
            catch { }

            try
            {
                var pf = tr.ParagraphFormat;
                pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                pf.LineRuleWithin = Office.MsoTriState.msoTrue;
                pf.SpaceWithin = 1f;
                pf.SpaceBefore = 0f;
                pf.SpaceAfter = 0f;
                var bullet = pf.Bullet;
                bullet.Visible = Office.MsoTriState.msoTrue;
                bullet.Type = PowerPoint.PpBulletType.ppBulletNumbered;
                try { bullet.Style = PowerPoint.PpNumberedBulletStyle.ppBulletArabicPeriod; } catch { }
                bullet.RelativeSize = 1f;
                try { bullet.UseTextColor = Office.MsoTriState.msoTrue; } catch { }
                try { bullet.Font.Name = bulletFont; } catch { }
                dynamic format = pf;
                try { format.LeftIndent = FootnoteIndent; } catch { }
                try { format.FirstLineIndent = -FootnoteIndent; } catch { }
            }
            catch { }

            try
            {
                var level = shape.TextFrame.Ruler.Levels[1];
                level.FirstMargin = 0f;
                level.LeftMargin = FootnoteIndent;
            }
            catch { }

            Place(shape, TopFromBottom(slideHeight, FooterTopOnRef));
        }

        private static void ApplySource(PowerPoint.Shape shape, PowerPoint.Slide slide, float slideHeight)
        {
            if (shape == null || !ShapeHelpers.HasTextFrame(shape)) return;
            Place(shape, TopFromBottom(slideHeight, SourceTopOnRef));
            var tf = shape.TextFrame;
            try { tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop; } catch { }
            ZeroMargins(tf);
            try { tf.WordWrap = Office.MsoTriState.msoTrue; } catch { }
            try { tf.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText; } catch { }

            var tr = tf.TextRange;
            try { if (string.IsNullOrWhiteSpace(tr.Text)) tr.Text = SourceName; } catch { }
            try
            {
                tr.Font.Name = "Arial";
                tr.Font.Size = 8f;
                tr.Font.Bold = Office.MsoTriState.msoFalse;
                TextBlockThemeHelper.ApplyText1Color(tr, slide);
            }
            catch { }

            try
            {
                var pf = tr.ParagraphFormat;
                pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                pf.LineRuleWithin = Office.MsoTriState.msoTrue;
                pf.SpaceWithin = 1f;
                pf.SpaceBefore = 0f;
                try { pf.LineRuleAfter = Office.MsoTriState.msoFalse; } catch { }
                pf.SpaceAfter = 6f;
                var bullet = pf.Bullet;
                bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
                bullet.Visible = Office.MsoTriState.msoFalse;
                dynamic format = pf;
                try { format.LeftIndent = 0f; } catch { }
                try { format.FirstLineIndent = 0f; } catch { }
            }
            catch { }

            Place(shape, TopFromBottom(slideHeight, SourceTopOnRef));
        }

        private static void Place(PowerPoint.Shape shape, float top)
        {
            try { shape.LockAspectRatio = Office.MsoTriState.msoFalse; } catch { }
            try { shape.Line.Visible = Office.MsoTriState.msoFalse; } catch { }
            try { shape.Fill.Visible = Office.MsoTriState.msoFalse; } catch { }
            shape.Left = BoxLeft;
            shape.Top = top;
            shape.Width = BoxWidth;
            shape.Height = BoxHeight;
        }

        private static void ZeroMargins(PowerPoint.TextFrame tf)
        {
            try
            {
                tf.MarginLeft = 0f;
                tf.MarginRight = 0f;
                tf.MarginTop = 0f;
                tf.MarginBottom = 0f;
            }
            catch { }
        }

        public static PowerPoint.Shape FindFooter(PowerPoint.Shapes shapes)
        {
            PowerPoint.Shape named = null;
            try
            {
                foreach (PowerPoint.Shape s in shapes)
                {
                    try
                    {
                        if (s.Type == Office.MsoShapeType.msoPlaceholder
                            && s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderFooter)
                            return s;
                    }
                    catch { }
                    try
                    {
                        if (named == null && string.Equals(s.Name, FooterName, StringComparison.OrdinalIgnoreCase))
                            named = s;
                    }
                    catch { }
                }
            }
            catch { }
            return named;
        }

        public static PowerPoint.Shape FindSource(PowerPoint.Shapes shapes)
        {
            PowerPoint.Shape named = null;
            try
            {
                foreach (PowerPoint.Shape s in shapes)
                {
                    if (!IsSourceShape(s)) continue;
                    if (IsPlaceholder(s)) return s;
                    if (named == null) named = s;
                }
            }
            catch { }
            return named;
        }

        private static PowerPoint.Shape AddSourcePlaceholder(PowerPoint.CustomLayout layout, float top)
        {
            PowerPoint.Shape source = null;
            try
            {
                source = layout.Shapes.AddPlaceholder(
                    PowerPoint.PpPlaceholderType.ppPlaceholderBody, BoxLeft, top, BoxWidth, BoxHeight);
            }
            catch
            {
                try
                {
                    source = layout.Shapes.AddPlaceholder(
                        PowerPoint.PpPlaceholderType.ppPlaceholderObject, BoxLeft, top, BoxWidth, BoxHeight);
                }
                catch { return null; }
            }
            MarkSource(source);
            return source;
        }

        private static PowerPoint.Shape PasteSourcePlaceholder(PowerPoint.Slide slide, PowerPoint.Shape layoutShape)
        {
            if (layoutShape == null) return null;
            try
            {
                layoutShape.Copy();
                var pasted = slide.Shapes.Paste()[1];
                MarkSource(pasted);
                return pasted;
            }
            catch { return null; }
        }

        private static void MarkSource(PowerPoint.Shape shape)
        {
            if (shape == null) return;
            try { shape.Name = SourceName; } catch { }
            try { shape.Tags.Add("CAT_ROLE", "Source"); } catch { }
        }

        private static bool IsPlaceholder(PowerPoint.Shape shape)
        {
            try { return shape != null && shape.Type == Office.MsoShapeType.msoPlaceholder; }
            catch { return false; }
        }

        private static bool IsSourceShape(PowerPoint.Shape shape)
        {
            try
            {
                string name = shape.Name ?? "";
                if (name.Equals(SourceName, StringComparison.OrdinalIgnoreCase)) return true;
                if (name.StartsWith(SourceName + " ", StringComparison.OrdinalIgnoreCase)) return true;
            }
            catch { }

            try
            {
                int n = shape.Tags.Count;
                for (int i = 1; i <= n; i++)
                {
                    if (string.Equals(shape.Tags.Name(i), "CAT_ROLE", StringComparison.OrdinalIgnoreCase)
                        && string.Equals(shape.Tags.Value(i), "Source", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch { }
            return false;
        }

        private static string ReadText(PowerPoint.Shape shape)
        {
            try
            {
                if (!ShapeHelpers.HasTextFrame(shape)) return "";
                return shape.TextFrame.TextRange.Text ?? "";
            }
            catch { return ""; }
        }

        private static void CopyShape(PowerPoint.Slide slide, PowerPoint.Shape layoutShape, string name)
        {
            if (layoutShape == null) return;
            float left = 0f, top = 0f, width = 72f, height = 20f;
            try { left = layoutShape.Left; top = layoutShape.Top; width = layoutShape.Width; height = layoutShape.Height; } catch { }

            var box = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            try { box.Name = name; } catch { }

            string text = name;
            try
            {
                if (ShapeHelpers.HasTextFrame(layoutShape))
                {
                    var existing = layoutShape.TextFrame.TextRange.Text;
                    if (!string.IsNullOrWhiteSpace(existing)) text = existing;
                }
            }
            catch { }
            TemplateElementHelper.ApplyAsFixedElement(box, layoutShape, text);
        }
    }

    public sealed class CopyPositionCommand : ICommand
    {
        public string Name => "Copy Position";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;
        public void Execute(CommandContext c)
        {
            var clip = ThisAddIn.Instance.PositionClipboard;
            clip.Clear();
            foreach (var s in c.OrderedSelection()) clip.Add(s.Left, s.Top);
        }
    }

    public sealed class PastePositionCommand : ICommand
    {
        public string Name => "Paste Position";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;
        public void Execute(CommandContext c)
        {
            var clip = ThisAddIn.Instance.PositionClipboard;
            if (!clip.HasData) { Notifier.Info("Copy a position first (Ctrl+1), then paste it (Ctrl+2)."); return; }
            var ordered = c.OrderedSelection();
            int n = Math.Min(ordered.Count, clip.Count);
            for (int i = 0; i < n; i++) { ordered[i].Left = clip[i].Left; ordered[i].Top = clip[i].Top; }
        }
    }

    public sealed class CopySpacingCommand : ICommand
    {
        public string Name => "Copy spacing";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count == 2;
        public void Execute(CommandContext c)
        {
            var shapes = ShapeHelpers.ToList(c.SelectedShapes);
            if (shapes.Count != 2)
            {
                Notifier.Info("Select exactly two objects to copy the spacing between them.");
                return;
            }
            ThisAddIn.Instance.SpacingClipboard.Set(HorizontalGap(shapes[0], shapes[1]), VerticalGap(shapes[0], shapes[1]));
        }

        internal static float HorizontalGap(PowerPoint.Shape a, PowerPoint.Shape b)
        {
            var left = a.Left <= b.Left ? a : b;
            var right = ReferenceEquals(left, a) ? b : a;
            return right.Left - (left.Left + left.Width);
        }

        internal static float VerticalGap(PowerPoint.Shape a, PowerPoint.Shape b)
        {
            var top = a.Top <= b.Top ? a : b;
            var bottom = ReferenceEquals(top, a) ? b : a;
            return bottom.Top - (top.Top + top.Height);
        }
    }

    public sealed class PasteVerticalSpacingCommand : ICommand
    {
        public string Name => "Paste vertical spacing";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;
        public void Execute(CommandContext c)
        {
            var clip = ThisAddIn.Instance.SpacingClipboard;
            if (!clip.HasData)
            {
                Notifier.Info("Copy spacing from two objects first.");
                return;
            }
            var ordered = ShapeHelpers.ToList(c.SelectedShapes).OrderBy(s => s.Top).ThenBy(s => s.Left).ToList();
            for (int i = 1; i < ordered.Count; i++)
                ordered[i].Top = ordered[i - 1].Top + ordered[i - 1].Height + clip.Vertical;
        }
    }

    public sealed class PasteHorizontalSpacingCommand : ICommand
    {
        public string Name => "Paste horizontal spacing";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;
        public void Execute(CommandContext c)
        {
            var clip = ThisAddIn.Instance.SpacingClipboard;
            if (!clip.HasData)
            {
                Notifier.Info("Copy spacing from two objects first.");
                return;
            }
            var ordered = ShapeHelpers.ToList(c.SelectedShapes).OrderBy(s => s.Left).ThenBy(s => s.Top).ToList();
            for (int i = 1; i < ordered.Count; i++)
                ordered[i].Left = ordered[i - 1].Left + ordered[i - 1].Width + clip.Horizontal;
        }
    }

    internal static class InsertConst
    {
        public const int StickyYellow = 0x00A8F2FF;
        public const float StickyWIn = 3f, StickyHIn = 2f;
    }

    public sealed class InsertTextboxCommand : ICommand
    {
        public string Name => "Insert Textbox";
        public bool CanExecute(CommandContext c) => c.ActiveSlide != null;
        public void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            float w = 3 * Cmd.Pt, h = 1 * Cmd.Pt;
            float left = 0.5f * Cmd.Pt;
            float top = 0.5f * Cmd.Pt;

            if (SlideLayoutHelper.TryGetTitleLayout(slide, out float tLeft, out float tTop, out float tWidth, out float tHeight))
            {
                left = tLeft;
                top = tTop + tHeight + 8f;
                if (tWidth > w) w = Math.Min(tWidth, 5f * Cmd.Pt);
            }

            var tb = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, w, h);
            try
            {
                dynamic fonts = slide.Design.SlideMaster.Theme.ThemeFontScheme.MinorFont;
                var tf = tb.TextFrame;
                var tr = tf.TextRange;
                tr.Text = Cmd.TextBoxPlaceholder;
                tr.Font.Name = fonts[1].Name;
                tr.Font.Size = 12f;
                tb.Line.Visible = Office.MsoTriState.msoFalse;
                Cmd.ClearTextFrameMargins(tf);
                ShapeHelpers.SetAutoSize(tb, PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText);
            }
            catch { }
            tb.Select();
        }
    }

    public sealed class InsertNewSlideCommand : ICommand
    {
        public string Name => "Insert New Slide";
        public bool CanExecute(CommandContext c) => c.ActiveSlide != null;
        public void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            var pres = (PowerPoint.Presentation)slide.Parent;
            pres.Slides.AddSlide(slide.SlideIndex + 1, slide.CustomLayout).Select();
        }
    }

    public sealed class InsertYellowStickyCommand : ICommand
    {
        public string Name => "Insert Yellow Sticky";
        public bool CanExecute(CommandContext c) => c.ActiveSlide != null;
        public void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            float w = InsertConst.StickyWIn * Cmd.Pt, h = InsertConst.StickyHIn * Cmd.Pt;
            float sw = CommandContext.SlideWidth(slide);
            var s = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFoldedCorner,
                sw - w / 2f, -h / 2f, w, h);
            s.Fill.Solid(); s.Fill.ForeColor.RGB = InsertConst.StickyYellow;
            s.Line.Visible = Office.MsoTriState.msoFalse;
            s.Shadow.Visible = Office.MsoTriState.msoFalse;
            try
            {
                if (ShapeHelpers.HasTextFrame(s))
                {
                    var tr = s.TextFrame.TextRange;
                    tr.Text = Cmd.TextBoxPlaceholder;
                    tr.Font.Size = 12f;
                    tr.Font.Color.RGB = 0x000000;
                }
            }
            catch { }
            s.Select();
        }
    }

    public sealed class PasteUnformattedTextCommand : ICommand
    {
        public string Name => "Paste Unformatted Text";
        public bool CanExecute(CommandContext c) => WinForms.Clipboard.ContainsText();
        public void Execute(CommandContext c)
        {
            string text = WinForms.Clipboard.GetText();
            if (string.IsNullOrEmpty(text)) return;
            if (c.HasTextSelection && c.SelectedText != null) c.SelectedText.Text = text;
            else if (c.HasShapeSelection && ShapeHelpers.HasTextFrame(c.SelectedShapes[1])) c.SelectedShapes[1].TextFrame.TextRange.Text = text;
            else
            {
                var tb = c.ActiveSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 60, 60, 300, 60);
                tb.TextFrame.TextRange.Text = Cmd.TextOrPlaceholder(text); tb.Select();
            }
        }
    }

    public sealed class CycleAccentColorsCommand : ICommand
    {
        public string Name => "Cycle Accent Colors";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;
        public void Execute(CommandContext c)
        { var slide = c.ActiveSlide; Cmd.ForEachSelected(c, s => ShapeHelpers.CycleAccentFill(s, slide)); }
    }

    public sealed class IncreaseListLevelCommand : ICommand
    {
        public string Name => "Increase List Level";
        public bool CanExecute(CommandContext c) => Cmd.TargetText(c) != null;
        public void Execute(CommandContext c)
        {
            var tr = Cmd.TargetText(c);
            for (int i = 1; i <= tr.Paragraphs().Count; i++)
            {
                var p = tr.Paragraphs(i);
                float size = 0;
                try { size = p.Font.Size; } catch { }
                p.IndentLevel = Cmd.Clamp(p.IndentLevel + 1, 1, 5);
                if (size > 0) try { p.Font.Size = size; } catch { }
            }
        }
    }

    public sealed class DecreaseListLevelCommand : ICommand
    {
        public string Name => "Decrease List Level";
        public bool CanExecute(CommandContext c) => Cmd.TargetText(c) != null;
        public void Execute(CommandContext c)
        {
            var tr = Cmd.TargetText(c);
            for (int i = 1; i <= tr.Paragraphs().Count; i++)
            {
                var p = tr.Paragraphs(i);
                float size = 0;
                try { size = p.Font.Size; } catch { }
                p.IndentLevel = Cmd.Clamp(p.IndentLevel - 1, 1, 5);
                if (size > 0) try { p.Font.Size = size; } catch { }
            }
        }
    }

    /// <summary>Strip box chrome; body font per run (sizes kept); plain text; Text 1; no highlight.</summary>
    public sealed class FixTextBoxCommand : ICommand
    {
        public string Name => "Fix Text Box";
        public bool CanExecute(CommandContext c) =>
            Cmd.EditableShapes(c).Any(ShapeHelpers.HasTextFrame);

        public void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            if (slide == null) return;
            Cmd.ForEachEditable(c, s => Cmd.FixTextBoxShape(s, slide));
        }
    }

    /// <summary>With a selection: delete all other shapes. With nothing selected: delete everything on the slide.</summary>
    public sealed class DeleteAllExceptSelectionCommand : ICommand
    {
        public string Name => "Delete All Except Selection";
        public bool CanExecute(CommandContext c) => c.ActiveSlide != null;
        public void Execute(CommandContext c)
        {
            var slide = c.ActiveSlide;
            var keep = new HashSet<int>();
            if (c.HasShapeSelection)
            {
                foreach (var s in ShapeHelpers.ToList(c.SelectedShapes)) keep.Add(s.Id);
            }

            var toDelete = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape sh in slide.Shapes)
            {
                if (keep.Count == 0 || !keep.Contains(sh.Id))
                    toDelete.Add(sh);
            }

            foreach (var sh in toDelete)
            {
                try { sh.Delete(); } catch { }
            }
        }
    }

    public sealed class WordWrapCommand : ICommand
    {
        public string Name => "Word Wrap";
        public bool CanExecute(CommandContext c) => Cmd.EditableShapes(c).Count > 0;
        public void Execute(CommandContext c) => Cmd.ForEachEditable(c, s => ShapeHelpers.SetWordWrap(s, true));
    }
    public sealed class DoNotWordWrapCommand : ICommand
    {
        public string Name => "Do Not Word Wrap";
        public bool CanExecute(CommandContext c) => Cmd.EditableShapes(c).Count > 0;
        public void Execute(CommandContext c) => Cmd.ForEachEditable(c, s => ShapeHelpers.SetWordWrap(s, false));
    }

    public sealed class ClearSelectedTextCommand : ICommand
    {
        public string Name => "Clear text in selection";
        public bool CanExecute(CommandContext c) => Cmd.EditableShapes(c).Any(ShapeHelpers.HasTextFrame);
        public void Execute(CommandContext c)
        {
            Cmd.ForEachEditable(c, s =>
            {
                if (!ShapeHelpers.HasTextFrame(s)) return;
                try { s.TextFrame.TextRange.Text = ""; } catch { }
            });
        }
    }

    public sealed class ListLineSpacingCommand : ICommand
    {
        public string Name => "List Line Spacing";
        public bool CanExecute(CommandContext c) =>
            Cmd.EditableShapes(c).Any(ShapeHelpers.HasTextFrame) || Cmd.TargetText(c) != null;

        public void Execute(CommandContext c)
        {
            var targets = Cmd.EditableShapes(c).Where(ShapeHelpers.HasTextFrame).ToList();
            PowerPoint.TextRange sample = targets.Count > 0
                ? targets[0].TextFrame.TextRange
                : Cmd.TargetText(c);
            if (sample == null) return;

            var r = Dialogs.ShowLineSpacing(Cmd.ReadLineSpacingInLines(sample));
            if (!r.Ok) return;

            if (targets.Count == 0)
                Apply(sample, r);
            else
                foreach (var s in targets)
                    try { Apply(s.TextFrame.TextRange, r); } catch { }
        }

        private static void Apply(PowerPoint.TextRange tr, LineSpacingResult r)
        {
            for (int i = 1; i <= tr.Paragraphs().Count; i++)
            {
                var p = tr.Paragraphs(i);
                var pf = p.ParagraphFormat;
                int level = Cmd.Clamp(p.IndentLevel, 1, 5);
                pf.LineRuleWithin = Office.MsoTriState.msoTrue;
                pf.SpaceWithin = r.Auto ? 1f : (float)r.Values[level - 1];
            }
        }
    }

    public sealed class TextBoxToShapeCommand : ICommand
    {
        public string Name => "Text box to shape";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;
        public void Execute(CommandContext c)
        {
            float m = 0.1f * Cmd.Pt;
            Cmd.ForEachSelected(c, s =>
            {
                if (!ShapeHelpers.HasTextFrame(s)) return;
                try { s.AutoShapeType = Office.MsoAutoShapeType.msoShapeRectangle; } catch { }
                s.TextFrame.MarginLeft = m; s.TextFrame.MarginRight = m; s.TextFrame.MarginTop = m; s.TextFrame.MarginBottom = m;
            });
        }
    }

    public sealed class ShapeToTextBoxCommand : ICommand
    {
        public string Name => "Shape to text box";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;
        public void Execute(CommandContext c)
        {
            float mx = 0.1f * Cmd.Pt, my = 0.05f * Cmd.Pt;
            Cmd.ForEachSelected(c, s =>
            {
                if (!ShapeHelpers.HasTextFrame(s)) return;
                s.Fill.Visible = Office.MsoTriState.msoFalse; s.Line.Visible = Office.MsoTriState.msoFalse;
                s.TextFrame.MarginLeft = mx; s.TextFrame.MarginRight = mx; s.TextFrame.MarginTop = my; s.TextFrame.MarginBottom = my;
                ShapeHelpers.SetAutoSize(s, PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText);
            });
        }
    }

    public sealed class SplitJoinTextboxesCommand : ICommand
    {
        public string Name => "Split/Join Textboxes";
        public bool CanExecute(CommandContext c) =>
            Cmd.EditableShapes(c).Any(ShapeHelpers.HasTextFrame);
        public void Execute(CommandContext c)
        {
            int count = Cmd.EditableShapes(c).Count;
            var r = Dialogs.ShowSplitJoin(count >= 2);
            if (!r.Ok) return;
            if (r.Mode == SplitJoinMode.Split) DoSplit(c, r.Level, r.KeepFormatting);
            else DoJoin(c, r.KeepFormatting);
        }

        private static void DoSplit(CommandContext c, int level, bool keep)
        {
            var shapes = Cmd.EditableShapes(c);
            if (shapes.Count < 1) return;
            var source = shapes[0];
            if (!ShapeHelpers.HasTextFrame(source)) return;
            var tr = source.TextFrame.TextRange;
            int paraCount = tr.Paragraphs().Count;
            if (paraCount <= 1) return;

            var chunks = new List<List<int>>();
            var cur = new List<int>();
            for (int i = 1; i <= paraCount; i++)
            {
                var p = tr.Paragraphs(i);
                bool boundary = (level == -1) || (p.IndentLevel <= level);
                if (boundary && cur.Count > 0) { chunks.Add(cur); cur = new List<int>(); }
                cur.Add(i);
            }
            if (cur.Count > 0) chunks.Add(cur);
            if (chunks.Count <= 1) return;

            var slide = c.ActiveSlide;
            float top = source.Top;
            foreach (var chunk in chunks)
            {
                PowerPoint.Shape box = keep
                    ? source.Duplicate()[1]
                    : slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, source.Left, top, source.Width, 20);
                box.Left = source.Left;
                box.Top = top;
                if (keep)
                    KeepParagraphs(box, chunk);
                else
                {
                    var texts = new List<string>();
                    foreach (int index in chunk)
                        texts.Add(tr.Paragraphs(index).Text.TrimEnd('\r', '\n'));
                    box.TextFrame.TextRange.Text = Cmd.TextOrPlaceholder(string.Join("\r", texts));
                }
                box.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                top = box.Top + box.Height + 6f;
            }
            source.Delete();
        }

        /// <summary>Drop every paragraph that is not in the chunk so each run keeps its own formatting.</summary>
        private static void KeepParagraphs(PowerPoint.Shape box, List<int> keepIndexes)
        {
            var keep = new HashSet<int>(keepIndexes);
            var range = box.TextFrame.TextRange;
            int count;
            try { count = range.Paragraphs().Count; }
            catch { return; }
            for (int i = count; i >= 1; i--)
            {
                if (keep.Contains(i)) continue;
                try { range.Paragraphs(i).Delete(); } catch { }
            }
        }

        private static void DoJoin(CommandContext c, bool keep)
        {
            var ordered = Cmd.EditableShapes(c).Where(ShapeHelpers.HasTextFrame).ToList();
            if (ordered.Count < 2) return;
            var first = ordered[0];
            var sb = new StringBuilder();
            foreach (var s in ordered) { if (sb.Length > 0) sb.Append('\r'); sb.Append(s.TextFrame.TextRange.Text.TrimEnd('\r', '\n')); }

            PowerPoint.Shape target = keep ? first
                : c.ActiveSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, first.Left, first.Top, first.Width, first.Height);
            target.TextFrame.TextRange.Text = Cmd.TextOrPlaceholder(sb.ToString());
            target.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;

            foreach (var s in ordered) { if (keep && s == first) continue; s.Delete(); }
        }
    }

    public sealed class NormalViewCommand : ICommand
    { public string Name => "Normal View";
      public bool CanExecute(CommandContext c) => c.ActiveWindow != null;
      public void Execute(CommandContext c) => c.ActiveWindow.ViewType = PowerPoint.PpViewType.ppViewNormal; }
    public sealed class SlideSorterViewCommand : ICommand
    { public string Name => "Slide Sorter View";
      public bool CanExecute(CommandContext c) => c.ActiveWindow != null;
      public void Execute(CommandContext c) => c.ActiveWindow.ViewType = PowerPoint.PpViewType.ppViewSlideSorter; }

    public sealed class SelectSimilarCommand : ICommand
    {
        private readonly bool _fill, _outline, _type;
        public SelectSimilarCommand(bool fill, bool outline, bool type) { _fill = fill; _outline = outline; _type = type; }

        public string Name => "Select Similar";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection && c.ActiveSlide != null;

        public void Execute(CommandContext c)
        {
            var reference = c.SelectedShapes[1];
            var slide = c.ActiveSlide;
            var matches = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape cand in slide.Shapes)
                if (Matches(reference, cand)) matches.Add(cand);
            if (matches.Count > 0) Cmd.Select(matches);
            else Notifier.Info("No similar shapes on this slide.");
        }

        private bool Matches(PowerPoint.Shape r, PowerPoint.Shape s)
        {
            if (_type && !SameType(r, s)) return false;
            if (_fill && !SameFill(r, s)) return false;
            if (_outline && !SameOutline(r, s)) return false;
            return true;
        }

        private static bool SameType(PowerPoint.Shape a, PowerPoint.Shape b)
        {
            try
            {
                if (a.Type != b.Type) return false;
                if (a.Type == Office.MsoShapeType.msoAutoShape) return a.AutoShapeType == b.AutoShapeType;
                return true;
            }
            catch { return false; }
        }

        private static bool SameFill(PowerPoint.Shape a, PowerPoint.Shape b)
        {
            try
            {
                if (a.Fill.Visible != b.Fill.Visible) return false;
                if (a.Fill.Visible == Office.MsoTriState.msoTrue) return a.Fill.ForeColor.RGB == b.Fill.ForeColor.RGB;
                return true;
            }
            catch { return false; }
        }

        private static bool SameOutline(PowerPoint.Shape a, PowerPoint.Shape b)
        {
            try
            {
                if (a.Line.Visible != b.Line.Visible) return false;
                if (a.Line.Visible == Office.MsoTriState.msoTrue)
                    return a.Line.ForeColor.RGB == b.Line.ForeColor.RGB && Math.Abs(a.Line.Weight - b.Line.Weight) < 0.01f;
                return true;
            }
            catch { return false; }
        }
    }

    /// <summary>Copy selected shape text as a tab-separated grid (rows by Top) for Excel paste.</summary>
    public sealed class CopyForExcelCommand : ICommand
    {
        public string Name => "Copy for Excel";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;
        public void Execute(CommandContext c)
        {
            var shapes = ShapeHelpers.ToList(c.SelectedShapes);
            var rows = GridDetector.Rows(shapes);
            if (rows.Count == 0) { Notifier.Info("Nothing to copy."); return; }

            int maxCols = rows.Max(r => r.Count);
            var sb = new StringBuilder();
            foreach (var row in rows)
            {
                for (int col = 0; col < maxCols; col++)
                {
                    if (col > 0) sb.Append('\t');
                    sb.Append(col < row.Count ? CellText(row[col]) : "");
                }
                sb.AppendLine();
            }

            try
            {
                WinForms.Clipboard.SetText(sb.ToString().TrimEnd('\r', '\n'));
                Notifier.Info($"Copied {rows.Count} row(s) × {maxCols} column(s) as tab-separated text for Excel.");
            }
            catch (Exception ex) { Notifier.Error($"Could not copy: {ex.Message}"); }
        }

        private static string CellText(PowerPoint.Shape s)
        {
            if (!ShapeHelpers.HasTextFrame(s)) return "";
            string t;
            try { t = s.TextFrame.TextRange.Text ?? ""; }
            catch { return ""; }
            t = t.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
            return t.Trim();
        }
    }

    /// <summary>
    /// Paste a tab-separated column, row, or table into the selected boxes by position.
    /// Row and column come from the same 8 pt grid used by Copy for Excel.
    /// </summary>
    public sealed class PasteForExcelCommand : ICommand
    {
        public string Name => "Paste into boxes";
        public bool CanExecute(CommandContext c) => true;

        public void Execute(CommandContext c)
        {
            if (!CanFill(c))
            {
                Notifier.Info("Select the destination boxes and copy a column, row, or table first.");
                return;
            }
            Fill(c);
        }

        public static bool CanFill(CommandContext c)
        {
            if (c == null || c.HasTextSelection) return false;
            if (TextShapes(c).Count < 2) return false;
            return TryReadMultiCell(out _);
        }

        public static void Fill(CommandContext c)
        {
            if (!TryReadMultiCell(out string[,] cells)) return;
            var rows = GridDetector.Rows(TextShapes(c));
            int gridRows = cells.GetLength(0);
            int gridCols = cells.GetLength(1);
            for (int r = 0; r < rows.Count && r < gridRows; r++)
            {
                var row = rows[r];
                for (int col = 0; col < row.Count && col < gridCols; col++)
                {
                    string cell = cells[r, col];
                    if (string.IsNullOrEmpty(cell)) continue;
                    try { ReplaceTextKeepFormat(row[col], cell); }
                    catch { }
                }
            }
        }

        private static bool TryReadMultiCell(out string[,] cells)
        {
            cells = null;
            if (!TableCreatorClipboardReader.TryReadGrid(out cells, out _)) return false;
            if (cells == null) return false;
            return cells.GetLength(0) * cells.GetLength(1) >= 2;
        }

        private static List<PowerPoint.Shape> TextShapes(CommandContext c)
        {
            return ShapeHelpers.ToList(c.SelectedShapes).Where(ShapeHelpers.HasTextFrame).ToList();
        }

        private static void ReplaceTextKeepFormat(PowerPoint.Shape shape, string text)
        {
            var range = shape.TextFrame.TextRange;
            string fontName = null;
            float size = 0f;
            var bold = Office.MsoTriState.msoTriStateMixed;
            var italic = Office.MsoTriState.msoTriStateMixed;
            int color = 0;
            try
            {
                if (range.Length > 0)
                {
                    var font = range.Characters(1, 1).Font;
                    fontName = font.Name;
                    size = font.Size;
                    bold = font.Bold;
                    italic = font.Italic;
                    try { color = font.Color.RGB; } catch { color = 0; }
                }
            }
            catch { }

            range.Text = text ?? "";

            try
            {
                var font = range.Font;
                if (!string.IsNullOrEmpty(fontName)) font.Name = fontName;
                if (size > 0f) font.Size = size;
                if (bold != Office.MsoTriState.msoTriStateMixed) font.Bold = bold;
                if (italic != Office.MsoTriState.msoTriStateMixed) font.Italic = italic;
                if (color != 0) font.Color.RGB = color;
            }
            catch { }
        }
    }

    internal static class ShapeCornerRadiusHelper
    {
        public static bool TryRead(PowerPoint.Shape shape, out float adjustment)
        {
            adjustment = 0f;
            try
            {
                if (shape?.Adjustments == null || shape.Adjustments.Count < 1) return false;
                adjustment = shape.Adjustments[1];
                return true;
            }
            catch { return false; }
        }

        public static bool TryApply(PowerPoint.Shape shape, float adjustment)
        {
            try
            {
                if (shape?.Adjustments == null || shape.Adjustments.Count < 1) return false;
                shape.Adjustments[1] = adjustment;
                return true;
            }
            catch { return false; }
        }
    }

    public sealed class SwapObjectPositionCommand : ICommand
    {
        public string Name => "Swap object position";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count == 2;
        public void Execute(CommandContext c)
        {
            if (!c.HasShapeSelection)
            {
                Notifier.Info("Select two shapes.");
                return;
            }

            if (c.SelectedShapes.Count != 2)
            {
                Notifier.Error("Please select only two objects");
                return;
            }

            var ordered = c.OrderedSelection();
            if (ordered.Count < 2)
            {
                Notifier.Error("Please select only two objects");
                return;
            }

            var a = ordered[0];
            var b = ordered[1];
            float aLeft = a.Left, aTop = a.Top;
            a.Left = b.Left;
            a.Top = b.Top;
            b.Left = aLeft;
            b.Top = aTop;
        }
    }

    public sealed class SameCornerRadiusCommand : ICommand
    {
        public string Name => "Same corner radius";
        public bool CanExecute(CommandContext c) => c.HasShapeSelection && c.SelectedShapes.Count >= 2;
        public void Execute(CommandContext c)
        {
            var ordered = c.OrderedSelection();
            if (ordered.Count < 2)
            {
                Notifier.Info("Select at least two shapes.");
                return;
            }

            if (!ShapeCornerRadiusHelper.TryRead(ordered[0], out float radius))
            {
                Notifier.Info("The first selected shape has no adjustable corner radius.");
                return;
            }

            int applied = 0;
            var unsupported = new List<string>();
            for (int i = 1; i < ordered.Count; i++)
            {
                if (ShapeCornerRadiusHelper.TryApply(ordered[i], radius)) applied++;
                else
                {
                    try { unsupported.Add(ordered[i].Name); }
                    catch { unsupported.Add("shape"); }
                }
            }

            if (applied == 0)
            {
                Notifier.Info("None of the other selected shapes support corner radius. Select rounded rectangles or similar shapes.");
                return;
            }

            if (unsupported.Count > 0)
                Notifier.Info($"Applied radius to {applied} shape(s). Skipped {unsupported.Count} without radius support.");
        }
    }

    public sealed class MakeLineVerticalCommand : ICommand
    {
        public string Name => "Make vertical";
        public bool CanExecute(CommandContext c) => HasLineSelection(c);
        public void Execute(CommandContext c)
        {
            if (!TryGetSingleLine(c, out var line)) return;
            try
            {
                float left = line.Left;
                line.Width = 0f;
                line.Left = left;
            }
            catch (Exception ex) { Notifier.Error(ex.Message); }
        }

        internal static bool HasLineSelection(CommandContext c)
        {
            if (!c.HasShapeSelection || c.SelectedShapes.Count != 1) return false;
            try { return IsLineLike(c.SelectedShapes[1]); }
            catch { return false; }
        }

        private static bool TryGetSingleLine(CommandContext c, out PowerPoint.Shape line)
        {
            line = null;
            if (!TryGetLineFromSelection(c, out line))
            {
                Notifier.Info("Select a single line or connector.");
                return false;
            }
            if (c.SelectedShapes.Count != 1)
            {
                Notifier.Info("Select a single line or connector.");
                return false;
            }
            return true;
        }

        private static bool TryGetLineFromSelection(CommandContext c, out PowerPoint.Shape line)
        {
            line = null;
            if (!c.HasShapeSelection) return false;
            try
            {
                line = c.SelectedShapes[1];
                return IsLineLike(line);
            }
            catch { return false; }
        }

        internal static bool IsLineLike(PowerPoint.Shape s)
        {
            try
            {
                var t = s.Type;
                if (t == Office.MsoShapeType.msoLine) return true;
                // Connectors (not always exposed in interop enum)
                return (int)t == 13;
            }
            catch { return false; }
        }
    }

    public sealed class MakeLineHorizontalCommand : ICommand
    {
        public string Name => "Make horizontal";
        public bool CanExecute(CommandContext c) => MakeLineVerticalCommand.HasLineSelection(c);
        public void Execute(CommandContext c)
        {
            if (!MakeLineVerticalCommand.HasLineSelection(c) || c.SelectedShapes.Count != 1)
            {
                Notifier.Info("Select a single line or connector.");
                return;
            }

            PowerPoint.Shape line;
            try { line = c.SelectedShapes[1]; }
            catch { Notifier.Info("Select a single line or connector."); return; }

            if (!MakeLineVerticalCommand.IsLineLike(line))
            {
                Notifier.Info("Select a single line or connector.");
                return;
            }

            try
            {
                float top = line.Top;
                line.Height = 0f;
                line.Top = top;
            }
            catch (Exception ex) { Notifier.Error(ex.Message); }
        }
    }
}
