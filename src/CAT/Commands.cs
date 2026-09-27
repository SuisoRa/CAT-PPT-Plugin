using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WinForms = System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Cat.Core;
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
                var s = Selection;
                try { return s != null && s.Type == PowerPoint.PpSelectionType.ppSelectionShapes && s.ShapeRange.Count > 0; }
                catch { return false; }
            }
        }

        public PowerPoint.ShapeRange SelectedShapes
        { get { try { return Selection?.ShapeRange; } catch { return null; } } }

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
        { foreach (var s in ShapeHelpers.ToList(ctx.SelectedShapes)) action(s); }

        /// <summary>Shape selection, or the text box that contains the caret (text-edit / soft selection).</summary>
        public static List<PowerPoint.Shape> EditableShapes(CommandContext ctx)
        {
            var list = new List<PowerPoint.Shape>();
            if (ctx == null) return list;
            if (ctx.HasShapeSelection)
            {
                list.AddRange(ShapeHelpers.ToList(ctx.SelectedShapes));
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
            if (tr == null) return;
            try
            {
                int n = tr.Paragraphs().Count;
                for (int i = 1; i <= n; i++)
                {
                    var pf = tr.Paragraphs(i).ParagraphFormat;
                    pf.LineRuleWithin = Office.MsoTriState.msoTrue;
                    pf.SpaceWithin = 1f;
                }
            }
            catch { }
        }

        /// <summary>One pass: body font, plain weight, Text 1, no highlight; keeps each fragment's point size.</summary>
        public static void FixTextRangeFormatting(PowerPoint.TextRange tr, PowerPoint.Slide slide)
        {
            if (tr == null || slide == null) return;
            string bodyFont = TextBlockThemeHelper.TryThemeMinorFont(slide);

            int len = 0;
            try { len = tr.Length; } catch { }

            bool anyRun = false;
            for (int i = 1; i <= 500; i++)
            {
                try
                {
                    var run = tr.Runs(i);
                    if (run == null || run.Length < 1) break;
                    NormalizeTextRun(run, bodyFont, slide);
                    anyRun = true;
                }
                catch { break; }
            }

            if (!anyRun && len > 0)
                NormalizeTextRun(tr, bodyFont, slide);

            if (len > 0 && len <= 10000)
            {
                for (int i = 1; i <= len; i++)
                {
                    try
                    {
                        var ch = tr.Characters(i, 1);
                        NormalizeTextRun(ch, bodyFont, slide);
                    }
                    catch { }
                }
            }
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
                Cmd.SetSingleLineSpacing(tr);
                FixTextRangeFormatting(tr, slide);
                ClearTextHighlights(shape);
            }
            catch { }
        }

        private static void NormalizeTextRun(PowerPoint.TextRange run, string bodyFont, PowerPoint.Slide slide)
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
                f.Name = bodyFont;
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
                try { df.Kernings = Office.MsoTriState.msoFalse; } catch { }
                try { df.Highlight = Office.MsoTriState.msoFalse; } catch { }
            }
            catch { }

            TextBlockThemeHelper.ApplyText1Color(run, slide);
            ClearRunHighlight(run);
        }

        private static void ClearRunHighlight(PowerPoint.TextRange run)
        {
            if (run == null) return;
            try
            {
                dynamic f = run.Font;
                try { f.Highlight = Office.MsoTriState.msoFalse; } catch { }
                try { f.Fill.Visible = Office.MsoTriState.msoFalse; } catch { }
            }
            catch { }
        }

        private static void ClearTextHighlights(PowerPoint.Shape shape)
        {
            try
            {
                dynamic tr2 = shape.TextFrame2.TextRange;
                int len = (int)tr2.Length;
                if (len < 1) return;
                int max = Math.Min(len, 10000);
                for (int i = 1; i <= max; i++)
                {
                    try
                    {
                        dynamic ch = tr2.Characters[i, 1];
                        dynamic font = ch.Font;
                        try { font.HighlightColor.ObjectThemeColor = Office.MsoThemeColorIndex.msoNotThemeColor; } catch { }
                        try { font.Fill.Visible = Office.MsoTriState.msoFalse; } catch { }
                    }
                    catch { }
                }
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
        }
        private static string Key(PowerPoint.Shape ph)
        { try { var pf = ph.PlaceholderFormat; return $"{(int)pf.Type}:{pf.ContainedType}"; } catch { return ph.Name; } }
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
        public bool CanExecute(CommandContext c) => c.HasShapeSelection;
        public void Execute(CommandContext c)
        {
            var r = Dialogs.ShowSplitJoin(c.SelectedShapes.Count >= 2);
            if (!r.Ok) return;
            if (r.Mode == SplitJoinMode.Split) DoSplit(c, r.Level, r.KeepFormatting);
            else DoJoin(c, r.KeepFormatting);
        }

        private static void DoSplit(CommandContext c, int level, bool keep)
        {
            var source = c.SelectedShapes[1];
            if (!ShapeHelpers.HasTextFrame(source)) return;
            var tr = source.TextFrame.TextRange;
            int paraCount = tr.Paragraphs().Count;
            if (paraCount <= 1) return;

            var chunks = new List<List<string>>(); var cur = new List<string>();
            for (int i = 1; i <= paraCount; i++)
            {
                var p = tr.Paragraphs(i);
                bool boundary = (level == -1) || (p.IndentLevel <= level);
                if (boundary && cur.Count > 0) { chunks.Add(cur); cur = new List<string>(); }
                cur.Add(p.Text.TrimEnd('\r', '\n'));
            }
            if (cur.Count > 0) chunks.Add(cur);
            if (chunks.Count <= 1) return;

            var slide = c.ActiveSlide; float top = source.Top;
            foreach (var chunk in chunks)
            {
                string text = string.Join("\r", chunk);
                PowerPoint.Shape box = keep ? source.Duplicate()[1]
                    : slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, source.Left, top, source.Width, 20);
                box.Left = source.Left; box.Top = top;
                box.TextFrame.TextRange.Text = Cmd.TextOrPlaceholder(text);
                box.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                top = box.Top + box.Height + 6f;
            }
            source.Delete();
        }

        private static void DoJoin(CommandContext c, bool keep)
        {
            var ordered = c.OrderedSelection().Where(ShapeHelpers.HasTextFrame).ToList();
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
