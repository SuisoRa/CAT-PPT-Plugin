using Cat.Commands;
using Cat.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.Tools
{
    /// <summary>
    /// Cycles font colour Text 1 → Text 2 → white. Not on the CAT tab;
    /// wire control id catFontAlt from Tools/HomeTab.ribbon.xml onto Home when ready.
    /// </summary>
    public sealed class FontColorAlternatorCommand : ICommand
    {
        private const string StepTag = "CAT_FontColorStep";

        public string Name => "Alternate font colour";

        public bool CanExecute(CommandContext c) =>
            Cmd.EditableShapes(c).Exists(ShapeHelpers.HasTextFrame) || Cmd.TargetText(c) != null;

        public void Execute(CommandContext c)
        {
            var shapes = Cmd.EditableShapes(c);
            bool any = false;
            foreach (var s in shapes)
            {
                if (!ShapeHelpers.HasTextFrame(s)) continue;
                try
                {
                    int next = NextStep(ReadStep(s));
                    Apply(s.TextFrame.TextRange, next);
                    WriteStep(s, next);
                    any = true;
                }
                catch { }
            }

            if (any) return;

            var tr = Cmd.TargetText(c);
            if (tr == null) return;
            try { Apply(tr, 0); } catch { }
        }

        private static int ReadStep(PowerPoint.Shape shape)
        {
            try
            {
                string raw = shape.Tags[StepTag];
                if (int.TryParse(raw, out int n) && n >= 0 && n <= 2) return n;
            }
            catch { }
            return InferStep(shape);
        }

        private static int InferStep(PowerPoint.Shape shape)
        {
            try
            {
                var fc = shape.TextFrame.TextRange.Font.Color;
                var theme = fc.ObjectThemeColor;
                if (theme == Office.MsoThemeColorIndex.msoThemeColorText1) return 0;
                if (theme == Office.MsoThemeColorIndex.msoThemeColorText2) return 1;
                if (fc.RGB == 0xFFFFFF) return 2;
            }
            catch { }
            return 2;
        }

        private static int NextStep(int current) => (current + 1) % 3;

        private static void WriteStep(PowerPoint.Shape shape, int step)
        {
            try { shape.Tags.Delete(StepTag); } catch { }
            try { shape.Tags.Add(StepTag, step.ToString()); } catch { }
        }

        private static void Apply(PowerPoint.TextRange tr, int step)
        {
            var fc = tr.Font.Color;
            if (step == 2)
            {
                fc.RGB = 0xFFFFFF;
                return;
            }
            fc.ObjectThemeColor = step == 0
                ? Office.MsoThemeColorIndex.msoThemeColorText1
                : Office.MsoThemeColorIndex.msoThemeColorText2;
            fc.TintAndShade = 0f;
        }
    }
}
