using Cat.Commands;
using Cat.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.Tools
{
    public sealed class ReduceTextMarginCommand : ICommand
    {
        private const float StepPt = 1f;

        public string Name => "Reduce margin";

        public bool CanExecute(CommandContext c) =>
            Cmd.EditableShapes(c).Exists(ShapeHelpers.HasTextFrame);

        public void Execute(CommandContext c)
        {
            Cmd.ForEachEditable(c, s =>
            {
                if (!ShapeHelpers.HasTextFrame(s)) return;
                try { Cmd.ReduceTextFrameMargins(s.TextFrame, StepPt); } catch { }
            });
        }
    }
}
