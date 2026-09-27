using Cat.Commands;
using Cat.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.TextBlocks
{
    public sealed class ChevronToggleCommand : ICommand
    {
        public string Name => "Toggle chevron text block";
        public bool CanExecute(CommandContext c)
        {
            try
            {
                var range = c.SelectedShapes;
                if (range == null) return false;
                foreach (PowerPoint.Shape s in range)
                    if (ChevronTextBlockService.IsChevronTextBlockSelection(s)) return true;
            }
            catch { }
            return false;
        }

        public void Execute(CommandContext c)
        {
            try
            {
                int n = ChevronTextBlockService.ToggleAll(c.SelectedShapes);
                if (n == 0)
                    Notifier.Info("Select one or more CAT chevron text blocks, then press Ctrl+T.");
            }
            catch { }
        }
    }
}
