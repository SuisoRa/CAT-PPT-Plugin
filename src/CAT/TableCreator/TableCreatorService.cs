using System;
using Cat.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.TableCreator
{
    public static class TableCreatorService
    {
        public static bool CreateEmpty(PowerPoint.Slide slide, int rows, int cols, bool hasHeader)
        {
            if (!ValidateDimensions(rows, cols, out string err))
            {
                Notifier.Error(err);
                return false;
            }

            var cells = EmptyGrid(rows, cols, TableCreatorSpec.DummyCellText);
            return Finish(slide, cells, hasHeader, "Table created.", useThemeText1Body: true);
        }

        public static bool CreateFromClipboard(PowerPoint.Slide slide, bool hasHeader)
        {
            if (!TableCreatorClipboardReader.TryReadGrid(out string[,] cells, out string error))
            {
                Notifier.Error(error);
                return false;
            }
            return Finish(slide, cells, hasHeader, "Table created from clipboard.", useThemeText1Body: false);
        }

        public static bool ConvertNativeTable(PowerPoint.Slide slide, PowerPoint.Shape tableShape, bool hasHeader)
        {
            if (!TableCreatorNativeTableReader.TryRead(tableShape, out string[,] cells, out string error))
            {
                Notifier.Error(error);
                return false;
            }

            try { tableShape.Delete(); }
            catch (Exception ex)
            {
                Notifier.Error("Could not remove original table: " + ex.Message);
                return false;
            }

            return Finish(slide, cells, hasHeader, "Table converted to text boxes.", useThemeText1Body: false);
        }

        private static bool Finish(
            PowerPoint.Slide slide,
            string[,] cells,
            bool hasHeader,
            string successMessage,
            bool useThemeText1Body)
        {
            try
            {
                var group = TableCreatorBuilder.Build(slide, cells, hasHeader, useThemeText1Body: useThemeText1Body);
                if (group == null)
                {
                    Notifier.Error("Could not create the table.");
                    return false;
                }
                group.Select();
                Notifier.Info(successMessage);
                return true;
            }
            catch (Exception ex)
            {
                Notifier.Error("Table Creator failed: " + ex.Message);
                return false;
            }
        }

        public static bool ValidateDimensions(int rows, int cols, out string error)
        {
            error = null;
            if (rows < TableCreatorSpec.MinDimension || cols < TableCreatorSpec.MinDimension)
            {
                error = $"Use at least {TableCreatorSpec.MinDimension} row and column.";
                return false;
            }
            if (rows > TableCreatorSpec.MaxDimension || cols > TableCreatorSpec.MaxDimension)
            {
                error = $"Maximum size is {TableCreatorSpec.MaxDimension}×{TableCreatorSpec.MaxDimension}.";
                return false;
            }
            return true;
        }

        public static string[,] EmptyGrid(int rows, int cols, string fill = "")
        {
            var cells = new string[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    cells[r, c] = fill ?? "";
            return cells;
        }
    }
}
