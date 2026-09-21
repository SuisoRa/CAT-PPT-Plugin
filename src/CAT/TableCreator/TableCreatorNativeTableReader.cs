using System;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.TableCreator
{
    public static class TableCreatorNativeTableReader
    {
        public static bool TryRead(PowerPoint.Shape shape, out string[,] cells, out string error)
        {
            cells = null;
            error = null;
            if (shape == null)
            {
                error = "Select a native PowerPoint table.";
                return false;
            }

            PowerPoint.Table table = null;
            try
            {
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                    table = shape.Table;
            }
            catch { }

            if (table == null)
            {
                error = "Selection is not a native PowerPoint table.";
                return false;
            }

            try
            {
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                if (rows < 1 || cols < 1)
                {
                    error = "Table has no cells.";
                    return false;
                }

                cells = new string[rows, cols];
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                        cells[r - 1, c - 1] = ReadCellText(table.Cell(r, c));
                }
                return true;
            }
            catch (Exception ex)
            {
                error = "Could not read table: " + ex.Message;
                return false;
            }
        }

        private static string ReadCellText(PowerPoint.Cell cell)
        {
            try
            {
                var tr = cell.Shape.TextFrame.TextRange;
                return (tr?.Text ?? "").Replace('\r', ' ').Replace('\n', ' ').Trim();
            }
            catch { return ""; }
        }
    }
}
