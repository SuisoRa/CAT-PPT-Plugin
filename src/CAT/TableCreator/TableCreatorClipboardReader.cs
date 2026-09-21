using System;
using System.Collections.Generic;
using System.Linq;
using WinForms = System.Windows.Forms;

namespace Cat.TableCreator
{
    public static class TableCreatorClipboardReader
    {
        public static bool TryReadGrid(out string[,] cells, out string error)
        {
            cells = null;
            error = null;

            if (!WinForms.Clipboard.ContainsText())
            {
                error = "Clipboard does not contain text. Copy an Excel range first.";
                return false;
            }

            string text;
            try { text = WinForms.Clipboard.GetText(); }
            catch (Exception ex)
            {
                error = "Could not read the clipboard: " + ex.Message;
                return false;
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                error = "Clipboard text is empty.";
                return false;
            }

            var rows = ParseRows(text);
            if (rows.Count == 0)
            {
                error = "No tabular data found on the clipboard.";
                return false;
            }

            int colCount = rows.Max(r => r.Count);
            int rowCount = rows.Count;
            cells = new string[rowCount, colCount];
            for (int r = 0; r < rowCount; r++)
            {
                var line = rows[r];
                for (int c = 0; c < colCount; c++)
                    cells[r, c] = c < line.Count ? line[c] : "";
            }
            return true;
        }

        private static List<List<string>> ParseRows(string text)
        {
            text = text.Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = text.Split(new[] { '\n' }, StringSplitOptions.None);
            var rows = new List<List<string>>();
            foreach (var line in lines)
            {
                if (rows.Count == 0 && string.IsNullOrWhiteSpace(line)) continue;
                if (rows.Count > 0 && rows[rows.Count - 1].Count == 1 && rows[rows.Count - 1][0] == "" && string.IsNullOrWhiteSpace(line))
                    continue;
                rows.Add(SplitColumns(line));
            }
            while (rows.Count > 0 && rows[rows.Count - 1].All(string.IsNullOrWhiteSpace))
                rows.RemoveAt(rows.Count - 1);
            return rows;
        }

        private static List<string> SplitColumns(string line)
        {
            if (line.IndexOf('\t') >= 0)
                return line.Split('\t').Select(NormalizeCell).ToList();
            return new List<string> { NormalizeCell(line) };
        }

        private static string NormalizeCell(string s)
        {
            if (s == null) return "";
            return s.Replace('\r', ' ').Replace('\n', ' ').Trim();
        }
    }
}
