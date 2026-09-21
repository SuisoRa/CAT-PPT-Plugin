using System;
using Cat.Core;
using Cat.Commands;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.TableCreator
{
    public sealed class TableCreatorLayout
    {
        public float TableLeft { get; private set; }
        public float TableTop { get; private set; }
        public float TableWidth { get; private set; }
        public float TableHeight { get; private set; }
        public float CellWidth { get; private set; }
        public float CellHeight { get; private set; }

        public static bool TryCompute(PowerPoint.Slide slide, int rows, int cols, out TableCreatorLayout layout)
        {
            layout = null;
            if (slide == null || rows < 1 || cols < 1) return false;

            float slideH = CommandContext.SlideHeight(slide);
            float slideW = CommandContext.SlideWidth(slide);

            float tableLeft;
            float tableWidth;
            float contentTop;

            if (SlideLayoutHelper.TryGetTitleLayout(slide, out float tLeft, out float tTop, out float tWidth, out float tHeight))
            {
                tableLeft = tLeft;
                tableWidth = tWidth;
                contentTop = tTop + tHeight + TableCreatorSpec.GapBelowTitlePt;
            }
            else
            {
                tableLeft = TableCreatorSpec.SlideMarginPt;
                tableWidth = Math.Max(1f, slideW - 2f * TableCreatorSpec.SlideMarginPt);
                contentTop = TableCreatorSpec.SlideMarginPt;
            }

            float contentBottom = slideH - TableCreatorSpec.SlideMarginPt;
            float available = Math.Max(1f, contentBottom - contentTop);
            float tableHeight = available * TableCreatorSpec.WhiteSpaceHeightFraction;

            float gap = TableCreatorSpec.CellGapPt;
            float hGaps = Math.Max(0, cols - 1) * gap;
            float vGaps = Math.Max(0, rows - 1) * gap;
            float cellW = (tableWidth - hGaps) / cols;
            float cellH = (tableHeight - vGaps) / rows;

            layout = new TableCreatorLayout
            {
                TableLeft = tableLeft,
                TableTop = contentTop,
                TableWidth = tableWidth,
                TableHeight = tableHeight,
                CellWidth = cellW,
                CellHeight = cellH,
                CellGap = gap
            };
            return true;
        }

        public float CellGap { get; private set; }

        public float CellLeft(int col) => TableLeft + col * (CellWidth + CellGap);
        public float CellTop(int row) => TableTop + row * (CellHeight + CellGap);
        public float RowSeparatorY(int rowIndex) =>
            CellTop(rowIndex) + CellHeight + CellGap * 0.5f;
    }
}
