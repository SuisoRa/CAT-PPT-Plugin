namespace Cat.TextBlocks
{
    using Office = Microsoft.Office.Core;
    /// <summary>Geometry from assets/samples/Chevron toggle asset sample.pptx (OOXML adj / 100000).</summary>
    internal static class ChevronTextBlockSpec
    {
        public const string BlockKindTag = "CAT_TextBlock";
        public const string BlockKindChevron = "Chevron";
        public const string PreviewFileName = "chevron-preview.png";
        public const string StateTag = "CAT_ChevronState";
        public const string ArrowMarkerTag = "CAT_ChevronArrow";
        public const string StateHome = "Home";
        public const string StateChevron = "Chevron";

        /// <summary>Pentagon (msoShapePentagon = 51) — flat left, point right. Matches sample adj 22880.</summary>
        public const float PentagonPerfectAdj = 22880f / 100000f;

        /// <summary>Chevron — notched left, point right.</summary>
        public const float ChevronNotchedAdj = 23007f / 100000f;

        public const float DefaultWidthPt = 2409034f / 914400f * 72f;
        public const float DefaultHeightPt = 719847f / 914400f * 72f;

        public const float LabelFontSizePt = 16f;

        /// <summary>One PowerPoint arrow-key nudge when snap-to-grid is off (1 screen pixel).</summary>
        public const float ArrowNudgePt = 72f / 96f;

        /// <summary>Stage-2 label sits this far to the right of the stage-1 label (17 arrow presses).</summary>
        public const float ChevronLabelShiftPt = ArrowNudgePt * 17f;

        /// <summary>Block arrows pentagon (VBA 51), not regular pentagon (12) or a right arrow.</summary>
        public static Office.MsoAutoShapeType PentagonAutoShapeType =>
            (Office.MsoAutoShapeType)51; // msoShapePentagon
    }
}
