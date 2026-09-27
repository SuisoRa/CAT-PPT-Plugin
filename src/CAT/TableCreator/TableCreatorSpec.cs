namespace Cat.TableCreator
{
    /// <summary>Fixed layout constants. Styling extensions plug in via <see cref="TableCreatorStyle"/> later.</summary>
    public static class TableCreatorSpec
    {
        public const float BodyFontSizePt = 12f;
        public const float HeaderFontSizePt = 14f;
        public const float SeparatorWeightPt = 0.5f;
        public const int SeparatorColorRgb = 0x7F7F7F;
        public const float WhiteSpaceHeightFraction = 0.70f;
        public const float GapBelowTitlePt = 12f;
        public const float SlideMarginPt = 36f;
        public const float CellGapPt = 2f;
        public const string DummyCellText = "Text";
        public const int MinDimension = 1;
        public const int MaxDimension = 50;
    }
}
