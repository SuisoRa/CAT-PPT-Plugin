namespace Cat.TableCreator
{
    /// <summary>Applied after geometry is fixed. Phase 2 can extend this without changing layout/build code.</summary>
    public sealed class TableCreatorStyle
    {
        public float BodyFontSizePt { get; set; } = TableCreatorSpec.BodyFontSizePt;
        public float HeaderFontSizePt { get; set; } = TableCreatorSpec.HeaderFontSizePt;
        public bool HeaderBold { get; set; } = true;
        public float SeparatorWeightPt { get; set; } = TableCreatorSpec.SeparatorWeightPt;
        public int SeparatorColorRgb { get; set; } = TableCreatorSpec.SeparatorColorRgb;

        public static TableCreatorStyle Default => new TableCreatorStyle();
    }
}
