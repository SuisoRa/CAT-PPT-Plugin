namespace Cat.TextBlocks
{
    /// <summary>Number + body stack from Samples for cursor.pptx (OOXML EMU → pt).</summary>
    internal static class ActionItemTextBlockSpec
    {
        public const string BlockKindActionItem = "ActionItem";
        public const string PreviewFileName = "action-item-preview.png";

        public const string DefaultNumberText = "1";
        public const string DefaultBodyText = "Action item here";

        public const float NumberFontSizePt = 36f;
        public const float BodyFontSizePt = 16f;

        private const float EmuPerPt = 914400f / 72f;

        public static float WidthPt => 1127760f / EmuPerPt;
        public static float NumberHeightPt => 553998f / EmuPerPt;
        public static float BodyHeightPt => 492443f / EmuPerPt;
        public static float BodyTopOffsetPt => (2270443f - 1683703f) / EmuPerPt;
        public static float TotalHeightPt => 1079183f / EmuPerPt;
    }
}
