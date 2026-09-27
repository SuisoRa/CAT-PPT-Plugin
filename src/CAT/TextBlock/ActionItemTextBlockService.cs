using System;
using Cat.Commands;
using Cat.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.TextBlocks
{
    public static class ActionItemTextBlockService
    {
        public static bool Insert(PowerPoint.Slide slide, string numberText = null, string bodyText = null)
        {
            if (slide == null) return false;
            numberText = string.IsNullOrWhiteSpace(numberText)
                ? ActionItemTextBlockSpec.DefaultNumberText
                : numberText.Trim();
            bodyText = string.IsNullOrWhiteSpace(bodyText)
                ? ActionItemTextBlockSpec.DefaultBodyText
                : bodyText.Trim();

            float w = ActionItemTextBlockSpec.WidthPt;
            float h = ActionItemTextBlockSpec.TotalHeightPt;
            float sw = CommandContext.SlideWidth(slide);
            float sh = CommandContext.SlideHeight(slide);
            float left = (sw - w) / 2f;
            float top = (sh - h) / 2f;

            if (SlideLayoutHelper.TryGetTitleLayout(slide, out float tLeft, out float tTop, out float tWidth, out float tHeight))
            {
                left = tLeft;
                top = tTop + tHeight + 12f;
                if (tWidth > 0 && tWidth < sw) w = Math.Min(tWidth, w);
            }

            try
            {
                var number = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    left, top, w, ActionItemTextBlockSpec.NumberHeightPt);
                StyleNumber(number, slide, numberText);

                float bodyTop = top + ActionItemTextBlockSpec.BodyTopOffsetPt;
                var body = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    left, bodyTop, w, ActionItemTextBlockSpec.BodyHeightPt);
                StyleBody(body, slide, bodyText);

                var names = new object[] { number.Name, body.Name };
                var group = slide.Shapes.Range(names).Group();
                TagShape(group, ChevronTextBlockSpec.BlockKindTag, ActionItemTextBlockSpec.BlockKindActionItem);
                group.Select();
                return true;
            }
            catch (Exception ex)
            {
                Notifier.Error("Could not insert action item text block: " + ex.Message);
                return false;
            }
        }

        private static void StyleNumber(PowerPoint.Shape tb, PowerPoint.Slide slide, string text)
        {
            tb.Line.Visible = Office.MsoTriState.msoFalse;
            tb.Fill.Visible = Office.MsoTriState.msoFalse;
            var tf = tb.TextFrame;
            Cmd.ClearTextFrameMargins(tf);
            try { tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop; } catch { }
            ShapeHelpers.SetWordWrap(tb, false);
            ShapeHelpers.SetAutoSize(tb, PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText);

            var tr = tf.TextRange;
            tr.Text = text;
            tr.Font.Name = TextBlockThemeHelper.TryThemeMajorFont(slide);
            tr.Font.Size = ActionItemTextBlockSpec.NumberFontSizePt;
            tr.Font.Bold = Office.MsoTriState.msoTrue;
            TextBlockThemeHelper.ApplyAccent1Color(tr, slide);
        }

        private static void StyleBody(PowerPoint.Shape tb, PowerPoint.Slide slide, string text)
        {
            tb.Line.Visible = Office.MsoTriState.msoFalse;
            tb.Fill.Visible = Office.MsoTriState.msoFalse;
            var tf = tb.TextFrame;
            Cmd.ClearTextFrameMargins(tf);
            try { tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop; } catch { }
            ShapeHelpers.SetWordWrap(tb, true);
            ShapeHelpers.SetAutoSize(tb, PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText);

            var tr = tf.TextRange;
            tr.Text = text;
            tr.Font.Name = TextBlockThemeHelper.TryThemeMinorFont(slide);
            tr.Font.Size = ActionItemTextBlockSpec.BodyFontSizePt;
            tr.Font.Bold = Office.MsoTriState.msoFalse;
            TextBlockThemeHelper.ApplyText1Color(tr, slide);
        }

        private static void TagShape(PowerPoint.Shape shape, string name, string value)
        {
            try { shape.Tags.Add(name, value); } catch { }
        }
    }
}
