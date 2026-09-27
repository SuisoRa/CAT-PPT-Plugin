using System;
using System.Collections.Generic;
using Cat.Commands;
using Cat.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.TextBlocks
{
    public static class ChevronTextBlockService
    {
        /// <summary>Updated on slide selection change so Ctrl+T is not stolen when no chevron block is selected.</summary>
        public static bool SelectionToggleAvailable { get; private set; }

        public static void RefreshSelectionToggleAvailability(PowerPoint.Selection selection)
        {
            SelectionToggleAvailable = false;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return;
            try
            {
                var range = selection.ShapeRange;
                if (range == null || range.Count < 1) return;
                foreach (PowerPoint.Shape s in range)
                {
                    if (IsChevronTextBlockSelection(s))
                    {
                        SelectionToggleAvailable = true;
                        return;
                    }
                }
            }
            catch { SelectionToggleAvailable = false; }
        }

        /// <summary>Toggles every selected CAT chevron block once. Other shapes are skipped.</summary>
        public static int ToggleAll(PowerPoint.ShapeRange range)
        {
            if (range == null) return 0;
            var seen = new HashSet<int>();
            int toggled = 0;
            foreach (PowerPoint.Shape s in range)
            {
                if (!TryResolveChevronBlock(s, out PowerPoint.Shape group, out PowerPoint.Shape arrow))
                    continue;
                int id;
                try { id = group.Id; }
                catch { continue; }
                if (!seen.Add(id)) continue;
                if (ToggleResolved(group, arrow)) toggled++;
            }
            return toggled;
        }

        public static bool InsertChevron(PowerPoint.Slide slide, string labelText = null)
        {
            if (slide == null) return false;
            labelText = string.IsNullOrWhiteSpace(labelText) ? Cmd.TextBoxPlaceholder : labelText.Trim();

            float w = ChevronTextBlockSpec.DefaultWidthPt;
            float h = ChevronTextBlockSpec.DefaultHeightPt;
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
                var arrow = slide.Shapes.AddShape(
                    ChevronTextBlockSpec.PentagonAutoShapeType,
                    left, top, w, h);
                ApplyArrowFill(arrow, slide);
                ApplyArrowOutline(arrow);
                ShapeCornerRadiusHelper.TryApply(arrow, ChevronTextBlockSpec.PentagonPerfectAdj);
                TagShape(arrow, ChevronTextBlockSpec.ArrowMarkerTag, "1");

                float labelLeft = left + w * 0.052f;
                float labelTop = top + h * 0.33f;
                float labelW = w * 0.835f;
                float labelH = h * 0.342f;
                var label = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    labelLeft, labelTop, labelW, labelH);
                StyleLabel(label, slide, labelText);

                var names = new object[] { arrow.Name, label.Name };
                var group = slide.Shapes.Range(names).Group();
                TagShape(group, ChevronTextBlockSpec.BlockKindTag, ChevronTextBlockSpec.BlockKindChevron);
                TagShape(group, ChevronTextBlockSpec.StateTag, ChevronTextBlockSpec.StateHome);
                group.Select();
                return true;
            }
            catch (Exception ex)
            {
                Notifier.Error("Could not insert chevron text block: " + ex.Message);
                return false;
            }
        }

        public static bool TryToggle(PowerPoint.Shape selected)
        {
            if (!TryResolveChevronBlock(selected, out PowerPoint.Shape group, out PowerPoint.Shape arrow))
            {
                Notifier.Info("Select a CAT chevron text block (insert from Text block), then press Ctrl+T.");
                return false;
            }
            return ToggleResolved(group, arrow);
        }

        private static bool ToggleResolved(PowerPoint.Shape group, PowerPoint.Shape arrow)
        {
            bool isChevron = ReadStateIsNotched(group);
            var slide = TryGetSlide(group);
            try
            {
                if (isChevron)
                {
                    EnsurePentagon(arrow);
                    ShapeCornerRadiusHelper.TryApply(arrow, ChevronTextBlockSpec.PentagonPerfectAdj);
                    if (slide != null) ApplyArrowFill(arrow, slide);
                    ShiftLabel(group, -ChevronTextBlockSpec.ChevronLabelShiftPt);
                    WriteState(group, ChevronTextBlockSpec.StateHome);
                }
                else
                {
                    EnsureChevron(arrow);
                    ShapeCornerRadiusHelper.TryApply(arrow, ChevronTextBlockSpec.ChevronNotchedAdj);
                    if (slide != null) ApplyArrowFill(arrow, slide);
                    ShiftLabel(group, ChevronTextBlockSpec.ChevronLabelShiftPt);
                    WriteState(group, ChevronTextBlockSpec.StateChevron);
                }
            }
            catch (Exception ex)
            {
                Notifier.Error("Chevron toggle failed: " + ex.Message);
                return false;
            }
            return true;
        }

        public static bool IsChevronTextBlockSelection(PowerPoint.Shape selected)
        {
            return TryResolveChevronBlock(selected, out _, out _);
        }

        private static void EnsurePentagon(PowerPoint.Shape arrow)
        {
            var pentagon = ChevronTextBlockSpec.PentagonAutoShapeType;
            if (arrow.AutoShapeType != pentagon)
                arrow.AutoShapeType = pentagon;
        }

        private static void EnsureChevron(PowerPoint.Shape arrow)
        {
            if (arrow.AutoShapeType != Office.MsoAutoShapeType.msoShapeChevron)
                arrow.AutoShapeType = Office.MsoAutoShapeType.msoShapeChevron;
        }

        private static void ShiftLabel(PowerPoint.Shape group, float dx)
        {
            var label = FindLabel(group);
            if (label == null) return;
            try { label.Left += dx; } catch { }
        }

        private static PowerPoint.Shape FindLabel(PowerPoint.Shape group)
        {
            try
            {
                var items = group.GroupItems;
                int n = items.Count;
                for (int i = 1; i <= n; i++)
                {
                    var s = items[i];
                    if (HasTag(s, ChevronTextBlockSpec.ArrowMarkerTag, "1")) continue;
                    if (ShapeHelpers.HasTextFrame(s)) return s;
                }
            }
            catch { }
            return null;
        }

        private static PowerPoint.Slide TryGetSlide(PowerPoint.Shape shape)
        {
            try
            {
                if (shape?.Parent is PowerPoint.Slide slide) return slide;
            }
            catch { }
            try
            {
                return ThisAddIn.Instance?.App?.ActiveWindow?.View?.Slide as PowerPoint.Slide;
            }
            catch { return null; }
        }

        private static bool ReadStateIsNotched(PowerPoint.Shape group)
        {
            try
            {
                string raw = group.Tags[ChevronTextBlockSpec.StateTag];
                if (string.Equals(raw, ChevronTextBlockSpec.StateChevron, StringComparison.OrdinalIgnoreCase))
                    return true;
                if (raw == "2") return true; // legacy numeric tag
            }
            catch { }
            return false;
        }

        private static void WriteState(PowerPoint.Shape group, string state)
        {
            try { group.Tags.Delete(ChevronTextBlockSpec.StateTag); } catch { }
            TagShape(group, ChevronTextBlockSpec.StateTag, state);
        }

        private static bool TryResolveChevronBlock(
            PowerPoint.Shape selected,
            out PowerPoint.Shape group,
            out PowerPoint.Shape arrow)
        {
            group = null;
            arrow = null;
            if (selected == null) return false;

            group = GetOwningGroup(selected);
            if (group == null) return false;

            if (!HasTag(group, ChevronTextBlockSpec.BlockKindTag, ChevronTextBlockSpec.BlockKindChevron))
                return false;

            arrow = FindArrowInGroup(group);
            return arrow != null;
        }

        private static PowerPoint.Shape GetOwningGroup(PowerPoint.Shape shape)
        {
            if (shape == null) return null;
            try
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                    return shape;
            }
            catch { }

            try
            {
                var parent = shape.ParentGroup;
                if (parent != null) return parent;
            }
            catch { }

            return null;
        }

        private static PowerPoint.Shape FindArrowInGroup(PowerPoint.Shape group)
        {
            try
            {
                var items = group.GroupItems;
                int n = items.Count;
                for (int i = 1; i <= n; i++)
                {
                    var s = items[i];
                    if (HasTag(s, ChevronTextBlockSpec.ArrowMarkerTag, "1"))
                        return s;
                }
                for (int i = 1; i <= n; i++)
                {
                    var s = items[i];
                    if (s.Type == Office.MsoShapeType.msoAutoShape)
                        return s;
                }
            }
            catch { }
            return null;
        }

        private static void StyleLabel(PowerPoint.Shape label, PowerPoint.Slide slide, string text)
        {
            label.Line.Visible = Office.MsoTriState.msoFalse;
            label.Fill.Visible = Office.MsoTriState.msoFalse;
            var tf = label.TextFrame;
            Cmd.ClearTextFrameMargins(tf);
            try { tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle; } catch { }
            ShapeHelpers.SetWordWrap(label, false);
            ShapeHelpers.SetAutoSize(label, PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText);

            var tr = tf.TextRange;
            tr.Text = text;
            tr.Font.Name = "Arial";
            tr.Font.Size = ChevronTextBlockSpec.LabelFontSizePt;
            tr.Font.Bold = Office.MsoTriState.msoTrue;
            try
            {
                tr.Font.Color.ObjectThemeColor = Office.MsoThemeColorIndex.msoThemeColorBackground1;
            }
            catch
            {
                try { tr.Font.Color.RGB = 0xFFFFFF; } catch { }
            }
            try { tr.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter; } catch { }
        }

        private static void ApplyArrowFill(PowerPoint.Shape arrow, PowerPoint.Slide slide)
        {
            arrow.Fill.Visible = Office.MsoTriState.msoTrue;
            arrow.Fill.Solid();
            try
            {
                var fc = arrow.Fill.ForeColor;
                fc.ObjectThemeColor = Office.MsoThemeColorIndex.msoThemeColorAccent4;
                fc.TintAndShade = -0.25f;
            }
            catch
            {
                try
                {
                    dynamic theme = slide.ThemeColorScheme;
                    arrow.Fill.ForeColor.RGB = theme[Office.MsoThemeColorSchemeIndex.msoThemeAccent4].RGB;
                }
                catch { }
            }
        }

        private static void ApplyArrowOutline(PowerPoint.Shape arrow)
        {
            try
            {
                arrow.Line.Visible = Office.MsoTriState.msoTrue;
                arrow.Line.Weight = 1f;
                arrow.Line.ForeColor.ObjectThemeColor = Office.MsoThemeColorIndex.msoThemeColorBackground1;
            }
            catch
            {
                try { arrow.Line.Visible = Office.MsoTriState.msoFalse; } catch { }
            }
        }

        private static void TagShape(PowerPoint.Shape shape, string name, string value)
        {
            try { shape.Tags.Add(name, value); } catch { }
        }

        private static bool HasTag(PowerPoint.Shape shape, string name, string expected)
        {
            try
            {
                string v = shape.Tags[name];
                return string.Equals(v, expected, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }
    }
}
