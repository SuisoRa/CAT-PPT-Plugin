using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Cat.Commands;
using Cat.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.UI
{
    public static class IconLibraryWindow
    {
        private static Window _open;

        public static void Show()
        {
            UiThreadRunner.Run(ShowOnUiThread);
        }

        private static void ShowOnUiThread()
        {
            if (_open != null) { _open.Activate(); return; }

            var win = Dialogs.NewShellWindow("Icon library", 860, 620);

            var root = new Grid { Margin = new Thickness(20) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var title = new TextBlock { Text = "Icon library", Style = (Style)win.Resources["Title"] };
            Grid.SetRow(title, 0);
            root.Children.Add(title);

            var search = new TextBox
            {
                Height = 36,
                Padding = new Thickness(10, 6, 10, 6),
                VerticalContentAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 10)
            };
            search.Tag = "Search by tags…";
            search.GotFocus += (s, e) => { if (search.Text == (string)search.Tag) search.Text = ""; };
            search.LostFocus += (s, e) => { if (string.IsNullOrWhiteSpace(search.Text)) search.Text = (string)search.Tag; };
            search.Text = (string)search.Tag;
            Grid.SetRow(search, 1);
            root.Children.Add(search);

            var options = new StackPanel { Margin = new Thickness(0, 0, 0, 10) };
            var colour = new ComboBox { Width = 160, Height = 32, VerticalContentAlignment = VerticalAlignment.Center };
            var iconHex = OutlineHexBox();
            var encColour = new ComboBox { Width = 160, Height = 32, VerticalContentAlignment = VerticalAlignment.Center };
            var encHex = FillHexBox();
            PopulateAccentCombo(colour);
            PopulateAccentCombo(encColour);

            var iconRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            iconRow.Children.Add(Label("Outline colour"));
            iconRow.Children.Add(colour);
            iconRow.Children.Add(Label("Hex", true));
            iconRow.Children.Add(iconHex);
            options.Children.Add(iconRow);

            var encRow = new StackPanel { Orientation = Orientation.Horizontal };
            var enclosed = new CheckBox
            {
                Content = "Enclosed in circle",
                Margin = new Thickness(0, 0, 12, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            encRow.Children.Add(enclosed);
            encRow.Children.Add(Label("Fill colour"));
            encRow.Children.Add(encColour);
            encRow.Children.Add(Label("Hex", true));
            encRow.Children.Add(encHex);
            options.Children.Add(encRow);
            Grid.SetRow(options, 2);
            root.Children.Add(options);

            var body = new Grid();
            body.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(168) });
            body.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var families = new ListBox { Margin = new Thickness(0, 0, 12, 0) };
            Grid.SetColumn(families, 0);
            body.Children.Add(families);

            var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            var iconGrid = new WrapPanel { Margin = new Thickness(4) };
            scroll.Content = iconGrid;
            Grid.SetColumn(scroll, 1);
            body.Children.Add(scroll);

            Grid.SetRow(body, 3);
            root.Children.Add(body);

            var hint = new TextBlock
            {
                Text = "Single-click to insert as SVG (vector — Convert to Shape / Ungroup in PowerPoint). Hover for tags.",
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x6B, 0x6B)),
                Margin = new Thickness(0, 10, 0, 0),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(hint, 4);
            root.Children.Add(hint);

            win.Content = root;

            List<IconEntry> all = new List<IconEntry>();
            string selectedFamily = null;

            void ReloadCatalog()
            {
                all = IconCatalog.Load(IconLibraryPaths.Root);
                families.Items.Clear();
                families.Items.Add("All categories");
                foreach (var f in all.Select(e => e.Family).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x))
                    families.Items.Add(f);
                families.SelectedIndex = 0;
            }

            void RenderIcons()
            {
                iconGrid.Children.Clear();
                string q = search.Text ?? "";
                if (q == (string)search.Tag) q = "";

                var filtered = IconCatalog.Filter(all, q, selectedFamily);
                if (filtered.Count == 0)
                {
                    iconGrid.Children.Add(new TextBlock
                    {
                        Text = "No icons found.",
                        Margin = new Thickness(8),
                        Foreground = Brushes.Gray,
                        TextWrapping = TextWrapping.Wrap,
                        MaxWidth = 500
                    });
                    return;
                }

                var normalBg = Brushes.White;
                var normalBorder = new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8));
                var pressedBg = new SolidColorBrush(Color.FromRgb(0xD0, 0xE7, 0xFA));
                var pressedBorder = new SolidColorBrush(Color.FromRgb(0x0F, 0x6C, 0xBD));
                var hoverBorder = new SolidColorBrush(Color.FromRgb(0x0F, 0x6C, 0xBD));

                foreach (var entry in filtered)
                {
                    var dip = new TranslateTransform(0, 0);
                    var tile = new Border
                    {
                        Width = 92,
                        Height = 92,
                        Margin = new Thickness(4),
                        Background = normalBg,
                        BorderBrush = normalBorder,
                        BorderThickness = new Thickness(1),
                        CornerRadius = new CornerRadius(8),
                        Cursor = Cursors.Hand,
                        RenderTransformOrigin = new Point(0.5, 0.5),
                        RenderTransform = dip,
                        SnapsToDevicePixels = true
                    };

                    var preview = new Image
                    {
                        Width = 72,
                        Height = 72,
                        Stretch = Stretch.Uniform,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        IsHitTestVisible = false
                    };
                    tile.Child = preview;

                    // Queue preview on the dedicated STA worker; apply bitmap on this WPF dispatcher.
                    string pathForPreview = entry.Path;
                    SvgPreviewQueue.Enqueue(pathForPreview, 128, preview);

                    string tagLine = string.Join(" - ", entry.Tags);
                    tile.ToolTip = new ToolTip
                    {
                        Content = tagLine,
                        Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom
                    };
                    ToolTipService.SetInitialShowDelay(tile, 200);
                    ToolTipService.SetShowDuration(tile, 12000);

                    void ResetTile()
                    {
                        tile.Background = normalBg;
                        tile.BorderBrush = normalBorder;
                        tile.BorderThickness = new Thickness(1);
                        dip.Y = 0;
                    }

                    tile.MouseEnter += (s, e) =>
                    {
                        if (Mouse.LeftButton != MouseButtonState.Pressed)
                        {
                            tile.BorderBrush = hoverBorder;
                            tile.BorderThickness = new Thickness(2);
                        }
                    };
                    tile.MouseLeave += (s, e) => ResetTile();

                    tile.PreviewMouseLeftButtonDown += (s, e) =>
                    {
                        e.Handled = true;
                        tile.CaptureMouse();
                        tile.Background = pressedBg;
                        tile.BorderBrush = pressedBorder;
                        tile.BorderThickness = new Thickness(2);
                        dip.Y = 3;
                    };

                    tile.PreviewMouseLeftButtonUp += (s, e) =>
                    {
                        bool inside = tile.IsMouseCaptured && tile.IsMouseOver;
                        if (tile.IsMouseCaptured) tile.ReleaseMouseCapture();
                        ResetTile();
                        e.Handled = true;
                        if (inside)
                        {
                            string path = entry.Path;
                            bool enc = enclosed.IsChecked == true;
                            if (!TryResolveColor(colour, iconHex, out int iconRgb, out string iconErr))
                            {
                                Notifier.Info(iconErr);
                                return;
                            }
                            int enclosureRgb = iconRgb;
                            if (enc && !TryResolveColor(encColour, encHex, out enclosureRgb, out string encErr))
                            {
                                Notifier.Info(encErr);
                                return;
                            }
                            UiThreadRunner.Run(() => IconCatalog.Insert(path, iconRgb, enc, enclosureRgb));
                        }
                    };

                    iconGrid.Children.Add(tile);
                }
            }

            search.TextChanged += (s, e) => RenderIcons();
            families.SelectionChanged += (s, e) =>
            {
                var item = families.SelectedItem as string;
                selectedFamily = item == "All categories" ? null : item;
                RenderIcons();
            };

            win.Closing += (s, e) => _open = null;
            ReloadCatalog();
            RenderIcons();

            _open = win;
            win.Show();
        }

        private static TextBlock Label(string text, bool gapBefore = false) => new TextBlock
        {
            Text = text,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(gapBefore ? 12 : 0, 0, 8, 0)
        };

        private static TextBox HexBox() => new TextBox
        {
            Width = 92,
            Height = 32,
            Margin = new Thickness(0, 0, 0, 0),
            VerticalContentAlignment = VerticalAlignment.Center,
            ToolTip = "Optional #RGB or #RRGGBB. Empty uses the accent picker."
        };

        private static TextBox OutlineHexBox()
        {
            var box = HexBox();
            box.ToolTip = "Outline (icon line). Optional hex; empty uses accent.";
            return box;
        }

        private static TextBox FillHexBox()
        {
            var box = HexBox();
            box.ToolTip = "Fill (circle behind icon). Optional hex; empty uses accent.";
            return box;
        }

        private static bool TryResolveColor(ComboBox accents, TextBox hex, out int rgb, out string error)
        {
            rgb = 0;
            error = null;
            string raw = hex?.Text?.Trim() ?? "";
            if (raw.Length > 0)
            {
                if (!SvgIconHelper.TryParseHexRgb(raw, out rgb))
                {
                    error = "Enter a hex colour as #RGB or #RRGGBB, or leave the box empty to use the accent.";
                    return false;
                }
                return true;
            }

            PowerPoint.Slide slide = null;
            try { slide = ThisAddIn.Instance?.App?.ActiveWindow?.View?.Slide as PowerPoint.Slide; } catch { }
            var accent = Office.MsoThemeColorSchemeIndex.msoThemeAccent1;
            if (accents?.SelectedItem is ComboBoxItem item && item.Tag is Office.MsoThemeColorSchemeIndex idx)
                accent = idx;
            rgb = slide != null ? SvgIconHelper.AccentRgb(slide, accent) : DefaultAccentRgb(accent);
            return true;
        }

        private static void PopulateAccentCombo(ComboBox colour)
        {
            PowerPoint.Slide slide = null;
            try { slide = ThisAddIn.Instance?.App?.ActiveWindow?.View?.Slide as PowerPoint.Slide; } catch { }

            for (int i = 1; i <= 6; i++)
            {
                var idx = (Office.MsoThemeColorSchemeIndex)((int)Office.MsoThemeColorSchemeIndex.msoThemeAccent1 + i - 1);
                colour.Items.Add(MakeAccentItem($"Accent {i}", idx, slide));
            }
            colour.SelectedIndex = 0;
        }

        private static ComboBoxItem MakeAccentItem(string label, Office.MsoThemeColorSchemeIndex? accent, PowerPoint.Slide slide)
        {
            var row = new StackPanel { Orientation = Orientation.Horizontal };
            var swatch = new Border
            {
                Width = 18,
                Height = 18,
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(0, 0, 10, 0),
                VerticalAlignment = VerticalAlignment.Center,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC)),
                BorderThickness = new Thickness(1)
            };
            int rgb = slide != null
                ? SvgIconHelper.AccentRgb(slide, accent ?? Office.MsoThemeColorSchemeIndex.msoThemeAccent1)
                : DefaultAccentRgb(accent ?? Office.MsoThemeColorSchemeIndex.msoThemeAccent1);
            swatch.Background = new SolidColorBrush(Color.FromRgb(
                (byte)(rgb & 0xFF),
                (byte)((rgb >> 8) & 0xFF),
                (byte)((rgb >> 16) & 0xFF)));
            row.Children.Add(swatch);
            row.Children.Add(new TextBlock { Text = label, VerticalAlignment = VerticalAlignment.Center });
            return new ComboBoxItem
            {
                Content = row,
                Tag = accent.HasValue ? (object)accent.Value : null,
                VerticalContentAlignment = VerticalAlignment.Center
            };
        }

        private static int DefaultAccentRgb(Office.MsoThemeColorSchemeIndex accent)
        {
            int i = (int)accent - (int)Office.MsoThemeColorSchemeIndex.msoThemeAccent1;
            // BGR values matching typical theme accents
            int[] defaults = { 0xC47244, 0x317DED, 0xA5A5A5, 0x00C0FF, 0xD59B5B, 0x47AD70 };
            if (i >= 0 && i < defaults.Length) return defaults[i];
            return 0xC47244;
        }
    }

    internal sealed class IconEntry
    {
        public string Family;
        public string Path;
        public IList<string> Tags;
    }

    internal static class IconCatalog
    {
        public static List<IconEntry> Load(string root)
        {
            var list = new List<IconEntry>();
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) return list;

            foreach (var familyDir in Directory.EnumerateDirectories(root))
            {
                string family = Path.GetFileName(familyDir);
                foreach (var file in Directory.EnumerateFiles(familyDir, "*.svg", SearchOption.TopDirectoryOnly))
                {
                    string name = Path.GetFileNameWithoutExtension(file);
                    var tags = name.Split(new[] { '-', '_', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(t => t.Trim())
                        .Where(t => t.Length > 0)
                        .ToList();
                    if (tags.Count == 0) tags.Add(name);
                    list.Add(new IconEntry { Family = family, Path = file, Tags = tags });
                }
            }

            return list.OrderBy(e => e.Family, StringComparer.OrdinalIgnoreCase)
                .ThenBy(e => Path.GetFileName(e.Path), StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public static List<IconEntry> Filter(IEnumerable<IconEntry> source, string query, string family)
        {
            var tokens = (query ?? "").Split(new[] { ' ', ',', ';', '-' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => t.Length > 0)
                .ToArray();

            IEnumerable<IconEntry> q = source;
            if (!string.IsNullOrEmpty(family))
                q = q.Where(e => string.Equals(e.Family, family, StringComparison.OrdinalIgnoreCase));

            if (tokens.Length == 0) return q.ToList();

            return q.Where(e =>
            {
                var hay = string.Join("-", e.Tags).ToLowerInvariant();
                foreach (var tok in tokens)
                    if (hay.IndexOf(tok.ToLowerInvariant(), StringComparison.Ordinal) < 0) return false;
                return true;
            }).ToList();
        }

        /// <summary>
        /// Insert as native SVG (vector). PowerPoint keeps it editable — Convert to Shape / Ungroup.
        /// Falls back to PNG only if SVG insert fails on the host.
        /// </summary>
        public static void Insert(string svgPath, int iconRgb, bool enclosed, int enclosureRgb)
        {
            string tempSvg = null;
            string tempPng = null;
            try
            {
                var app = ThisAddIn.Instance?.App;
                var slide = app?.ActiveWindow?.View?.Slide as PowerPoint.Slide;
                if (slide == null) { Notifier.Info("Open a slide first."); return; }

                string pathToUse = svgPath;
                tempSvg = SvgIconHelper.WriteOutlineTintedTempFile(svgPath, SvgIconHelper.RgbToHex(iconRgb));
                pathToUse = tempSvg;

                float sw = CommandContext.SlideWidth(slide);
                float sh = CommandContext.SlideHeight(slide);
                const float outer = 72f;
                float left = sw / 2f - outer / 2f;
                float top = sh / 2f - outer / 2f;

                PowerPoint.Shape iconShape;
                try
                {
                    iconShape = AddPictureEmbedded(slide, pathToUse, left, top, outer, outer);
                }
                catch
                {
                    // Older hosts / blocked SVG → raster fallback
                    tempPng = SvgIconHelper.CreatePngFile(pathToUse, null, 512);
                    iconShape = AddPictureEmbedded(slide, tempPng, left, top, outer, outer);
                }

                iconShape.LockAspectRatio = Office.MsoTriState.msoTrue;

                if (!enclosed)
                {
                    iconShape.Select();
                    return;
                }

                // Circle behind icon; group so user can Ungroup later and keep both pieces.
                var circle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, outer, outer);
                circle.Fill.Visible = Office.MsoTriState.msoTrue;
                circle.Fill.Solid();
                circle.Fill.ForeColor.RGB = enclosureRgb;
                circle.Line.Visible = Office.MsoTriState.msoFalse;
                try { circle.Line.Weight = 0f; } catch { }
                try { circle.ZOrder(Office.MsoZOrderCmd.msoSendToBack); } catch { }

                const float inset = 14f;
                float inner = outer - inset * 2f;
                iconShape.Left = left + inset;
                iconShape.Top = top + inset;
                iconShape.Width = inner;
                iconShape.Height = inner;

                try
                {
                    var names = new object[] { circle.Name, iconShape.Name };
                    slide.Shapes.Range(names).Group().Select();
                }
                catch
                {
                    iconShape.Select();
                }
            }
            catch (Exception ex) { Notifier.Error($"Could not insert icon: {ex.Message}"); }
            finally
            {
                // Delay delete slightly so PowerPoint finishes embedding the file.
                if (tempSvg != null) TryDeleteLater(tempSvg);
                if (tempPng != null) TryDeleteLater(tempPng);
            }
        }

        private static PowerPoint.Shape AddPictureEmbedded(
            PowerPoint.Slide slide, string path, float left, float top, float width, float height)
        {
            return slide.Shapes.AddPicture(
                path,
                Office.MsoTriState.msoFalse,  // do not link — embed
                Office.MsoTriState.msoTrue,   // save with document
                left, top, width, height);
        }

        private static void TryDeleteLater(string path)
        {
            try
            {
                var t = new System.Threading.Thread(() =>
                {
                    System.Threading.Thread.Sleep(1500);
                    try { File.Delete(path); } catch { }
                });
                t.IsBackground = true;
                t.Start();
            }
            catch { }
        }
    }
}
