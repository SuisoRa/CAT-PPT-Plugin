using System;
using System.Collections.Generic;
using System.Globalization;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Media.Imaging;
using WinForms = System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.UI
{
    internal static class SvgIconHelper
    {
        private static readonly Dictionary<string, byte[]> PngCache =
            new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

        public static string BuildPreviewHtml(string svgMarkup)
        {
            svgMarkup = svgMarkup.Replace("\"", "'");
            return $@"<!DOCTYPE html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'/>
<style>
html,body{{margin:0;padding:0;width:100%;height:100%;overflow:hidden;background:#FFFFFF}}
body{{display:flex;align-items:center;justify-content:center}}
svg{{max-width:88%;max-height:88%;width:auto;height:auto}}
</style></head>
<body>{svgMarkup}</body></html>";
        }

        public static string WriteTintedTempFile(string svgPath, string tintHex)
        {
            string svg = ApplyTint(File.ReadAllText(svgPath), tintHex);
            string temp = Path.Combine(Path.GetTempPath(), "cat-icon-" + Guid.NewGuid().ToString("N") + ".svg");
            File.WriteAllText(temp, svg, new UTF8Encoding(false));
            return temp;
        }

        /// <summary>Stroke-only tint for icon library inserts (no fill).</summary>
        public static string WriteOutlineTintedTempFile(string svgPath, string tintHex)
        {
            string svg = ApplyOutlineTint(File.ReadAllText(svgPath), tintHex);
            string temp = Path.Combine(Path.GetTempPath(), "cat-icon-" + Guid.NewGuid().ToString("N") + ".svg");
            File.WriteAllText(temp, svg, new UTF8Encoding(false));
            return temp;
        }

        public static string ApplyOutlineTint(string svg, string hex)
        {
            if (string.IsNullOrEmpty(hex)) return svg;
            hex = NormalizeHex(hex);

            const string marker = "<!--cat-outline-tint-->";
            if (!svg.Contains(marker))
            {
                int svgTag = svg.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
                if (svgTag >= 0)
                {
                    int close = svg.IndexOf('>', svgTag);
                    if (close > 0)
                    {
                        svg = svg.Insert(close + 1,
                            marker + $"<style type='text/css'>*{{fill:none!important;stroke:{hex}!important;color:{hex}!important}}</style>");
                    }
                }
            }

            svg = Regex.Replace(svg,
                @"(fill\s*=\s*[""'])(?!none|transparent|url)([^""']*)([""'])",
                "$1none$3",
                RegexOptions.IgnoreCase);
            svg = Regex.Replace(svg,
                @"(stroke\s*=\s*[""'])(?!none|transparent|url)([^""']*)([""'])",
                "$1" + hex + "$3",
                RegexOptions.IgnoreCase);
            svg = Regex.Replace(svg,
                @"(stroke\s*:\s*)(?!none|transparent)(#[0-9A-Fa-f]{3,8}|rgb\([^)]+\)|[a-zA-Z]+)(\s*[;""'])",
                "$1" + hex + "$3",
                RegexOptions.IgnoreCase);
            svg = Regex.Replace(svg,
                @"(fill\s*:\s*)(?!none|transparent)(#[0-9A-Fa-f]{3,8}|rgb\([^)]+\)|[a-zA-Z]+)(\s*[;""'])",
                "$1none$3",
                RegexOptions.IgnoreCase);

            return svg;
        }

        public static string ApplyTint(string svg, string hex)
        {
            if (string.IsNullOrEmpty(hex)) return svg;
            hex = NormalizeHex(hex);

            const string marker = "<!--cat-tint-->";
            if (!svg.Contains(marker))
            {
                int svgTag = svg.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
                if (svgTag >= 0)
                {
                    int close = svg.IndexOf('>', svgTag);
                    if (close > 0)
                    {
                        svg = svg.Insert(close + 1,
                            marker + $"<style type='text/css'>*{{fill:{hex}!important;stroke:{hex}!important;color:{hex}!important}}</style>");
                    }
                }
            }

            svg = Regex.Replace(svg,
                @"(fill\s*=\s*[""'])(?!none|transparent|url)([^""']*)([""'])",
                "$1" + hex + "$3",
                RegexOptions.IgnoreCase);
            svg = Regex.Replace(svg,
                @"(stroke\s*=\s*[""'])(?!none|transparent|url)([^""']*)([""'])",
                "$1" + hex + "$3",
                RegexOptions.IgnoreCase);
            svg = Regex.Replace(svg,
                @"(stroke\s*:\s*)(?!none|transparent)(#[0-9A-Fa-f]{3,8}|rgb\([^)]+\)|[a-zA-Z]+)(\s*[;""'])",
                "$1" + hex + "$3",
                RegexOptions.IgnoreCase);
            svg = Regex.Replace(svg,
                @"(fill\s*:\s*)(?!none|transparent)(#[0-9A-Fa-f]{3,8}|rgb\([^)]+\)|[a-zA-Z]+)(\s*[;""'])",
                "$1" + hex + "$3",
                RegexOptions.IgnoreCase);

            return svg;
        }

        private static string NormalizeHex(string hex)
        {
            hex = hex.Trim().TrimStart('#');
            if (hex.Length == 6) return "#" + hex.ToUpperInvariant();
            if (hex.Length == 3)
                return "#" + hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
            return "#" + hex;
        }

        public static int AccentRgb(PowerPoint.Slide slide, Office.MsoThemeColorSchemeIndex accent)
        {
            try
            {
                dynamic themeColors = slide.ThemeColorScheme;
                return themeColors[accent].RGB;
            }
            catch { return 0xC47244; }
        }

        /// <summary>Parses #RGB or #RRGGBB into a PowerPoint BGR colour. Empty input is not a colour.</summary>
        public static bool TryParseHexRgb(string text, out int bgr)
        {
            bgr = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            text = text.Trim();
            if (text.StartsWith("#", StringComparison.Ordinal)) text = text.Substring(1);
            if (text.Length == 3)
                text = string.Concat(text[0], text[0], text[1], text[1], text[2], text[2]);
            if (text.Length != 6) return false;
            if (!int.TryParse(text, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int rgb)) return false;
            int r = (rgb >> 16) & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = rgb & 0xFF;
            bgr = r | (g << 8) | (b << 16);
            return true;
        }

        public static string RgbToHex(int rgb)
        {
            byte r = (byte)(rgb & 0xFF);
            byte g = (byte)((rgb >> 8) & 0xFF);
            byte b = (byte)((rgb >> 16) & 0xFF);
            return $"#{r:X2}{g:X2}{b:X2}";
        }

        /// <summary>PNG bytes for a tile preview. Uses cache; may spawn STA thread if needed.</summary>
        public static byte[] GetPreviewPngBytes(string svgPath, int size = 128) =>
            RenderPreviewPngBytes(svgPath, size);

        /// <summary>Called from the dedicated SVG preview worker or STA callers.</summary>
        internal static byte[] RenderPreviewPngBytes(string svgPath, int size = 128)
        {
            string key = svgPath + "|" + size;
            lock (PngCache)
            {
                if (PngCache.TryGetValue(key, out var cached)) return cached;
            }

            byte[] png = null;
            try
            {
                string path = CreatePngFile(svgPath, null, size);
                try { png = File.ReadAllBytes(path); }
                finally { try { File.Delete(path); } catch { } }
            }
            catch { png = null; }

            if (png == null || png.Length < 32 || IsBlankPng(png))
                png = LetterFallbackPng(Path.GetFileNameWithoutExtension(svgPath), size);

            lock (PngCache)
            {
                PngCache[key] = png;
            }
            return png;
        }

        /// <summary>Must run on the WPF UI thread.</summary>
        public static BitmapImage BytesToBitmapImage(byte[] png)
        {
            var bmp = new BitmapImage();
            using (var ms = new MemoryStream(png))
            {
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.StreamSource = ms;
                bmp.EndInit();
            }
            bmp.Freeze();
            return bmp;
        }

        private static bool IsBlankPng(byte[] png)
        {
            try
            {
                using (var ms = new MemoryStream(png))
                using (var bmp = new Bitmap(ms))
                {
                    int samples = 0, white = 0;
                    for (int y = 0; y < bmp.Height; y += 6)
                    for (int x = 0; x < bmp.Width; x += 6)
                    {
                        var c = bmp.GetPixel(x, y);
                        samples++;
                        if (c.R > 248 && c.G > 248 && c.B > 248) white++;
                    }
                    return samples > 0 && white * 100 / samples > 96;
                }
            }
            catch { return true; }
        }

        private static byte[] LetterFallbackPng(string name, int size)
        {
            string letter = "?";
            if (!string.IsNullOrWhiteSpace(name))
            {
                var parts = name.Split(new[] { '-', '_', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                letter = parts.Length > 0 ? parts[0].Substring(0, 1).ToUpperInvariant() : name.Substring(0, 1).ToUpperInvariant();
            }

            using (var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb))
            using (var g = Graphics.FromImage(bmp))
            using (var ms = new MemoryStream())
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                g.Clear(Color.White);
                using (var fill = new SolidBrush(Color.FromArgb(0xE8, 0xF3, 0xFC)))
                    g.FillEllipse(fill, size / 8, size / 8, size * 3 / 4, size * 3 / 4);
                using (var pen = new Pen(Color.FromArgb(0x0F, 0x6C, 0xBD), Math.Max(1.5f, size / 48f)))
                    g.DrawEllipse(pen, size / 8, size / 8, size * 3 / 4, size * 3 / 4);
                using (var font = new Font("Segoe UI Semibold", size * 0.32f, FontStyle.Bold, GraphicsUnit.Pixel))
                using (var brush = new SolidBrush(Color.FromArgb(0x0F, 0x6C, 0xBD)))
                {
                    var sz = g.MeasureString(letter, font);
                    g.DrawString(letter, font, brush, (size - sz.Width) / 2f, (size - sz.Height) / 2f);
                }
                bmp.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }

        public static string CreatePngFile(string svgPath, string tintHex, int size = 256)
        {
            string outPath = Path.Combine(Path.GetTempPath(), "cat-icon-" + Guid.NewGuid().ToString("N") + ".png");
            Exception err = null;

            void Run()
            {
                try { Rasterize(svgPath, tintHex, size, outPath); }
                catch (Exception ex) { err = ex; }
            }

            if (Thread.CurrentThread.Name == "CAT.SvgPreview"
                || Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                Run();
            }
            else
            {
                var gate = new ManualResetEvent(false);
                var t = new Thread(() => { try { Run(); } finally { gate.Set(); } });
                t.SetApartmentState(ApartmentState.STA);
                t.IsBackground = true;
                t.Start();
                if (!gate.WaitOne(TimeSpan.FromSeconds(20)))
                    throw new TimeoutException("Icon preview timed out.");
            }

            if (err != null) throw err;
            if (!File.Exists(outPath) || new FileInfo(outPath).Length < 32)
                throw new InvalidOperationException("Could not render icon preview.");
            return outPath;
        }

        private static void Rasterize(string svgPath, string tintHex, int size, string outPath)
        {
            string svg = File.ReadAllText(svgPath);
            if (!string.IsNullOrEmpty(tintHex))
                svg = ApplyTint(svg, tintHex);

            // Give IE an explicit box when the file has no width/height attributes.
            int svgOpen = svg.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
            if (svgOpen >= 0)
            {
                int svgClose = svg.IndexOf('>', svgOpen);
                if (svgClose > svgOpen)
                {
                    string openTag = svg.Substring(svgOpen, svgClose - svgOpen);
                    if (openTag.IndexOf("width", StringComparison.OrdinalIgnoreCase) < 0)
                        svg = svg.Insert(svgClose, $" width='{size}' height='{size}'");
                }
            }

            bool finished = false;
            Exception captureError = null;
            string html = BuildPreviewHtml(svg);

            using (var form = new WinForms.Form
            {
                Width = size + 8,
                Height = size + 8,
                FormBorderStyle = WinForms.FormBorderStyle.None,
                ShowInTaskbar = false,
                StartPosition = WinForms.FormStartPosition.Manual,
                Location = new Point(64, 64),
                BackColor = Color.White,
                TopMost = false,
                Opacity = 1
            })
            {
                var wb = new WinForms.WebBrowser
                {
                    Left = 0,
                    Top = 0,
                    Width = size,
                    Height = size,
                    ScriptErrorsSuppressed = true,
                    ScrollBarsEnabled = false,
                    AllowNavigation = true,
                    IsWebBrowserContextMenuEnabled = false
                };
                form.Controls.Add(wb);

                wb.DocumentCompleted += (s, e) =>
                {
                    if (finished) return;
                    if (wb.ReadyState != WinForms.WebBrowserReadyState.Complete) return;
                    try
                    {
                        for (int i = 0; i < 8; i++)
                        {
                            try { WinForms.Application.DoEvents(); }
                            catch (InvalidOperationException) { }
                            Thread.Sleep(40);
                        }
                        CaptureBrowser(wb, size, outPath);
                        finished = true;
                    }
                    catch (Exception ex) { captureError = ex; finished = true; }
                    try { form.Close(); } catch { }
                };

                form.Show();
                wb.DocumentText = html;

                var deadline = DateTime.UtcNow.AddSeconds(12);
                while (!finished && DateTime.UtcNow < deadline)
                {
                    try { WinForms.Application.DoEvents(); }
                    catch (InvalidOperationException) { Thread.Sleep(15); }
                    Thread.Sleep(15);
                }
            }

            if (captureError != null) throw captureError;
            if (!File.Exists(outPath))
                throw new InvalidOperationException("Icon preview did not finish loading.");
        }

        private static void CaptureBrowser(WinForms.WebBrowser wb, int size, string outPath)
        {
            using (var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb))
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                bool ok = false;
                IntPtr hwnd = wb.Handle;
                if (hwnd != IntPtr.Zero)
                {
                    IntPtr hdc = g.GetHdc();
                    try { ok = PrintWindow(hwnd, hdc, 2); }
                    finally { g.ReleaseHdc(hdc); }
                }
                if (!ok || IsMostlyBlank(bmp))
                {
                    g.Clear(Color.White);
                    try { wb.DrawToBitmap(bmp, new Rectangle(0, 0, size, size)); }
                    catch { }
                }
                bmp.Save(outPath, ImageFormat.Png);
            }
        }

        private static bool IsMostlyBlank(Bitmap bmp)
        {
            int samples = 0, white = 0;
            for (int y = 0; y < bmp.Height; y += 8)
            for (int x = 0; x < bmp.Width; x += 8)
            {
                var c = bmp.GetPixel(x, y);
                samples++;
                if (c.R > 250 && c.G > 250 && c.B > 250) white++;
            }
            return samples > 0 && white * 100 / samples > 97;
        }
    }
}
