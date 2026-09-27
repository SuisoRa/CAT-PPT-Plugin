using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Cat.Commands;
using Cat.TextBlocks;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WinForms = System.Windows.Forms;
using Keys = System.Windows.Forms.Keys;

namespace Cat.Core
{
    public struct KeyCombo : IEquatable<KeyCombo>
    {
        public bool HasCtrl { get; }
        public bool HasAlt { get; }
        public bool HasShift { get; }
        public Keys Key { get; }

        public KeyCombo(bool ctrl, bool alt, bool shift, Keys key)
        { HasCtrl = ctrl; HasAlt = alt; HasShift = shift; Key = key; }

        public static KeyCombo Ctrl(Keys k)         => new KeyCombo(true, false, false, k);
        public static KeyCombo Alt(Keys k)          => new KeyCombo(false, true, false, k);
        public static KeyCombo Shift(Keys k)        => new KeyCombo(false, false, true, k);
        public static KeyCombo CtrlAlt(Keys k)      => new KeyCombo(true, true, false, k);
        public static KeyCombo CtrlShift(Keys k)    => new KeyCombo(true, false, true, k);
        public static KeyCombo AltShift(Keys k)     => new KeyCombo(false, true, true, k);
        public static KeyCombo CtrlAltShift(Keys k) => new KeyCombo(true, true, true, k);

        /// <summary>Human-readable shortcut for help UI.</summary>
        public string DisplayLabel()
        {
            string key = FormatKey(Key);
            return (HasCtrl ? "Ctrl+" : "") + (HasAlt ? "Alt+" : "") + (HasShift ? "Shift+" : "") + key;
        }

        private static string FormatKey(Keys key)
        {
            switch (key)
            {
                case Keys.OemOpenBrackets: return "[";
                case Keys.OemCloseBrackets: return "]";
                case Keys.D0: return "0";
                case Keys.D1: return "1";
                case Keys.D2: return "2";
                case Keys.D3: return "3";
                case Keys.D4: return "4";
                case Keys.D5: return "5";
                case Keys.D7: return "7";
                case Keys.D8: return "8";
                case Keys.Delete: return "Delete";
                case Keys.Up: return "↑";
                case Keys.Down: return "↓";
                case Keys.Left: return "←";
                case Keys.Right: return "→";
                default:
                    string s = key.ToString();
                    if (s.StartsWith("D", StringComparison.Ordinal) && s.Length == 2 && char.IsDigit(s[1]))
                        return s.Substring(1);
                    return s;
            }
        }

        public bool Equals(KeyCombo o) =>
            HasCtrl == o.HasCtrl && HasAlt == o.HasAlt && HasShift == o.HasShift && Key == o.Key;
        public override bool Equals(object obj) => obj is KeyCombo o && Equals(o);
        public override int GetHashCode()
        {
            int h = (int)Key;
            if (HasCtrl) h |= 0x10000;
            if (HasAlt) h |= 0x20000;
            if (HasShift) h |= 0x40000;
            return h;
        }
        public override string ToString() =>
            (HasCtrl ? "Ctrl+" : "") + (HasAlt ? "Alt+" : "") + (HasShift ? "Shift+" : "") + Key;
    }

    internal static class NativeMethods
    {
        public const int WH_KEYBOARD = 2;
        public const int VK_CONTROL = 0x11, VK_MENU = 0x12, VK_SHIFT = 0x10;
        public const long KeyUpBit = 0x80000000;

        public delegate IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, KeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int nVirtKey);
        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();
    }

    public sealed class KeyboardHook
    {
        private readonly ShortcutDispatcher _dispatcher;
        private readonly NativeMethods.KeyboardProc _proc;
        private IntPtr _hookId = IntPtr.Zero;

        public KeyboardHook(ShortcutDispatcher dispatcher)
        { _dispatcher = dispatcher; _proc = HookCallback; }

        public void Install()
        {
            if (_hookId != IntPtr.Zero) return;

            // Warm the managed path so the first real keystroke is not a huge JIT
            // (Windows silently drops hooks that stall).
            _dispatcher.WillHandle(new KeyCombo(true, false, false, Keys.F24));
            IsDown(NativeMethods.VK_CONTROL);

            _hookId = NativeMethods.SetWindowsHookEx(
                NativeMethods.WH_KEYBOARD, _proc, IntPtr.Zero, NativeMethods.GetCurrentThreadId());

            if (_hookId == IntPtr.Zero)
            {
                int err = Marshal.GetLastWin32Error();
                Notifier.Error($"CAT could not install keyboard shortcuts (Win32 {err}).");
            }
        }

        public void Uninstall()
        {
            if (_hookId == IntPtr.Zero) return;
            NativeMethods.UnhookWindowsHookEx(_hookId);
            _hookId = IntPtr.Zero;
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                bool keyUp = (lParam.ToInt64() & NativeMethods.KeyUpBit) != 0;
                if (!keyUp)
                {
                    Keys key = (Keys)(wParam.ToInt32() & 0xFFFF);
                    if (!IsModifierKey(key))
                    {
                        bool ctrl = IsDown(NativeMethods.VK_CONTROL);
                        bool alt = IsDown(NativeMethods.VK_MENU);
                        bool shift = IsDown(NativeMethods.VK_SHIFT);
                        if (ctrl || alt || shift)
                        {
                            var combo = new KeyCombo(ctrl, alt, shift, key);
                            if (_dispatcher.WillHandle(combo))
                            {
                                _dispatcher.QueueExecute(combo);
                                return (IntPtr)1;
                            }
                        }
                    }
                }
            }
            return NativeMethods.CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private static bool IsDown(int vk) => (NativeMethods.GetAsyncKeyState(vk) & 0x8000) != 0;

        private static bool IsModifierKey(Keys k) =>
            k == Keys.ControlKey || k == Keys.LControlKey || k == Keys.RControlKey ||
            k == Keys.Menu || k == Keys.LMenu || k == Keys.RMenu ||
            k == Keys.ShiftKey || k == Keys.LShiftKey || k == Keys.RShiftKey;
    }

    public sealed class ShortcutDispatcher
    {
        private readonly WinForms.Control _ui;
        private ICommand _pending;

        public ShortcutDispatcher(PowerPoint.Application app)
        {
            _ui = new WinForms.Control();
            var unused = _ui.Handle;
            UiThreadRunner.Install(_ui);
        }

        public bool WillHandle(KeyCombo combo)
        {
            _pending = null;
            if (!ShortcutMap.Bindings.TryGetValue(combo, out ICommand command)) return false;

            // Ctrl+M must not fire when Alt is still down (Ctrl+Alt+M is align middle).
            if (command is InsertNewSlideCommand
                && combo.HasCtrl && !combo.HasAlt && !combo.HasShift
                && (NativeMethods.GetAsyncKeyState(NativeMethods.VK_MENU) & 0x8000) != 0)
                return false;

            if (command is ChevronToggleCommand && !ChevronTextBlockService.SelectionToggleAvailable)
                return false;

            _pending = command;
            return true;
        }

        public void QueueExecute(KeyCombo combo)
        {
            var command = _pending;
            _pending = null;
            if (command == null) return;

            // Always post — never run PowerPoint COM inside the hook callback.
            _ui.BeginInvoke(new WinForms.MethodInvoker(() =>
            {
                try { CommandRunner.Run(command); }
                catch (Exception ex) { Notifier.Error($"'{combo}' could not run: {ex.Message}"); }
            }));
        }
    }

    /// <summary>WinForms STA anchor for WPF windows, COM, and icon preview callbacks.</summary>
    public static class UiThreadRunner
    {
        private static WinForms.Control _anchor;

        public static void Install(WinForms.Control anchor) => _anchor = anchor;

        public static void Run(Action action)
        {
            if (action == null) return;
            if (_anchor == null || !_anchor.InvokeRequired)
            {
                action();
                return;
            }
            _anchor.Invoke(action);
        }

        public static void RunAsync(Action action)
        {
            if (action == null) return;
            if (_anchor == null || !_anchor.InvokeRequired)
            {
                action();
                return;
            }
            _anchor.BeginInvoke(action);
        }

        public static bool IsUiThread =>
            _anchor == null || !_anchor.InvokeRequired;
    }

    public static class CommandRunner
    {
        public static void Run(ICommand command)
        {
            var app = ThisAddIn.Instance?.App;
            if (app == null) return;
            var ctx = new CommandContext(app);
            if (!command.CanExecute(ctx))
            {
                Notifier.Info($"'{command.Name}' isn't available for the current selection.");
                return;
            }

            try { app.StartNewUndoEntry(); } catch { }
            command.Execute(ctx);
        }
    }

    public static class BuildInfo
    {
        public const string Version = "1.0.0";
        public const string DisplayName = "CAT";
        public static string VersionLabel => $"{DisplayName}  \u2022  {Version}";
    }

    public static class AddInSettings
    {
        private const string RegKey = @"Software\CAT";
        public static string IconFolderPath { get; set; }
        /// <summary>When true, align commands use the first-selected shape as the anchor.</summary>
        public static bool AlignByFirst { get; set; }

        public static void Load()
        {
            bool persistPath = false;
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegKey))
                {
                    if (key == null)
                    {
                        IconFolderPath = DefaultSubfolder("icons");
                        AlignByFirst = false;
                        persistPath = true;
                    }
                    else
                    {
                        IconFolderPath = key.GetValue("IconFolderPath") as string;
                        var abf = key.GetValue("AlignByFirst");
                        AlignByFirst = abf != null && Convert.ToInt32(abf) != 0;
                    }
                }
            }
            catch { }

            if (string.IsNullOrWhiteSpace(IconFolderPath) || !Directory.Exists(IconFolderPath))
            {
                IconFolderPath = DefaultSubfolder("icons");
                persistPath = true;
            }

            if (persistPath) Save();
            IconRepository.ClearCache();
        }

        public static void Save()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(RegKey))
                {
                    if (!string.IsNullOrEmpty(IconFolderPath)) key.SetValue("IconFolderPath", IconFolderPath);
                    key.SetValue("AlignByFirst", AlignByFirst ? 1 : 0);
                }
            }
            catch { }
        }

        private static string DefaultSubfolder(string name)
        {
            var baseDir = Path.GetDirectoryName(typeof(AddInSettings).Assembly.Location) ?? "";
            string nextToDll = Path.Combine(baseDir, name);
            if (Directory.Exists(nextToDll)) return nextToDll;

            try
            {
                var dir = new DirectoryInfo(baseDir);
                for (int i = 0; i < 6 && dir != null; i++, dir = dir.Parent)
                {
                    string assets = Path.Combine(dir.FullName, "assets", name);
                    if (Directory.Exists(assets)) return assets;
                }
            }
            catch { }

            return nextToDll;
        }
    }

    public static class ShortcutMap
    {
        public static readonly IReadOnlyDictionary<KeyCombo, ICommand> Bindings =
            new Dictionary<KeyCombo, ICommand>
            {
                { KeyCombo.Alt(Keys.G),                     new AlignAndGroupCommand() },
                { KeyCombo.CtrlAltShift(Keys.R),            new AlignRowsAndGroupCommand() },
                { KeyCombo.CtrlAltShift(Keys.C),            new AlignColumnsAndGroupCommand() },
                { KeyCombo.CtrlAlt(Keys.Down),              new AlignBottomCommand() },
                { KeyCombo.CtrlAlt(Keys.C),                 new AlignCenterCommand() },
                { KeyCombo.CtrlAlt(Keys.Left),              new AlignLeftCommand() },
                { KeyCombo.CtrlAlt(Keys.M),                 new AlignMiddleCommand() },
                { KeyCombo.CtrlAlt(Keys.Right),             new AlignRightCommand() },
                { KeyCombo.CtrlAlt(Keys.Up),                new AlignTopCommand() },
                { KeyCombo.AltShift(Keys.H),                new DistributeHorizontallyCommand() },
                { KeyCombo.AltShift(Keys.V),                new DistributeVerticallyCommand() },
                { KeyCombo.Alt(Keys.OemCloseBrackets),      new BringForwardCommand() },
                { KeyCombo.AltShift(Keys.OemCloseBrackets), new BringToFrontCommand() },
                { KeyCombo.Alt(Keys.OemOpenBrackets),       new SendBackwardCommand() },
                { KeyCombo.AltShift(Keys.OemOpenBrackets),  new SendToBackCommand() },
                { KeyCombo.Ctrl(Keys.D1),                   new CopyPositionCommand() },
                { KeyCombo.Ctrl(Keys.D2),                   new PastePositionCommand() },
                { KeyCombo.CtrlShift(Keys.D8),              new DoNotResizeCommand() },
                { KeyCombo.Ctrl(Keys.D8),                   new ResizeCommand() },
                { KeyCombo.Ctrl(Keys.D5),                   new FitToWindowCommand() },
                { KeyCombo.Shift(Keys.Delete),              new ClearSelectedTextCommand() },
                { KeyCombo.CtrlShift(Keys.E),               new MakeSameHeightCommand() },
                { KeyCombo.Alt(Keys.Z),                     new MakeSameSizeCommand() },
                { KeyCombo.CtrlAlt(Keys.E),                 new MakeSameWidthCommand() },
                { KeyCombo.CtrlAlt(Keys.R),                 new ResetFixedElementsCommand() },
                { KeyCombo.CtrlAlt(Keys.J),                 new SplitJoinTextboxesCommand() },
                { KeyCombo.Ctrl(Keys.M),                    new InsertNewSlideCommand() },
                { KeyCombo.Alt(Keys.Q),                     new InsertTextboxCommand() },
                { KeyCombo.Ctrl(Keys.D0),                   new InsertYellowStickyCommand() },
                { KeyCombo.Ctrl(Keys.T),                   new ChevronToggleCommand() },
                { KeyCombo.CtrlAlt(Keys.T),                 new PasteUnformattedTextCommand() },
                { KeyCombo.AltShift(Keys.A),                new CycleAccentColorsCommand() },
                { KeyCombo.AltShift(Keys.Right),            new IncreaseListLevelCommand() },
                { KeyCombo.AltShift(Keys.Left),             new DecreaseListLevelCommand() },
                { KeyCombo.Ctrl(Keys.D7),                   new WordWrapCommand() },
                { KeyCombo.CtrlShift(Keys.D7),              new DoNotWordWrapCommand() },
                { KeyCombo.CtrlShift(Keys.L),               new ListLineSpacingCommand() },
                { KeyCombo.CtrlAlt(Keys.Delete),            new DeleteAllExceptSelectionCommand() },
                { KeyCombo.CtrlShift(Keys.Delete),          new DeleteAllExceptSelectionCommand() },
                { KeyCombo.Ctrl(Keys.D3),                   new NormalViewCommand() },
                { KeyCombo.Ctrl(Keys.D4),                   new SlideSorterViewCommand() },
            };
    }

    public static class Notifier
    {
        public static void Info(string m) =>
            WinForms.MessageBox.Show(m, BuildInfo.DisplayName, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);

        public static void Error(string m) =>
            WinForms.MessageBox.Show(m, BuildInfo.DisplayName, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
    }

    public sealed class SelectionOrderTracker : IDisposable
    {
        private readonly PowerPoint.Application _app;
        private readonly List<int> _order = new List<int>();

        public SelectionOrderTracker(PowerPoint.Application app)
        { _app = app; _app.WindowSelectionChange += OnSelectionChange; }

        private void OnSelectionChange(PowerPoint.Selection sel)
        {
            try { ChevronTextBlockService.RefreshSelectionToggleAvailability(sel); } catch { }
            try
            {
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                { _order.Clear(); return; }

                var current = new HashSet<int>();
                var currentInApiOrder = new List<int>();
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    int id = s.Id;
                    if (current.Add(id)) currentInApiOrder.Add(id);
                }

                _order.RemoveAll(id => !current.Contains(id));
                var known = new HashSet<int>(_order);
                foreach (int id in currentInApiOrder)
                    if (known.Add(id)) _order.Add(id);
            }
            catch { _order.Clear(); }
        }

        public IList<PowerPoint.Shape> GetOrderedSelection(PowerPoint.ShapeRange range)
        {
            var byId = new Dictionary<int, PowerPoint.Shape>();
            if (range != null)
                foreach (PowerPoint.Shape s in range) byId[s.Id] = s;

            var ordered = new List<PowerPoint.Shape>();
            foreach (int id in _order)
                if (byId.TryGetValue(id, out var s)) { ordered.Add(s); byId.Remove(id); }
            foreach (var leftover in byId.Values) ordered.Add(leftover);
            return ordered;
        }

        public void Dispose()
        { try { _app.WindowSelectionChange -= OnSelectionChange; } catch { } }
    }

    public sealed class PositionClipboard
    {
        public struct Slot { public float Left, Top; public Slot(float l, float t) { Left = l; Top = t; } }
        private readonly List<Slot> _slots = new List<Slot>();
        public bool HasData => _slots.Count > 0;
        public int Count => _slots.Count;
        public void Clear() => _slots.Clear();
        public void Add(float l, float t) => _slots.Add(new Slot(l, t));
        public Slot this[int i] => _slots[i];
    }

    /// <summary>Gap between two shapes, in points. Horizontal is edge-to-edge left-to-right; vertical is top-to-bottom.</summary>
    public sealed class SpacingClipboard
    {
        public bool HasData { get; private set; }
        public float Horizontal { get; private set; }
        public float Vertical { get; private set; }

        public void Set(float horizontal, float vertical)
        {
            Horizontal = horizontal;
            Vertical = vertical;
            HasData = true;
        }
    }

    public static class IconRepository
    {
        private const int RibbonIconSize = 32;
        private static readonly Assembly IconAssembly = typeof(IconRepository).Assembly;
        private static string[] _embeddedIconNames;

        private static readonly Dictionary<string, string> _pathCache =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static List<string> _searchDirs;
        private static string _searchDirsKey;

        /// <summary>Bitmap sized for the Office ribbon; caller must keep it alive with IPictureDisp.</summary>
        public static Bitmap CreateRibbonBitmap(string controlId)
        {
            if (string.IsNullOrWhiteSpace(controlId)) return null;

            using (var src = OpenSourceImage(controlId))
            {
                if (src == null) return null;
                return ScaleToRibbonSize(src);
            }
        }

        private static Image OpenSourceImage(string controlId)
        {
            using (var stream = OpenEmbeddedStream(controlId))
            {
                if (stream != null)
                {
                    using (var decoded = Image.FromStream(stream, useEmbeddedColorManagement: false, validateImageData: true))
                        return new Bitmap(decoded);
                }
            }

            string path = ResolveFilePath(controlId);
            if (path == null) return null;
            using (var decoded = Image.FromFile(path))
                return new Bitmap(decoded);
        }

        private static Stream OpenEmbeddedStream(string controlId)
        {
            EnsureEmbeddedIndex();
            string suffix = "." + controlId + ".png";
            foreach (string name in _embeddedIconNames)
            {
                if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                    return IconAssembly.GetManifestResourceStream(name);
            }
            return null;
        }

        private static void EnsureEmbeddedIndex()
        {
            if (_embeddedIconNames != null) return;
            _embeddedIconNames = IconAssembly
                .GetManifestResourceNames()
                .Where(n => n.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                .ToArray();
        }

        private static Bitmap ScaleToRibbonSize(Image src)
        {
            int w = Math.Max(1, src.Width);
            int h = Math.Max(1, src.Height);
            float scale = Math.Min(RibbonIconSize / (float)w, RibbonIconSize / (float)h);
            int tw = Math.Max(1, (int)Math.Round(w * scale));
            int th = Math.Max(1, (int)Math.Round(h * scale));
            int ox = (RibbonIconSize - tw) / 2;
            int oy = (RibbonIconSize - th) / 2;

            var bmp = new Bitmap(RibbonIconSize, RibbonIconSize, PixelFormat.Format32bppArgb);
            bmp.SetResolution(96f, 96f);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.DrawImage(src, new Rectangle(ox, oy, tw, th));
            }
            return bmp;
        }

        public static string ResolveFilePath(string controlId)
        {
            if (string.IsNullOrWhiteSpace(controlId)) return null;
            EnsureIconSettingsLoaded();
            if (_pathCache.TryGetValue(controlId, out string cached)) return cached;

            foreach (string dir in GetSearchDirectories())
            {
                string candidate = Path.Combine(dir, controlId + ".png");
                if (!File.Exists(candidate)) continue;
                _pathCache[controlId] = candidate;
                return candidate;
            }

            _pathCache[controlId] = null;
            return null;
        }

        public static Bitmap LoadBitmap(string path)
        {
            using (var src = Image.FromFile(path))
                return ScaleToRibbonSize(src);
        }

        public static string Resolve(string controlId) => ResolveFilePath(controlId);

        private static IEnumerable<string> GetSearchDirectories()
        {
            string settingsKey = AddInSettings.IconFolderPath ?? "";
            if (_searchDirs != null && string.Equals(settingsKey, _searchDirsKey, StringComparison.Ordinal))
                return _searchDirs;

            var list = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddDir(string path)
            {
                if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path)) return;
                try
                {
                    path = Path.GetFullPath(path);
                    if (seen.Add(path)) list.Add(path);
                }
                catch { }
            }

            var baseDir = Path.GetDirectoryName(IconAssembly.Location) ?? "";
            AddDir(Path.Combine(baseDir, "icons"));

            try
            {
                var dir = new DirectoryInfo(baseDir);
                for (int i = 0; i < 6 && dir != null; i++, dir = dir.Parent)
                    AddDir(Path.Combine(dir.FullName, "assets", "icons"));
            }
            catch { }

            AddDir(AddInSettings.IconFolderPath);

            _searchDirs = list;
            _searchDirsKey = settingsKey;
            return list;
        }

        public static void ClearCache()
        {
            _pathCache.Clear();
            _searchDirs = null;
            _searchDirsKey = null;
        }

        private static void EnsureIconSettingsLoaded()
        {
            if (string.IsNullOrWhiteSpace(AddInSettings.IconFolderPath) || !Directory.Exists(AddInSettings.IconFolderPath))
                AddInSettings.Load();
        }
    }

    public static class ShapeHelpers
    {
        public const float PointsPerInch = 72f;

        public static List<PowerPoint.Shape> ToList(PowerPoint.ShapeRange range)
        {
            var list = new List<PowerPoint.Shape>();
            if (range == null) return list;
            foreach (PowerPoint.Shape s in range) list.Add(s);
            return list;
        }

        public static bool HasTextFrame(PowerPoint.Shape s)
        { try { return s.HasTextFrame == Office.MsoTriState.msoTrue; } catch { return false; } }

        public static void SetAutoSize(PowerPoint.Shape s, PowerPoint.PpAutoSize mode)
        { if (HasTextFrame(s)) s.TextFrame.AutoSize = mode; }

        public static void SetWordWrap(PowerPoint.Shape s, bool wrap)
        { if (HasTextFrame(s)) s.TextFrame.WordWrap = wrap ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; }

        public static readonly Office.MsoThemeColorSchemeIndex[] AccentCycle =
        {
            Office.MsoThemeColorSchemeIndex.msoThemeAccent1,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent2,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent3,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent4,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent5,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent6,
        };
        private const string AccentTag = "CAT_ACCENT_INDEX";

        public static void CycleAccentFill(PowerPoint.Shape shape, PowerPoint.Slide slide)
        {
            dynamic themeColors = slide.ThemeColorScheme;
            int current = -1;

            try
            {
                if (shape.Fill.Visible == Office.MsoTriState.msoTrue)
                {
                    int rgb = shape.Fill.ForeColor.RGB;
                    for (int i = 0; i < AccentCycle.Length; i++)
                    {
                        try
                        {
                            if (themeColors[AccentCycle[i]].RGB == rgb)
                            { current = i; break; }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            int next = current >= 0 ? (current + 1) % AccentCycle.Length : 0;
            var accent = AccentCycle[next];
            var fillRgb = themeColors[accent].RGB;
            shape.Fill.Visible = Office.MsoTriState.msoTrue;
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = fillRgb;
            try { shape.Tags.Add(AccentTag, next.ToString()); } catch { }
        }
    }

    /// <summary>Title placeholder geometry from the slide layout (for aligning new text boxes).</summary>
    public static class SlideLayoutHelper
    {
        public static bool TryGetTitleLayout(PowerPoint.Slide slide, out float left, out float top, out float width, out float height)
        {
            left = top = width = height = 0f;
            try
            {
                if (slide?.CustomLayout == null) return false;
                foreach (PowerPoint.Shape s in slide.CustomLayout.Shapes)
                {
                    if (s.Type != Office.MsoShapeType.msoPlaceholder) continue;
                    var t = s.PlaceholderFormat.Type;
                    if (t != PowerPoint.PpPlaceholderType.ppPlaceholderTitle
                        && t != PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle)
                        continue;
                    left = s.Left;
                    top = s.Top;
                    width = s.Width;
                    height = s.Height;
                    return true;
                }
            }
            catch { }
            return false;
        }
    }

    /// <summary>Cluster shapes into grid rows/columns (8 pt tolerance) for matrix copy tools.</summary>
    public static class GridDetector
    {
        public const float Tolerance = 8f;

        private struct Item
        {
            public PowerPoint.Shape Shape;
            public float Top, Left;
        }

        public static List<List<PowerPoint.Shape>> Rows(IEnumerable<PowerPoint.Shape> shapes) =>
            Cluster(shapes, byTop: true);

        private static List<List<PowerPoint.Shape>> Cluster(IEnumerable<PowerPoint.Shape> shapes, bool byTop)
        {
            var items = new List<Item>();
            foreach (var s in shapes)
            {
                try { items.Add(new Item { Shape = s, Top = s.Top, Left = s.Left }); }
                catch { }
            }

            Func<Item, float> primary = byTop ? (Func<Item, float>)(i => i.Top) : (i => i.Left);
            Func<Item, float> secondary = byTop ? (Func<Item, float>)(i => i.Left) : (i => i.Top);
            items.Sort((a, b) => primary(a).CompareTo(primary(b)));

            var clusters = new List<List<Item>>();
            List<Item> current = null;
            float anchor = 0f;
            foreach (var item in items)
            {
                float v = primary(item);
                if (current == null || Math.Abs(v - anchor) > Tolerance)
                {
                    current = new List<Item>();
                    clusters.Add(current);
                    anchor = v;
                }
                current.Add(item);
            }

            var result = new List<List<PowerPoint.Shape>>(clusters.Count);
            foreach (var c in clusters)
            {
                c.Sort((a, b) => secondary(a).CompareTo(secondary(b)));
                result.Add(c.Select(i => i.Shape).ToList());
            }
            return result;
        }
    }
}
