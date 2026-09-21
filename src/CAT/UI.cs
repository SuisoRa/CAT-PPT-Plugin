using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using WinForms = System.Windows.Forms;
using System.Windows.Markup;
using System.Windows.Media;
using Cat.TableCreator;
using Cat.AutoFormat;
using Cat.Commands;
using Cat.Core;
using Office = Microsoft.Office.Core;

namespace Cat.UI
{
    public enum SplitJoinMode { Split, Join }

    public sealed class SplitJoinResult
    { public bool Ok; public SplitJoinMode Mode = SplitJoinMode.Split; public int Level = 1; public bool KeepFormatting = true; }

    public sealed class LineSpacingResult
    { public bool Ok; public bool Auto; public bool InLines = true; public double[] Values = { 1, 1, 1, 1, 1 }; }

    public sealed class SelectSimilarResult
    { public bool Ok; public bool Fill = true; public bool Outline = true; public bool Type = true; }

    internal static class ModernTheme
    {
        private static ResourceDictionary _cached;

        public static ResourceDictionary Get()
        {
            if (_cached != null) return _cached;
            _cached = (ResourceDictionary)XamlReader.Parse(Xaml);
            return _cached;
        }

        private const string Xaml = @"
<ResourceDictionary xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
                    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
  <Style TargetType='TextBlock'>
    <Setter Property='FontFamily' Value='Segoe UI'/>
    <Setter Property='FontSize' Value='13'/>
    <Setter Property='Foreground' Value='#1F1F1F'/>
  </Style>
  <Style TargetType='RadioButton'>
    <Setter Property='FontFamily' Value='Segoe UI'/><Setter Property='FontSize' Value='13'/>
    <Setter Property='Margin' Value='0,0,18,0'/>
  </Style>
  <Style TargetType='CheckBox'>
    <Setter Property='FontFamily' Value='Segoe UI'/><Setter Property='FontSize' Value='13'/>
    <Setter Property='Margin' Value='0,3,0,3'/>
  </Style>
  <Style TargetType='TextBox'>
    <Setter Property='FontFamily' Value='Segoe UI'/><Setter Property='FontSize' Value='13'/>
    <Setter Property='Padding' Value='6,4'/><Setter Property='BorderBrush' Value='#E2E2E2'/>
    <Setter Property='VerticalContentAlignment' Value='Center'/>
  </Style>
  <Style TargetType='ComboBox'>
    <Setter Property='FontFamily' Value='Segoe UI'/><Setter Property='FontSize' Value='13'/>
    <Setter Property='Height' Value='32'/>
  </Style>
  <Style x:Key='Title' TargetType='TextBlock'>
    <Setter Property='FontFamily' Value='Segoe UI Semibold'/><Setter Property='FontSize' Value='16'/>
    <Setter Property='Foreground' Value='#0F6CBD'/><Setter Property='Margin' Value='0,0,0,14'/>
  </Style>
  <Style x:Key='Primary' TargetType='Button'>
    <Setter Property='Foreground' Value='White'/><Setter Property='Background' Value='#0F6CBD'/>
    <Setter Property='FontFamily' Value='Segoe UI'/><Setter Property='FontSize' Value='13'/>
    <Setter Property='Height' Value='34'/><Setter Property='Cursor' Value='Hand'/>
    <Setter Property='Template'>
      <Setter.Value>
        <ControlTemplate TargetType='Button'>
          <Border x:Name='b' Background='{TemplateBinding Background}' CornerRadius='6'>
            <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='b' Property='Opacity' Value='0.9'/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style x:Key='Secondary' TargetType='Button'>
    <Setter Property='Foreground' Value='#1F1F1F'/><Setter Property='Background' Value='#F5F5F5'/>
    <Setter Property='FontFamily' Value='Segoe UI'/><Setter Property='FontSize' Value='13'/>
    <Setter Property='Height' Value='34'/><Setter Property='Cursor' Value='Hand'/>
    <Setter Property='Template'>
      <Setter.Value>
        <ControlTemplate TargetType='Button'>
          <Border x:Name='b' Background='{TemplateBinding Background}' CornerRadius='6' BorderBrush='#E2E2E2' BorderThickness='1'>
            <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='b' Property='Background' Value='#ECECEC'/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style x:Key='Stepper' TargetType='Button'>
    <Setter Property='Width' Value='26'/><Setter Property='Height' Value='26'/>
    <Setter Property='FontSize' Value='15'/><Setter Property='Foreground' Value='#6B6B6B'/>
    <Setter Property='Cursor' Value='Hand'/>
    <Setter Property='Template'>
      <Setter.Value>
        <ControlTemplate TargetType='Button'>
          <Border x:Name='b' Background='#F0F0F0' CornerRadius='13'>
            <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='b' Property='Background' Value='#E2E2E2'/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style x:Key='AfTool' TargetType='Button'>
    <Setter Property='Width' Value='46'/><Setter Property='Height' Value='46'/>
    <Setter Property='Margin' Value='3'/><Setter Property='Cursor' Value='Hand'/>
    <Setter Property='Background' Value='White'/>
    <Setter Property='Template'>
      <Setter.Value>
        <ControlTemplate TargetType='Button'>
          <Border x:Name='b' Background='{TemplateBinding Background}' CornerRadius='10' BorderBrush='#E2E2E2' BorderThickness='1'>
            <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property='IsMouseOver' Value='True'><Setter TargetName='b' Property='Background' Value='#E8F3FC'/></Trigger>
            <Trigger Property='IsPressed' Value='True'><Setter TargetName='b' Property='Background' Value='#D0E7FA'/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
</ResourceDictionary>";
    }

    public static class Dialogs
    {
        internal static Window NewShellWindow(string title, double width, double height)
        {
            var w = new Window
            {
                Title = title,
                Width = width,
                Height = height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = Brushes.White,
                ShowInTaskbar = false
            };
            w.Resources.MergedDictionaries.Add(ModernTheme.Get());
            try
            {
                IntPtr hwnd = (IntPtr)ThisAddIn.Instance.App.HWND;
                if (hwnd != IntPtr.Zero)
                    new System.Windows.Interop.WindowInteropHelper(w).Owner = hwnd;
            }
            catch { w.WindowStartupLocation = WindowStartupLocation.CenterScreen; }
            return w;
        }

        private static Window NewWindow(string title, double width)
        {
            var w = new Window
            {
                Title = title,
                Width = width,
                SizeToContent = SizeToContent.Height,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowStyle = WindowStyle.SingleBorderWindow,
                Background = Brushes.White,
                ShowInTaskbar = false
            };
            w.Resources.MergedDictionaries.Add(ModernTheme.Get());

            try
            {
                IntPtr hwnd = (IntPtr)ThisAddIn.Instance.App.HWND;
                if (hwnd != IntPtr.Zero)
                    new System.Windows.Interop.WindowInteropHelper(w).Owner = hwnd;
                else
                    w.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
            catch { w.WindowStartupLocation = WindowStartupLocation.CenterScreen; }

            return w;
        }

        private static Button Styled(string content, string styleKey, Window host) =>
            new Button { Content = content, Style = (Style)host.Resources[styleKey] };

        public static SplitJoinResult ShowSplitJoin(bool canJoin)
        {
            var result = new SplitJoinResult();
            var win = NewWindow("Split / Join", 300);

            var root = new StackPanel { Margin = new Thickness(20) };
            root.Children.Add(new TextBlock { Text = "Split / Join", Style = (Style)win.Resources["Title"] });

            var rbSplit = new RadioButton { Content = "Split", GroupName = "mode", IsChecked = true };
            var rbJoin = new RadioButton { Content = "Join", GroupName = "mode", IsEnabled = canJoin };
            var modeRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 16) };
            modeRow.Children.Add(rbSplit); modeRow.Children.Add(rbJoin);
            root.Children.Add(modeRow);

            var cbLevel = new ComboBox { Margin = new Thickness(0, 0, 0, 16) };
            AddLevel(cbLevel, "1 - First level", 1);
            AddLevel(cbLevel, "2 - Second level", 2);
            AddLevel(cbLevel, "3 - Third level", 3);
            AddLevel(cbLevel, "4 - Fourth level", 4);
            AddLevel(cbLevel, "5 - Fifth level", 5);
            AddLevel(cbLevel, "All (every line)", -1);
            cbLevel.SelectedIndex = 0;
            root.Children.Add(cbLevel);

            var chkKeep = new CheckBox { Content = "Keep formatting", IsChecked = true, Margin = new Thickness(0, 0, 0, 18) };
            root.Children.Add(chkKeep);

            var btnGo = Styled("Split", "Primary", win);
            root.Children.Add(btnGo);

            RoutedEventHandler updateMode = (s, e) =>
            {
                bool split = rbSplit.IsChecked == true;
                cbLevel.IsEnabled = split;
                btnGo.Content = split ? "Split" : "Join";
            };
            rbSplit.Checked += updateMode; rbJoin.Checked += updateMode;

            btnGo.Click += (s, e) =>
            {
                result.Ok = true;
                result.Mode = rbSplit.IsChecked == true ? SplitJoinMode.Split : SplitJoinMode.Join;
                result.KeepFormatting = chkKeep.IsChecked == true;
                if (cbLevel.SelectedItem is ComboBoxItem it && it.Tag is int tag) result.Level = tag;
                win.DialogResult = true;
                win.Close();
            };

            win.Content = root;
            win.ShowDialog();
            return result;
        }

        private static void AddLevel(ComboBox cb, string text, int tag) =>
            cb.Items.Add(new ComboBoxItem { Content = text, Tag = tag });

        public static LineSpacingResult ShowLineSpacing()
        {
            var result = new LineSpacingResult();
            var win = NewWindow("List Line Spacing", 340);

            var root = new StackPanel { Margin = new Thickness(20) };
            root.Children.Add(new TextBlock { Text = "List Line Spacing", Style = (Style)win.Resources["Title"] });

            var chkAuto = new CheckBox { Content = "Auto (single spacing)", Margin = new Thickness(0, 0, 0, 12) };
            root.Children.Add(chkAuto);

            var rbLines = new RadioButton { Content = "Lines", GroupName = "units", IsChecked = true };
            var rbPoints = new RadioButton { Content = "Points", GroupName = "units" };
            var unitRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 12) };
            unitRow.Children.Add(rbLines); unitRow.Children.Add(rbPoints);
            root.Children.Add(unitRow);

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var fields = new TextBox[5];
            for (int i = 0; i < 5; i++)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                var lbl = new TextBlock { Text = $"Level {i + 1}", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 4, 0, 4) };
                Grid.SetRow(lbl, i); Grid.SetColumn(lbl, 0); grid.Children.Add(lbl);

                var box = new TextBox { Width = 56, Text = "1.0", TextAlignment = TextAlignment.Center, Margin = new Thickness(6, 0, 6, 0) };
                var dec = Styled("\u2212", "Stepper", win);
                var inc = Styled("+", "Stepper", win);
                dec.Click += (s, e) => Step(box, -0.1);
                inc.Click += (s, e) => Step(box, +0.1);

                var rowPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 4, 0, 4) };
                rowPanel.Children.Add(dec); rowPanel.Children.Add(box); rowPanel.Children.Add(inc);
                Grid.SetRow(rowPanel, i); Grid.SetColumn(rowPanel, 1); grid.Children.Add(rowPanel);

                fields[i] = box;
            }
            root.Children.Add(grid);

            var setRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 14, 0, 4) };
            var setZero = Styled("Set all to 0", "Secondary", win); setZero.Width = 110; setZero.Margin = new Thickness(0, 0, 10, 0);
            var setOne = Styled("Set all to 1", "Secondary", win); setOne.Width = 110;
            setZero.Click += (s, e) => FillAll(fields, 0);
            setOne.Click += (s, e) => FillAll(fields, 1);
            setRow.Children.Add(setZero); setRow.Children.Add(setOne);
            root.Children.Add(setRow);

            var btnGrid = new Grid { Margin = new Thickness(0, 18, 0, 0) };
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(10) });
            btnGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var cancel = Styled("Cancel", "Secondary", win);
            var apply = Styled("Apply", "Primary", win);
            Grid.SetColumn(cancel, 0); Grid.SetColumn(apply, 2);
            btnGrid.Children.Add(cancel); btnGrid.Children.Add(apply);
            root.Children.Add(btnGrid);

            RoutedEventHandler autoToggle = (s, e) =>
            {
                bool auto = chkAuto.IsChecked == true;
                grid.IsEnabled = !auto; unitRow.IsEnabled = !auto; setRow.IsEnabled = !auto;
            };
            chkAuto.Checked += autoToggle; chkAuto.Unchecked += autoToggle;

            cancel.Click += (s, e) => { win.DialogResult = false; win.Close(); };
            apply.Click += (s, e) =>
            {
                result.Ok = true;
                result.Auto = chkAuto.IsChecked == true;
                result.InLines = rbLines.IsChecked == true;
                for (int i = 0; i < 5; i++) result.Values[i] = Parse(fields[i].Text);
                win.DialogResult = true; win.Close();
            };

            win.Content = root;
            win.ShowDialog();
            return result;
        }

        private static void Step(TextBox box, double delta)
        {
            double v = Math.Max(0, Math.Round(Parse(box.Text) + delta, 2));
            box.Text = v.ToString("0.0#", CultureInfo.InvariantCulture);
        }
        private static void FillAll(TextBox[] fields, double value)
        {
            foreach (var f in fields) f.Text = value.ToString("0.0#", CultureInfo.InvariantCulture);
        }
        private static double Parse(string s) =>
            double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double v) ? v : 1.0;

        public static SelectSimilarResult ShowSelectSimilar()
        {
            var result = new SelectSimilarResult();
            var win = NewWindow("Select similar", 280);
            var root = new StackPanel { Margin = new Thickness(20) };
            root.Children.Add(new TextBlock { Text = "Select similar", Style = (Style)win.Resources["Title"] });
            root.Children.Add(new TextBlock
            {
                Text = "Match shapes on the current slide by:",
                Margin = new Thickness(0, 0, 0, 10),
                Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x6B, 0x6B))
            });
            var chkFill = new CheckBox { Content = "Shape fill", IsChecked = true };
            var chkOutline = new CheckBox { Content = "Shape outline", IsChecked = true };
            var chkType = new CheckBox { Content = "Shape type", IsChecked = true };
            root.Children.Add(chkFill);
            root.Children.Add(chkOutline);
            root.Children.Add(chkType);
            var apply = Styled("Apply", "Primary", win);
            apply.Margin = new Thickness(0, 18, 0, 0);
            apply.Click += (s, e) =>
            {
                result.Ok = true;
                result.Fill = chkFill.IsChecked == true;
                result.Outline = chkOutline.IsChecked == true;
                result.Type = chkType.IsChecked == true;
                win.DialogResult = true;
                win.Close();
            };
            root.Children.Add(apply);
            win.Content = root;
            win.ShowDialog();
            return result;
        }

        public static void ShowShortcutsHelp()
        {
            var win = NewShellWindow("CAT shortcuts", 520, 560);
            var root = new Grid { Margin = new Thickness(20) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            root.Children.Add(new TextBlock
            {
                Text = "Keyboard shortcuts",
                Style = (Style)win.Resources["Title"]
            });

            var list = new ListView
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8)),
                BorderThickness = new Thickness(1)
            };
            var view = new GridView();
            view.Columns.Add(new GridViewColumn
            {
                Header = "Shortcut",
                Width = 160,
                DisplayMemberBinding = new System.Windows.Data.Binding("Shortcut")
            });
            view.Columns.Add(new GridViewColumn
            {
                Header = "Action",
                Width = 300,
                DisplayMemberBinding = new System.Windows.Data.Binding("Action")
            });
            list.View = view;

            var rows = ShortcutMap.Bindings
                .Select(kv => new { Shortcut = kv.Key.DisplayLabel(), Action = kv.Value.Name })
                .OrderBy(r => r.Action, StringComparer.OrdinalIgnoreCase)
                .ToList();
            foreach (var r in rows) list.Items.Add(r);

            Grid.SetRow(list, 1);
            root.Children.Add(list);

            var close = Styled("Close", "Primary", win);
            close.Width = 100;
            close.HorizontalAlignment = HorizontalAlignment.Right;
            close.Margin = new Thickness(0, 14, 0, 0);
            close.Click += (s, e) => win.Close();
            Grid.SetRow(close, 2);
            root.Children.Add(close);

            win.Content = root;
            win.ShowDialog();
        }
    }

    [ComVisible(true)]
    public sealed class RibbonController : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private readonly Dictionary<string, Func<ICommand>> _actions = new Dictionary<string, Func<ICommand>>
        {
            { "catAlignRows", () => new AlignRowsAndGroupCommand() },
            { "catAlignCols", () => new AlignColumnsAndGroupCommand() },
            { "catDistH", () => new DistributeHorizontallyCommand() },
            { "catDistV", () => new DistributeVerticallyCommand() },
            { "catCopyPos", () => new CopyPositionCommand() },
            { "catPastePos", () => new PastePositionCommand() },
            { "catSameH", () => new MakeSameHeightCommand() },
            { "catSameW", () => new MakeSameWidthCommand() },
            { "catSameSize", () => new MakeSameSizeCommand() },
            { "catResize", () => new ResizeCommand() },
            { "catNoResize", () => new DoNotResizeCommand() },
            { "catReset", () => new ResetFixedElementsCommand() },
            { "catSplitJoin", () => new SplitJoinTextboxesCommand() },
            { "catWrap", () => new WordWrapCommand() },
            { "catNoWrap", () => new DoNotWordWrapCommand() },
            { "catLineSpacing", () => new ListLineSpacingCommand() },
            { "catFixTextBox", () => new FixTextBoxCommand() },
            { "catSelectSimilar", () => new SelectSimilarCommand(true, true, true) },
            { "catCopyExcel", () => new CopyForExcelCommand() },
            { "catSwapPosition", () => new SwapObjectPositionCommand() },
            { "catSameCornerRadius", () => new SameCornerRadiusCommand() },
            { "catMakeVertical", () => new MakeLineVerticalCommand() },
            { "catMakeHorizontal", () => new MakeLineHorizontalCommand() },
        };

        private static readonly Dictionary<string, stdole.IPictureDisp> _iconCache =
            new Dictionary<string, stdole.IPictureDisp>(StringComparer.OrdinalIgnoreCase);
        /// <summary>Office holds IPictureDisp without copying pixels; keep bitmaps alive.</summary>
        private static readonly Dictionary<string, System.Drawing.Bitmap> _iconBitmapKeepAlive =
            new Dictionary<string, System.Drawing.Bitmap>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Every ribbon control that uses getImage="GetImage".</summary>
        private static readonly string[] RibbonIconControlIds =
        {
            "catAutoFormat", "catTableCreator", "catIconLibrary",
            "catMakeVertical", "catMakeHorizontal",
            "catSelectSimilar", "catSelectSimilarOptions", "catCopyExcel",
            "catAlignRows", "catAlignCols", "catDistH", "catDistV",
            "catCopyPos", "catPastePos", "catSameH", "catSameW", "catSameSize",
            "catResize", "catNoResize", "catSplitJoin", "catWrap", "catNoWrap",
            "catLineSpacing", "catFixTextBox", "catReset",
            "catSwapPosition", "catSameCornerRadius", "catHelp"
        };

        public string GetCustomUI(string ribbonID)
        {
            try { return GetResourceText("Cat.Ribbon.xml"); }
            catch { return null; }
        }

        public void OnLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            AddInSettings.Load();
            RefreshImages(ribbonUI);
        }

        public void RefreshRibbonImages() => RefreshImages(_ribbon);

        internal static void RefreshImages(Office.IRibbonUI ribbonUi)
        {
            _iconCache.Clear();
            _iconBitmapKeepAlive.Clear();
            IconRepository.ClearCache();
            AddInSettings.Load();
            PreloadRibbonIcons();
            try { ribbonUi?.Invalidate(); } catch { }
        }

        private static void PreloadRibbonIcons()
        {
            foreach (string id in RibbonIconControlIds)
                TryCacheRibbonImage(id);
        }

        private static stdole.IPictureDisp TryCacheRibbonImage(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            if (_iconCache.TryGetValue(id, out var existing)) return existing;

            try
            {
                var bmp = IconRepository.CreateRibbonBitmap(id);
                if (bmp == null) return null;
                var pic = RibbonPictureConverter.FromBitmap(bmp);
                if (pic == null) return null;
                _iconBitmapKeepAlive[id] = bmp;
                _iconCache[id] = pic;
                return pic;
            }
            catch { return null; }
        }

        public bool GetAlignByFirst(Office.IRibbonControl control) => AddInSettings.AlignByFirst;

        public void OnAlignByFirst(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                AddInSettings.AlignByFirst = pressed;
                AddInSettings.Save();
                try { _ribbon?.InvalidateControl("catAlignByFirst"); } catch { }
            }
            catch (Exception ex) { Notifier.Error($"Could not update Align by first: {ex.Message}"); }
        }

        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                if (control.Id == "catAutoFormat")
                {
                    AutoFormatToolbar.Toggle();
                    return;
                }
                if (control.Id == "catIconLibrary")
                {
                    IconLibraryWindow.Show();
                    return;
                }
                if (control.Id == "catTableCreator")
                {
                    TableCreatorWindow.Show();
                    return;
                }
                if (control.Id == "catSelectSimilarOptions")
                {
                    var r = Dialogs.ShowSelectSimilar();
                    if (r.Ok) CommandRunner.Run(new SelectSimilarCommand(r.Fill, r.Outline, r.Type));
                    return;
                }
                if (control.Id == "catHelp")
                {
                    Dialogs.ShowShortcutsHelp();
                    return;
                }
                if (_actions.TryGetValue(control.Id, out var factory))
                    CommandRunner.Run(factory());
            }
            catch (Exception ex) { Notifier.Error($"Action failed: {ex.Message}"); }
        }

        public string GetVersionLabel(Office.IRibbonControl c) => "";

        public stdole.IPictureDisp GetImage(Office.IRibbonControl c)
        {
            try
            {
                string id = c?.Id ?? "";
                if (string.IsNullOrEmpty(id)) return null;
                if (_iconCache.TryGetValue(id, out var cached)) return cached;
                return TryCacheRibbonImage(id);
            }
            catch { return null; }
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                using (var reader = new StreamReader(stream)) return reader.ReadToEnd();
            }
        }
    }

    internal sealed class RibbonPictureConverter : WinForms.AxHost
    {
        private RibbonPictureConverter() : base(string.Empty) { }

        public static stdole.IPictureDisp FromBitmap(System.Drawing.Bitmap bitmap)
        {
            if (bitmap == null) return null;
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(bitmap);
        }
    }

    internal sealed class PictureConverter : WinForms.AxHost
    {
        private PictureConverter() : base(string.Empty) { }
        public static stdole.IPictureDisp ToPicture(System.Drawing.Image image) =>
            RibbonPictureConverter.FromBitmap(image as System.Drawing.Bitmap ?? new System.Drawing.Bitmap(image));
    }

    public static class AutoFormatToolbar
    {
        private static Window _window;
        private const string RegKey = @"Software\CAT";

        public static void Toggle()
        {
            if (_window != null) { Close(); return; }
            Show();
        }

        public static void Close()
        {
            try { _window?.Close(); } catch { }
            _window = null;
        }

        private static void Show()
        {
            var win = new Window
            {
                Title = "",
                SizeToContent = SizeToContent.WidthAndHeight,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = Brushes.Transparent,
                ShowInTaskbar = false,
                Topmost = true
            };
            win.Resources.MergedDictionaries.Add(ModernTheme.Get());

            try
            {
                IntPtr hwnd = (IntPtr)ThisAddIn.Instance.App.HWND;
                if (hwnd != IntPtr.Zero)
                    new System.Windows.Interop.WindowInteropHelper(win).Owner = hwnd;
            }
            catch { }

            RestorePosition(win);

            var shell = new Border
            {
                Background = Brushes.White,
                CornerRadius = new CornerRadius(12),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(14, 12, 14, 12),
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    BlurRadius = 18,
                    ShadowDepth = 2,
                    Opacity = 0.24,
                    Color = Colors.Black
                }
            };

            var row = new StackPanel { Orientation = Orientation.Horizontal };

            Add(row, win, "catAfLinesV", "0.5pt grey vertical lines in horizontal gaps", () => new AfInsertVerticalGapLinesCommand());
            Add(row, win, "catAfLinesH", "0.5pt grey horizontal lines in vertical gaps", () => new AfInsertHorizontalGapLinesCommand());
            Sep(row);
            Add(row, win, "catAfBracketLeft", "Brace on the left", () => new AfBracketCommand(BracketSide.Left));
            Add(row, win, "catAfBracketRight", "Brace on the right", () => new AfBracketCommand(BracketSide.Right));
            Add(row, win, "catAfBracketTop", "Brace above", () => new AfBracketCommand(BracketSide.Top));
            Add(row, win, "catAfBracketBottom", "Brace below", () => new AfBracketCommand(BracketSide.Bottom));
            Sep(row);
            Add(row, win, "catAfCalloutUp", "Callout pointing up", () => new AfCalloutCommand(CalloutDirection.Up));
            Sep(row);
            Add(row, win, "catAfKeyTakeaway", "Key takeaway: light grey fill, no outline", () => new AfKeyTakeawayCommand());
            Sep(row);
            Add(row, win, "catAfNumbers", "Numbered circles, top to bottom", () => new AfNumberCircleCommand(false));
            Add(row, win, "catAfLetters", "Lettered circles (A, B, C…)", () => new AfNumberCircleCommand(true));
            Sep(row);
            Add(row, win, "catAfRectangle", "Reshape picture from circle to rectangle", () => new AfPictureToRectangleCommand());
            Add(row, win, "catAfSquare", "Make the selection square (1:1)", () => new AfMakeSquareCommand());
            Sep(row);
            Add(row, win, "catAfTitle", "Apply slide title placeholder (fixed element)", () => new AfAdoptTemplateElementCommand(TemplateElement.Title));
            Add(row, win, "catAfSubtitle", "Apply slide subtitle placeholder (fixed element)", () => new AfAdoptTemplateElementCommand(TemplateElement.Subtitle));
            Add(row, win, "catAfFootnote", "Apply slide footnote placeholder (fixed element)", () => new AfAdoptTemplateElementCommand(TemplateElement.Footnote));
            Add(row, win, "catAfSource", "Apply slide source placeholder (fixed element)", () => new AfAdoptTemplateElementCommand(TemplateElement.Source));

            shell.Child = row;
            win.Content = shell;

            shell.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
                {
                    try { win.DragMove(); } catch { }
                }
            };

            win.KeyDown += (s, e) => { if (e.Key == System.Windows.Input.Key.Escape) win.Close(); };
            win.Closing += (s, e) => { SavePosition(win); _window = null; };

            _window = win;
            win.Show();
        }

        private static void Sep(StackPanel row) =>
            row.Children.Add(new Border
            {
                Width = 1,
                Height = 36,
                Background = new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8)),
                Margin = new Thickness(8, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center
            });

        private static void Add(StackPanel row, Window host, string id, string tip, Func<ICommand> factory)
        {
            var b = new Button
            {
                ToolTip = tip,
                Style = (Style)host.Resources["AfTool"]
            };
            b.Content = LoadIcon(id);
            b.Click += (s, e) => CommandRunner.Run(factory());
            row.Children.Add(b);
        }

        private static System.Windows.Controls.Image LoadIcon(string id)
        {
            const double size = 22;
            try
            {
                using (var bmp = IconRepository.CreateRibbonBitmap(id))
                {
                    if (bmp == null)
                        return new System.Windows.Controls.Image { Width = size, Height = size };

                    using (var ms = new MemoryStream())
                    {
                        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        ms.Position = 0;
                        var wpf = new System.Windows.Media.Imaging.BitmapImage();
                        wpf.BeginInit();
                        wpf.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                        wpf.StreamSource = ms;
                        wpf.EndInit();
                        wpf.Freeze();
                        return new System.Windows.Controls.Image
                        {
                            Source = wpf,
                            Width = size,
                            Height = size,
                            Stretch = Stretch.Uniform
                        };
                    }
                }
            }
            catch { }

            return new System.Windows.Controls.Image { Width = size, Height = size };
        }

        private static void RestorePosition(Window win)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegKey))
                {
                    if (key == null) { Centre(win); return; }
                    var l = key.GetValue("ToolbarLeft") as string;
                    var t = key.GetValue("ToolbarTop") as string;
                    if (double.TryParse(l, NumberStyles.Any, CultureInfo.InvariantCulture, out double left)
                        && double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out double top)
                        && left > -2000 && top > -2000)
                    { win.Left = left; win.Top = top; return; }
                }
            }
            catch { }
            Centre(win);
        }

        private static void Centre(Window win) =>
            win.WindowStartupLocation = WindowStartupLocation.CenterOwner;

        private static void SavePosition(Window win)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegKey))
                {
                    key.SetValue("ToolbarLeft", win.Left.ToString(CultureInfo.InvariantCulture));
                    key.SetValue("ToolbarTop", win.Top.ToString(CultureInfo.InvariantCulture));
                }
            }
            catch { }
        }
    }
}
