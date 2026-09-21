using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using Cat;
using Cat.Commands;
using Cat.Core;
using Cat.UI;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Cat.TableCreator
{
    public static class TableCreatorWindow
    {
        private static Window _window;

        public static void Show()
        {
            if (_window != null)
            {
                try
                {
                    _window.Activate();
                    _window.Focus();
                }
                catch { Close(); Show(); }
                return;
            }

            var win = new Window
            {
                Title = "Table Creator",
                Width = 360,
                SizeToContent = SizeToContent.Height,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Background = System.Windows.Media.Brushes.White
            };
            win.Resources.MergedDictionaries.Add(ModernTheme.Get());

            try
            {
                IntPtr hwnd = (IntPtr)ThisAddIn.Instance.App.HWND;
                if (hwnd != IntPtr.Zero)
                    new System.Windows.Interop.WindowInteropHelper(win).Owner = hwnd;
            }
            catch { }

            var root = new Grid { Margin = new Thickness(22) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var title = new TextBlock { Text = "Table Creator", Style = (Style)win.Resources["Title"] };
            Grid.SetRow(title, 0);
            root.Children.Add(title);

            var sizePanel = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };
            sizePanel.Children.Add(new TextBlock { Text = "Table size", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 8) });

            var rowGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            rowGrid.Children.Add(new TextBlock { Text = "Rows", VerticalAlignment = VerticalAlignment.Center });
            var rowsBox = new TextBox { Text = "3" };
            Grid.SetColumn(rowsBox, 1);
            rowGrid.Children.Add(rowsBox);
            sizePanel.Children.Add(rowGrid);

            var colGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            colGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            colGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            colGrid.Children.Add(new TextBlock { Text = "Columns", VerticalAlignment = VerticalAlignment.Center });
            var colsBox = new TextBox { Text = "3" };
            Grid.SetColumn(colsBox, 1);
            colGrid.Children.Add(colsBox);
            sizePanel.Children.Add(colGrid);

            var hasHeader = new CheckBox { Content = "Table has header", IsChecked = true };
            sizePanel.Children.Add(hasHeader);

            Grid.SetRow(sizePanel, 1);
            root.Children.Add(sizePanel);

            var actions = new StackPanel { Margin = new Thickness(0, 4, 0, 0) };
            actions.Children.Add(new TextBlock { Text = "Actions", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 8) });

            var btnCreate = new Button { Content = "Create Table", Style = (Style)win.Resources["Primary"], Margin = new Thickness(0, 0, 0, 8) };
            var btnClipboard = new Button { Content = "Create Table from Clipboard", Style = (Style)win.Resources["Primary"], Margin = new Thickness(0, 0, 0, 8) };
            var btnConvert = new Button { Content = "Convert Native Table → Text Boxes", Style = (Style)win.Resources["Primary"] };

            actions.Children.Add(btnCreate);
            actions.Children.Add(btnClipboard);
            actions.Children.Add(btnConvert);

            Grid.SetRow(actions, 2);
            root.Children.Add(actions);

            var close = new Button
            {
                Content = "Close",
                Margin = new Thickness(0, 16, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Right,
                MinWidth = 80
            };
            close.Click += (s, e) => win.Close();
            Grid.SetRow(close, 3);
            root.Children.Add(close);

            bool headerChecked() => hasHeader.IsChecked == true;

            btnCreate.Click += (s, e) =>
            {
                if (!TryParseDimensions(rowsBox.Text, colsBox.Text, out int rows, out int cols, out string err))
                {
                    Notifier.Error(err);
                    return;
                }
                RunOnSlide(slide => TableCreatorService.CreateEmpty(slide, rows, cols, headerChecked()));
            };

            btnClipboard.Click += (s, e) =>
            {
                RunOnSlide(slide => TableCreatorService.CreateFromClipboard(slide, headerChecked()));
            };

            btnConvert.Click += (s, e) =>
            {
                RunConvert(headerChecked());
            };

            win.Content = root;
            win.Closing += (s, e) => { _window = null; };
            _window = win;
            win.Show();
        }

        public static void Close()
        {
            try { _window?.Close(); } catch { }
            _window = null;
        }

        private static void RunOnSlide(Func<PowerPoint.Slide, bool> action)
        {
            try
            {
                var slide = new CommandContext(ThisAddIn.Instance.App).ActiveSlide;
                if (slide == null)
                {
                    Notifier.Info("Open a slide first.");
                    return;
                }
                try { ThisAddIn.Instance.App.StartNewUndoEntry(); } catch { }
                action(slide);
            }
            catch (Exception ex) { Notifier.Error(ex.Message); }
        }

        private static void RunConvert(bool hasHeader)
        {
            try
            {
                var ctx = new CommandContext(ThisAddIn.Instance.App);
                var slide = ctx.ActiveSlide;
                if (slide == null)
                {
                    Notifier.Info("Open a slide first.");
                    return;
                }
                if (!ctx.HasShapeSelection || ctx.SelectedShapes.Count != 1)
                {
                    Notifier.Info("Select a single native PowerPoint table.");
                    return;
                }

                PowerPoint.Shape tableShape;
                try { tableShape = ctx.SelectedShapes[1]; }
                catch
                {
                    Notifier.Info("Select a single native PowerPoint table.");
                    return;
                }

                try { ThisAddIn.Instance.App.StartNewUndoEntry(); } catch { }
                TableCreatorService.ConvertNativeTable(slide, tableShape, hasHeader);
            }
            catch (Exception ex) { Notifier.Error(ex.Message); }
        }

        private static bool TryParseDimensions(string rowsText, string colsText, out int rows, out int cols, out string error)
        {
            rows = cols = 0;
            error = null;
            if (!int.TryParse(rowsText?.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out rows)
                || !int.TryParse(colsText?.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out cols))
            {
                error = "Rows and columns must be whole numbers.";
                return false;
            }
            return TableCreatorService.ValidateDimensions(rows, cols, out error);
        }
    }
}
