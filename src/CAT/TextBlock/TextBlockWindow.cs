using System;

using System.Windows;

using System.Windows.Controls;

using System.Windows.Input;

using System.Windows.Media;

using Cat.Core;

using Cat.UI;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using WpfTextBlock = System.Windows.Controls.TextBlock;



namespace Cat.TextBlocks

{

    public static class TextBlockWindow

    {

        private static Window _window;

        private static readonly SolidColorBrush FrameBorder = new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8));

        private static readonly SolidColorBrush FrameBackground = new SolidColorBrush(Color.FromRgb(0xFA, 0xFA, 0xFA));

        private static readonly SolidColorBrush HoverBackground = new SolidColorBrush(Color.FromRgb(0xF0, 0xF4, 0xFF));

        private static readonly SolidColorBrush MutedForeground = new SolidColorBrush(Color.FromRgb(0x6B, 0x6B, 0x6B));



        public static void Show()

        {

            if (_window != null)

            {

                try { _window.Activate(); return; }

                catch { Close(); }

            }



            var win = new Window

            {

                Title = "Text block",

                Width = 420,

                SizeToContent = SizeToContent.Height,

                ResizeMode = ResizeMode.NoResize,

                WindowStartupLocation = WindowStartupLocation.CenterScreen,

                Background = Brushes.White

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



            var title = new WpfTextBlock { Text = "Text block", Style = (Style)win.Resources["Title"] };

            Grid.SetRow(title, 0);

            root.Children.Add(title);



            var hint = new WpfTextBlock

            {

                Text = "Click a preview to insert that block on the active slide. Select chevron blocks and press Ctrl+T to toggle shape.",

                TextWrapping = TextWrapping.Wrap,

                Foreground = MutedForeground,

                Margin = new Thickness(0, 0, 0, 12)

            };

            Grid.SetRow(hint, 1);

            root.Children.Add(hint);



            var catalog = new StackPanel();

            catalog.Children.Add(BuildClickablePreview(

                ChevronTextBlockSpec.PreviewFileName,

                "Chevron text block",

                () => TryInsert(slide => ChevronTextBlockService.InsertChevron(slide), null)));

            catalog.Children.Add(BuildClickablePreview(

                ActionItemTextBlockSpec.PreviewFileName,

                "Numbered action item",

                () => TryInsert(slide => ActionItemTextBlockService.Insert(slide), null)));

            Grid.SetRow(catalog, 2);

            root.Children.Add(catalog);



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



            win.Content = root;

            win.Closing += (s, e) => _window = null;

            _window = win;

            win.Show();

        }



        private static void TryInsert(Func<PowerPoint.Slide, bool> insert, string successMessage)

        {

            try

            {

                var slide = ThisAddIn.Instance?.App?.ActiveWindow?.View?.Slide as PowerPoint.Slide;

                if (slide == null)

                {

                    Notifier.Info("Open a slide first.");

                    return;

                }

                if (insert(slide) && !string.IsNullOrEmpty(successMessage))

                    Notifier.Info(successMessage);

            }

            catch (Exception ex) { Notifier.Error(ex.Message); }

        }



        private static FrameworkElement BuildClickablePreview(string previewFileName, string label, Action onInsert)

        {

            var card = new Border

            {

                BorderBrush = FrameBorder,

                BorderThickness = new Thickness(1),

                Background = FrameBackground,

                Padding = new Thickness(12),

                MinHeight = 100,

                Margin = new Thickness(0, 0, 0, 10),

                Cursor = Cursors.Hand,

                ToolTip = "Click to insert"

            };



            card.MouseEnter += (s, e) => card.Background = HoverBackground;

            card.MouseLeave += (s, e) => card.Background = FrameBackground;

            card.MouseLeftButtonUp += (s, e) => onInsert();



            var stack = new StackPanel();

            stack.Children.Add(new WpfTextBlock

            {

                Text = label,

                FontWeight = FontWeights.SemiBold,

                Margin = new Thickness(0, 0, 0, 8)

            });



            var bmp = TextBlockPreviewHelper.TryLoadPreviewImage(previewFileName);

            if (bmp != null)

            {

                stack.Children.Add(new Image

                {

                    Source = bmp,

                    Stretch = Stretch.Uniform,

                    HorizontalAlignment = HorizontalAlignment.Center,

                    MaxHeight = 120

                });

            }

            else

            {

                stack.Children.Add(new WpfTextBlock

                {

                    Text = $"Preview: assets\\text-block\\{previewFileName}",

                    TextWrapping = TextWrapping.Wrap,

                    Foreground = MutedForeground,

                    HorizontalAlignment = HorizontalAlignment.Center,

                    TextAlignment = TextAlignment.Center

                });

            }



            card.Child = stack;

            return card;

        }



        public static void Close()

        {

            try { _window?.Close(); } catch { }

            _window = null;

        }

    }

}


