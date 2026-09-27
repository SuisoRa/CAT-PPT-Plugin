using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Cat.Core;
using Cat.TableCreator;
using Cat.TextBlocks;
using Cat.UI;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

[assembly: AssemblyTitle("CAT")]
[assembly: AssemblyDescription("PowerPoint keyboard shortcuts and Auto format toolbar")]
[assembly: AssemblyProduct("CAT")]
[assembly: ComVisible(false)]
[assembly: Guid("c8e4b2a1-7d3f-4c9e-8a12-5b6d0e1f3a88")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace Cat
{
    public partial class ThisAddIn
    {
        public static ThisAddIn Instance { get; private set; }

        private KeyboardHook _hook;
        private ShortcutDispatcher _dispatcher;

        public SelectionOrderTracker SelectionTracker { get; private set; }
        public PositionClipboard PositionClipboard { get; private set; }
        public SpacingClipboard SpacingClipboard { get; private set; }
        public RibbonController Ribbon { get; private set; }

        public PowerPoint.Application App => this.Application;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Instance = this;
            AddInSettings.Load();

            SelectionTracker = new SelectionOrderTracker(this.Application);
            PositionClipboard = new PositionClipboard();
            SpacingClipboard = new SpacingClipboard();

            _dispatcher = new ShortcutDispatcher(this.Application);
            _hook = new KeyboardHook(_dispatcher);
            _hook.Install();

            try { Ribbon?.RefreshRibbonImages(); } catch { }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try { AutoFormatToolbar.Close(); } catch { }
            try { TableCreatorWindow.Close(); } catch { }
            try { TextBlockWindow.Close(); } catch { }
            try { _hook?.Uninstall(); } catch { }
            try { SelectionTracker?.Dispose(); } catch { }
            AddInSettings.Save();
            Instance = null;
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Ribbon = new RibbonController();
            return Ribbon;
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
        #endregion
    }
}
