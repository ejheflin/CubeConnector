using System;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace CubeConnector
{
    /// <summary>
    /// Debounced automatic refresh. When enabled, any workbook edit (SheetChange) restarts a
    /// 3-second timer; when it elapses, the standard refresh runs (updating only cells that need
    /// it). Loop-safe: the refresh's own writes are guarded so they don't re-trigger a refresh.
    /// On/off state persists to %LOCALAPPDATA%\CubeConnector\autorefresh.txt (default off).
    /// </summary>
    internal static class AutoRefreshManager
    {
        private const int DebounceMs = 3000;
        private static Excel.Application _app;
        private static Timer _timer;          // System.Windows.Forms.Timer — ticks on the UI thread
        private static bool _attached;
        private static bool _refreshing;

        private static string StateFile =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "CubeConnector", "autorefresh.txt");

        /// <summary>Persisted on/off state. Default off when missing/unreadable.</summary>
        public static bool Enabled
        {
            get
            {
                try { return File.Exists(StateFile) &&
                             File.ReadAllText(StateFile).Trim().Equals("on", StringComparison.OrdinalIgnoreCase); }
                catch { return false; }
            }
        }

        /// <summary>Called once from AutoOpen. Attaches the change handler if persisted on.</summary>
        public static void Initialize(Excel.Application app)
        {
            _app = app;
            if (Enabled) Attach();
        }

        /// <summary>Persist the new state and attach/detach the handler live.</summary>
        public static void SetEnabled(bool on)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(StateFile));
                File.WriteAllText(StateFile, on ? "on" : "off");
            }
            catch { /* non-fatal: toggle just won't persist */ }
            if (on) Attach(); else Detach();
        }

        public static void Attach()
        {
            if (_attached || _app == null) return;
            if (_timer == null)
            {
                _timer = new Timer { Interval = DebounceMs };
                _timer.Tick += OnTick;
            }
            _app.SheetChange += OnSheetChange;
            _attached = true;
        }

        public static void Detach()
        {
            if (!_attached || _app == null) return;
            try { _app.SheetChange -= OnSheetChange; } catch { }
            try { _timer?.Stop(); } catch { }
            _attached = false;
        }

        private static void OnSheetChange(object sh, Excel.Range target)
        {
            if (_refreshing) return;       // ignore the refresh's own writes
            _timer.Stop();
            _timer.Start();                // (re)start the 3-second debounce window
        }

        private static void OnTick(object sender, EventArgs e)
        {
            _timer.Stop();
            ExcelAsyncUtil.QueueAsMacro(() => RunRefresh());   // run in a macro-safe context
        }

        private static void RunRefresh()
        {
            if (_refreshing || _app == null) return;
            _refreshing = true;
            bool prevEvents = true;
            try
            {
                prevEvents = _app.EnableEvents;
                _app.EnableEvents = false;     // suppress events while the refresh writes cells
                DynamicFunctionRegistration.RefreshCore(silent: true);   // no popups/clipboard during background refresh
            }
            catch { /* auto-refresh must never throw a modal error mid-edit */ }
            finally
            {
                try { _app.EnableEvents = prevEvents; } catch { }
                _refreshing = false;
            }
        }
    }
}
