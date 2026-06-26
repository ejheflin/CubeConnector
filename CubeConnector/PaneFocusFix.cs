using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CubeConnector
{
    /// <summary>
    /// Works around WebView2Feedback #951: a WebView2 hosted in an Excel task pane keeps Win32
    /// keyboard focus after the user clicks back into the grid, so typing doesn't start cell
    /// editing (you must double-click). While the pane is open we run an in-process
    /// EVENT_OBJECT_FOCUS WinEvent hook: when an Excel grid window (class "EXCEL7") raises a focus
    /// event while keyboard focus is still stuck on our WebView2 subtree, we push focus back to the
    /// grid. Hook is scoped to this process and active only while installed; inert otherwise.
    /// </summary>
    internal static class PaneFocusFix
    {
        private const uint EVENT_OBJECT_FOCUS = 0x8005;
        private const uint WINEVENT_OUTOFCONTEXT = 0x0000;

        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
            WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);
        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);
        [DllImport("user32.dll")]
        private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct GUITHREADINFO
        {
            public int cbSize;
            public uint flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        private static IntPtr _hook;
        private static IntPtr _webViewHwnd;
        private static WinEventDelegate _callback;   // field-rooted so the GC can't collect it while hooked
        private static bool _busy;

        public static void Install(IntPtr webViewHwnd)
        {
            if (_hook != IntPtr.Zero) Uninstall();
            _webViewHwnd = webViewHwnd;
            _callback = WinEventProc;
            uint pid = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            _hook = SetWinEventHook(EVENT_OBJECT_FOCUS, EVENT_OBJECT_FOCUS, IntPtr.Zero,
                _callback, pid, 0, WINEVENT_OUTOFCONTEXT);
        }

        public static void Uninstall()
        {
            if (_hook != IntPtr.Zero) { try { UnhookWinEvent(_hook); } catch { } _hook = IntPtr.Zero; }
            _callback = null;
            _webViewHwnd = IntPtr.Zero;
        }

        private static void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (_busy || _webViewHwnd == IntPtr.Zero || hwnd == IntPtr.Zero) return;
            try
            {
                // Only react to focus events from an Excel worksheet grid window.
                if (GetClass(hwnd) != "EXCEL7") return;

                // Is keyboard focus currently stuck on our WebView2 subtree?
                var gti = new GUITHREADINFO { cbSize = Marshal.SizeOf(typeof(GUITHREADINFO)) };
                if (!GetGUIThreadInfo(0, ref gti)) return;
                IntPtr focus = gti.hwndFocus;
                bool stuckOnPane = focus != IntPtr.Zero &&
                                   (focus == _webViewHwnd || IsChild(_webViewHwnd, focus));
                if (!stuckOnPane) return;

                // Hand keyboard focus back to the Excel grid.
                _busy = true;
                uint myThread = GetCurrentThreadId();
                uint gridThread = GetWindowThreadProcessId(hwnd, out _);
                bool attached = false;
                try
                {
                    if (gridThread != myThread) attached = AttachThreadInput(myThread, gridThread, true);
                    SetFocus(hwnd);
                }
                finally
                {
                    if (attached) AttachThreadInput(myThread, gridThread, false);
                    _busy = false;
                }
            }
            catch { _busy = false; }
        }

        private static string GetClass(IntPtr hwnd)
        {
            var sb = new StringBuilder(64);
            int n = GetClassName(hwnd, sb, sb.Capacity);
            return n > 0 ? sb.ToString() : "";
        }
    }
}
