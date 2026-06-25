/*
 * CubeConnector - OfficeTheme
 *
 * Maps the current Office UI theme to a WebView2 color scheme, so the task pane follows
 * the user's *Office* theme (File > Account > Office Theme) rather than the OS theme.
 *
 * Office stores the choice in HKCU\Software\Microsoft\Office\16.0\Common\UI Theme:
 *   0 = Colorful, 3 = Dark Gray, 5 = White  -> light content area
 *   4 = Black                               -> dark content area
 * Only "Black" gives Office a dark work area, so only it maps to Dark. If the value is
 * missing or unrecognized (e.g. "use system setting"), fall back to Auto (follow the OS).
 */

using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;

namespace CubeConnector
{
    internal static class OfficeTheme
    {
        public static CoreWebView2PreferredColorScheme Scheme()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Common"))
                {
                    object v = key?.GetValue("UI Theme");
                    if (v is int t)
                    {
                        if (t == 4) return CoreWebView2PreferredColorScheme.Dark;
                        if (t == 0 || t == 3 || t == 5) return CoreWebView2PreferredColorScheme.Light;
                    }
                }
            }
            catch { /* fall through to Auto */ }
            return CoreWebView2PreferredColorScheme.Auto;
        }
    }
}
