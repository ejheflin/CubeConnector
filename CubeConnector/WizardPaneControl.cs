/*
 * CubeConnector - WizardPaneControl
 *
 * A WinForms UserControl that hosts the WebView2-based wizard UI inside an
 * Excel Custom Task Pane (CTP).  The CTP factory requires a parameterless
 * constructor and a UserControl subclass — this class provides both.
 *
 * On Load it:
 *   1. Initialises WebView2 exactly as WizardWindow does (same userData dir,
 *      same virtual-host mapping, same WizardBridge registration).
 *   2. Kicks off a background prefetch of the dataset list so the model
 *      dropdown is populated before the user clicks it.
 *
 * On failure it shows a friendly label inside the pane instead of a MessageBox.
 */

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace CubeConnector
{
    // Excel Custom Task Panes host an ActiveX/COM control, so the UserControl must be
    // COM-visible with a stable CLSID — otherwise CreateCustomTaskPane fails with
    // "Unable to create specified ActiveX control". (On .NET Framework, ComVisible + Guid
    // is sufficient; the ComDefaultInterface dance is only needed on .NET 6+.)
    [ComVisible(true)]
    [Guid("7F3A2B14-9C6D-4E58-B1A2-3D4E5F6A7B8C")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class WizardPaneControl : UserControl
    {
        private readonly WebView2 _web = new WebView2();

        public WizardPaneControl()
        {
            _web.Dock = DockStyle.Fill;
            Controls.Add(_web);
            Load += async (s, e) => await InitAsync();
            HandleDestroyed += (s, e) => PaneFocusFix.Uninstall();
        }

        private async System.Threading.Tasks.Task InitAsync()
        {
            // Kick off dataset prefetch immediately — best-effort, fire and forget.
            // Discard is intentional: we never await the cache warm-up.
            var _ = System.Threading.Tasks.Task.Run(() => PowerBiRestClient.WarmDatasetCache());

            try
            {
                string userData = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "CubeConnector", "WebView2");
                Directory.CreateDirectory(userData);

                var env = await CoreWebView2Environment.CreateAsync(null, userData);
                await _web.EnsureCoreWebView2Async(env);

                // Match the user's Office theme (not the OS theme) for the UI's prefers-color-scheme.
                try { _web.CoreWebView2.Profile.PreferredColorScheme = OfficeTheme.Scheme(); } catch { }

                _web.CoreWebView2.AddHostObjectToScript("cc", new WizardBridge());

                UiHost.AttachEmbeddedUi(_web.CoreWebView2, env);
                _web.CoreWebView2.Navigate("https://cubeconnector.ui/index.html");

                // WinEvent OUTOFCONTEXT callbacks are delivered via the installing thread's message
                // loop. This async continuation may be off the UI thread (which would explain zero
                // callbacks), so install on the control's UI thread.
                IntPtr h = _web.Handle;
                if (this.InvokeRequired) this.BeginInvoke((Action)(() => PaneFocusFix.Install(h)));
                else PaneFocusFix.Install(h);
            }
            catch (Exception ex)
            {
                // Do NOT use MessageBox from a UserControl Load handler — add a label instead.
                _web.Visible = false;
                var lbl = new Label
                {
                    Text = "Couldn't start the formula manager. This feature needs the Microsoft " +
                           "Edge WebView2 Runtime, which is normally already installed.\n\n" +
                           "Details: " + ex.Message,
                    Dock = DockStyle.Fill,
                    Padding = new Padding(12),
                    TextAlign = System.Drawing.ContentAlignment.TopLeft
                };
                Controls.Add(lbl);
            }
        }
    }
}
