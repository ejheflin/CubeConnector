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
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace CubeConnector
{
    public class WizardPaneControl : UserControl
    {
        private readonly WebView2 _web = new WebView2();

        public WizardPaneControl()
        {
            _web.Dock = DockStyle.Fill;
            Controls.Add(_web);
            Load += async (s, e) => await InitAsync();
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

                _web.CoreWebView2.AddHostObjectToScript("cc", new WizardBridge());

                string uiDir = Path.Combine(
                    Path.GetDirectoryName(ExcelDna.Integration.ExcelDnaUtil.XllPath), "ui");
                _web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "cubeconnector.ui", uiDir, CoreWebView2HostResourceAccessKind.Allow);
                _web.CoreWebView2.Navigate("https://cubeconnector.ui/index.html");
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
