using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace CubeConnector
{
    public class WizardWindow : Form
    {
        private readonly WebView2 _web = new WebView2();

        public WizardWindow()
        {
            Text = "CubeConnector — Manage Formulas";
            Width = 520; Height = 760; StartPosition = FormStartPosition.CenterScreen;
            _web.Dock = DockStyle.Fill;
            Controls.Add(_web);
            Load += async (s, e) => await InitAsync();
        }

        private async Task InitAsync()
        {
            try
            {
                string userData = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "CubeConnector", "WebView2");
                Directory.CreateDirectory(userData);
                var env = await CoreWebView2Environment.CreateAsync(null, userData);
                await _web.EnsureCoreWebView2Async(env);

                // Follow the OS/Office light-or-dark theme so the UI's prefers-color-scheme matches.
                try { _web.CoreWebView2.Profile.PreferredColorScheme = CoreWebView2PreferredColorScheme.Auto; } catch { }

                _web.CoreWebView2.AddHostObjectToScript("cc", new WizardBridge());

                string uiDir = Path.Combine(Path.GetDirectoryName(
                    ExcelDna.Integration.ExcelDnaUtil.XllPath), "ui");
                _web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "cubeconnector.ui", uiDir, CoreWebView2HostResourceAccessKind.Allow);
                _web.CoreWebView2.Navigate("https://cubeconnector.ui/index.html");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Couldn't start the formula manager. This feature needs the Microsoft Edge WebView2 Runtime, " +
                    "which is normally already installed.\n\nDetails: " + ex.Message,
                    "CubeConnector", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Close();
            }
        }

        private static WizardWindow _open;
        public static void ShowSingleton()
        {
            if (_open == null || _open.IsDisposed) { _open = new WizardWindow(); _open.Show(); }
            else _open.BringToFront();
        }
    }
}
