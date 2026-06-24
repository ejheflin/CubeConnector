/*
 * CubeConnector - EnumerateModelsSmokeTest (TEMPORARY)
 *
 * Single remaining diagnostic after probe cleanup: verifies the consolidated
 * production services still work end to end -- PowerBiAuth (silent token) +
 * PowerBiRestClient (enumerate workspaces/datasets). Remove once the wizard
 * exercises these services itself.
 */

using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CubeConnector
{
    public static class EnumerateModelsSmokeTest
    {
        private const string OutputFilePath = @"C:\dev\CubeConnector_gh\enumerate_smoketest.txt";

        public static void RunEnumerateModelsSmokeTest()
        {
            var sb = new StringBuilder();
            try
            {
                string token = PowerBiAuth.GetAccessToken(out string src, out string err);
                if (string.IsNullOrEmpty(token))
                {
                    Finish(sb.Append("AUTH FAILED: " + err).ToString(), "Auth failed - see file.");
                    return;
                }
                sb.AppendLine("Token via: " + src + "   identity: " + (PowerBiAuth.GetUpnFromToken(token) ?? "(unknown)"));
                sb.AppendLine();

                var datasets = PowerBiRestClient.GetAllDatasets(token);
                sb.AppendLine("Accessible datasets: " + datasets.Count);
                sb.AppendLine();

                string lastWs = null;
                foreach (var d in datasets)
                {
                    if (d.WorkspaceName != lastWs) { sb.AppendLine("[" + d.WorkspaceName + "]"); lastWs = d.WorkspaceName; }
                    sb.AppendLine("    " + d.Name + "  (" + d.Id + ")" +
                        (d.IsRefreshable ? "" : "  [not refreshable]"));
                }

                Finish(sb.ToString(), "Enumerated " + datasets.Count + " datasets via consolidated services.\n\nSee file.");
            }
            catch (Exception ex)
            {
                sb.AppendLine("FAILED: " + ex.GetType().Name + ": " + ex.Message);
                Finish(sb.ToString(), "Smoke test failed - see file.");
            }
        }

        private static void Finish(string content, string dialog)
        {
            string path = OutputFilePath;
            try { File.WriteAllText(path, content); }
            catch { path = Path.Combine(Path.GetTempPath(), "cc_enumerate_smoketest.txt"); File.WriteAllText(path, content); }
            MessageBox.Show(dialog + "\n\nReport: " + path, "Enumerate Models (TEST)", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
