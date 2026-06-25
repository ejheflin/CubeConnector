using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CubeConnector
{
    /// <summary>
    /// Materializes native runtime dependencies that cannot be loaded from inside the packed
    /// .xll (see Excel-DNA issue #488). Extracts the embedded WebView2Loader.dll to
    /// %LOCALAPPDATA%\CubeConnector\runtime and puts that folder on the DLL search path so
    /// WebView2 finds the loader. Best-effort: failures fall back to a loose loader if present.
    /// </summary>
    internal static class RuntimeBootstrap
    {
        private const string LoaderResource = "CubeConnector.native.WebView2Loader.dll";

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        private static string RuntimeDir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "CubeConnector", "runtime");

        public static void EnsureWebView2Loader()
        {
            try
            {
                string dir = RuntimeDir;
                Directory.CreateDirectory(dir);
                string dest = Path.Combine(dir, "WebView2Loader.dll");

                byte[] embedded = ReadEmbedded();
                if (embedded == null) return; // nothing to extract; rely on a loose loader if present

                if (!File.Exists(dest) || new FileInfo(dest).Length != embedded.Length)
                    File.WriteAllBytes(dest, embedded);

                SetDllDirectory(dir);   // add the runtime folder to the loader search path
                LoadLibrary(dest);      // pre-resident so WebView2's by-name LoadLibrary resolves to this
            }
            catch { /* best-effort: a loose WebView2Loader.dll beside the .xll still works */ }
        }

        private static byte[] ReadEmbedded()
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var s = asm.GetManifestResourceStream(LoaderResource))
            {
                if (s == null) return null;
                using (var ms = new MemoryStream()) { s.CopyTo(ms); return ms.ToArray(); }
            }
        }
    }
}
