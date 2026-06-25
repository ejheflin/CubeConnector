using System;
using System.IO;
using System.Reflection;
using Microsoft.Web.WebView2.Core;

namespace CubeConnector
{
    /// <summary>
    /// Serves the wizard UI (index.html, app.js, styles.css, logo) to a WebView2 from
    /// assembly-embedded resources, so no on-disk ui\ folder needs to ship beside the .xll.
    /// Requests to https://cubeconnector.ui/* are answered from embedded resources named
    /// "CubeConnector.ui.<relative-path-with-dots>".
    /// </summary>
    internal static class UiHost
    {
        private const string Origin = "https://cubeconnector.ui/";
        private const string ResourcePrefix = "CubeConnector.ui.";

        public static void AttachEmbeddedUi(CoreWebView2 core, CoreWebView2Environment env)
        {
            core.AddWebResourceRequestedFilter(Origin + "*", CoreWebView2WebResourceContext.All);
            core.WebResourceRequested += (s, e) =>
            {
                try
                {
                    string uri = e.Request.Uri;
                    string rel = uri.Length > Origin.Length ? uri.Substring(Origin.Length) : "";
                    int cut = rel.IndexOfAny(new[] { '?', '#' });
                    if (cut >= 0) rel = rel.Substring(0, cut);
                    if (string.IsNullOrEmpty(rel)) rel = "index.html";

                    byte[] bytes = ReadResource(rel);
                    if (bytes == null) { e.Response = env.CreateWebResourceResponse(null, 404, "Not Found", ""); return; }
                    e.Response = env.CreateWebResourceResponse(new MemoryStream(bytes), 200, "OK", "Content-Type: " + ContentType(rel));
                }
                catch
                {
                    e.Response = env.CreateWebResourceResponse(null, 500, "Error", "");
                }
            };
        }

        private static byte[] ReadResource(string rel)
        {
            string name = ResourcePrefix + rel.Replace('/', '.');
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(name))
            {
                if (stream == null) return null;
                using (var ms = new MemoryStream()) { stream.CopyTo(ms); return ms.ToArray(); }
            }
        }

        private static string ContentType(string rel)
        {
            switch (Path.GetExtension(rel).ToLowerInvariant())
            {
                case ".html": return "text/html; charset=utf-8";
                case ".js":   return "text/javascript; charset=utf-8";
                case ".css":  return "text/css; charset=utf-8";
                case ".png":  return "image/png";
                case ".svg":  return "image/svg+xml";
                case ".json": return "application/json; charset=utf-8";
                default:      return "application/octet-stream";
            }
        }
    }
}
