/*
 * CubeConnector - PowerBiAuth
 *
 * Reusable Power BI token provider with one-time sign-in + silent refresh,
 * with NO Azure AD app registration and NO admin consent:
 *
 *   - Uses a well-known, pre-consented first-party PUBLIC client ID (Azure CLI).
 *   - First call: browser authorization-code + PKCE flow (loopback TcpListener).
 *     This passes Conditional Access where device-code is blocked.
 *   - Caches the resulting REFRESH token via Windows DPAPI (CurrentUser scope),
 *     so later calls redeem it silently (no browser). Falls back to interactive
 *     only if the refresh token is missing/expired/revoked.
 *
 * Pure System.Net + System.Security (DPAPI). No MSAL, no native DLLs.
 * Tokens are held only in memory + the DPAPI-encrypted refresh-token file.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace CubeConnector
{
    public static class PowerBiAuth
    {
        private const string ClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"; // Azure CLI (loopback redirect)
        private const string Authority = "https://login.microsoftonline.com/organizations/oauth2/v2.0";
        private const string Scope = "https://analysis.windows.net/powerbi/api/.default offline_access openid profile";

        static PowerBiAuth()
        {
            // .NET Framework defaults can negotiate a TLS version Azure rejects in a fresh
            // process; ensure TLS 1.2 before any HTTPS token call.
            try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12; } catch { }
        }

        // Single scope (no extras) for the WAM helper.
        private const string WamScope = "https://analysis.windows.net/powerbi/api/.default";
        // Dev fallback path for the out-of-process WAM helper.
        private const string WamHelperDevPath = @"C:\dev\CubeConnector_gh\WamHelper\bin\Release\WamHelper.exe";

        private static string AppDataDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "CubeConnector");
            }
        }

        private static string CacheFile { get { return Path.Combine(AppDataDir, "pbi_refresh.bin"); } }
        private static string ModeFile { get { return Path.Combine(AppDataDir, "pbi_mode.txt"); } }

        // Account mode: "wam" = use Windows logon identity (zero-click, default);
        //               "browser" = use a specific account the user signed into via browser.
        private static string GetMode()
        {
            try { return File.Exists(ModeFile) ? File.ReadAllText(ModeFile).Trim() : "wam"; }
            catch { return "wam"; }
        }

        private static void SetMode(string mode)
        {
            try { Directory.CreateDirectory(AppDataDir); File.WriteAllText(ModeFile, mode); } catch { }
        }

        /// <summary>
        /// Returns a Power BI access token. Tries the cached refresh token first (silent);
        /// falls back to interactive browser sign-in. Returns null on failure.
        /// </summary>
        public static string GetAccessToken(out string source, out string error)
        {
            source = null; error = null;
            bool preferBrowser = GetMode() == "browser";

            // 1. Zero-click SSO via WAM (Windows session identity) -- unless the user has
            //    explicitly chosen a specific (different) account.
            if (!preferBrowser)
            {
                string wamTok = TryWamSilent(out _);
                if (!string.IsNullOrEmpty(wamTok)) { source = "WAM zero-click SSO"; return wamTok; }
            }

            // 2. Silent: redeem a cached refresh token (the chosen browser account).
            string cachedTok = TryCachedRefresh();
            if (!string.IsNullOrEmpty(cachedTok)) { source = "silent (cached refresh token)"; return cachedTok; }

            // 3. Interactive: browser auth-code + PKCE.
            string itok = Interactive(out string ierr);
            if (!string.IsNullOrEmpty(itok)) { source = "interactive (browser)"; return itok; }
            error = ierr ?? "Authentication failed.";
            return null;
        }

        /// <summary>
        /// Force sign-in as a specific (possibly different) account via the browser, and
        /// REMEMBER that choice so future silent calls use it instead of the Windows identity.
        /// Returns the access token, or null (with error) on failure.
        /// </summary>
        public static string SignInAsDifferentAccount(out string error)
        {
            error = null;
            SignOut();           // clear any cached refresh token
            SetMode("browser");  // stick to the browser-chosen account
            string tok = Interactive(out string ierr);
            if (string.IsNullOrEmpty(tok)) { error = ierr ?? "Sign-in failed."; return null; }
            return tok;
        }

        /// <summary>Revert to using the Windows logon identity (zero-click WAM) going forward.</summary>
        public static void UseWindowsAccount()
        {
            SignOut();
            SetMode("wam");
        }

        private static string TryCachedRefresh()
        {
            string rt = LoadRefreshToken();
            if (string.IsNullOrEmpty(rt)) return null;
            string body = PostForm(Authority + "/token", new Dictionary<string, string>
            {
                { "client_id", ClientId },
                { "grant_type", "refresh_token" },
                { "refresh_token", rt },
                { "scope", Scope }
            }, out _);
            string tok = JsonStr(body, "access_token");
            if (!string.IsNullOrEmpty(tok))
            {
                string newRt = JsonStr(body, "refresh_token"); // AAD rotates these
                if (!string.IsNullOrEmpty(newRt)) SaveRefreshToken(newRt);
                return tok;
            }
            SignOut(); // expired/revoked
            return null;
        }

        /// <summary>Current account mode: "wam" (Windows identity) or "browser" (chosen account).</summary>
        public static string AccountMode { get { return GetMode(); } }

        /// <summary>
        /// Decode the user identity (UPN / email) from a Power BI access token (a JWT).
        /// No signature validation -- display only. Returns null if not extractable.
        /// </summary>
        public static string GetUpnFromToken(string accessToken)
        {
            try
            {
                var parts = (accessToken ?? "").Split('.');
                if (parts.Length < 2) return null;
                string p = parts[1].Replace('-', '+').Replace('_', '/');
                switch (p.Length % 4) { case 2: p += "=="; break; case 3: p += "="; break; }
                string json = Encoding.UTF8.GetString(Convert.FromBase64String(p));
                return JsonStr(json, "upn") ?? JsonStr(json, "unique_name")
                    ?? JsonStr(json, "preferred_username") ?? JsonStr(json, "email");
            }
            catch { return null; }
        }

        /// <summary>
        /// Extract the tenant id (tid claim) from a Power BI access token (a JWT).
        /// Returns "common" if not extractable. No signature validation -- display/routing only.
        /// </summary>
        public static string GetTidFromTokenPublic(string token)
        {
            try {
                var parts = (token ?? "").Split('.');
                if (parts.Length < 2) return "common";
                string p = parts[1].Replace('-', '+').Replace('_', '/');
                switch (p.Length % 4) { case 2: p += "=="; break; case 3: p += "="; break; }
                string json = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(p));
                var m = System.Text.RegularExpressions.Regex.Match(json, "\"tid\"\\s*:\\s*\"([^\"]*)\"");
                return m.Success ? m.Groups[1].Value : "common";
            } catch { return "common"; }
        }

        /// <summary>Delete the cached refresh token (forces interactive sign-in next time).</summary>
        public static void SignOut()
        {
            try { if (File.Exists(CacheFile)) File.Delete(CacheFile); } catch { }
        }

        public static bool HasCachedSignIn { get { return File.Exists(CacheFile); } }

        // ---- WAM zero-click SSO (out-of-process helper) ----

        /// <summary>
        /// Attempt a zero-click token via WamHelper.exe (MSAL + Windows broker, PRT).
        /// Runs out-of-process so the native broker never loads inside Excel. Returns
        /// null on any failure (helper missing, broker unavailable, no PRT, etc.).
        /// </summary>
        private static string TryWamSilent(out string error)
        {
            error = null;
            string exe = ResolveWamHelperPath();
            if (exe == null) { error = "WamHelper.exe not found"; return null; }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = "silent " + ClientId + " organizations " + WamScope,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(psi))
                {
                    string stdout = p.StandardOutput.ReadToEnd();
                    p.WaitForExit(60000);

                    foreach (var raw in stdout.Split('\n'))
                    {
                        string line = raw.Trim();
                        if (line.StartsWith("ACCESS_TOKEN=")) return line.Substring("ACCESS_TOKEN=".Length);
                        if (line.StartsWith("ERROR=")) { error = line.Substring("ERROR=".Length); return null; }
                    }
                    error = "no token in helper output";
                    return null;
                }
            }
            catch (Exception ex) { error = ex.GetType().Name + ": " + ex.Message; return null; }
        }

        /// <summary>Find WamHelper.exe next to the add-in, else fall back to the dev build path.</summary>
        private static string ResolveWamHelperPath()
        {
            try
            {
                string xll = ExcelDna.Integration.ExcelDnaUtil.XllPath;
                string dir = Path.GetDirectoryName(xll);
                if (dir != null)
                {
                    string beside = Path.Combine(dir, "WamHelper.exe");
                    if (File.Exists(beside)) return beside;
                    string sub = Path.Combine(dir, "WamHelper", "WamHelper.exe");
                    if (File.Exists(sub)) return sub;
                }
            }
            catch { }
            return File.Exists(WamHelperDevPath) ? WamHelperDevPath : null;
        }

        // ---- interactive browser auth-code flow ----

        private static string Interactive(out string error)
        {
            error = null;
            TcpListener listener = null;
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                listener = new TcpListener(IPAddress.Loopback, 0);
                listener.Start();
                int port = ((IPEndPoint)listener.LocalEndpoint).Port;
                string redirectUri = "http://localhost:" + port;

                string verifier = Base64Url(RandomBytes(32));
                string challenge = Base64Url(Sha256(Encoding.ASCII.GetBytes(verifier)));
                string state = Base64Url(RandomBytes(16));

                string authUrl = Authority + "/authorize?client_id=" + ClientId +
                    "&response_type=code&redirect_uri=" + Uri.EscapeDataString(redirectUri) +
                    "&response_mode=query&scope=" + Uri.EscapeDataString(Scope) +
                    "&state=" + state +
                    "&code_challenge=" + challenge + "&code_challenge_method=S256";

                var acceptTask = listener.AcceptTcpClientAsync();
                Process.Start(new ProcessStartInfo(authUrl) { UseShellExecute = true });

                if (!acceptTask.Wait(TimeSpan.FromSeconds(180)))
                {
                    error = "Timed out waiting for the browser redirect.";
                    return null;
                }

                string requestLine;
                using (var client = acceptTask.Result)
                using (var stream = client.GetStream())
                {
                    var buf = new byte[8192];
                    int read = stream.Read(buf, 0, buf.Length);
                    requestLine = Encoding.ASCII.GetString(buf, 0, read).Split('\n')[0];

                    string html = "<html><body style='font-family:sans-serif'><h3>CubeConnector</h3>" +
                        "<p>Sign-in complete. You can close this tab and return to Excel.</p></body></html>";
                    byte[] hb = Encoding.UTF8.GetBytes(html);
                    byte[] hdr = Encoding.ASCII.GetBytes(
                        "HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: "
                        + hb.Length + "\r\nConnection: close\r\n\r\n");
                    stream.Write(hdr, 0, hdr.Length);
                    stream.Write(hb, 0, hb.Length);
                    stream.Flush();
                }

                string query = "";
                var qm = Regex.Match(requestLine, @"GET\s+\/\?([^\s]*)\s");
                if (qm.Success) query = qm.Groups[1].Value;

                string code = QueryParam(query, "code");
                string err = QueryParam(query, "error");
                if (!string.IsNullOrEmpty(err) || string.IsNullOrEmpty(code))
                {
                    error = "Authorize step failed: " + err + " " +
                        Uri.UnescapeDataString(QueryParam(query, "error_description") ?? "");
                    return null;
                }

                string tkBody = PostForm(Authority + "/token", new Dictionary<string, string>
                {
                    { "client_id", ClientId },
                    { "grant_type", "authorization_code" },
                    { "code", code },
                    { "redirect_uri", redirectUri },
                    { "code_verifier", verifier },
                    { "scope", Scope }
                }, out _);

                string token = JsonStr(tkBody, "access_token");
                if (string.IsNullOrEmpty(token))
                {
                    error = "Token exchange failed: " + (JsonStr(tkBody, "error_description") ?? tkBody);
                    return null;
                }

                string refresh = JsonStr(tkBody, "refresh_token");
                if (!string.IsNullOrEmpty(refresh)) SaveRefreshToken(refresh);
                return token;
            }
            catch (Exception ex)
            {
                error = ex.GetType().Name + ": " + ex.Message;
                return null;
            }
            finally
            {
                try { if (listener != null) listener.Stop(); } catch { }
            }
        }

        // ---- DPAPI refresh-token cache ----

        private static void SaveRefreshToken(string rt)
        {
            try
            {
                string dir = Path.GetDirectoryName(CacheFile);
                Directory.CreateDirectory(dir);
                byte[] enc = ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(rt), null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(CacheFile, enc);
            }
            catch { /* non-fatal: silent re-auth just won't be available */ }
        }

        private static string LoadRefreshToken()
        {
            try
            {
                if (!File.Exists(CacheFile)) return null;
                byte[] dec = ProtectedData.Unprotect(
                    File.ReadAllBytes(CacheFile), null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(dec);
            }
            catch { return null; }
        }

        // ---- helpers ----

        private static string PostForm(string url, Dictionary<string, string> form, out HttpStatusCode status)
        {
            var sb = new StringBuilder();
            foreach (var kv in form)
            {
                if (sb.Length > 0) sb.Append('&');
                sb.Append(Uri.EscapeDataString(kv.Key)).Append('=').Append(Uri.EscapeDataString(kv.Value));
            }
            byte[] data = Encoding.UTF8.GetBytes(sb.ToString());
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = "application/x-www-form-urlencoded";
            req.ContentLength = data.Length;
            using (var rs = req.GetRequestStream()) rs.Write(data, 0, data.Length);
            try
            {
                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    status = resp.StatusCode;
                    using (var sr = new StreamReader(resp.GetResponseStream())) return sr.ReadToEnd();
                }
            }
            catch (WebException wex) when (wex.Response is HttpWebResponse er)
            {
                status = er.StatusCode;
                using (var sr = new StreamReader(er.GetResponseStream())) return sr.ReadToEnd();
            }
        }

        private static string JsonStr(string json, string key)
        {
            var m = Regex.Match(json ?? "", "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : null;
        }

        private static string QueryParam(string query, string key)
        {
            var m = Regex.Match(query ?? "", "(?:^|&)" + Regex.Escape(key) + "=([^&]*)");
            return m.Success ? m.Groups[1].Value : null;
        }

        private static byte[] RandomBytes(int n)
        {
            var b = new byte[n];
            using (var rng = new RNGCryptoServiceProvider()) rng.GetBytes(b);
            return b;
        }

        private static byte[] Sha256(byte[] data)
        {
            using (var sha = SHA256.Create()) return sha.ComputeHash(data);
        }

        private static string Base64Url(byte[] b)
        {
            return Convert.ToBase64String(b).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        }
    }
}
