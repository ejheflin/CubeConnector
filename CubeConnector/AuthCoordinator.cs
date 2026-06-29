using System;
using System.Threading;
using System.Threading.Tasks;

namespace CubeConnector
{
    /// <summary>
    /// Runs the Power BI sign-in cascade on a background thread so the Excel UI thread is never
    /// blocked. Holds sign-in state + a cached access token. All members are fast and safe to call
    /// from the UI thread (WebView2 host-object calls). The UI polls Snapshot() for live state.
    /// </summary>
    internal static class AuthCoordinator
    {
        private enum St { SignedOut, SigningIn, Ready, Error }

        private static readonly object _lock = new object();
        private static St _status = St.SignedOut;
        private static string _upn, _source, _error, _accessToken;
        private static DateTime _expiresUtc = DateTime.MinValue;
        private static CancellationTokenSource _cts;

        internal sealed class Snap { public string status; public string upn; public string source; public string error; }

        public static Snap Snapshot()
        {
            lock (_lock)
            {
                string s = _status == St.SigningIn ? "signing-in"
                         : _status == St.Ready ? "ready"
                         : _status == St.Error ? "error" : "signed-out";
                return new Snap { status = s, upn = _upn, source = _source, error = _error };
            }
        }

        /// <summary>Cached, unexpired access token (2-min skew), or null. Never does network.</summary>
        public static string TokenIfReady()
        {
            lock (_lock)
            {
                if (_status == St.Ready && _accessToken != null && DateTime.UtcNow < _expiresUtc.AddMinutes(-2))
                    return _accessToken;
                return null;
            }
        }

        public static void EnsureSignedIn() { Start(false); }
        public static void SignInDifferent() { Cancel(); PowerBiAuth.PrepareDifferentAccount(); Start(true); }
        public static void UseWindowsAccount() { Cancel(); PowerBiAuth.PrepareWindowsAccount(); Start(false); }

        public static void Cancel()
        {
            CancellationTokenSource cts;
            lock (_lock) { cts = _cts; if (_status == St.SigningIn) { _status = St.SignedOut; _error = null; } }
            try { cts?.Cancel(); } catch { }
        }

        private static void Start(bool forceInteractive)
        {
            CancellationToken ct;
            lock (_lock)
            {
                if (!forceInteractive && _status == St.Ready && _accessToken != null
                    && DateTime.UtcNow < _expiresUtc.AddMinutes(-2)) return;   // already good
                if (_status == St.SigningIn) return;                            // already in flight
                _status = St.SigningIn; _error = null;
                _cts = new CancellationTokenSource();
                ct = _cts.Token;
            }
            Task.Run(() => RunCascade(forceInteractive, ct));
        }

        private static void RunCascade(bool forceInteractive, CancellationToken ct)
        {
            string token = null, source = null, error = null;
            try { token = PowerBiAuth.AcquireToken(forceInteractive, ct, out source, out error); }
            catch (OperationCanceledException) { error = "canceled"; }
            catch (Exception ex) { error = ex.Message; }

            lock (_lock)
            {
                if (ct.IsCancellationRequested)
                {
                    _status = St.SignedOut; _error = null; _accessToken = null;
                    return;
                }
                if (!string.IsNullOrEmpty(token))
                {
                    _accessToken = token;
                    _expiresUtc = ExpiryFromJwt(token);
                    _upn = PowerBiAuth.GetUpnFromToken(token);
                    _source = source; _error = null;
                    _status = St.Ready;
                }
                else
                {
                    _accessToken = null;
                    _error = error ?? "Sign-in failed.";
                    _status = St.Error;
                }
            }
        }

        private static DateTime ExpiryFromJwt(string token)
        {
            try
            {
                var parts = (token ?? "").Split('.');
                if (parts.Length >= 2)
                {
                    string p = parts[1].Replace('-', '+').Replace('_', '/');
                    switch (p.Length % 4) { case 2: p += "=="; break; case 3: p += "="; break; }
                    string json = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(p));
                    var m = System.Text.RegularExpressions.Regex.Match(json, "\"exp\"\\s*:\\s*(\\d+)");
                    if (m.Success && long.TryParse(m.Groups[1].Value, out long exp))
                        return new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(exp);
                }
            }
            catch { }
            return DateTime.UtcNow.AddMinutes(55);
        }
    }
}
