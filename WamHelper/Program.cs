/*
 * CubeConnector WamHelper - out-of-process WAM (Web Account Manager) token helper.
 *
 * Runs MSAL + the Windows broker (and its native msalruntime.dll) in ITS OWN
 * process so the heavy/native dependency never loads inside EXCEL.EXE. The Excel
 * add-in shells out to this exe and reads one line of stdout:
 *
 *     ACCESS_TOKEN=<token>      (success)
 *     ERROR=<category>: <msg>   (failure -> add-in falls back to browser flow)
 *
 * Mode (arg 0):
 *   silent      - zero-click SSO: AcquireTokenSilent(OperatingSystemAccount). No UI.
 *   interactive - AcquireTokenInteractive via broker (account picker). Optional.
 *
 * Args: <mode> <clientId> <tenant> <scope>
 */

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;

internal static class Program
{
    private static int Main(string[] args)
    {
        try
        {
            string mode = args.Length > 0 ? args[0] : "silent";
            string clientId = args.Length > 1 ? args[1] : "04b07795-8ddb-461a-bbee-02f9e1bf7b46"; // Azure CLI
            string tenant = args.Length > 2 ? args[2] : "organizations";
            string scope = args.Length > 3 ? args[3] : "https://analysis.windows.net/powerbi/api/.default";

            string token = RunAsync(mode, clientId, tenant, scope).GetAwaiter().GetResult();
            Console.Out.WriteLine("ACCESS_TOKEN=" + token);
            return 0;
        }
        catch (MsalException mex)
        {
            Console.Out.WriteLine("ERROR=" + mex.ErrorCode + ": " + Flatten(mex.Message));
            return 2;
        }
        catch (Exception ex)
        {
            Console.Out.WriteLine("ERROR=" + ex.GetType().Name + ": " + Flatten(ex.Message));
            return 3;
        }
    }

    private static async Task<string> RunAsync(string mode, string clientId, string tenant, string scope)
    {
        var app = PublicClientApplicationBuilder
            .Create(clientId)
            .WithAuthority("https://login.microsoftonline.com/" + tenant)
            .WithDefaultRedirectUri()
            .WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows))
            .Build();

        var scopes = new[] { scope };

        if (mode == "interactive")
        {
            var resultI = await app.AcquireTokenInteractive(scopes)
                .WithAccount(PublicClientApplication.OperatingSystemAccount)
                .ExecuteAsync().ConfigureAwait(false);
            return resultI.AccessToken;
        }

        // silent / zero-click: use the Windows session account via the broker (PRT).
        var result = await app.AcquireTokenSilent(scopes, PublicClientApplication.OperatingSystemAccount)
            .ExecuteAsync(CancellationToken.None).ConfigureAwait(false);
        return result.AccessToken;
    }

    private static string Flatten(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return s.Replace("\r", " ").Replace("\n", " ").Trim();
    }
}
