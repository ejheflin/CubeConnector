/*
 * CubeConnector - PowerBiRestClient
 *
 * Reusable Power BI REST API client for enumerating the workspaces and semantic
 * models (datasets) the signed-in user can access. Pro-compatible (the list
 * endpoints work on a Pro license with a delegated user token).
 *
 * Auth is delegated to PowerBiAuth (WAM zero-click -> cached refresh -> browser).
 * JSON is parsed with DataContractJsonSerializer (same approach as ConfigurationStore).
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace CubeConnector
{
    public class WorkspaceInfo
    {
        public string Id;
        public string Name;
        public bool IsOnDedicatedCapacity;
    }

    public class DatasetInfo
    {
        public string Id;
        public string Name;
        public string WorkspaceId;
        public string WorkspaceName;
        public string ConfiguredBy;
        public bool IsRefreshable;
        public string TargetStorageMode;
    }

    public static class PowerBiRestClient
    {
        private const string PbiApi = "https://api.powerbi.com/v1.0/myorg";

        // ---- in-memory dataset cache (prefetch on pane open) ----

        private static volatile List<DatasetInfo> _cache;
        private static volatile bool _warming;

        /// <summary>
        /// Best-effort silent prefetch: if a token is already available (WAM/cached),
        /// populate the dataset cache in the background so the UI dropdown is instant.
        /// Safe to call from a background thread and idempotent — if the cache is already
        /// populated or a warm is in flight, it returns immediately (so firing it from both
        /// the ribbon click and the pane load never double-fetches).
        /// </summary>
        public static void WarmDatasetCache()
        {
            if (_cache != null || _warming) return;
            _warming = true;
            try
            {
                string token = PowerBiAuth.GetAccessToken(out _, out _); // silent path if available
                if (!string.IsNullOrEmpty(token))
                    _cache = GetAllDatasets(token);
            }
            catch { /* prefetch is best-effort — never surface errors here */ }
            finally { _warming = false; }
        }

        /// <summary>
        /// Returns all datasets, using the in-memory cache when it has been warmed,
        /// otherwise falls through to the live REST call.
        /// </summary>
        public static List<DatasetInfo> GetAllDatasetsCached(string accessToken)
        {
            if (_cache != null) return _cache;
            _cache = GetAllDatasets(accessToken);
            return _cache;
        }

        /// <summary>Drop the dataset cache (e.g. after switching accounts).</summary>
        public static void ClearCache() { _cache = null; }

        // ---- DataContract response shapes ----

        [DataContract]
        private class GroupsResponse { [DataMember(Name = "value")] public List<GroupJson> value { get; set; } }

        [DataContract]
        private class GroupJson
        {
            [DataMember(Name = "id")] public string id { get; set; }
            [DataMember(Name = "name")] public string name { get; set; }
            [DataMember(Name = "isOnDedicatedCapacity")] public bool isOnDedicatedCapacity { get; set; }
        }

        [DataContract]
        private class DatasetsResponse { [DataMember(Name = "value")] public List<DatasetJson> value { get; set; } }

        [DataContract]
        private class DatasetJson
        {
            [DataMember(Name = "id")] public string id { get; set; }
            [DataMember(Name = "name")] public string name { get; set; }
            [DataMember(Name = "configuredBy")] public string configuredBy { get; set; }
            [DataMember(Name = "isRefreshable")] public bool isRefreshable { get; set; }
            [DataMember(Name = "targetStorageMode")] public string targetStorageMode { get; set; }
        }

        // ---- public API ----

        /// <summary>All workspaces (groups) the signed-in user can access.</summary>
        public static List<WorkspaceInfo> GetWorkspaces(string accessToken)
        {
            var resp = GetJson<GroupsResponse>(PbiApi + "/groups", accessToken);
            var list = new List<WorkspaceInfo>();
            if (resp?.value != null)
                foreach (var g in resp.value)
                    list.Add(new WorkspaceInfo { Id = g.id, Name = g.name, IsOnDedicatedCapacity = g.isOnDedicatedCapacity });
            return list;
        }

        /// <summary>Datasets within a workspace.</summary>
        public static List<DatasetInfo> GetDatasets(string accessToken, WorkspaceInfo workspace)
        {
            var resp = GetJson<DatasetsResponse>(PbiApi + "/groups/" + workspace.Id + "/datasets", accessToken);
            var list = new List<DatasetInfo>();
            if (resp?.value != null)
                foreach (var d in resp.value)
                    list.Add(new DatasetInfo
                    {
                        Id = d.id,
                        Name = d.name,
                        WorkspaceId = workspace.Id,
                        WorkspaceName = workspace.Name,
                        ConfiguredBy = d.configuredBy,
                        IsRefreshable = d.isRefreshable,
                        TargetStorageMode = d.targetStorageMode
                    });
            return list;
        }

        /// <summary>Datasets in the user's personal "My workspace".</summary>
        public static List<DatasetInfo> GetMyWorkspaceDatasets(string accessToken)
        {
            var resp = GetJson<DatasetsResponse>(PbiApi + "/datasets", accessToken);
            var list = new List<DatasetInfo>();
            if (resp?.value != null)
                foreach (var d in resp.value)
                    list.Add(new DatasetInfo
                    {
                        Id = d.id, Name = d.name, WorkspaceId = null, WorkspaceName = "My workspace",
                        ConfiguredBy = d.configuredBy, IsRefreshable = d.isRefreshable, TargetStorageMode = d.targetStorageMode
                    });
            return list;
        }

        /// <summary>
        /// Every accessible dataset across all workspaces (+ My workspace), sorted by
        /// workspace then dataset name. One REST call per workspace.
        /// </summary>
        public static List<DatasetInfo> GetAllDatasets(string accessToken)
        {
            var all = new List<DatasetInfo>();
            try { all.AddRange(GetMyWorkspaceDatasets(accessToken)); } catch { }
            foreach (var ws in GetWorkspaces(accessToken))
            {
                try { all.AddRange(GetDatasets(accessToken, ws)); } catch { }
            }
            return all.OrderBy(d => d.WorkspaceName, StringComparer.OrdinalIgnoreCase)
                      .ThenBy(d => d.Name, StringComparer.OrdinalIgnoreCase)
                      .ToList();
        }

        // ---- executeQueries introspection ----

        public static ModelMetadata ExecuteQueriesIntrospect(string accessToken, string groupId, string datasetId)
        {
            string baseUrl = string.IsNullOrEmpty(groupId)
                ? PbiApi + "/datasets/" + datasetId + "/executeQueries"
                : PbiApi + "/groups/" + groupId + "/datasets/" + datasetId + "/executeQueries";

            var md = new ModelMetadata();
            foreach (var r in RunDax(accessToken, baseUrl, "EVALUATE INFO.VIEW.TABLES()"))
                md.Tables.Add(Val(r, "Name"));
            foreach (var r in RunDax(accessToken, baseUrl, "EVALUATE INFO.VIEW.COLUMNS()"))
            {
                string type = Val(r, "Type"); string cat = Val(r, "DataCategory");
                if (type == "RowNumber" || cat == "RowNumber") continue;
                md.Columns.Add(new ModelColumn { Table = Val(r, "Table"), Name = Val(r, "Name"),
                    DataType = Val(r, "DataType"), IsHidden = Val(r, "IsHidden") == "True", Description = Val(r, "Description") });
            }
            foreach (var r in RunDax(accessToken, baseUrl, "EVALUATE INFO.VIEW.MEASURES()"))
                md.Measures.Add(new ModelMeasure { Table = Val(r, "Table"), Name = Val(r, "Name"), Description = Val(r, "Description") });
            return md;
        }

        // executeQueries returns results[0].tables[0].rows: [{ "INFO.VIEW.COLUMNS()[Name]": val, ... }]
        private static List<Dictionary<string,string>> RunDax(string token, string url, string dax)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST"; req.ContentType = "application/json";
            req.Headers["Authorization"] = "Bearer " + token; req.Accept = "application/json";
            string body = "{\"queries\":[{\"query\":\"" + dax.Replace("\"","\\\"") + "\"}],"
                + "\"serializerSettings\":{\"includeNulls\":true}}";
            byte[] data = Encoding.UTF8.GetBytes(body);
            req.ContentLength = data.Length;
            using (var rs = req.GetRequestStream()) rs.Write(data, 0, data.Length);

            string resp;
            try { using (var r = (HttpWebResponse)req.GetResponse())
                  using (var sr = new StreamReader(r.GetResponseStream())) resp = sr.ReadToEnd(); }
            catch (WebException wex) when (wex.Response is HttpWebResponse er)
            { using (var sr = new StreamReader(er.GetResponseStream()))
                throw new InvalidOperationException("executeQueries failed (" + (int)er.StatusCode + "): " + sr.ReadToEnd()); }

            return ParseRows(resp);
        }

        // Minimal parse of results[0].tables[0].rows using JavaScriptSerializer.
        private static List<Dictionary<string,string>> ParseRows(string json)
        {
            var ser = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var root = (Dictionary<string,object>)ser.DeserializeObject(json);
            var rows = new List<Dictionary<string,string>>();
            var results = root.TryGetValue("results", out var ro) ? ro as object[] : null;
            if (results == null || results.Length == 0) return rows;
            var res0 = (Dictionary<string,object>)results[0];
            var tables = res0.TryGetValue("tables", out var to) ? to as object[] : null;
            if (tables == null || tables.Length == 0) return rows;
            var t0 = (Dictionary<string,object>)tables[0];
            var rowArr = t0.TryGetValue("rows", out var rr) ? rr as object[] : null;
            if (rowArr == null) return rows;
            foreach (Dictionary<string,object> row in rowArr.Cast<Dictionary<string,object>>())
            {
                var d = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in row)
                {
                    // keys look like "[Name]" or "Table[Name]" or "INFO.VIEW.COLUMNS()[Name]" -> take inside last [ ]
                    string key = kv.Key; int lb = key.LastIndexOf('['), rb = key.LastIndexOf(']');
                    if (lb >= 0 && rb > lb) key = key.Substring(lb + 1, rb - lb - 1);
                    d[key] = kv.Value?.ToString() ?? "";
                }
                rows.Add(d);
            }
            return rows;
        }

        private static string Val(Dictionary<string,string> row, string key)
            => row.TryGetValue(key, out var v) ? v : "";

        // ---- HTTP / JSON ----

        private static T GetJson<T>(string url, string bearer) where T : class
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "GET";
            req.Headers["Authorization"] = "Bearer " + bearer;
            req.Accept = "application/json";

            string body;
            HttpStatusCode status;
            try
            {
                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    status = resp.StatusCode;
                    using (var sr = new StreamReader(resp.GetResponseStream())) body = sr.ReadToEnd();
                }
            }
            catch (WebException wex) when (wex.Response is HttpWebResponse er)
            {
                using (var sr = new StreamReader(er.GetResponseStream())) body = sr.ReadToEnd();
                throw new InvalidOperationException("Power BI REST call failed (" + (int)er.StatusCode + "): " + body);
            }

            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(body)))
            {
                var ser = new DataContractJsonSerializer(typeof(T));
                return (T)ser.ReadObject(ms);
            }
        }
    }
}
