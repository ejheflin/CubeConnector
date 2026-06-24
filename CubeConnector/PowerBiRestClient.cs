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
