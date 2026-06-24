using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

namespace CubeConnector
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class WizardBridge
    {
        private static readonly JavaScriptSerializer J = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private static string Ok(object o) => J.Serialize(o);
        private static string Err(Exception e) => J.Serialize(new { error = e.Message });

        public string GetAccount()
        {
            try {
                string token = PowerBiAuth.GetAccessToken(out string src, out string err);
                if (token == null) return J.Serialize(new { error = err ?? "sign-in failed" });
                return Ok(new { upn = PowerBiAuth.GetUpnFromToken(token), mode = PowerBiAuth.AccountMode, source = src });
            } catch (Exception e) { return Err(e); }
        }

        public string SignInDifferent()
        {
            try { string t = PowerBiAuth.SignInAsDifferentAccount(out string err);
                  return t == null ? J.Serialize(new { error = err }) : Ok(new { upn = PowerBiAuth.GetUpnFromToken(t) }); }
            catch (Exception e) { return Err(e); }
        }

        public string UseWindowsAccount()
        {
            try { PowerBiAuth.UseWindowsAccount(); return GetAccount(); } catch (Exception e) { return Err(e); }
        }

        public string ListDatasets()
        {
            try {
                string token = PowerBiAuth.GetAccessToken(out _, out string err);
                if (token == null) return J.Serialize(new { error = err ?? "sign-in failed" });
                var ds = PowerBiRestClient.GetAllDatasets(token)
                    .Select(d => new { d.Id, d.Name, d.WorkspaceId, d.WorkspaceName, d.IsRefreshable }).ToList();
                return Ok(new { datasets = ds });
            } catch (Exception e) { return Err(e); }
        }

        public string GetModel(string datasetId, string groupId)
        {
            try {
                string token = PowerBiAuth.GetAccessToken(out _, out string err);
                if (token == null) return J.Serialize(new { error = err ?? "sign-in failed" });
                ModelMetadata md;
                try { md = PowerBiRestClient.ExecuteQueriesIntrospect(token, groupId, datasetId); }
                catch {
                    // Fallback: MSOLAP via ModelIntrospector (may prompt once).
                    string tenantId = PowerBiAuth.GetTidFromTokenPublic(token);
                    var app = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                    var info = ModelIntrospector.IntrospectDataset(app.ActiveWorkbook ?? app.Workbooks.Add(), datasetId, tenantId);
                    md = new ModelMetadata();
                    md.Tables.AddRange(info.Tables.Select(t => t.Name));
                    md.Measures.AddRange(info.Measures.Select(m => new ModelMeasure { Table = m.Table, Name = m.Name }));
                    md.Columns.AddRange(info.Columns.Select(c => new ModelColumn { Table = c.Table, Name = c.Name, DataType = c.DataType, IsHidden = c.IsHidden }));
                }
                return Ok(new {
                    measures = md.Measures.Select(m => new { m.Table, m.Name }),
                    columns = md.Columns.Where(c => !c.IsHidden).Select(c => new { c.Table, c.Name, c.DataType })
                });
            } catch (Exception e) { return Err(e); }
        }

        public string GetFunctions()
        {
            try { return Ok(new { functions = FunctionStore.GetAll() }); } catch (Exception e) { return Err(e); }
        }

        // Loosely-typed save DTO: FilterType arrives as a STRING from the UI.
        // JavaScriptSerializer can't deserialize the FilterType enum from a name,
        // so we receive all-primitive fields and map to UDFConfig via Enum.TryParse.
        private class SaveDto
        {
            public string FunctionName { get; set; }
            public string MeasureName { get; set; }
            public string DatasetId { get; set; }
            public string TenantId { get; set; }
            public string DatasetPrefix { get; set; }
            public List<SaveParam> Parameters { get; set; }
        }
        private class SaveParam
        {
            public string Name { get; set; }
            public int Position { get; set; }
            public string TableName { get; set; }
            public string FieldName { get; set; }
            public string DataType { get; set; }
            public string FilterType { get; set; }
            public bool IsOptional { get; set; }
        }

        public string SaveFunction(string json)
        {
            try {
                var dto = J.Deserialize<SaveDto>(json);
                if (dto == null) throw new ArgumentException("No function data.");
                var config = new UDFConfig
                {
                    FunctionName = dto.FunctionName,
                    MeasureName = dto.MeasureName,
                    DatasetId = dto.DatasetId,
                    TenantId = dto.TenantId,
                    DatasetPrefix = dto.DatasetPrefix,
                    Parameters = new List<ParameterConfig>()
                };
                foreach (var p in dto.Parameters ?? new List<SaveParam>())
                {
                    var pc = new ParameterConfig
                    {
                        Name = p.Name, Position = p.Position, TableName = p.TableName,
                        FieldName = p.FieldName, DataType = p.DataType ?? "text", IsOptional = p.IsOptional
                    };
                    if (!string.IsNullOrEmpty(p.FilterType) && Enum.TryParse(p.FilterType, true, out FilterType ft))
                        pc.FilterType = ft;
                    config.Parameters.Add(pc);
                }
                // The UI doesn't collect a tenant id; derive it from the signed-in token so the
                // saved function can build its pbiazure connection. The user is already signed
                // in, so this is a silent (cached) token acquisition.
                if (string.IsNullOrEmpty(config.TenantId))
                {
                    string token = PowerBiAuth.GetAccessToken(out _, out _);
                    if (!string.IsNullOrEmpty(token)) config.TenantId = PowerBiAuth.GetTidFromTokenPublic(token);
                }
                FunctionStore.Save(config);
                return Ok(new { ok = true });
            } catch (Exception e) { return Err(e); }
        }

        public string DeleteFunction(string name)
        {
            try { FunctionStore.Delete(name); return Ok(new { ok = true }); } catch (Exception e) { return Err(e); }
        }

        public string ExportFunctions(string namesJson, string path)
        {
            try { var names = J.Deserialize<List<string>>(namesJson) ?? new List<string>();
                  FunctionStore.Export(names, path); return Ok(new { ok = true, path }); }
            catch (Exception e) { return Err(e); }
        }

        public string ImportFunctions(string path, string policy)
        {
            try {
                var p = (ImportPolicy)Enum.Parse(typeof(ImportPolicy), policy, true);
                var r = FunctionStore.Import(path, p);
                return Ok(new { added = r.Added, overwritten = r.Overwritten, skipped = r.Skipped });
            } catch (Exception e) { return Err(e); }
        }

        // Dev-only; removed before completion (Task 7).
        public string SelfCheck()
        {
            try {
                string n = FunctionStore.SanitizeName("Net Amount 2025!");
                return Ok(new { sanitize = n });
            } catch (Exception e) { return Err(e); }
        }
    }
}
