# CubeConnector Configuration Wizard — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers-extended-cc:subagent-driven-development (recommended) or superpowers-extended-cc:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace hand-edited `CubeConnectorConfig.json` with an in-Excel WebView2 wizard that lets non-technical users create, manage, import/export, and share CubeConnector formulas, backed by the proven silent-auth / REST-enumeration / introspection services.

**Architecture:** A ribbon button opens a WinForms `WizardWindow` hosting a `WebView2` control that renders an HTML/CSS/JS UI (evolved from `docs/index.html`). A COM-visible `WizardBridge` exposes JSON-in/JSON-out methods to the UI; it orchestrates `FunctionStore` (per-user `functions.json` CRUD + import/export + migration), `PowerBiAuth`, `PowerBiRestClient`, and introspection (REST `executeQueries` with `ModelIntrospector`/MSOLAP fallback). `AutoOpen` keeps registering UDFs at startup from the same per-user file.

**Tech Stack:** C# / .NET Framework 4.7.2, Excel-DNA 1.9, WinForms, Microsoft.Web.WebView2, Power BI REST + MSOLAP, DataContractJsonSerializer.

**Verification approach (read this):** This repo is a COM-bound Excel-DNA add-in with **no automated test harness**, and behavior depends on live Excel + Power BI. Verification is therefore: (a) `msbuild` build succeeds, and (b) explicit **manual checks** in Excel via the established close-Excel → build → reopen loop. Pure-logic pieces (name sanitization, import merge) get a temporary self-check exposed through the bridge during development and removed before completion. Build command used throughout:

```
"C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\MSBuild\Current\Bin\MSBuild.exe" \
  "C:\dev\CubeConnector_gh\CubeConnector\CubeConnector.csproj" /t:Rebuild /p:Configuration=Debug /p:Platform=AnyCPU /v:minimal /nologo
```
Excel locks the `.xll`; fully close Excel (watch for orphan `EXCEL.EXE`) before each rebuild.

---

## File Structure

- `CubeConnector/FunctionStore.cs` *(new)* — owns `%LOCALAPPDATA%\CubeConnector\functions.json`: load/save, CRUD, name sanitization, export, import+merge, legacy migration. Single source of truth for config persistence.
- `CubeConnector/ConfigurationStore.cs` *(modify)* — `GetAllConfigs()` delegates to `FunctionStore`; remove its own file path + JSON-file reading.
- `CubeConnector/PowerBiRestClient.cs` *(modify)* — add `ExecuteQueriesIntrospect(token, groupId, datasetId)` returning model metadata.
- `CubeConnector/ModelMetadata.cs` *(new)* — shared DTO (`ModelMetadata` with tables/columns/measures) returned by both the REST and MSOLAP introspection paths (so the bridge has one shape).
- `CubeConnector/WizardBridge.cs` *(new)* — COM-visible, JSON-in/JSON-out methods called from the WebView2 UI.
- `CubeConnector/WizardWindow.cs` *(new)* — WinForms window hosting WebView2, environment setup, virtual-host mapping, bridge wiring.
- `CubeConnector/CubeConnectorRibbon.cs` *(modify)* — add "Manage Formulas" button; remove the temporary "Enumerate Models (TEST)" button + handler.
- `CubeConnector/ui/index.html`, `ui/app.js`, `ui/styles.css` *(new)* — the wizard UI, evolved from `docs/index.html`, deployed beside the add-in (post-build copy, like `WamHelper`).
- `CubeConnector/CubeConnector.csproj` *(modify)* — add WebView2 reference, compile entries, `ui/` deploy target.
- `CubeConnector/EnumerateModelsSmokeTest.cs` *(delete)* — superseded.

---

## Task 1: FunctionStore — per-user config persistence, CRUD, import/export, migration

**Goal:** A single class that owns reading/writing `functions.json` in `%LOCALAPPDATA%\CubeConnector`, plus name sanitization and import-merge logic.

**Files:**
- Create: `CubeConnector/FunctionStore.cs`
- Modify: `CubeConnector/CubeConnector.csproj` (add `<Compile Include="FunctionStore.cs" />`)

**Acceptance Criteria:**
- [ ] `GetAll()` returns `List<UDFConfig>` from the per-user file (empty list if none).
- [ ] `Save(UDFConfig)` upserts by `FunctionName` (case-insensitive) and persists.
- [ ] `Delete(name)` removes by name and persists.
- [ ] `SanitizeName(friendly)` turns `"Net Amount"` → `"CC.NetAmount"` (prefix `CC.`, strip invalid chars, no leading digit).
- [ ] `Export(names)` writes a `{functions:[...]}` JSON to a chosen path.
- [ ] `Import(path, policy)` merges and returns counts `{added, overwritten, skipped}`.
- [ ] `MigrateLegacyIfNeeded()` copies a legacy file next to the `.xll` into the per-user path when no per-user file exists.

**Verify:** Build succeeds; logic exercised via the dev self-check added in Task 4 and manual import/export in Task 7.

**Steps:**

- [ ] **Step 1: Create `FunctionStore.cs` with the full implementation**

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;

namespace CubeConnector
{
    public enum ImportPolicy { Overwrite, Skip, KeepBoth }

    public class ImportResult { public int Added; public int Overwritten; public int Skipped; }

    public static class FunctionStore
    {
        public const string FunctionPrefix = "CC.";

        private static string Dir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CubeConnector");
        private static string FilePath => Path.Combine(Dir, "functions.json");

        // ---- DataContract shapes (the on-disk format) ----
        [DataContract] public class FileWrapper { [DataMember(Name = "functions")] public List<FuncJson> functions { get; set; } }
        [DataContract] public class FuncJson
        {
            [DataMember(Name = "functionName")] public string functionName { get; set; }
            [DataMember(Name = "tenantId")] public string tenantId { get; set; }
            [DataMember(Name = "datasetPrefix")] public string datasetPrefix { get; set; }
            [DataMember(Name = "datasetId")] public string datasetId { get; set; }
            [DataMember(Name = "measureName")] public string measureName { get; set; }
            [DataMember(Name = "parameters")] public List<ParamJson> parameters { get; set; }
        }
        [DataContract] public class ParamJson
        {
            [DataMember(Name = "name")] public string name { get; set; }
            [DataMember(Name = "position")] public int position { get; set; }
            [DataMember(Name = "tableName")] public string tableName { get; set; }
            [DataMember(Name = "fieldName")] public string fieldName { get; set; }
            [DataMember(Name = "dataType")] public string dataType { get; set; }
            [DataMember(Name = "filterType")] public string filterType { get; set; }
            [DataMember(Name = "isOptional")] public bool isOptional { get; set; }
        }

        // ---- public API ----

        public static List<UDFConfig> GetAll()
        {
            var wrapper = ReadFile();
            var list = new List<UDFConfig>();
            if (wrapper?.functions == null) return list;
            foreach (var f in wrapper.functions) list.Add(ToConfig(f));
            return list;
        }

        public static void Save(UDFConfig config)
        {
            if (config == null || string.IsNullOrWhiteSpace(config.FunctionName))
                throw new ArgumentException("Function must have a name.");
            var wrapper = ReadFile() ?? new FileWrapper { functions = new List<FuncJson>() };
            wrapper.functions.RemoveAll(f =>
                string.Equals(f.functionName, config.FunctionName, StringComparison.OrdinalIgnoreCase));
            wrapper.functions.Add(ToJson(config));
            WriteFile(wrapper);
        }

        public static void Delete(string functionName)
        {
            var wrapper = ReadFile();
            if (wrapper?.functions == null) return;
            wrapper.functions.RemoveAll(f =>
                string.Equals(f.functionName, functionName, StringComparison.OrdinalIgnoreCase));
            WriteFile(wrapper);
        }

        /// <summary>"Net Amount" -> "CC.NetAmount". Strips invalid chars, avoids leading digit.</summary>
        public static string SanitizeName(string friendly)
        {
            string core = (friendly ?? "").Trim();
            if (core.StartsWith(FunctionPrefix, StringComparison.OrdinalIgnoreCase))
                core = core.Substring(FunctionPrefix.Length);
            core = Regex.Replace(core, @"[^A-Za-z0-9_]", "");
            if (string.IsNullOrEmpty(core)) core = "Formula";
            if (char.IsDigit(core[0])) core = "_" + core;
            return FunctionPrefix + core;
        }

        public static void Export(IEnumerable<string> names, string path)
        {
            var wrapper = ReadFile() ?? new FileWrapper { functions = new List<FuncJson>() };
            var set = new HashSet<string>(names ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            var chosen = wrapper.functions.Where(f => set.Count == 0 || set.Contains(f.functionName)).ToList();
            WriteWrapperTo(new FileWrapper { functions = chosen }, path);
        }

        public static ImportResult Import(string path, ImportPolicy policy)
        {
            var incoming = ReadWrapperFrom(path) ?? new FileWrapper { functions = new List<FuncJson>() };
            var current = ReadFile() ?? new FileWrapper { functions = new List<FuncJson>() };
            var result = new ImportResult();

            foreach (var inc in incoming.functions ?? new List<FuncJson>())
            {
                int idx = current.functions.FindIndex(f =>
                    string.Equals(f.functionName, inc.functionName, StringComparison.OrdinalIgnoreCase));
                if (idx < 0) { current.functions.Add(inc); result.Added++; continue; }
                switch (policy)
                {
                    case ImportPolicy.Overwrite: current.functions[idx] = inc; result.Overwritten++; break;
                    case ImportPolicy.Skip: result.Skipped++; break;
                    case ImportPolicy.KeepBoth:
                        inc.functionName = MakeUniqueName(inc.functionName, current.functions);
                        current.functions.Add(inc); result.Added++; break;
                }
            }
            WriteFile(current);
            return result;
        }

        public static void MigrateLegacyIfNeeded()
        {
            try
            {
                if (File.Exists(FilePath)) return;
                string xll = ExcelDna.Integration.ExcelDnaUtil.XllPath;
                string legacy = Path.Combine(Path.GetDirectoryName(xll), "CubeConnectorConfig.json");
                if (File.Exists(legacy))
                {
                    Directory.CreateDirectory(Dir);
                    File.Copy(legacy, FilePath);
                }
            }
            catch { /* migration is best-effort */ }
        }

        // ---- helpers ----

        private static string MakeUniqueName(string name, List<FuncJson> existing)
        {
            string baseName = name; int n = 2;
            while (existing.Any(f => string.Equals(f.functionName, name, StringComparison.OrdinalIgnoreCase)))
                name = baseName + n++;
            return name;
        }

        private static FileWrapper ReadFile()
        {
            try { return File.Exists(FilePath) ? ReadWrapperFrom(FilePath) : null; }
            catch { return null; }
        }

        private static FileWrapper ReadWrapperFrom(string path)
        {
            using (var fs = File.OpenRead(path))
                return (FileWrapper)new DataContractJsonSerializer(typeof(FileWrapper)).ReadObject(fs);
        }

        private static void WriteFile(FileWrapper wrapper)
        {
            Directory.CreateDirectory(Dir);
            WriteWrapperTo(wrapper, FilePath);
        }

        private static void WriteWrapperTo(FileWrapper wrapper, string path)
        {
            using (var ms = new MemoryStream())
            {
                var ser = new DataContractJsonSerializer(typeof(FileWrapper));
                ser.WriteObject(ms, wrapper);
                File.WriteAllBytes(path, ms.ToArray());
            }
        }

        private static UDFConfig ToConfig(FuncJson f)
        {
            var c = new UDFConfig
            {
                FunctionName = f.functionName,
                TenantId = f.tenantId,
                DatasetPrefix = f.datasetPrefix,
                DatasetId = f.datasetId,
                MeasureName = f.measureName,
                Parameters = new List<ParameterConfig>()
            };
            if (!string.IsNullOrEmpty(c.DatasetPrefix) && !string.IsNullOrEmpty(c.DatasetId)
                && Guid.TryParse(c.DatasetId, out _) && !c.DatasetId.StartsWith(c.DatasetPrefix))
                c.DatasetId = c.DatasetPrefix + c.DatasetId;
            if (f.parameters != null)
                foreach (var p in f.parameters)
                {
                    var pc = new ParameterConfig
                    {
                        Name = p.name, Position = p.position, TableName = p.tableName,
                        FieldName = p.fieldName, DataType = p.dataType ?? "text", IsOptional = p.isOptional
                    };
                    if (!string.IsNullOrEmpty(p.filterType) && Enum.TryParse(p.filterType, true, out FilterType ft))
                        pc.FilterType = ft;
                    c.Parameters.Add(pc);
                }
            return c;
        }

        private static FuncJson ToJson(UDFConfig c)
        {
            return new FuncJson
            {
                functionName = c.FunctionName, tenantId = c.TenantId, datasetPrefix = c.DatasetPrefix,
                datasetId = c.DatasetId, measureName = c.MeasureName,
                parameters = (c.Parameters ?? new List<ParameterConfig>()).Select(p => new ParamJson
                {
                    name = p.Name, position = p.Position, tableName = p.TableName, fieldName = p.FieldName,
                    dataType = p.DataType, filterType = p.FilterType.ToString(), isOptional = p.IsOptional
                }).ToList()
            };
        }
    }
}
```

- [ ] **Step 2: Add to csproj** — add `<Compile Include="FunctionStore.cs" />` next to the other `<Compile>` entries.

- [ ] **Step 3: Build** — run the MSBuild command. Expected: exit 0.

- [ ] **Step 4: Commit**

```bash
git add CubeConnector/FunctionStore.cs CubeConnector/CubeConnector.csproj
git commit -m "feat: FunctionStore for per-user config (CRUD, import/export, migration)"
```

---

## Task 2: Repoint ConfigurationStore to FunctionStore + migration

**Goal:** `AutoOpen` reads functions from the per-user file via `FunctionStore`; legacy file migrates once.

**Files:**
- Modify: `CubeConnector/ConfigurationStore.cs` (`GetAllConfigs`, `InitializeConfigs`)
- Modify: `CubeConnector/DynamicFunctionRegistration.cs` (`AutoOpen` calls migration first)

**Acceptance Criteria:**
- [ ] `ConfigurationStore.GetAllConfigs()` returns `FunctionStore.GetAll()` (falling back to the existing hardcoded fallback only if empty).
- [ ] `AutoOpen` calls `FunctionStore.MigrateLegacyIfNeeded()` before reading configs.
- [ ] Existing functions still register at startup from the per-user file.

**Verify:** Build succeeds. Manual: place a `functions.json` with one function in `%LOCALAPPDATA%\CubeConnector`, start Excel, confirm the function is callable in a cell.

**Steps:**

- [ ] **Step 1: Replace `InitializeConfigs()` body in `ConfigurationStore.cs`**

Replace the existing `InitializeConfigs` method with:

```csharp
private static void InitializeConfigs()
{
    _configs = FunctionStore.GetAll();
    if (_configs == null || _configs.Count == 0)
        _configs = GetFallbackConfigs();
}
```

Leave `GetFallbackConfigs()`, `GetConfig()`, `GetAllConfigs()` as-is. The old `LoadFromJson`/`ConvertToUDFConfig`/contract classes in this file become dead code — delete them to avoid duplication with `FunctionStore`.

- [ ] **Step 2: Call migration in `AutoOpen`** — in `DynamicFunctionRegistration.cs`, at the very start of the `try` in `AutoOpen()` (before `ConfigurationStore.GetAllConfigs()`), add:

```csharp
FunctionStore.MigrateLegacyIfNeeded();
```

- [ ] **Step 3: Build** — MSBuild. Expected exit 0.

- [ ] **Step 4: Manual check** — copy `CubeConnector/CubeConnectorConfig.json` to `%LOCALAPPDATA%\CubeConnector\functions.json`, start Excel, type `=CC.AmtNet()` (or the configured function) and confirm it resolves (no #NAME?). Close Excel.

- [ ] **Step 5: Commit**

```bash
git add CubeConnector/ConfigurationStore.cs CubeConnector/DynamicFunctionRegistration.cs
git commit -m "refactor: ConfigurationStore reads per-user store via FunctionStore + migration"
```

---

## Task 3: Silent introspection via REST executeQueries (+ shared ModelMetadata)

**Goal:** Read a model's tables/columns/measures silently with the held token; expose one DTO shape used by the bridge, falling back to `ModelIntrospector` (MSOLAP).

**Files:**
- Create: `CubeConnector/ModelMetadata.cs`
- Modify: `CubeConnector/PowerBiRestClient.cs` (add `ExecuteQueriesIntrospect`)
- Modify: `CubeConnector/CubeConnector.csproj` (compile `ModelMetadata.cs`)

**Acceptance Criteria:**
- [ ] `ModelMetadata` holds `Tables`, `Columns` (Table, Name, DataType, IsHidden), `Measures` (Table, Name).
- [ ] `PowerBiRestClient.ExecuteQueriesIntrospect(token, groupId, datasetId)` POSTs `EVALUATE INFO.VIEW.COLUMNS()` / `MEASURES()` / `TABLES()` and parses results into `ModelMetadata`.
- [ ] RowNumber system columns filtered out.
- [ ] On HTTP error, throws `InvalidOperationException` with the response body (caller falls back).

**Verify:** Build succeeds. Manual via Task 7 (pick a model in the builder; fields populate with no sign-in popup).

**Steps:**

- [ ] **Step 1: Create `ModelMetadata.cs`**

```csharp
using System.Collections.Generic;

namespace CubeConnector
{
    public class ModelMetadata
    {
        public List<string> Tables = new List<string>();
        public List<ModelColumn> Columns = new List<ModelColumn>();
        public List<ModelMeasure> Measures = new List<ModelMeasure>();
    }
    public class ModelColumn { public string Table; public string Name; public string DataType; public bool IsHidden;
        public string Qualified => "[" + Table + "].[" + Name + "]"; }
    public class ModelMeasure { public string Table; public string Name; }
}
```

- [ ] **Step 2: Add `ExecuteQueriesIntrospect` to `PowerBiRestClient.cs`**

Add inside the `PowerBiRestClient` class:

```csharp
public static ModelMetadata ExecuteQueriesIntrospect(string accessToken, string groupId, string datasetId)
{
    string baseUrl = string.IsNullOrEmpty(groupId)
        ? PbiApi + "/datasets/" + datasetId + "/executeQueries"
        : PbiApi + "/groups/" + groupId + "/datasets/" + datasetId + "/executeQueries";

    var md = new ModelMetadata();
    foreach (var rows in new[] {
        RunDax(accessToken, baseUrl, "EVALUATE INFO.VIEW.TABLES()"),
        RunDax(accessToken, baseUrl, "EVALUATE INFO.VIEW.COLUMNS()"),
        RunDax(accessToken, baseUrl, "EVALUATE INFO.VIEW.MEASURES()") })
    { /* assigned below per call */ }

    foreach (var r in RunDax(accessToken, baseUrl, "EVALUATE INFO.VIEW.TABLES()"))
        md.Tables.Add(Val(r, "Name"));
    foreach (var r in RunDax(accessToken, baseUrl, "EVALUATE INFO.VIEW.COLUMNS()"))
    {
        string type = Val(r, "Type"); string cat = Val(r, "DataCategory");
        if (type == "RowNumber" || cat == "RowNumber") continue;
        md.Columns.Add(new ModelColumn { Table = Val(r, "Table"), Name = Val(r, "Name"),
            DataType = Val(r, "DataType"), IsHidden = Val(r, "IsHidden") == "True" });
    }
    foreach (var r in RunDax(accessToken, baseUrl, "EVALUATE INFO.VIEW.MEASURES()"))
        md.Measures.Add(new ModelMeasure { Table = Val(r, "Table"), Name = Val(r, "Name") });
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
```

Add `using System.Linq;` to the file if not present. Add a project reference to **`System.Web.Extensions`** (for `JavaScriptSerializer`) in csproj `<ItemGroup>` of references.

- [ ] **Step 3: csproj** — add `<Compile Include="ModelMetadata.cs" />` and `<Reference Include="System.Web.Extensions" />`.

- [ ] **Step 4: Build** — MSBuild. Expected exit 0.

- [ ] **Step 5: Commit**

```bash
git add CubeConnector/ModelMetadata.cs CubeConnector/PowerBiRestClient.cs CubeConnector/CubeConnector.csproj
git commit -m "feat: silent REST executeQueries introspection + ModelMetadata DTO"
```

---

## Task 4: WizardBridge — JSON-in/JSON-out service facade

**Goal:** A COM-visible class the WebView2 UI calls; each method returns a JSON string, errors as `{"error":"..."}`.

**Files:**
- Create: `CubeConnector/WizardBridge.cs`
- Modify: `CubeConnector/CubeConnector.csproj` (compile)

**Acceptance Criteria:**
- [ ] Methods: `GetAccount()`, `SignInDifferent()`, `UseWindowsAccount()`, `ListDatasets()`, `GetModel(datasetId, groupId)`, `GetFunctions()`, `SaveFunction(json)`, `DeleteFunction(name)`, `ExportFunctions(namesJson, path)`, `ImportFunctions(path, policy)`, and dev-only `SelfCheck()`.
- [ ] All return JSON strings; exceptions are caught and returned as `{"error": "..."}`.
- [ ] `GetModel` tries `ExecuteQueriesIntrospect`; on failure falls back to `ModelIntrospector.IntrospectDataset` (deriving tenantId from the token).

**Verify:** Build succeeds; `SelfCheck()` (Task 7) exercises FunctionStore round-trip + sanitize + import merge.

**Steps:**

- [ ] **Step 1: Create `WizardBridge.cs`**

```csharp
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

        public string SaveFunction(string json)
        {
            try {
                var dto = J.Deserialize<UDFConfig>(json);
                FunctionStore.Save(dto);
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
```

- [ ] **Step 2: Add `GetTidFromTokenPublic` to `PowerBiAuth.cs`** — the fallback needs the tenant id. Add a public wrapper:

```csharp
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
```

- [ ] **Step 3: csproj** — add `<Compile Include="WizardBridge.cs" />`.

- [ ] **Step 4: Build** — MSBuild. Expected exit 0.

- [ ] **Step 5: Commit**

```bash
git add CubeConnector/WizardBridge.cs CubeConnector/PowerBiAuth.cs CubeConnector/CubeConnector.csproj
git commit -m "feat: WizardBridge JSON facade over auth/REST/introspection/store"
```

---

## Task 5: WizardWindow + WebView2 host + ribbon entry

**Goal:** A ribbon button opens a WinForms window hosting WebView2, mapped to the `ui/` folder, with the bridge registered.

**Files:**
- Create: `CubeConnector/WizardWindow.cs`
- Modify: `CubeConnector/CubeConnectorRibbon.cs` (button + handler; remove smoke-test)
- Delete: `CubeConnector/EnumerateModelsSmokeTest.cs`
- Modify: `CubeConnector/CubeConnector.csproj` (WebView2 package, compile, ui deploy, drop smoke-test compile)

**Acceptance Criteria:**
- [ ] "Manage Formulas" button appears in the Data ▸ CubeConnector group.
- [ ] Clicking it opens a window that loads `ui/index.html` via virtual host mapping.
- [ ] `window.chrome.webview.hostObjects.cc.GetAccount()` works from the page.
- [ ] If WebView2 Runtime is missing, a friendly MessageBox explains it.
- [ ] Smoke-test button + class removed.

**Verify:** Build succeeds; manual: open Excel, click Manage Formulas, the page loads and shows the signed-in account.

**Steps:**

- [ ] **Step 1: Add WebView2 NuGet** — install into the solution packages folder:

```bash
"$TEMP/nuget.exe" install Microsoft.Web.WebView2 -OutputDirectory "C:/dev/CubeConnector_gh/packages" -DependencyVersion Highest -Framework net472
```

Then add to csproj references (adjust version to the installed folder):

```xml
<Reference Include="Microsoft.Web.WebView2.Core">
  <HintPath>..\packages\Microsoft.Web.WebView2.<VER>\lib\net462\Microsoft.Web.WebView2.Core.dll</HintPath>
</Reference>
<Reference Include="Microsoft.Web.WebView2.WinForms">
  <HintPath>..\packages\Microsoft.Web.WebView2.<VER>\lib\net462\Microsoft.Web.WebView2.WinForms.dll</HintPath>
</Reference>
```

Add a post-build copy of the WebView2 native loader (`WebView2Loader.dll`) and the `ui/` folder beside the add-in, alongside the existing WamHelper copy target:

```xml
<None Include="ui\**\*.*"><CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory></None>
```

- [ ] **Step 2: Create `WizardWindow.cs`**

```csharp
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
```

- [ ] **Step 3: Ribbon — replace the smoke-test button.** In `CubeConnectorRibbon.cs` GetCustomUI, replace the `EnumerateModelsTestBtn` button XML with:

```xml
<button id='ManageFormulasBtn'
        label='Manage Formulas'
        onAction='OnManageFormulasClicked'
        imageMso='TableInsertDialog' />
```

Replace the `OnEnumerateModelsClicked` handler with:

```csharp
public void OnManageFormulasClicked(IRibbonControl control)
{
    WizardWindow.ShowSingleton();
}
```

- [ ] **Step 4: Delete smoke-test** — delete `CubeConnector/EnumerateModelsSmokeTest.cs` and its `<Compile Include="EnumerateModelsSmokeTest.cs" />` line.

- [ ] **Step 5: Build** — MSBuild. Expected exit 0. Confirm `bin\Debug\ui\index.html` and WebView2 DLLs are present (create a placeholder `ui/index.html` containing `<h1>ok</h1>` before this build if Task 6 isn't done yet).

- [ ] **Step 6: Manual check** — open Excel, Data ▸ CubeConnector ▸ Manage Formulas; the window opens and shows the placeholder (or, after Task 6, the UI).

- [ ] **Step 7: Commit**

```bash
git add CubeConnector/WizardWindow.cs CubeConnector/CubeConnectorRibbon.cs CubeConnector/CubeConnector.csproj
git rm CubeConnector/EnumerateModelsSmokeTest.cs
git commit -m "feat: WizardWindow WebView2 host + Manage Formulas ribbon button"
```

---

## Task 6: The wizard UI (library, import, builder)

**Goal:** The HTML/CSS/JS UI evolved from `docs/index.html`, calling the bridge, with plain-language labels, "?" helpers, the consumer Import hero path, and the builder + preview.

**Files:**
- Create: `CubeConnector/ui/index.html`, `CubeConnector/ui/app.js`, `CubeConnector/ui/styles.css`

**Acceptance Criteria:**
- [ ] Library lists functions (from `GetFunctions`), shows the signed-in account, New/Import/Export, restart banner.
- [ ] Import flow: choose file (via bridge), pick collision policy, show result counts, restart banner.
- [ ] Builder: ① data (model) ② the number you want (measure) ③ filters (match / date range) ④ name; only ②+④ required.
- [ ] Plain-primary labels with muted technical term + "?" helper on data/measure/filter.
- [ ] Live preview: plain sentence + template + filled example.
- [ ] Save calls `SaveFunction`; the function appears in the library; restart banner shown.

**Verify:** Manual end-to-end in Task 7.

**Steps:**

- [ ] **Step 1: Create `ui/styles.css`** — copy the `<style>` block from `docs/index.html` into `ui/styles.css` (the design system is already there). Add a `.helper` tooltip style:

```css
.helper{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:50%;
  background:linear-gradient(135deg,var(--gradient-start),var(--gradient-end));color:#fff;font-size:11px;
  font-weight:700;cursor:help;margin-left:6px}
.tech-term{color:var(--text-tertiary);font-weight:400}
.restart-banner{background:#fef3c7;border-left:4px solid #f59e0b;color:#92400e;padding:10px 14px;border-radius:8px;margin:10px 0}
```

- [ ] **Step 2: Create `ui/index.html`** — a single page with two views (library + editor) toggled by a class. Bridge wrapper inlined.

```html
<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CubeConnector</title>
<link href="styles.css" rel="stylesheet"></head>
<body>
  <div id="app">
    <div id="account" class="subtitle"></div>
    <div id="restart" class="restart-banner" style="display:none">✓ Saved. Close and reopen Excel to use your changes.</div>

    <!-- LIBRARY VIEW -->
    <section id="libraryView">
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn btn-primary" onclick="newFunction()">+ New formula</button>
        <button class="btn" onclick="doImport()">Import formulas someone shared</button>
        <button class="btn" onclick="doExport()">Export</button>
      </div>
      <h3 class="section-title" style="margin-top:14px">Your formulas</h3>
      <div id="functionList"></div>
    </section>

    <!-- EDITOR VIEW -->
    <section id="editorView" style="display:none">
      <a href="#" onclick="showLibrary();return false;">‹ Back</a>
      <h3>① What data? <span class="tech-term">(model)</span><span class="helper" title="The Power BI dataset your formula pulls from. Pick the one that has the numbers you need.">?</span></h3>
      <select id="modelSelect" class="cc-input" onchange="onModelChange()"></select>
      <h3>② The number you want <span class="tech-term">(measure)</span><span class="helper" title="A ready-made calculation in your data, like Net Amount or Total Volume. Pick the one this formula returns.">?</span></h3>
      <select id="measureSelect" class="cc-input" onchange="renderPreview()"></select>
      <h3>③ Let people filter by… <span class="subtitle">(optional)</span><span class="helper" title="A filter lets whoever uses your formula narrow the result — e.g. to an account or a date range.">?</span></h3>
      <div id="filterList"></div>
      <button class="btn" onclick="addFilter()">+ Add a filter</button>
      <h3>④ Name it</h3>
      <input id="friendlyName" class="cc-input" placeholder="e.g. Net Amount" oninput="renderPreview()">
      <div id="nameHint" class="subtitle"></div>
      <div id="preview" class="formula-preview"></div>
      <button class="btn btn-primary" onclick="saveFunction()" style="margin-top:12px">Save formula</button>
    </section>
  </div>
  <script src="app.js"></script>
</body></html>
```

- [ ] **Step 3: Create `ui/app.js`** — bridge wrapper + state + rendering. (Complete logic; `cc` is the host object.)

```javascript
const cc = window.chrome.webview.hostObjects.cc;
let MODEL = { measures: [], columns: [] };
let CURRENT = null; // function being edited

async function call(p){ const s = await p; const o = JSON.parse(s); if(o.error) throw new Error(o.error); return o; }

async function boot(){
  try { const a = await call(cc.GetAccount()); document.getElementById('account').textContent = 'Signed in: ' + (a.upn||'(unknown)'); }
  catch(e){ document.getElementById('account').textContent = 'Not signed in'; }
  await refreshLibrary();
}

async function refreshLibrary(){
  const o = await call(cc.GetFunctions());
  const list = document.getElementById('functionList'); list.innerHTML = '';
  (o.functions||[]).forEach(f => {
    const div = document.createElement('div'); div.className='function-item';
    div.innerHTML = `<div class="function-name">${f.FunctionName}</div>
      <div class="function-meta">${f.MeasureName||''} · ${(f.Parameters||[]).length} filters
      <a href="#" onclick="editFunction('${f.FunctionName}');return false;">Edit</a>
      <a href="#" onclick="delFunction('${f.FunctionName}');return false;">Delete</a></div>`;
    list.appendChild(div);
  });
}

function showLibrary(){ document.getElementById('editorView').style.display='none';
  document.getElementById('libraryView').style.display='block'; }
function showEditor(){ document.getElementById('libraryView').style.display='none';
  document.getElementById('editorView').style.display='block'; }

async function newFunction(){
  CURRENT = { FunctionName:'', MeasureName:'', DatasetId:'', TenantId:'', Parameters:[] };
  showEditor(); await loadModels(); document.getElementById('friendlyName').value=''; renderFilters(); renderPreview();
}

async function loadModels(){
  const sel = document.getElementById('modelSelect'); sel.innerHTML = '<option>Loading…</option>';
  const o = await call(cc.ListDatasets());
  sel.innerHTML='';
  (o.datasets||[]).forEach(d => { const opt=document.createElement('option');
    opt.value = JSON.stringify({id:d.Id, group:d.WorkspaceId});
    opt.textContent = (d.WorkspaceName||'') + ' ▸ ' + d.Name; sel.appendChild(opt); });
  if (sel.options.length) await onModelChange();
}

async function onModelChange(){
  const sel = document.getElementById('modelSelect'); if(!sel.value) return;
  const {id, group} = JSON.parse(sel.value);
  CURRENT.DatasetId = id; CURRENT._group = group;
  const ms = document.getElementById('measureSelect'); ms.innerHTML='<option>Loading…</option>';
  try { MODEL = await call(cc.GetModel(id, group||'')); }
  catch(e){ ms.innerHTML='<option>Couldn\'t read this data</option>'; return; }
  ms.innerHTML='';
  MODEL.measures.forEach(m => { const o=document.createElement('option'); o.value=m.Name; o.textContent=m.Name; ms.appendChild(o); });
  renderFilters(); renderPreview();
}

function addFilter(){
  CURRENT.Parameters.push({ Name:'', TableName:'', FieldName:'', DataType:'text', FilterType:'List', IsOptional:true, _kind:'match' });
  renderFilters(); renderPreview();
}
function renderFilters(){
  const wrap = document.getElementById('filterList'); wrap.innerHTML='';
  CURRENT.Parameters.filter(p=>!p._isEnd).forEach((p) => {
    const idx = CURRENT.Parameters.indexOf(p);
    const fields = MODEL.columns.map(c => `<option value='${c.Table}||${c.Name}||${c.DataType}'
      ${p.TableName===c.Table&&p.FieldName===c.Name?'selected':''}>${c.Table} · ${c.Name}</option>`).join('');
    const div = document.createElement('div'); div.className='parameter-card';
    div.innerHTML = `<select class="cc-input" onchange="setField(${idx}, this.value)"><option value="">choose a field…</option>${fields}</select>
      <label><input type="radio" name="kind${idx}" ${p._kind!=='range'?'checked':''} onchange="setKind(${idx},'match')"> Match value(s)</label>
      <label><input type="radio" name="kind${idx}" ${p._kind==='range'?'checked':''} onchange="setKind(${idx},'range')"> Date range</label>
      <input class="cc-input" placeholder="filter name" value="${p.Name||''}" oninput="setName(${idx}, this.value)">
      <a href="#" onclick="removeFilter(${idx});return false;">remove</a>`;
    wrap.appendChild(div);
  });
}
function setField(i,v){ const [t,f,dt]=v.split('||'); CURRENT.Parameters[i].TableName=t; CURRENT.Parameters[i].FieldName=f;
  CURRENT.Parameters[i].DataType = mapType(dt); if(!CURRENT.Parameters[i].Name) CURRENT.Parameters[i].Name=suggest(f); renderPreview(); }
function setName(i,v){ CURRENT.Parameters[i].Name=v; renderPreview(); }
function setKind(i,k){ CURRENT.Parameters[i]._kind=k; renderPreview(); }
function removeFilter(i){ CURRENT.Parameters.splice(i,1); renderFilters(); renderPreview(); }
function mapType(dt){ dt=(dt||'').toLowerCase(); if(dt.includes('date')||dt.includes('time'))return 'date';
  if(['integer','int64','number','double','decimal','currency'].includes(dt))return 'number'; return 'text'; }
function suggest(f){ return (f||'param').replace(/[^A-Za-z0-9]/g,'').toLowerCase(); }

function paramNames(){
  const out=[];
  CURRENT.Parameters.filter(p=>!p._isEnd).forEach(p=>{
    if(p._kind==='range'){ out.push((p.Name||'from')+'_start',(p.Name||'to')+'_end'); }
    else out.push(p.Name||'value');
  });
  return out;
}
function renderPreview(){
  const friendly = document.getElementById('friendlyName').value || 'Formula';
  const name = 'CC.' + friendly.replace(/[^A-Za-z0-9_]/g,'') ;
  document.getElementById('nameHint').innerHTML = "In Excel you'll type: <b>="+name+"(…)</b>";
  const measure = document.getElementById('measureSelect').value || 'the value';
  const names = paramNames();
  const tmpl = '='+name+'('+names.join(', ')+')';
  const ex = '='+name+'('+names.map(n=>n.includes('date')||n.includes('start')||n.includes('end')?'"1/1/2025"':'"4000"').join(',')+')';
  document.getElementById('preview').innerHTML =
    `<div><b>How you'll use it</b></div><div>Returns <b>${measure}</b>${names.length?`, filtered by ${names.join(', ')}`:''}.</div>
     <code>${tmpl}</code><div class="subtitle">Example:</div><code>${ex}</code>`;
}

async function saveFunction(){
  const friendly = document.getElementById('friendlyName').value.trim();
  if(!document.getElementById('measureSelect').value || !friendly){ alert('Pick a number and give it a name.'); return; }
  // Expand range filters into RangeStart/RangeEnd pairs; assign positions.
  const params=[]; let pos=0;
  CURRENT.Parameters.filter(p=>!p._isEnd).forEach(p=>{
    if(p._kind==='range'){
      params.push({Name:(p.Name||'from')+'_start',Position:pos++,TableName:p.TableName,FieldName:p.FieldName,DataType:'date',FilterType:'RangeStart',IsOptional:true});
      params.push({Name:(p.Name||'to')+'_end',Position:pos++,TableName:p.TableName,FieldName:p.FieldName,DataType:'date',FilterType:'RangeEnd',IsOptional:true});
    } else {
      params.push({Name:p.Name||'value',Position:pos++,TableName:p.TableName,FieldName:p.FieldName,DataType:p.DataType||'text',FilterType:'List',IsOptional:true});
    }
  });
  const dto = { FunctionName:'CC.'+friendly.replace(/[^A-Za-z0-9_]/g,''), MeasureName:'['+document.getElementById('measureSelect').value+']',
    DatasetId:CURRENT.DatasetId, TenantId:CURRENT.TenantId||'', Parameters:params };
  await call(cc.SaveFunction(JSON.stringify(dto)));
  document.getElementById('restart').style.display='block';
  await refreshLibrary(); showLibrary();
}

async function editFunction(name){
  const o = await call(cc.GetFunctions());
  const f = (o.functions||[]).find(x=>x.FunctionName===name); if(!f) return;
  CURRENT = JSON.parse(JSON.stringify(f)); CURRENT._group='';
  showEditor(); await loadModels();
  // best-effort select existing dataset; then set name/measure
  document.getElementById('friendlyName').value = name.replace(/^CC\./,'');
  // collapse RangeStart/RangeEnd pairs back to one 'range' filter for editing
  const collapsed=[]; (f.Parameters||[]).forEach(p=>{ if(p.FilterType==='RangeEnd')return;
    collapsed.push({...p,_kind:p.FilterType==='RangeStart'?'range':'match'}); });
  CURRENT.Parameters = collapsed; renderFilters(); renderPreview();
}
async function delFunction(name){ if(!confirm('Delete '+name+'?'))return; await call(cc.DeleteFunction(name));
  document.getElementById('restart').style.display='block'; await refreshLibrary(); }

async function doImport(){
  const path = prompt('Path to the shared formulas file (.json):'); if(!path) return;
  const policy = confirm('Overwrite formulas that already exist? (Cancel = keep both)') ? 'Overwrite' : 'KeepBoth';
  const r = await call(cc.ImportFunctions(path, policy));
  alert(`Imported: ${r.added} new, ${r.overwritten} replaced, ${r.skipped} skipped.`);
  document.getElementById('restart').style.display='block'; await refreshLibrary();
}
async function doExport(){
  const path = prompt('Save shared file to (.json):'); if(!path) return;
  await call(cc.ExportFunctions(JSON.stringify([]), path)); alert('Exported to '+path);
}

boot();
```

> Note: `prompt()`-based file paths are a v1 placeholder for the file dialog. A later refinement replaces `doImport`/`doExport` with native `OpenFileDialog`/`SaveFileDialog` invoked through two new bridge methods (`PickOpenFile()`, `PickSaveFile()`); out of scope here to keep the task focused.

- [ ] **Step 4: Add `.cc-input` style** to `ui/styles.css`:

```css
.cc-input{width:100%;padding:.6rem .8rem;border:1.5px solid var(--border);border-radius:.5rem;
  font-family:'DM Sans',sans-serif;font-size:.9rem;margin:4px 0}
```

- [ ] **Step 5: Build + manual** — rebuild; open Manage Formulas; create a formula end-to-end (covered in Task 7).

- [ ] **Step 6: Commit**

```bash
git add CubeConnector/ui
git commit -m "feat: wizard UI (library, import, builder) on WebView2 bridge"
```

---

## Task 7: End-to-end verification + remove dev self-check

**Goal:** Prove the full loop and clean up dev-only code.

**Files:**
- Modify: `CubeConnector/WizardBridge.cs` (remove `SelfCheck`)

**Acceptance Criteria:**
- [ ] Build → open Excel → Manage Formulas → create a formula (model → measure → one match filter + one date range) → Save → restart banner shows.
- [ ] Close + reopen Excel → the new `=CC.<Name>(...)` resolves in a cell (not #NAME?).
- [ ] Export the function to a file; delete it; Import it back (KeepBoth and Overwrite both tested); restart; it works.
- [ ] Switch account (Sign in as different) and confirm the dataset list changes.
- [ ] `SelfCheck` removed; build still green.

**Verify:** The manual sequence above, performed once.

**Steps:**

- [ ] **Step 1: Run the full manual sequence** (build with the MSBuild command; close Excel between rebuilds). Record any failures and fix in the relevant task.
- [ ] **Step 2: Remove `SelfCheck()`** from `WizardBridge.cs`.
- [ ] **Step 3: Build** — MSBuild. Expected exit 0.
- [ ] **Step 4: Commit**

```bash
git add CubeConnector/WizardBridge.cs
git commit -m "chore: remove dev self-check; wizard end-to-end verified"
```

---

## Self-Review

**Spec coverage:** §2 hosting → Tasks 5,6. §3 store/migration → Tasks 1,2. §4 flows (consumer import + builder) → Task 6. §5 builder/labels/preview/executeQueries → Tasks 3,6. §6 restart UX → Task 6 (banner). §7 errors → bridge `{error}` envelope (Task 4) + UI handling (Task 6). §8 testing → Task 7. §10 components → Tasks 1–6. All covered.

**Placeholder scan:** The `prompt()` file paths and the `executeQueries` "validate or fall back" are explicitly bounded with concrete fallbacks, not vague TODOs. No "add error handling"–style gaps (bridge envelope + UI try/catch are concrete).

**Type consistency:** `UDFConfig`/`ParameterConfig`/`FilterType` reused from existing code; `WizardBridge` (de)serializes `UDFConfig` (PascalCase properties → matches JS `dto` keys `FunctionName`/`MeasureName`/`Parameters`). `ModelMetadata`/`ModelColumn`/`ModelMeasure` consistent across Tasks 3,4,6. Bridge method names match JS calls in `app.js`.
