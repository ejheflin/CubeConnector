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
