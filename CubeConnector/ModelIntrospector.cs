/*
 * CubeConnector - ModelIntrospector
 *
 * Reusable API for enumerating the tables, fields, and measures of a Power BI
 * semantic model using ONLY the existing Analyze-in-Excel (MSOLAP) workbook
 * connection -- no XMLA endpoint, no fresh ADOMD.NET connection, Pro-compatible.
 *
 * Validated approach (see IntrospectionProbe results):
 *   1. Discover the OLAP connection by TYPE (its MSOLAP/pbiazure connection string),
 *      not by a hardcoded name.
 *   2. Parse tenantId / datasetId out of that connection string.
 *   3. Build a DEFAULT-mode connection from the same string (the AiE connection is
 *      cube-mode, which rejects DAX) and run INFO.VIEW.TABLES/COLUMNS/MEASURES
 *      through a hidden ListObject/QueryTable.
 *
 * This class is UI-free and is NOT wired into any production code path. The core
 * never touches the production "CubeConnector" connection or the JSON config; it
 * creates its own uniquely-named temp connection + hidden sheet and cleans them up.
 */

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace CubeConnector
{
    public class TableInfo
    {
        public string Name;
        public bool IsHidden;
        public string Description;
        public string StorageMode;
    }

    public class ColumnInfo
    {
        public string Table;
        public string Name;
        public string DataType;     // Power BI data type: Text, Integer, Number, DateTime, Boolean, ...
        public bool IsHidden;
        public string Description;
        public string SortByColumn;
        public string FormatString;

        /// <summary>Fully-qualified MDX-style reference, e.g. [Account].[AccountID].</summary>
        public string QualifiedName { get { return "[" + Table + "].[" + Name + "]"; } }

        /// <summary>Maps the model data type to the config's coarse dataType (text/number/date).</summary>
        public string ConfigDataType
        {
            get
            {
                string t = (DataType ?? "").ToLowerInvariant();
                if (t.Contains("date") || t.Contains("time")) return "date";
                if (t == "integer" || t == "int64" || t == "number" || t == "double" ||
                    t == "decimal" || t == "currency") return "number";
                return "text";
            }
        }
    }

    public class MeasureInfo
    {
        public string Table;        // home table
        public string Name;
        public string DataType;
        public string Description;
        public string DisplayFolder;
        public bool IsHidden;
        public string FormatString;

        /// <summary>Measure reference as used in DAX/config, e.g. [AmtNet].</summary>
        public string BracketName { get { return "[" + Name + "]"; } }
    }

    public class SemanticModelInfo
    {
        public string ConnectionName;
        public string TenantId;
        public string DatasetId;
        public string ConnectionString;
        public List<TableInfo> Tables = new List<TableInfo>();
        public List<ColumnInfo> Columns = new List<ColumnInfo>();
        public List<MeasureInfo> Measures = new List<MeasureInfo>();

        /// <summary>Visible (non-hidden) columns for a given table.</summary>
        public List<ColumnInfo> ColumnsFor(string tableName)
        {
            var list = new List<ColumnInfo>();
            foreach (var c in Columns)
                if (string.Equals(c.Table, tableName, StringComparison.OrdinalIgnoreCase))
                    list.Add(c);
            return list;
        }
    }

    public static class ModelIntrospector
    {
        private const string TempConnName = "CC_Model_Conn";
        private const string TempSheetName = "__CC_Model_Query__";
        private const string TempListObjName = "CC_Model_QT";

        /// <summary>
        /// Enumerate the semantic model reachable through the workbook's existing OLAP connection.
        /// Throws InvalidOperationException if no OLAP connection is present.
        /// </summary>
        public static SemanticModelInfo Introspect(Excel.Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            string connName, rawConnString;
            if (FindOlapConnection(workbook, out connName, out rawConnString) == null)
                throw new InvalidOperationException(
                    "No OLAP (MSOLAP / Power BI) connection found in this workbook. " +
                    "Establish an Analyze-in-Excel connection first.");

            string connString = EnsureOleDbPrefix(rawConnString);

            var info = new SemanticModelInfo
            {
                ConnectionName = connName,
                ConnectionString = connString,
                TenantId = ExtractTenantId(rawConnString),
                DatasetId = ExtractDatasetId(rawConnString)
            };

            PopulateFromConnString(workbook, info, connString);
            return info;
        }

        /// <summary>
        /// Fixed Microsoft "Analyze in Excel" client id used in the MSOLAP Identity Provider
        /// field. This is the SAME for every user/tenant — it is NOT a tenant id. (The third
        /// Identity Provider value is an application/client id; the /common authority + the
        /// user's interactive sign-in resolve the actual tenant.)
        /// </summary>
        public const string AnalyzeInExcelClientId = "929d0ec0-7a41-4b1e-bc7c-b754a28bddcc";

        /// <summary>
        /// Build the Pro-compatible MSOLAP connection string for a dataset from its catalog
        /// (e.g. "sobe_wowvirtualserver-&lt;guid&gt;" or a bare GUID). The tenantId parameter is
        /// accepted for compatibility but intentionally NOT used in the Identity Provider field
        /// (that field needs the fixed AiE client id, not a tenant).
        /// </summary>
        public static string BuildConnectionString(string initialCatalog, string tenantId)
        {
            string catalog = initialCatalog;
            // Accept a bare dataset GUID and add the standard Power BI catalog prefix.
            if (Guid.TryParse(initialCatalog, out _))
                catalog = "sobe_wowvirtualserver-" + initialCatalog;

            return "OLEDB;Provider=MSOLAP.8;Integrated Security=ClaimsToken;Persist Security Info=True;" +
                   "Initial Catalog=" + catalog + ";" +
                   "Data Source=pbiazure://api.powerbi.com;" +
                   "MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;" +
                   "Identity Provider=https://login.microsoftonline.com/common, " +
                   "https://analysis.windows.net/powerbi/api, " + AnalyzeInExcelClientId + ";" +
                   "Update Isolation Level=2";
        }

        /// <summary>
        /// Introspect a semantic model from a dataset GUID + tenant id by building the
        /// connection from scratch -- the wizard path (no pre-existing Get Data connection).
        /// </summary>
        public static SemanticModelInfo IntrospectDataset(Excel.Workbook workbook, string datasetId, string tenantId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            string connString = BuildConnectionString(datasetId, tenantId);
            var info = new SemanticModelInfo
            {
                ConnectionName = "(built from dataset id)",
                ConnectionString = connString,
                TenantId = tenantId,
                DatasetId = ExtractDatasetId(connString)
            };
            PopulateFromConnString(workbook, info, connString);
            return info;
        }

        /// <summary>Run the three INFO.VIEW queries over a connection string into <paramref name="info"/>.</summary>
        private static void PopulateFromConnString(Excel.Workbook workbook, SemanticModelInfo info, string connString)
        {
            CleanupArtifacts(workbook);
            Excel.ListObject lo = null;
            try
            {
                lo = SetupQueryTable(workbook, connString);

                foreach (var row in Query(lo, "EVALUATE INFO.VIEW.TABLES()"))
                {
                    info.Tables.Add(new TableInfo
                    {
                        Name = Get(row, "name"),
                        IsHidden = ParseBool(Get(row, "ishidden")),
                        Description = Get(row, "description"),
                        StorageMode = Get(row, "storagemode")
                    });
                }

                foreach (var row in Query(lo, "EVALUATE INFO.VIEW.COLUMNS()"))
                {
                    // Skip system RowNumber columns.
                    string dataCat = Get(row, "datacategory");
                    string type = Get(row, "type");
                    if (string.Equals(dataCat, "RowNumber", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(type, "RowNumber", StringComparison.OrdinalIgnoreCase))
                        continue;

                    info.Columns.Add(new ColumnInfo
                    {
                        Table = Get(row, "table"),
                        Name = Get(row, "name"),
                        DataType = Get(row, "datatype"),
                        IsHidden = ParseBool(Get(row, "ishidden")),
                        Description = Get(row, "description"),
                        SortByColumn = Get(row, "sortbycolumn"),
                        FormatString = Get(row, "formatstring")
                    });
                }

                foreach (var row in Query(lo, "EVALUATE INFO.VIEW.MEASURES()"))
                {
                    info.Measures.Add(new MeasureInfo
                    {
                        Table = Get(row, "table"),
                        Name = Get(row, "name"),
                        DataType = Get(row, "datatype"),
                        Description = Get(row, "description"),
                        DisplayFolder = Get(row, "displayfolder"),
                        IsHidden = ParseBool(Get(row, "ishidden")),
                        FormatString = Get(row, "formatstring")
                    });
                }
            }
            finally
            {
                CleanupArtifacts(workbook);
            }
        }

        // ---- connection discovery / parsing ----

        private static Excel.WorkbookConnection FindOlapConnection(Excel.Workbook wb, out string name, out string connString)
        {
            name = null; connString = null;
            foreach (Excel.WorkbookConnection c in wb.Connections)
            {
                try
                {
                    Excel.OLEDBConnection oledb = c.OLEDBConnection;
                    if (oledb == null) continue;
                    string cs = oledb.Connection as string;
                    if (string.IsNullOrEmpty(cs)) continue;
                    if (cs.IndexOf("MSOLAP", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        cs.IndexOf("pbiazure", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        name = c.Name; connString = cs; return c;
                    }
                }
                catch { }
            }
            return null;
        }

        private static string EnsureOleDbPrefix(string cs)
        {
            if (string.IsNullOrEmpty(cs)) return cs;
            return cs.StartsWith("OLEDB;", StringComparison.OrdinalIgnoreCase) ? cs : "OLEDB;" + cs;
        }

        public static string ExtractDatasetId(string connString)
        {
            var m = Regex.Match(connString ?? "", @"Initial\s+Catalog\s*=\s*([^;]+)", RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value.Trim() : null;
        }

        public static string ExtractTenantId(string connString)
        {
            var m = Regex.Match(connString ?? "",
                @"Identity\s+Provider\s*=.*?,\s*[^,;]+,\s*([0-9a-fA-F\-]{36})", RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value.Trim();
            var all = Regex.Matches(connString ?? "", @"[0-9a-fA-F\-]{36}");
            return all.Count > 0 ? all[all.Count - 1].Value : null;
        }

        // ---- query execution ----

        private static Excel.ListObject SetupQueryTable(Excel.Workbook wb, string connString)
        {
            Excel.WorkbookConnection probeConn = null;
            try { probeConn = wb.Connections[TempConnName]; } catch { }
            if (probeConn == null)
            {
                probeConn = wb.Connections.Add2(
                    Name: TempConnName,
                    Description: "Temporary CubeConnector model introspection",
                    ConnectionString: connString,
                    CommandText: "Model",
                    lCmdtype: Excel.XlCmdType.xlCmdDefault,
                    CreateModelConnection: Type.Missing,
                    ImportRelationships: Type.Missing);
            }

            Excel.Worksheet sheet = wb.Worksheets.Add();
            sheet.Name = TempSheetName;
            sheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            Excel.ListObject lo = sheet.ListObjects.Add(
                SourceType: Excel.XlListObjectSourceType.xlSrcExternal,
                Source: probeConn,
                LinkSource: true,
                XlListObjectHasHeaders: Excel.XlYesNoGuess.xlYes,
                Destination: sheet.Range["A1"]);
            lo.Name = TempListObjName;

            Excel.QueryTable qt = lo.QueryTable;
            qt.CommandType = Excel.XlCmdType.xlCmdDefault;
            qt.CommandText = "EVALUATE { 1 }";
            qt.Refresh(BackgroundQuery: false);
            return lo;
        }

        /// <summary>
        /// Run a DAX query and return each data row as a dictionary keyed by normalized
        /// column name (lowercased, brackets stripped). Parsing by name is robust to
        /// column-order changes between engine versions.
        /// </summary>
        private static List<Dictionary<string, string>> Query(Excel.ListObject lo, string dax)
        {
            Excel.QueryTable qt = lo.QueryTable;
            qt.CommandType = Excel.XlCmdType.xlCmdDefault;
            qt.CommandText = dax;
            qt.Refresh(BackgroundQuery: false);

            var rows = new List<Dictionary<string, string>>();
            object raw = lo.Range.Value2;
            object[,] grid = raw as object[,];
            if (grid == null) return rows;

            int r0 = grid.GetLowerBound(0), r1 = grid.GetUpperBound(0);
            int c0 = grid.GetLowerBound(1), c1 = grid.GetUpperBound(1);

            var headers = new List<string>();
            for (int c = c0; c <= c1; c++)
            {
                object v = grid[r0, c];
                headers.Add(Normalize(v == null ? "" : v.ToString()));
            }

            for (int r = r0 + 1; r <= r1; r++)
            {
                var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int c = c0; c <= c1; c++)
                {
                    object v = grid[r, c];
                    d[headers[c - c0]] = v == null ? "" : v.ToString();
                }
                rows.Add(d);
            }
            return rows;
        }

        private static string Normalize(string s)
        {
            if (s == null) return "";
            return s.Trim().Trim('[', ']').Trim().ToLowerInvariant();
        }

        private static string Get(Dictionary<string, string> row, string key)
        {
            string v;
            return (row != null && row.TryGetValue(key, out v)) ? v : "";
        }

        private static bool ParseBool(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            return s.Equals("True", StringComparison.OrdinalIgnoreCase) || s == "1" ||
                   s.Equals("TRUE", StringComparison.OrdinalIgnoreCase);
        }

        private static void CleanupArtifacts(Excel.Workbook wb)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            bool prev = app.DisplayAlerts;
            app.DisplayAlerts = false;
            try { wb.Worksheets[TempSheetName].Delete(); } catch { }
            try { wb.Connections[TempConnName].Delete(); } catch { }
            app.DisplayAlerts = prev;
        }
    }
}
