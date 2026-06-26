/*
 * CubeConnector - Excel-DNA add-in for querying Power BI datasets
 * Copyright (C) 2026
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 *
 * For enterprise licensing options, please contact the project maintainers.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using Microsoft.Office.Core;
using ExcelDna.Integration.CustomUI;

namespace CubeConnector
{
    /// <summary>
    /// Excel-DNA add-in that dynamically registers functions based on JSON configuration
    /// </summary>
    /// 
    public class ReloadResult
    {
        public int Reloaded;
        public bool RemovedNeedRestart;
    }

    public class DynamicFunctionRegistration : IExcelAddIn
    {
        // Static reference to Excel Application for cache access
        public static Microsoft.Office.Interop.Excel.Application ExcelApp { get; private set; }

        // Track registered function name -> parameter count (arity), to detect new / removed /
        // arity-changed functions across a runtime reload.
        private static readonly System.Collections.Generic.Dictionary<string, int> _registeredArity =
            new System.Collections.Generic.Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);

        public void AutoOpen()  // NOT static anymore
        {
            try
            {
                RuntimeBootstrap.EnsureWebView2Loader();
                FunctionStore.MigrateLegacyIfNeeded();

                // Store Excel Application reference for cache access
                ExcelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

                // Start auto-refresh if the user left it enabled (persists across restarts).
                AutoRefreshManager.Initialize(ExcelApp);

                var configs = ConfigurationStore.GetAllConfigs();
                if (configs == null || configs.Count == 0) return;

                RegisterFunctionsFromConfig(configs);
                foreach (var c in configs) _registeredArity[c.FunctionName] = c.Parameters?.Count ?? 0;
                AddContextMenuItems();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }

        /// <summary>
        /// Invalidate config cache, re-register all functions, detect deletions.
        /// Safe to call at runtime — Excel-DNA can register/update but not unregister.
        /// </summary>
        public static ReloadResult ReloadFunctions()
        {
            ConfigurationStore.Invalidate();
            var configs = ConfigurationStore.GetAllConfigs() ?? new System.Collections.Generic.List<UDFConfig>();

            // Excel-DNA can register NEW functions at runtime, but only from a macro context
            // (xlfRegister), and it cannot unregister or change the arity of one already
            // registered. So: register only the new ones (via QueueAsMacro); a removal or an
            // arity change requires a restart. A same-arity edit needs nothing here — the
            // delegate reads fresh config at call time and we already invalidated the cache.
            var newConfigs = new System.Collections.Generic.List<UDFConfig>();
            var currentNames = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            bool needRestart = false;

            foreach (var c in configs)
            {
                currentNames.Add(c.FunctionName);
                int arity = c.Parameters?.Count ?? 0;
                if (!_registeredArity.TryGetValue(c.FunctionName, out int prevArity))
                    newConfigs.Add(c);                 // brand new -> register live
                else if (prevArity != arity)
                    needRestart = true;                // arity changed -> old signature lingers until restart
            }
            foreach (var prev in _registeredArity.Keys)
                if (!currentNames.Contains(prev)) { needRestart = true; break; }  // removed -> lingers until restart

            if (newConfigs.Count > 0)
                ExcelAsyncUtil.QueueAsMacro(() => RegisterFunctionsFromConfig(newConfigs));

            _registeredArity.Clear();
            foreach (var c in configs) _registeredArity[c.FunctionName] = c.Parameters?.Count ?? 0;

            return new ReloadResult { Reloaded = configs.Count, RemovedNeedRestart = needRestart };
        }

        public void AutoClose()
        {
            try
            {
                AutoRefreshManager.Detach();

                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var cellMenu = app.CommandBars["Cell"];

                try { cellMenu.Controls["CubeConnector - Drill to Details"].Delete(); } catch { }
                try { cellMenu.Controls["CubeConnector - Drill to Pivot"].Delete(); } catch { }
                try { cellMenu.Controls["CubeConnector - Refresh Cache"].Delete(); } catch { }
            }
            catch { }
        }
        /// <summary>
        /// The shared refresh pipeline used by both the ribbon Refresh button and auto-refresh:
        /// ensure prerequisites, then refresh only the cells that need it. Throws on failure —
        /// callers decide how to surface errors.
        /// </summary>
        public static void RefreshCore()
        {
            EnsureConnectionExists();
            EnsureCacheExists();
            var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            var workbook = app.ActiveWorkbook;
            new RefreshManager(app, workbook).RefreshAll();
        }
        /// <summary>
        /// Register all functions from configuration
        /// </summary>
        private static void RegisterFunctionsFromConfig(List<UDFConfig> configs)
        {
            //System.Windows.Forms.MessageBox.Show("RegisterFunctionsFromConfig STARTING", "Debug");

            var registrationItems = new List<ExcelFunctionRegistration>();

            foreach (var config in configs)
            {
                //System.Windows.Forms.MessageBox.Show($"Creating registration for: {config.FunctionName}", "Debug");

                var registration = CreateFunctionRegistration(config);
                if (registration != null)
                {
                    registrationItems.Add(registration);
                    //System.Windows.Forms.MessageBox.Show($"Added to list: {config.FunctionName}", "Debug");
                }
                // else: CreateFunctionRegistration already shows a diagnostic MessageBox for config errors
            }

            //System.Windows.Forms.MessageBox.Show($"About to register {registrationItems.Count} functions", "Debug");

            // Register each function
            foreach (var registration in registrationItems)
            {
                var attr = (ExcelFunctionAttribute)registration.FunctionAttributes;
                //System.Windows.Forms.MessageBox.Show($"Calling ExcelIntegration.RegisterDelegates for: {attr.Name}", "Debug");

                try
                {
                    ExcelIntegration.RegisterDelegates(
                        new List<Delegate> { registration.FunctionDelegate },
                        new List<object> { registration.FunctionAttributes },
                        new List<List<object>> { registration.ParameterAttributes }
                    );

                    //System.Windows.Forms.MessageBox.Show($"SUCCESS: {attr.Name}", "Debug");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(
                        $"EXCEPTION for {attr.Name}:\n\n{ex.Message}",
                        "Error");
                }
            }

            //System.Windows.Forms.MessageBox.Show("RegisterFunctionsFromConfig COMPLETE", "Debug");
        }        /// <summary>
                 /// Create a function registration for a specific config
                 /// </summary>
        private static ExcelFunctionRegistration CreateFunctionRegistration(UDFConfig config)
        {
            try
            {
                // Determine how many parameters this function needs
                int paramCount = config.Parameters?.Count ?? 0;

                // Excel-DNA supports up to 15 parameters
                if (paramCount > 15)
                {
                    System.Windows.Forms.MessageBox.Show(
                        $"Function '{config.FunctionName}' has {paramCount} parameters. Maximum is 15.",
                        "Configuration Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return null;
                }

                // Create delegate based on parameter count
                Delegate functionDelegate = CreateDelegateForParameterCount(config.FunctionName, paramCount);

                // Create function attributes
                var functionAttr = new ExcelFunctionAttribute
                {
                    Name = config.FunctionName,
                    Description = $"Retrieves {config.MeasureName} from Power BI dataset",
                    Category = "CubeConnector",
                    IsMacroType = false // Allow use in arithmetic expressions
                };

                // Create parameter attributes
                var parameterAttrs = new List<object>();
                if (config.Parameters != null)
                {
                    foreach (var param in config.Parameters.OrderBy(p => p.Position))
                    {
                        string description = BuildParameterDescription(param);
                        
                        parameterAttrs.Add(new ExcelArgumentAttribute
                        {
                            Name = param.Name,
                            Description = description,
                            AllowReference = false
                        });
                    }
                }

                return new ExcelFunctionRegistration
                {
                    FunctionDelegate = functionDelegate,
                    FunctionAttributes = functionAttr,
                    ParameterAttributes = parameterAttrs
                };
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error creating registration for '{config.FunctionName}':\n\n{ex.Message}",
                    "Registration Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return null;
            }
        }

        /// <summary>
        /// Build a descriptive parameter description for IntelliSense
        /// </summary>
        private static string BuildParameterDescription(ParameterConfig param)
        {
            string desc = $"{param.TableName}.{param.FieldName}";
            
            if (!string.IsNullOrEmpty(param.DataType))
            {
                desc += $" ({param.DataType})";
            }

            if (param.FilterType != FilterType.List)
            {
                desc += $" [{param.FilterType}]";
            }

            if (param.IsOptional)
            {
                desc += " - Optional";
            }

            return desc;
        }

        /// <summary>
        /// Create the appropriate delegate for the given parameter count
        /// </summary>
        private static Delegate CreateDelegateForParameterCount(string functionName, int paramCount)
        {
            // We create delegates that call a generic executor function
            switch (paramCount)
            {
                case 0:
                    return new Func<object>(() => ExecuteFunction(functionName));
                case 1:
                    return new Func<object, object>((p0) => ExecuteFunction(functionName, p0));
                case 2:
                    return new Func<object, object, object>((p0, p1) => ExecuteFunction(functionName, p0, p1));
                case 3:
                    return new Func<object, object, object, object>((p0, p1, p2) => ExecuteFunction(functionName, p0, p1, p2));
                case 4:
                    return new Func<object, object, object, object, object>((p0, p1, p2, p3) => ExecuteFunction(functionName, p0, p1, p2, p3));
                case 5:
                    return new Func<object, object, object, object, object, object>((p0, p1, p2, p3, p4) => ExecuteFunction(functionName, p0, p1, p2, p3, p4));
                case 6:
                    return new Func<object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5));
                case 7:
                    return new Func<object, object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5, p6) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5, p6));
                case 8:
                    return new Func<object, object, object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5, p6, p7) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5, p6, p7));
                case 9:
                    return new Func<object, object, object, object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5, p6, p7, p8) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5, p6, p7, p8));
                case 10:
                    return new Func<object, object, object, object, object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5, p6, p7, p8, p9) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5, p6, p7, p8, p9));
                case 11:
                    return new Func<object, object, object, object, object, object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10));
                case 12:
                    return new Func<object, object, object, object, object, object, object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11));
                case 13:
                    return new Func<object, object, object, object, object, object, object, object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12));
                case 14:
                    return new Func<object, object, object, object, object, object, object, object, object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13));
                case 15:
                    return new Func<object, object, object, object, object, object, object, object, object, object, object, object, object, object, object, object>((p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14) => ExecuteFunction(functionName, p0, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14));
                default:
                    throw new ArgumentException($"Unsupported parameter count: {paramCount}");
            }
        }

        /// <summary>
        /// Generic function executor - routes to cache lookup
        /// </summary>
        private static object ExecuteFunction(string functionName, params object[] args)
        {
            try
            {
                // Build cache key from function name and parameters
                string cacheKey = CacheKey.Build(functionName, args);

                // Look up in cache
                return CacheManager.Lookup(cacheKey);
            }
            catch (Exception ex)
            {
                return $"#ERROR: {ex.Message}";
            }
        }
        private static void AddContextMenuItems()
        {
            //System.Windows.Forms.MessageBox.Show("AddContextMenuItems is running!", "Debug");
            var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            var cellMenu = app.CommandBars["Cell"];

            // Remove old items if they exist
            try { cellMenu.Controls["CubeConnector - Drill to Details"].Delete(); } catch { }
            try { cellMenu.Controls["CubeConnector - Drill to Pivot"].Delete(); } catch { }
            try { cellMenu.Controls["CubeConnector - Refresh"].Delete(); } catch { }

            // Add "Drill to Details"
            var detailsButton = (Microsoft.Office.Core.CommandBarButton)
                cellMenu.Controls.Add(Type: Microsoft.Office.Core.MsoControlType.msoControlButton, Temporary: true);
            detailsButton.Caption = "CubeConnector - Drill to Details";
            detailsButton.OnAction = "DrillToDetailsHandler";

            // "Drill to Pivot" temporarily removed pending further testing. The delete
            // calls above still run so any stale item from a prior version is cleaned up.
            // Re-add the button here (and the ribbon button in CubeConnectorRibbon) to re-enable.

            // Add "Refresh"
            var refreshButton = (Microsoft.Office.Core.CommandBarButton)
                cellMenu.Controls.Add(Type: Microsoft.Office.Core.MsoControlType.msoControlButton, Temporary: true);
            refreshButton.Caption = "CubeConnector - Refresh";
            refreshButton.OnAction = "RefreshCacheHandler";
        }

        public static void DrillToDetailsHandler()
        {
            try
            {
                EnsureConnectionExists();
                EnsureCacheExists();

                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var activeCell = app.ActiveCell;
                var workbook = app.ActiveWorkbook;

                // Check if cell contains multiple CubeConnector functions
                if (activeCell.HasFormula)
                {
                    string formula = activeCell.Formula.ToString();
                    var allFunctionNames = ConfigurationStore.GetAllConfigs().Select(c => c.FunctionName).ToList();
                    int udfCount = CountUDFsInFormula(formula, allFunctionNames);

                    if (udfCount > 1)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            $"The selected cell contains {udfCount} CubeConnector functions.\n\n" +
                            "Drill to Details only supports cells with a single CubeConnector function.\n\n" +
                            "Please select a cell with only one function and try again.",
                            "Multiple Functions Detected",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        return;
                    }
                }

                var manager = new DrillthroughManager(app, workbook);
                manager.ExecuteDrillthrough(activeCell);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }

        /// <summary>
        /// Count how many CubeConnector UDFs appear in a formula
        /// </summary>
        private static int CountUDFsInFormula(string formula, List<string> functionNames)
        {
            int count = 0;
            string upperFormula = formula.ToUpper();

            foreach (var funcName in functionNames)
            {
                string searchPattern = funcName.ToUpper() + "(";
                int index = 0;

                // Count all occurrences of "FunctionName("
                while ((index = upperFormula.IndexOf(searchPattern, index)) != -1)
                {
                    count++;
                    index += searchPattern.Length;
                }
            }

            return count;
        }

        public static void DrillToPivotHandler()
        {
            try
            {
                EnsureConnectionExists();
                EnsureCacheExists();
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var activeCell = app.ActiveCell;
                var workbook = app.ActiveWorkbook;

                // Check if cell contains multiple CubeConnector functions
                if (activeCell.HasFormula)
                {
                    string formula = activeCell.Formula.ToString();
                    var allFunctionNames = ConfigurationStore.GetAllConfigs().Select(c => c.FunctionName).ToList();
                    int udfCount = CountUDFsInFormula(formula, allFunctionNames);

                    if (udfCount > 1)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            $"The selected cell contains {udfCount} CubeConnector functions.\n\n" +
                            "Drill to Pivot only supports cells with a single CubeConnector function.\n\n" +
                            "Please select a cell with only one function and try again.",
                            "Multiple Functions Detected",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        return;
                    }
                }

                // Create manager instance and execute
                var manager = new PivotManager(app, workbook);
                manager.ExecuteDrillToPivot(activeCell);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error drilling to pivot:\n\n{ex.Message}",
                    "CubeConnector Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public static void RefreshCacheHandler()
        {
            try
            {
                EnsureConnectionExists();
                EnsureCacheExists();
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var workbook = app.ActiveWorkbook;

                // Create manager instance and execute
                var manager = new RefreshManager(app, workbook);
                manager.RefreshAll();

                //System.Windows.Forms.MessageBox.Show(
                //    "Cache refreshed successfully!",
                //    "CubeConnector",
                //    System.Windows.Forms.MessageBoxButtons.OK,
                //    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error refreshing cache:\n\n{ex.Message}",
                    "CubeConnector Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        /// <summary>Human-readable workbook-connection name for a dataset: "&lt;ModelName&gt; (&lt;shortId&gt;)".</summary>
        internal static string ConnectionNameForDataset(UDFConfig config)
        {
            string model = (config != null ? config.ModelName : null) ?? "";
            model = model.Trim();
            string label;
            if (model.Length == 0) label = "CubeConnector Data";
            else
            {
                var sb = new System.Text.StringBuilder();
                foreach (char c in model)
                    if (char.IsLetterOrDigit(c) || c == ' ' || c == '-' || c == '_') sb.Append(c);
                label = sb.ToString().Trim();
                if (label.Length == 0) label = "CubeConnector Data";
            }
            return label + " (" + ShortDatasetId(config != null ? config.DatasetId : null) + ")";
        }

        /// <summary>Short stable id (first 8 alphanumerics of the dataset GUID) for sheet/listobject names (Excel sheet names max 31 chars).</summary>
        internal static string ShortDatasetId(string datasetId)
        {
            var sb = new System.Text.StringBuilder();
            foreach (char c in datasetId ?? "") { if (char.IsLetterOrDigit(c)) sb.Append(c); if (sb.Length >= 8) break; }
            return sb.Length > 0 ? sb.ToString() : "ds";
        }

        /// <summary>Create the workbook connection for one dataset if it doesn't already exist.</summary>
        internal static void EnsureConnectionForDataset(UDFConfig config)
        {
            var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            var workbook = app.ActiveWorkbook;
            string connName = ConnectionNameForDataset(config);
            try { var existing = workbook.Connections[connName]; return; } catch { }
            string connectionString = ModelIntrospector.BuildConnectionString(config.DatasetId, config.TenantId);
            workbook.Connections.Add2(
                Name: connName,
                Description: "CubeConnector dataset connection",
                ConnectionString: connectionString,
                CommandText: "Model",
                lCmdtype: Microsoft.Office.Interop.Excel.XlCmdType.xlCmdDefault,
                CreateModelConnection: Type.Missing,
                ImportRelationships: Type.Missing);
        }

        internal static void EnsureConnectionExists()
        {
            try
            {
                var configs = ConfigurationStore.GetAllConfigs();
                if (configs == null || configs.Count == 0)
                    throw new Exception("No configuration found. Cannot create connection.");
                var seen = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
                foreach (var c in configs)
                    if (!string.IsNullOrEmpty(c.DatasetId) && seen.Add(c.DatasetId))
                        EnsureConnectionForDataset(c);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Failed to create connection:\n\n{ex.Message}",
                    "Connection Error");
                throw;
            }
        }
        internal static void EnsureCacheExists()
        {
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var workbook = app.ActiveWorkbook;

                // Check if cache sheet exists
                Microsoft.Office.Interop.Excel.Worksheet cacheSheet;
                try
                {
                    cacheSheet = workbook.Worksheets["__CubeConnector_Cache__"];
                }
                catch
                {
                    // Create cache sheet
                    cacheSheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.Add();
                    cacheSheet.Name = "__CubeConnector_Cache__";
                }

                // Check if cache table exists
                Microsoft.Office.Interop.Excel.ListObject cacheTable;
                try
                {
                    cacheTable = cacheSheet.ListObjects["CubeConnector_CacheTable"];
                    cacheSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;
                    return; // Cache exists
                }
                catch
                {
                    // Create cache table structure
                    cacheSheet.Range["A1"].Value2 = "CacheKey";
                    cacheSheet.Range["B1"].Value2 = "Result";
                    cacheSheet.Range["C1"].Value2 = "Timestamp";
                    cacheSheet.Range["D1"].Value2 = "FunctionSignature";

                    var headerRange = cacheSheet.Range["A1:D1"];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    cacheTable = cacheSheet.ListObjects.Add(
                        SourceType: Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange,
                        Source: cacheSheet.Range["A1:D1"],
                        XlListObjectHasHeaders: Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes
                    );

                    cacheTable.Name = "CubeConnector_CacheTable";
                    cacheTable.TableStyle = "TableStyleMedium2";

                    cacheSheet.Columns["A:D"].AutoFit();
                    cacheSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Failed to create cache:\n\n{ex.Message}", "Cache Error");
                throw;
            }
        }
    }

    /// <summary>
    /// Helper class to hold function registration information
    /// </summary>
    internal class ExcelFunctionRegistration
    {
        public Delegate FunctionDelegate { get; set; }
        public ExcelFunctionAttribute FunctionAttributes { get; set; }
        public List<object> ParameterAttributes { get; set; }
    }
}
