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
    public class DynamicFunctionRegistration : IExcelAddIn
    {
        [ExcelFunction(Name = "CC.Test", Description = "Test function", Category = "CubeConnector")]
        public static string TestFunction()
        {
            return "Test works!";
        }

        // Remove static constructor entirely!

        public void AutoOpen()  // NOT static anymore
        {
            try
            {
                var configs = ConfigurationStore.GetAllConfigs();
                if (configs == null || configs.Count == 0) return;

                RegisterFunctionsFromConfig(configs);
                AddContextMenuItems();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }

        public void AutoClose()
        {
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var cellMenu = app.CommandBars["Cell"];

                try { cellMenu.Controls["CubeConnector - Drill to Details"].Delete(); } catch { }
                try { cellMenu.Controls["CubeConnector - Drill to Pivot"].Delete(); } catch { }
                try { cellMenu.Controls["CubeConnector - Refresh Cache"].Delete(); } catch { }
            }
            catch { }
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
                else
                {
                    System.Windows.Forms.MessageBox.Show($"Registration was NULL for: {config.FunctionName}", "Debug");
                }
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
                    IsMacroType = true // Required to call Excel Application
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

            // Add "Drill to Pivot"
            var pivotButton = (Microsoft.Office.Core.CommandBarButton)
                cellMenu.Controls.Add(Type: Microsoft.Office.Core.MsoControlType.msoControlButton, Temporary: true);
            pivotButton.Caption = "CubeConnector - Drill to Pivot";
            pivotButton.OnAction = "DrillToPivotHandler";

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

                var manager = new DrillthroughManager(app, workbook);
                manager.ExecuteDrillthrough(activeCell);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Error");
            }
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
        public static void OnRefreshClicked(ExcelDna.Integration.CustomUI.IRibbonControl control)
        {
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var workbook = app.ActiveWorkbook;
                var manager = new RefreshManager(app, workbook);
                manager.RefreshAll();
                //System.Windows.Forms.MessageBox.Show(
                //    "CubeConnector cache refreshed successfully!",
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
        internal static void EnsureConnectionExists()
        {
            try
            {
                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var workbook = app.ActiveWorkbook;
                string connName = "CubeConnector";

                // Check if connection already exists
                try
                {
                    var existingConn = workbook.Connections[connName];
                    return; // Connection exists, we're done
                }
                catch
                {
                    // Connection doesn't exist, create it
                }

                // Get first config to extract connection details
                var configs = ConfigurationStore.GetAllConfigs();
                if (configs == null || configs.Count == 0)
                {
                    throw new Exception("No configuration found. Cannot create connection.");
                }

                var config = configs[0]; // Use first config for tenant/dataset

                // Build Power BI connection string (same format as your old ThisAddIn)
                string connectionString = $"OLEDB;Provider=MSOLAP.8;Integrated Security=ClaimsToken;Persist Security Info=True;" +
                    $"Initial Catalog={config.DatasetId};" +
                    $"Data Source=pbiazure://api.powerbi.com;" +
                    $"MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;" +
                    $"Identity Provider=https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, {config.TenantId};" +
                    $"Update Isolation Level=2";

                // Create the connection
                workbook.Connections.Add2(
                    Name: connName,
                    Description: "Auto-created by CubeConnector",
                    ConnectionString: connectionString,
                    CommandText: "Model",
                    lCmdtype: Microsoft.Office.Interop.Excel.XlCmdType.xlCmdDefault,
                    CreateModelConnection: Type.Missing,
                    ImportRelationships: Type.Missing
                );
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
