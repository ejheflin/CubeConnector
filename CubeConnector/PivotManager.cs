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
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CubeConnector
{
    /// <summary>
    /// Manages drill-to-pivot functionality - creates filtered pivot tables from aggregate values
    /// </summary>
    public class PivotManager
    {
        private Excel.Application xlApp;
        private Excel.Workbook workbook;

        // Store pending pivot configuration
        private static Excel.Application _staticXlApp;
        private static Excel.Worksheet _pendingSheet;
        private static string _pendingPivotName;
        private static RefreshItem _pendingItem;
        private static int _retryCount = 0;
        private const int MAX_RETRIES = 10;

        public PivotManager(Excel.Application application, Excel.Workbook workbook)
        {
            this.xlApp = application;
            this.workbook = workbook;
        }

        /// <summary>
        /// Test method - check if template pivot has fields available
        /// </summary>
        public void TestTemplatePivotFields()
        {
            try
            {
                Excel.Worksheet templateSheet = workbook.Worksheets["_PivotTemplate"];
                Excel.PivotTable templatePivot = templateSheet.PivotTables("TemplatePivot");

                string results = "Testing different field access methods:\n\n";

                // Find CubeFields (OLAP-specific)
                try
                {
                    int count = templatePivot.CubeFields.Count;
                    results += $"CubeFields.Count: {count}\n";

                    // Try to list some
                    string cubeFields = "";
                    int listed = 0;
                    foreach (Excel.CubeField field in templatePivot.CubeFields)
                    {
                        cubeFields += field.Name + "\n";
                        listed++;
                        if (listed >= 10) break;
                    }
                    results += $"\nFirst 10 CubeFields:\n{cubeFields}";
                }
                catch (Exception ex)
                {
                    results += $"CubeFields: ERROR - {ex.Message}\n";
                }

                MessageBox.Show(results, "Field Access Test", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Template test failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Execute drill-to-pivot from the active cell
        /// </summary>
        public void ExecuteDrillToPivot(Excel.Range cell)
        {
            try
            {
                // Validate cell has a formula
                if (!cell.HasFormula)
                {
                    MessageBox.Show("Please select a cell with a CubeConnector formula.", "No Formula", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string formula = cell.Formula.ToString();

                // Extract the actual function name from the formula
                string functionName = ExtractFunctionName(formula);
                if (string.IsNullOrEmpty(functionName))
                {
                    MessageBox.Show("Please select a cell with a CubeConnector formula.", "Invalid Formula", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Check if this function name matches one of our configured functions
                var allFunctionNames = ConfigurationStore.GetAllConfigs().Select(c => c.FunctionName).ToList();
                bool isCCFunction = allFunctionNames.Any(name => name.Equals(functionName, StringComparison.OrdinalIgnoreCase));

                if (!isCCFunction)
                {
                    MessageBox.Show("Please select a cell with a CubeConnector formula.", "Invalid Formula", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Parse the formula
                var item = ParseFormulaToRefreshItem(formula, cell);
                if (item == null)
                {
                    MessageBox.Show("Could not parse the formula.", "Parse Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Create pivot setup sheet
                CreatePivotSetupSheet(item);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Drill to Pivot failed:\n\n{ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Called when user clicks the "Configure Pivot" button
        /// </summary>
        public static void ConfigurePivotNow()
        {
            try
            {
                MessageBox.Show(
                    $"DEBUG: Button clicked, checking variables:\n" +
                    $"_pendingSheet: {(_pendingSheet == null ? "NULL" : _pendingSheet.Name)}\n" +
                    $"_pendingPivotName: {(_pendingPivotName ?? "NULL")}\n" +
                    $"_pendingItem: {(_pendingItem == null ? "NULL" : "EXISTS")}",
                    "Debug Retrieval",
                    MessageBoxButtons.OK);

                if (_pendingSheet == null || string.IsNullOrEmpty(_pendingPivotName) || _pendingItem == null)
                {
                    MessageBox.Show("No pending pivot configuration found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Look up pivot
                Excel.PivotTable pivotTable = _pendingSheet.PivotTables(_pendingPivotName);

                if (pivotTable == null)
                {
                    MessageBox.Show("Could not find pivot table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Check if fields are available
                int fieldCount = 0;
                try
                {
                    fieldCount = pivotTable.PivotFields().Count;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Pivot fields not ready yet:\n\n{ex.Message}", "Not Ready", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (fieldCount == 0)
                {
                    MessageBox.Show("Pivot has no fields yet. Please wait for the field list to load.", "Not Ready", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Configure the pivot
                var manager = new PivotManager(_staticXlApp, _pendingSheet.Parent as Excel.Workbook);
                manager.ConfigurePivotTableLayout(pivotTable);
                manager.AddParameterFieldsToRows(pivotTable, _pendingItem);
                manager.AddMeasureToValues(pivotTable, _pendingItem);

                try { manager.ApplyFiltersToRowFields(pivotTable, _pendingItem); } catch { }
                try { manager.SetTabularFormat(pivotTable); } catch { }

                // Delete the button now that we're done
                try
                {
                    foreach (Excel.Button btn in _pendingSheet.Buttons())
                    {
                        if (btn.Text.Contains("Configure Pivot"))
                        {
                            btn.Delete();
                            break;
                        }
                    }

                    // Clear instructions
                    _pendingSheet.Range["E3:E4"].ClearContents();
                }
                catch { }

                //MessageBox.Show("Pivot configured successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CleanupPendingPivot();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Configuration failed:\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Background task that waits for pivot fields to load, then configures (NOT USED - keeping for reference)
        /// </summary>
        private static void ConfigurePivotWhenReady()
        {
            try
            {
                // Wait 2 seconds before first check (let Excel initialize)
                System.Threading.Thread.Sleep(2000);

                for (int attempt = 1; attempt <= MAX_RETRIES; attempt++)
                {
                    try
                    {
                        // Must marshal back to UI thread for COM calls
                        bool success = false;
                        int fieldCount = 0;

                        _staticXlApp.Application.Wait(DateTime.Now.AddMilliseconds(100));

                        // Look up pivot
                        Excel.PivotTable pivotTable = _pendingSheet.PivotTables(_pendingPivotName);

                        if (pivotTable != null)
                        {
                            try
                            {
                                fieldCount = pivotTable.PivotFields().Count;

                                if (fieldCount > 0)
                                {
                                    // Fields loaded! Configure the pivot
                                    var manager = new PivotManager(_staticXlApp, _pendingSheet.Parent as Excel.Workbook);
                                    manager.ConfigurePivotTableLayout(pivotTable);
                                    manager.AddParameterFieldsToRows(pivotTable, _pendingItem);
                                    manager.AddMeasureToValues(pivotTable, _pendingItem);

                                    try { manager.ApplyFiltersToRowFields(pivotTable, _pendingItem); } catch { }
                                    try { manager.SetTabularFormat(pivotTable); } catch { }

                                    //MessageBox.Show("Pivot configured successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    success = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Attempt {attempt}: {ex.Message}");
                            }
                        }

                        if (success)
                        {
                            CleanupPendingPivot();
                            return;
                        }

                        System.Diagnostics.Debug.WriteLine($"Attempt {attempt}/{MAX_RETRIES}: {fieldCount} fields found");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Attempt {attempt} error: {ex.Message}");
                    }

                    // Wait before next attempt
                    if (attempt < MAX_RETRIES)
                    {
                        System.Threading.Thread.Sleep(2000);
                    }
                }

                // Timeout
                MessageBox.Show(
                    "Could not configure pivot automatically.\n\nPlease drag fields manually from the field list.",
                    "Timeout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                CleanupPendingPivot();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Background task failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CleanupPendingPivot();
            }
        }

        /// <summary>
        /// Public static callback method for Excel.OnTime (keeping for reference, not used)
        /// </summary>
        public static void ConfigurePivotCallback()
        {
            try
            {
                if (_pendingSheet == null || string.IsNullOrEmpty(_pendingPivotName) || _pendingItem == null)
                {
                    return;
                }

                _retryCount++;

                // Look up the pivot table fresh
                Excel.PivotTable pivotTable = null;
                try
                {
                    pivotTable = _pendingSheet.PivotTables(_pendingPivotName);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Attempt {_retryCount}: Could not find pivot: {ex.Message}");
                }

                if (pivotTable == null)
                {
                    if (_retryCount >= MAX_RETRIES)
                    {
                        MessageBox.Show("Could not find pivot table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        CleanupPendingPivot();
                    }
                    else
                    {
                        // Retry
                        DateTime nextTime = DateTime.Now.AddSeconds(2);
                        _staticXlApp.OnTime(nextTime, "ConfigurePivotCallback", Type.Missing, Type.Missing);
                    }
                    return;
                }

                // Check if fields are available
                int fieldCount = 0;
                try
                {
                    fieldCount = pivotTable.PivotFields().Count;
                    MessageBox.Show($"Attempt {_retryCount}: Found {fieldCount} fields!", "Debug", MessageBoxButtons.OK);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Attempt {_retryCount}: PivotFields error: {ex.Message}", "Debug", MessageBoxButtons.OK);
                }

                if (fieldCount == 0)
                {
                    if (_retryCount >= MAX_RETRIES)
                    {
                        MessageBox.Show(
                            "Pivot fields did not load in time.\n\n" +
                            "The pivot was created, but you'll need to manually configure it.",
                            "Timeout",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        CleanupPendingPivot();
                    }
                    else
                    {
                        // Retry
                        DateTime nextTime = DateTime.Now.AddSeconds(2);
                        _staticXlApp.OnTime(nextTime, "ConfigurePivotCallback", Type.Missing, Type.Missing);
                    }
                    return;
                }

                // Success! Configure the pivot
                MessageBox.Show($"Fields loaded! Configuring pivot with {fieldCount} fields...", "Success", MessageBoxButtons.OK);

                var manager = new PivotManager(_staticXlApp, _pendingSheet.Parent as Excel.Workbook);
                manager.ConfigurePivotTableLayout(pivotTable);
                manager.AddParameterFieldsToRows(pivotTable, _pendingItem);
                manager.AddMeasureToValues(pivotTable, _pendingItem);

                try
                {
                    manager.ApplyFiltersToRowFields(pivotTable, _pendingItem);
                }
                catch { }

                try
                {
                    manager.SetTabularFormat(pivotTable);
                }
                catch { }

                //MessageBox.Show("Pivot configured successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CleanupPendingPivot();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Callback failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CleanupPendingPivot();
            }
        }

        /// <summary>
        /// Clean up pending pivot state
        /// </summary>
        private static void CleanupPendingPivot()
        {
            _staticXlApp = null;
            _pendingSheet = null;
            _pendingPivotName = null;
            _pendingItem = null;
            _retryCount = 0;
        }

        /// <summary>
        /// Create a pivot table sheet connected to the Power BI model with pre-filtered parameters
        /// </summary>
        private void CreatePivotSetupSheet(RefreshItem item)
        {
            try
            {
                // Get the existing connection to extract connection details
                Excel.WorkbookConnection existingConn;
                try
                {
                    existingConn = workbook.Connections["CubeConnector"];
                }
                catch
                {
                    throw new Exception("Connection 'CubeConnector' not found.");
                }

                // Extract connection string
                string connectionString = existingConn.OLEDBConnection.Connection;

                // Generate unique names
                string baseName = $"Pivot - {item.Config.FunctionName}";
                string sheetName = GetUniqueSheetName(baseName);
                string pivotConnName = $"PivotConn_{Guid.NewGuid().ToString().Substring(0, 8)}";

                // Create new sheet
                Excel.Worksheet pivotSheet = workbook.Worksheets.Add();
                pivotSheet.Name = sheetName;

                // Create a NEW connection for the pivot in CUBE mode
                Excel.WorkbookConnection pivotConn = workbook.Connections.Add2(
                    Name: pivotConnName,
                    Description: "Pivot connection to Power BI cube",
                    ConnectionString: connectionString,
                    CommandText: "Model",
                    lCmdtype: Excel.XlCmdType.xlCmdCube,
                    CreateModelConnection: Type.Missing,
                    ImportRelationships: Type.Missing
                );

                // Create pivot cache from the new cube connection
                Excel.PivotCache pivotCache = workbook.PivotCaches().Create(
                    SourceType: Excel.XlPivotTableSourceType.xlExternal,
                    SourceData: pivotConn
                );

                // Create pivot table at A3
                Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: pivotSheet.Range["A3"],
                    TableName: "PivotResults"
                );

                // CRITICAL: Disable auto-refresh while we configure
                pivotTable.ManualUpdate = true;

                try
                {
                    // Configure the pivot using CubeFields (available immediately!)
                    int fieldCount = pivotTable.CubeFields.Count;

                    if (fieldCount > 0)
                    {
                        ConfigurePivotTableLayout(pivotTable);

                        // Add fields to rows AND apply filters BEFORE refresh
                        AddParameterFieldsWithFilters(pivotTable, item);

                        // Add the measure
                        AddMeasureToValues(pivotTable, item);

                        try { SetTabularFormat(pivotTable); } catch { }
                    }
                    else
                    {
                        MessageBox.Show(
                            "Pivot created but no CubeFields found.\n\nYou may need to configure it manually.",
                            "Warning",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                }
                finally
                {
                    // Re-enable refresh and refresh now (with filters applied)
                    pivotTable.ManualUpdate = false;
                }

                // Now refresh with filters applied
                try
                {
                    pivotTable.RefreshTable();
                }
                catch (Exception refreshEx)
                {
                    System.Diagnostics.Debug.WriteLine($"Refresh error: {refreshEx.Message}");
                }

                // After refresh, move page filters to rows (they keep their filters)
                //try
                //{
                //    MovePageFieldsToRows(pivotTable, item);
                //}
                //catch (Exception moveEx)
                //{
                //    System.Diagnostics.Debug.WriteLine($"Could not move fields to rows: {moveEx.Message}");
                //}

                pivotSheet.Activate();

                MessageBox.Show(
                    "Pivot table created and configured!",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create pivot: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Move page/filter fields to rows after data has been filtered and loaded
        /// </summary>
        //private void MovePageFieldsToRows(Excel.PivotTable pivotTable, RefreshItem item)
        //{
        //    var config = item.Config;

        //    for (int i = 0; i < config.Parameters.Count && i < item.Parameters.Length; i++)
        //    {
        //        var paramConfig = config.Parameters[i];
        //        var paramValue = item.Parameters[i];

        //        // Only move fields that have values (those we filtered)
        //        if (string.IsNullOrWhiteSpace(paramValue))
        //        {
        //            continue;
        //        }

        //        string hierarchyName = $"[{paramConfig.TableName}].[{paramConfig.FieldName}]";

        //        try
        //        {
        //            Excel.CubeField cubeField = pivotTable.CubeFields[hierarchyName];

        //            // Move from Page to Row (filter should be preserved)
        //            if (cubeField.Orientation == Excel.XlPivotFieldOrientation.xlPageField)
        //            {
        //                cubeField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            System.Diagnostics.Debug.WriteLine($"Could not move {hierarchyName} to rows: {ex.Message}");
        //        }
        //    }
        //}

        /// <summary>
        /// Apply filters to fields that are already in the rows area
        /// </summary>
        private void ApplyFiltersToRowFields(Excel.PivotTable pivotTable, RefreshItem item)
        {
            // This method is deprecated - use AddParameterFieldsWithFilters instead
            System.Diagnostics.Debug.WriteLine("ApplyFiltersToRowFields called - use AddParameterFieldsWithFilters instead");
        }

        /// <summary>
        /// Add parameter fields to rows AND apply filters BEFORE refresh
        /// This is critical for OLAP pivots - filters must be set before data loads
        /// </summary>
        private void AddParameterFieldsWithFilters(Excel.PivotTable pivotTable, RefreshItem item)
        {
            var config = item.Config;

            for (int i = 0; i < config.Parameters.Count && i < item.Parameters.Length; i++)
            {
                var paramConfig = config.Parameters[i];
                var paramValue = item.Parameters[i];

                // ONLY add fields that have values
                if (string.IsNullOrWhiteSpace(paramValue))
                {
                    continue;
                }

                string hierarchyName = $"[{paramConfig.TableName}].[{paramConfig.FieldName}]";

                try
                {
                    // Get the CubeField
                    Excel.CubeField cubeField = pivotTable.CubeFields[hierarchyName];

                    // Check if we have multiple comma-separated values
                    var values = paramValue.Split(',')
                        .Select(v => v.Trim())
                        .Where(v => !string.IsNullOrEmpty(v))
                        .ToList();

                    try
                    {
                        if (values.Count == 1)
                        {
                            // Single value - use page field with CurrentPageName (this works great!)
                            cubeField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                            string memberName = BuildMemberName(paramConfig, values[0]);
                            cubeField.CurrentPageName = memberName;
                            System.Diagnostics.Debug.WriteLine($"Filtered {hierarchyName} = {memberName}");
                        }
                        else if (values.Count > 1)
                        {
                            // Multiple values - add to rows first, then we'll filter after refresh
                            cubeField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                            // Store this info so we can filter it after the pivot loads
                            System.Diagnostics.Debug.WriteLine($"Added {hierarchyName} to rows with {values.Count} values (will filter after load)");
                        }
                    }
                    catch (Exception pageEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"Field setup failed for {hierarchyName}: {pageEx.Message}");

                        // Fall back to just adding to rows without filter
                        try
                        {
                            cubeField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                        }
                        catch (Exception rowEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"RowField also failed: {rowEx.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not get CubeField {hierarchyName}: {ex.Message}");
                }
            }
        }        /// <summary>
                 /// Apply filters to page fields (this filters the cube query before data is pulled)
                 /// </summary>
        private void ApplyFiltersToPageFields(Excel.PivotTable pivotTable, RefreshItem item)
        {
            var config = item.Config;

            for (int i = 0; i < config.Parameters.Count && i < item.Parameters.Length; i++)
            {
                var paramConfig = config.Parameters[i];
                var paramValue = item.Parameters[i];

                // ONLY apply filters for parameters with values
                if (string.IsNullOrWhiteSpace(paramValue))
                {
                    continue; // Skip empty parameters
                }

                try
                {
                    // Build the MDX field name for the hierarchy
                    string hierarchyName = $"[{paramConfig.TableName}].[{paramConfig.FieldName}]";

                    // Get the pivot field (works for both regular and OLAP pivots)
                    Excel.PivotField pivotField = pivotTable.PivotFields(hierarchyName);

                    // DON'T add to page field - we'll add to rows instead
                    // Just note that we want to filter this field
                    System.Diagnostics.Debug.WriteLine($"Will filter field {hierarchyName} with value: {paramValue}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not find field [{paramConfig.TableName}].[{paramConfig.FieldName}]: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Apply filter to a pivot field
        /// </summary>
        private void ApplyFilterToPivotField(Excel.PivotField pivotField, ParameterConfig paramConfig, string paramValue)
        {
            // Clear existing filters
            try
            {
                pivotField.ClearAllFilters();
            }
            catch { }

            // Handle different filter types
            switch (paramConfig.FilterType)
            {
                case FilterType.List:
                    // Split comma-separated values
                    var values = paramValue.Split(',').Select(v => v.Trim()).Where(v => !string.IsNullOrEmpty(v)).ToList();

                    if (values.Count == 1)
                    {
                        // Single value - set CurrentPage
                        string memberName = BuildMemberName(paramConfig, values[0]);
                        try
                        {
                            pivotField.CurrentPage = memberName;
                        }
                        catch
                        {
                            // Try without full member path
                            pivotField.CurrentPage = values[0];
                        }
                    }
                    else if (values.Count > 1)
                    {
                        // Multiple values - enable multiple page items
                        try
                        {
                            pivotField.EnableMultiplePageItems = true;
                        }
                        catch { }

                        // Hide all items first
                        try
                        {
                            foreach (Excel.PivotItem item in pivotField.PivotItems())
                            {
                                item.Visible = false;
                            }

                            // Show only the ones we want
                            foreach (var value in values)
                            {
                                string memberName = BuildMemberName(paramConfig, value);
                                try
                                {
                                    Excel.PivotItem pivotItem = pivotField.PivotItems(memberName);
                                    pivotItem.Visible = true;
                                }
                                catch
                                {
                                    // Try with just the value
                                    try
                                    {
                                        Excel.PivotItem pivotItem = pivotField.PivotItems(value);
                                        pivotItem.Visible = true;
                                    }
                                    catch
                                    {
                                        // Member doesn't exist, skip
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Could not set multiple values: {ex.Message}");
                        }
                    }
                    break;

                case FilterType.RangeStart:
                case FilterType.RangeEnd:
                    // For range filters with OLAP, just set as page field
                    // User can adjust the filter manually if needed
                    System.Diagnostics.Debug.WriteLine($"Range filters require manual adjustment for OLAP pivots: {paramConfig.Name}");
                    break;
            }
        }

        /// <summary>
        /// Build the full MDX member name for a value
        /// </summary>
        private string BuildMemberName(ParameterConfig paramConfig, string value)
        {
            // MDX member format: [Table].[Hierarchy].[Field].&[Value]
            // Or: [Table].[Field].&[Value]

            // Format the value based on data type
            string formattedValue = FormatMemberValue(value, paramConfig.DataType);

            // Build the member name
            return $"[{paramConfig.TableName}].[{paramConfig.FieldName}].&[{formattedValue}]";
        }

        /// <summary>
        /// Format a value for MDX member name
        /// </summary>
        private string FormatMemberValue(string value, string dataType)
        {
            switch (dataType.ToLower())
            {
                case "date":
                case "datetime":
                    // Check if it's a year value
                    if (IsYearValue(value))
                    {
                        return value; // Years can be used as-is in some models
                    }

                    // Try to parse as Excel date number
                    if (double.TryParse(value, out double excelDateNumber))
                    {
                        try
                        {
                            DateTime dt = DateTime.FromOADate(excelDateNumber);
                            return dt.ToString("yyyy-MM-dd");
                        }
                        catch { }
                    }

                    // Try to parse as date string
                    if (DateTime.TryParse(value, out DateTime parsedDt))
                    {
                        return parsedDt.ToString("yyyy-MM-dd");
                    }
                    break;

                case "number":
                case "integer":
                    return value;

                case "text":
                default:
                    return value;
            }

            return value;
        }

        /// <summary>
        /// Configure pivot table layout settings
        /// </summary>
        private void ConfigurePivotTableLayout(Excel.PivotTable pivotTable)
        {
            // General settings for OLAP pivot
            pivotTable.RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow);
            pivotTable.TableStyle2 = "PivotStyleMedium2";
            pivotTable.HasAutoFormat = false;
            pivotTable.DisplayFieldCaptions = true;
            pivotTable.ShowDrillIndicators = true;
        }

        /// <summary>
        /// Add parameter fields to the rows area (only for parameters with values)
        /// </summary>
        private void AddParameterFieldsToRows(Excel.PivotTable pivotTable, RefreshItem item)
        {
            var config = item.Config;
            var fieldsAdded = new List<string>();
            var fieldsSkipped = new List<string>();
            var fieldsFailed = new List<string>();

            // FIRST: List all available CubeFields (OLAP-specific)
            var availableFields = new List<string>();
            try
            {
                foreach (Excel.CubeField field in pivotTable.CubeFields)
                {
                    try
                    {
                        availableFields.Add(field.Name);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not enumerate cube fields: {ex.Message}",
                    "Field Enumeration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // NOW try to add our fields
            for (int i = 0; i < config.Parameters.Count && i < item.Parameters.Length; i++)
            {
                var paramConfig = config.Parameters[i];
                var paramValue = item.Parameters[i];

                // ONLY add to rows if user provided a value for this parameter
                if (string.IsNullOrWhiteSpace(paramValue))
                {
                    fieldsSkipped.Add($"{paramConfig.Name} (empty)");
                    continue; // Skip parameters without values
                }

                // Try to find a matching field
                string hierarchyName = $"[{paramConfig.TableName}].[{paramConfig.FieldName}]";

                // Look for exact match or partial match
                var matchingField = availableFields.FirstOrDefault(f =>
                    f == hierarchyName ||
                    f.Contains(paramConfig.FieldName) ||
                    f.Contains(paramConfig.TableName));

                if (matchingField != null)
                {
                    try
                    {
                        // For OLAP pivots, use CubeFields to add to rows
                        Excel.CubeField cubeField = pivotTable.CubeFields[matchingField];
                        cubeField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                        fieldsAdded.Add($"{matchingField} = {paramValue}");
                    }
                    catch (Exception ex)
                    {
                        fieldsFailed.Add($"{matchingField}: {ex.Message}");
                    }
                }
                else
                {
                    fieldsFailed.Add($"No match found for {paramConfig.TableName}.{paramConfig.FieldName}");
                }
            }

            // Show diagnostic message
            string message = "Field Addition Results:\n\n";

            if (fieldsAdded.Any())
            {
                message += "✓ Added:\n" + string.Join("\n", fieldsAdded) + "\n\n";
            }

            if (fieldsSkipped.Any())
            {
                message += "○ Skipped:\n" + string.Join("\n", fieldsSkipped) + "\n\n";
            }

            if (fieldsFailed.Any())
            {
                message += "✗ Failed:\n" + string.Join("\n", fieldsFailed);
            }

            MessageBox.Show(message, "Add Fields Diagnostic", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Try alternate field name formats if the standard format doesn't work
        /// </summary>
        private bool TryAlternateFieldFormats(Excel.PivotTable pivotTable, ParameterConfig paramConfig, string paramValue, List<string> fieldsAdded)
        {
            // Try different naming conventions
            string[] alternateFormats = new[]
            {
                $"{paramConfig.TableName}.{paramConfig.FieldName}",           // No brackets
                $"[{paramConfig.TableName}].[{paramConfig.FieldName}].[{paramConfig.FieldName}]", // Hierarchy format
                $"{paramConfig.FieldName}",                                   // Just field name
                $"[{paramConfig.FieldName}]"                                  // Just field name with brackets
            };

            foreach (var format in alternateFormats)
            {
                try
                {
                    Excel.PivotField field = pivotTable.PivotFields(format);
                    if (field.Orientation == Excel.XlPivotFieldOrientation.xlHidden)
                    {
                        field.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                        fieldsAdded.Add($"{format} (alternate) = {paramValue}");
                        System.Diagnostics.Debug.WriteLine($"Successfully added {format} to rows");
                        return true; // Success!
                    }
                }
                catch
                {
                    // This format didn't work, try next
                }
            }

            return false; // None of the formats worked
        }

        /// <summary>
        /// Add the measure to the values area
        /// </summary>
        private void AddMeasureToValues(Excel.PivotTable pivotTable, RefreshItem item)
        {
            var config = item.Config;

            // The measure name from config is in format [MeasureName]
            string measureName = config.MeasureName.Trim('[', ']');

            // Try different MDX measure name formats
            string[] possibleNames = new[]
            {
                $"[Measures].[{measureName}]",  // Standard OLAP format
                config.MeasureName,              // As-is from config
                $"[{measureName}]"               // Simple bracket format
            };

            foreach (var name in possibleNames)
            {
                try
                {
                    // For OLAP pivots, use CubeFields
                    Excel.CubeField measureField = pivotTable.CubeFields[name];
                    measureField.Orientation = Excel.XlPivotFieldOrientation.xlDataField;

                    System.Diagnostics.Debug.WriteLine($"Successfully added measure: {name}");
                    return; // Success!
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not add measure {name}: {ex.Message}");
                }
            }

            // If we get here, none of the formats worked
            System.Diagnostics.Debug.WriteLine($"Failed to add measure with any format");
        }

        /// <summary>
        /// Set the pivot table to tabular format
        /// </summary>
        private void SetTabularFormat(Excel.PivotTable pivotTable)
        {
            try
            {
                // Set all row fields to tabular layout
                foreach (Excel.PivotField field in pivotTable.PivotFields())
                {
                    if (field.Orientation == Excel.XlPivotFieldOrientation.xlRowField)
                    {
                        try
                        {
                            field.LayoutForm = Excel.XlLayoutFormType.xlTabular;
                            field.Subtotals = new bool[12]; // Turn off all subtotals
                        }
                        catch { }
                    }
                }

                // Turn off grand totals
                pivotTable.ColumnGrand = false;
                pivotTable.RowGrand = true; // Keep row grand total
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not set tabular format: {ex.Message}");
            }
        }


        /// <summary>
        /// Wait for the pivot table to load its field list from the cube
        /// </summary>
        private bool WaitForPivotToLoad(Excel.PivotTable pivotTable)
        {
            int maxAttempts = 20; // 10 seconds total
            int attemptDelay = 500; // 500ms between attempts

            for (int i = 0; i < maxAttempts; i++)
            {
                try
                {
                    // Try to count the fields
                    int fieldCount = pivotTable.PivotFields().Count;

                    if (fieldCount > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"Pivot loaded with {fieldCount} fields after {i * attemptDelay}ms");
                        return true; // Success!
                    }
                }
                catch
                {
                    // PivotFields not ready yet
                }

                // Wait and try again
                System.Threading.Thread.Sleep(attemptDelay);

                // Process Windows messages to allow Excel to update
                System.Windows.Forms.Application.DoEvents();
            }

            return false; // Timeout
        }

        /// <summary>
        /// Create a visible template pivot that's already connected and loaded
        /// Call this during refresh when we have time
        /// </summary>
        public static void EnsurePivotTemplateExists(Excel.Application xlApp, Excel.Workbook workbook)
        {
            try
            {
                // Check if template already exists
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == "_PivotTemplate")
                    {
                        return; // Already exists
                    }
                }

                // Get the connection
                Excel.WorkbookConnection existingConn;
                try
                {
                    existingConn = workbook.Connections["CubeConnector"];
                }
                catch
                {
                    return; // No connection yet
                }

                // Create VISIBLE template sheet (so Excel loads the fields)
                Excel.Worksheet templateSheet = workbook.Worksheets.Add();
                templateSheet.Name = "_PivotTemplate";
                templateSheet.Tab.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                // Add note at top
                templateSheet.Range["A1"].Value2 = "Pivot Template - Do Not Delete";
                templateSheet.Range["A1"].Font.Bold = true;
                templateSheet.Range["A1"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                templateSheet.Range["A2"].Value2 = "This sheet is used by the Drill-to-Pivot feature.";

                // Extract connection string
                string connectionString = existingConn.OLEDBConnection.Connection;
                string templateConnName = "PivotTemplate_Conn";

                // Create pivot connection
                Excel.WorkbookConnection pivotConn = workbook.Connections.Add2(
                    Name: templateConnName,
                    Description: "Template pivot connection",
                    ConnectionString: connectionString,
                    CommandText: "Model",
                    lCmdtype: Excel.XlCmdType.xlCmdCube,
                    CreateModelConnection: Type.Missing,
                    ImportRelationships: Type.Missing
                );

                // Create pivot cache
                Excel.PivotCache pivotCache = workbook.PivotCaches().Create(
                    SourceType: Excel.XlPivotTableSourceType.xlExternal,
                    SourceData: pivotConn
                );

                // Create template pivot at A4 (below the note)
                Excel.PivotTable templatePivot = pivotCache.CreatePivotTable(
                    TableDestination: templateSheet.Range["A4"],
                    TableName: "TemplatePivot"
                );

                // Try to force Excel to load the pivot by selecting it
                try
                {
                    templateSheet.Activate();
                    templateSheet.Range["A4"].Select(); // Select the pivot
                    System.Threading.Thread.Sleep(2000); // Wait 2 seconds
                    templatePivot.RefreshTable();
                    System.Threading.Thread.Sleep(2000); // Wait 2 more seconds
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error activating template: {ex.Message}");
                }

                System.Diagnostics.Debug.WriteLine("Pivot template created successfully (visible sheet)");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not create pivot template: {ex.Message}");
            }
        }

        /// <summary>
        /// Get a unique sheet name, avoiding conflicts
        /// </summary>
        private string GetUniqueSheetName(string baseName)
        {
            // Excel sheet names limited to 31 characters
            if (baseName.Length > 31)
            {
                baseName = baseName.Substring(0, 31);
            }

            string sheetName = baseName;
            int counter = 1;

            while (SheetNameExists(sheetName))
            {
                string suffix = $" ({counter})";
                int maxBaseLength = 31 - suffix.Length;

                if (baseName.Length > maxBaseLength)
                {
                    sheetName = baseName.Substring(0, maxBaseLength) + suffix;
                }
                else
                {
                    sheetName = baseName + suffix;
                }

                counter++;

                if (counter > 100)
                {
                    sheetName = $"Pivot_{Guid.NewGuid().ToString().Substring(0, 8)}";
                    break;
                }
            }

            return sheetName;
        }

        /// <summary>
        /// Check if a sheet name already exists
        /// </summary>
        private bool SheetNameExists(string name)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Check if value looks like a year
        /// </summary>
        private bool IsYearValue(string value)
        {
            if (int.TryParse(value, out int year))
            {
                return year >= 1900 && year <= 2150;
            }
            return false;
        }

        /// <summary>
        /// Parse a formula into a RefreshItem
        /// </summary>
        private RefreshItem ParseFormulaToRefreshItem(string formula, Excel.Range cell)
        {
            try
            {
                string functionName = ExtractFunctionName(formula);
                if (string.IsNullOrEmpty(functionName))
                {
                    return null;
                }

                var config = ConfigurationStore.GetConfig(functionName);
                if (config == null)
                {
                    return null;
                }

                var parameters = ExtractParametersFromCell(cell);
                string cacheKey = CacheKey.BuildFromStrings(functionName, parameters);

                return new RefreshItem
                {
                    CacheKey = cacheKey,
                    Config = config,
                    Parameters = parameters,
                    FunctionSignature = formula,
                    Cell = cell
                };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Extract function name from formula
        /// </summary>
        private string ExtractFunctionName(string formula)
        {
            if (!formula.StartsWith("=")) return null;

            int parenIndex = formula.IndexOf('(');
            if (parenIndex < 0) return null;

            return formula.Substring(1, parenIndex - 1).Trim();
        }

        /// <summary>
        /// Extract parameters from cell formula
        /// </summary>
        private string[] ExtractParametersFromCell(Excel.Range cell)
        {
            var parameters = new List<string>();

            try
            {
                string formula = cell.Formula.ToString();
                int startIdx = formula.IndexOf('(');
                int endIdx = formula.LastIndexOf(')');

                if (startIdx < 0 || endIdx < 0 || endIdx <= startIdx)
                {
                    return new string[0];
                }

                string argsString = formula.Substring(startIdx + 1, endIdx - startIdx - 1);
                var argTokens = SplitFormulaArguments(argsString);

                foreach (var token in argTokens)
                {
                    string cleanToken = token.Trim();

                    if (string.IsNullOrEmpty(cleanToken))
                    {
                        parameters.Add("");
                        continue;
                    }

                    if ((cleanToken.StartsWith("\"") && cleanToken.EndsWith("\"")) ||
                        (cleanToken.StartsWith("'") && cleanToken.EndsWith("'")))
                    {
                        parameters.Add(cleanToken.Trim('"', '\''));
                    }
                    else if (double.TryParse(cleanToken, out double numValue))
                    {
                        parameters.Add(cleanToken);
                    }
                    else
                    {
                        try
                        {
                            Excel.Range refRange = cell.Worksheet.Range[cleanToken];
                            var value = refRange.Value2;

                            if (refRange.Cells.Count > 1)
                            {
                                var values = new List<string>();
                                foreach (Excel.Range c in refRange.Cells)
                                {
                                    var v = c.Value2;
                                    if (v != null) values.Add(v.ToString());
                                }
                                parameters.Add(string.Join(",", values));
                            }
                            else
                            {
                                parameters.Add(value?.ToString() ?? "");
                            }
                        }
                        catch
                        {
                            parameters.Add(cleanToken);
                        }
                    }
                }
            }
            catch { }

            while (parameters.Count < 3)
            {
                parameters.Add("");
            }

            return parameters.ToArray();
        }

        /// <summary>
        /// Split formula arguments by comma
        /// </summary>
        private List<string> SplitFormulaArguments(string argsString)
        {
            var arguments = new List<string>();
            var currentArg = new System.Text.StringBuilder();
            int parenDepth = 0;
            bool inQuotes = false;
            char quoteChar = '\0';

            for (int i = 0; i < argsString.Length; i++)
            {
                char c = argsString[i];

                if ((c == '"' || c == '\'') && (i == 0 || argsString[i - 1] != '\\'))
                {
                    if (!inQuotes)
                    {
                        inQuotes = true;
                        quoteChar = c;
                    }
                    else if (c == quoteChar)
                    {
                        inQuotes = false;
                    }
                    currentArg.Append(c);
                }
                else if (c == '(' && !inQuotes)
                {
                    parenDepth++;
                    currentArg.Append(c);
                }
                else if (c == ')' && !inQuotes)
                {
                    parenDepth--;
                    currentArg.Append(c);
                }
                else if (c == ',' && !inQuotes && parenDepth == 0)
                {
                    arguments.Add(currentArg.ToString());
                    currentArg.Clear();
                }
                else
                {
                    currentArg.Append(c);
                }
            }

            if (currentArg.Length > 0)
            {
                arguments.Add(currentArg.ToString());
            }

            return arguments;
        }
        /// <summary>
        /// Filter row fields after the pivot has loaded (for multiple values)
        /// </summary>
        private void FilterRowFieldsAfterLoad(Excel.PivotTable pivotTable, RefreshItem item)
        {
            var config = item.Config;

            for (int i = 0; i < config.Parameters.Count && i < item.Parameters.Length; i++)
            {
                var paramConfig = config.Parameters[i];
                var paramValue = item.Parameters[i];

                if (string.IsNullOrWhiteSpace(paramValue))
                {
                    continue;
                }

                // Check if we have multiple comma-separated values
                var values = paramValue.Split(',')
                    .Select(v => v.Trim())
                    .Where(v => !string.IsNullOrEmpty(v))
                    .ToList();

                // Only filter if we have multiple values
                if (values.Count <= 1)
                {
                    continue;
                }

                string hierarchyName = $"[{paramConfig.TableName}].[{paramConfig.FieldName}]";

                try
                {
                    // Get the PivotField (after refresh, we can access it this way)
                    Excel.PivotField pivotField = pivotTable.PivotFields(hierarchyName);

                    // Hide all items first
                    foreach (Excel.PivotItem pivotItem in pivotField.PivotItems())
                    {
                        pivotItem.Visible = false;
                    }

                    // Show only the items we want
                    foreach (var value in values)
                    {
                        string memberName = BuildMemberName(paramConfig, value);
                        try
                        {
                            Excel.PivotItem pivotItem = pivotField.PivotItems(memberName);
                            pivotItem.Visible = true;
                        }
                        catch
                        {
                            // Try without full member path
                            try
                            {
                                Excel.PivotItem pivotItem = pivotField.PivotItems(value);
                                pivotItem.Visible = true;
                            }
                            catch (Exception itemEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"Could not show item {value}: {itemEx.Message}");
                            }
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"Filtered {hierarchyName} to {values.Count} items");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not filter {hierarchyName}: {ex.Message}");
                }
            }
        }
    }
}