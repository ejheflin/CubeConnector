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
    /// Manages drillthrough functionality - shows detail rows behind aggregate values
    /// </summary>
    public class DrillthroughManager
    {
        private Excel.Application xlApp;
        private Excel.Workbook workbook;
        private const int MAX_ROWS = 1000;

        public DrillthroughManager(Excel.Application application, Excel.Workbook workbook)
        {
            this.xlApp = application;
            this.workbook = workbook;
        }

        /// <summary>
        /// Execute drillthrough from the active cell
        /// </summary>
        public void ExecuteDrillthrough(Excel.Range cell)
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

                // Build drillthrough query
                string daxQuery = BuildDrillthroughQuery(item);

                //MessageBox.Show($"Drillthrough query:\n\n{daxQuery}\n\nExecuting...", "Drillthrough", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Execute query and create sheet
                var drillthroughSheet = ExecuteDrillthroughQuery(daxQuery, item);

                if (drillthroughSheet == null)
                {
                    MessageBox.Show("Query returned no results.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Check if results were truncated by seeing if we got MAX_ROWS
                var queryTable = drillthroughSheet.ListObjects["DrillthroughResults"];
                int rowCount = queryTable.DataBodyRange.Rows.Count;
                bool wasTruncated = rowCount >= MAX_ROWS;

                // Format the sheet
                FormatDrillthroughSheet(item, drillthroughSheet, wasTruncated);

                // Show truncation warning
                if (wasTruncated)
                {
                    MessageBox.Show(
                        $"Results limited to {MAX_ROWS:N0} rows.\n\nThe actual query may have returned more data.",
                        "Results Truncated",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Drillthrough failed:\n\n{ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Build a DETAILROWS query with the same filters as the aggregate
        /// </summary>
        private string BuildDrillthroughQuery(RefreshItem item)
        {
            var config = item.Config;
            var filters = new List<string>();

            // Build filters from parameters
            for (int i = 0; i < config.Parameters.Count && i < item.Parameters.Length; i++)
            {
                var paramConfig = config.Parameters[i];
                var paramValue = item.Parameters[i];

                if (string.IsNullOrWhiteSpace(paramValue))
                {
                    continue;
                }

                string filter = BuildDrillthroughFilter(paramConfig, paramValue);
                if (!string.IsNullOrEmpty(filter))
                {
                    filters.Add(filter);
                }
            }

            // Build CALCULATETABLE with DETAILROWS
            string calculateTable;
            if (filters.Count > 0)
            {
                calculateTable = $"CALCULATETABLE(DETAILROWS({config.MeasureName}),{string.Join(",", filters)})";
            }
            else
            {
                calculateTable = $"DETAILROWS({config.MeasureName})";
            }

            // Wrap with TOPN to limit rows
            string query = $"EVALUATE TOPN({MAX_ROWS},{calculateTable})";

            return query;
        }

        /// <summary>
        /// Execute the drillthrough query and return the sheet with results
        /// </summary>
        private Excel.Worksheet ExecuteDrillthroughQuery(string daxQuery, RefreshItem item)
        {
            try
            {
                // Get the Power BI connection
                Excel.WorkbookConnection conn;
                try
                {
                    conn = workbook.Connections["CubeConnector"];
                }
                catch
                {
                    throw new Exception("Connection 'CubeConnector' not found.");
                }

                // Generate unique sheet name
                string baseName = $"Drillthrough - {item.Config.FunctionName}";
                string sheetName = GetUniqueSheetName(baseName);

                // Create the drillthrough sheet (not temp - this is the final sheet)
                Excel.Worksheet drillthroughSheet = workbook.Worksheets.Add();
                drillthroughSheet.Name = sheetName;

                // Create ListObject with the connection - this is the query table
                Excel.ListObject queryTable = drillthroughSheet.ListObjects.Add(
                    SourceType: Excel.XlListObjectSourceType.xlSrcExternal,
                    Source: conn,
                    LinkSource: true,
                    XlListObjectHasHeaders: Excel.XlYesNoGuess.xlGuess,
                    Destination: drillthroughSheet.Range["A3"]  // Start at A3 to leave room for header
                );

                queryTable.Name = "DrillthroughResults";

                // Set up the query
                Excel.QueryTable qt = queryTable.QueryTable;
                qt.CommandType = conn.OLEDBConnection.CommandType;
                qt.CommandText = daxQuery;

                // Execute
                qt.Refresh(BackgroundQuery: false);

                // Check if we got results
                if (qt.ResultRange == null)
                {
                    drillthroughSheet.Delete();
                    return null;
                }

                // Apply table styling
                queryTable.TableStyle = "TableStyleMedium2";

                // Auto-fit columns
                qt.ResultRange.Columns.AutoFit();

                return drillthroughSheet;
            }
            catch (Exception ex)
            {
                throw new Exception($"Query execution failed: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Format the drillthrough sheet with header and warnings
        /// </summary>
        private void FormatDrillthroughSheet(RefreshItem item, Excel.Worksheet drillthroughSheet, bool wasTruncated)
        {
            try
            {
                // Add a header with context at row 1
                Excel.Range headerCell = drillthroughSheet.Range["A1"];
                headerCell.Value2 = $"Drillthrough: {item.FunctionSignature}";
                headerCell.Font.Bold = true;
                headerCell.Font.Size = 12;

                if (wasTruncated)
                {
                    Excel.Range warningCell = drillthroughSheet.Range["A2"];
                    warningCell.Value2 = $"⚠ Results limited to {MAX_ROWS:N0} rows";
                    warningCell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    warningCell.Font.Italic = true;
                }

                // Activate the sheet
                drillthroughSheet.Activate();

                //MessageBox.Show($"Drillthrough complete!\n\nSheet created: {drillthroughSheet.Name}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to format drillthrough sheet: {ex.Message}", ex);
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

            // Check if name exists
            while (SheetNameExists(sheetName))
            {
                // Add counter suffix
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

                // Safety: don't loop forever
                if (counter > 100)
                {
                    sheetName = $"Drillthrough_{Guid.NewGuid().ToString().Substring(0, 8)}";
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
        /// Parse a formula into a RefreshItem
        /// </summary>
        private RefreshItem ParseFormulaToRefreshItem(string formula, Excel.Range cell)
        {
            try
            {
                // Extract function name
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

                // Extract parameters
                var parameters = ExtractParametersFromCell(cell);

                // Build cache key
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
            catch
            {
                // Ignore errors
            }

            // Pad to 3 parameters
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
        /// Build a filter expression for drillthrough (simple version)
        /// </summary>
        private string BuildDrillthroughFilter(ParameterConfig paramConfig, string paramValue)
        {
            string table = paramConfig.TableName;
            string field = paramConfig.FieldName;
            string fullField = $"{table}[{field}]";

            // Handle different filter types
            switch (paramConfig.FilterType)
            {
                case FilterType.List:
                    var values = paramValue.Split(',').Select(v => v.Trim()).Where(v => !string.IsNullOrEmpty(v));
                    if (!values.Any()) return "";

                    var formattedValues = values.Select(v => FormatDaxValue(v, paramConfig.DataType));

                    if (values.Count() == 1)
                        return $"{fullField}={formattedValues.First()}";
                    else
                        return $"{fullField} IN {{{string.Join(",", formattedValues)}}}";

                case FilterType.RangeStart:
                    // Check for year value
                    if ((paramConfig.DataType == "date" || paramConfig.DataType == "datetime") && IsYearValue(paramValue))
                    {
                        return $"{fullField}>=DATE({paramValue},1,1)";
                    }
                    return $"{fullField}>={FormatDaxValue(paramValue, paramConfig.DataType)}";

                case FilterType.RangeEnd:
                    // Check for year value
                    if ((paramConfig.DataType == "date" || paramConfig.DataType == "datetime") && IsYearValue(paramValue))
                    {
                        return $"{fullField}<=DATE({paramValue},12,31)";
                    }
                    return $"{fullField}<={FormatDaxValue(paramValue, paramConfig.DataType)}";

                default:
                    return "";
            }
        }

        /// <summary>
        /// Format a value for DAX
        /// </summary>
        private string FormatDaxValue(string value, string dataType)
        {
            switch (dataType.ToLower())
            {
                case "text":
                    return $"\"{value.Replace("\"", "\"\"")}\"";

                case "number":
                case "integer":
                    return value.Trim();

                case "date":
                case "datetime":
                    // Try to parse as Excel date number
                    if (double.TryParse(value, out double excelDateNumber))
                    {
                        try
                        {
                            DateTime dt = DateTime.FromOADate(excelDateNumber);
                            return $"DATE({dt.Year},{dt.Month},{dt.Day})";
                        }
                        catch
                        {
                            // Fall through to regular date parsing
                        }
                    }
                    // Try to parse as date string
                    if (DateTime.TryParse(value, out DateTime parsedDt))
                    {
                        return $"DATE({parsedDt.Year},{parsedDt.Month},{parsedDt.Day})";
                    }
                    return value;

                default:
                    return $"\"{value}\"";
            }
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
    }
}