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
    /// Manages refreshing UDF cells by querying Power BI
    /// </summary>
    public class RefreshManager
    {
        private Excel.Application xlApp;
        private Excel.Workbook workbook;

        // Maximum DAX query length (Excel QueryTable.CommandText limit)
        private const int MAX_QUERY_LENGTH = 30000;

        // Minimum queries needed to create a pool (otherwise treat as orphans)
        private const int MIN_POOL_SIZE = 3;

        public RefreshManager(Excel.Application application, Excel.Workbook workbook)
        {
            this.xlApp = application;
            this.workbook = workbook;
        }

        public void RefreshAll()
        {
            RefreshInternal(null, null);
        }

        public void ClearCacheAndRefresh()
        {
            // Step 1: Delete the cache to force all cells to refresh
            DeleteCacheWorksheet();

            // Step 2: Recreate the cache (empty) so we can store results later
            DynamicFunctionRegistration.EnsureCacheExists();

            // Step 3: Trigger recalc so all UDF cells show #REFRESH
            var app = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            app.CalculateFullRebuild();

            // Step 4: Call normal refresh to refill the cache
            RefreshInternal(null, null);
        }

        public void RefreshSheet(Excel.Worksheet sheet)
        {
            if (sheet == null)
            {
                MessageBox.Show("Invalid sheet.", "Error");
                return;
            }

            RefreshInternal(sheet, null);
        }

        public void RefreshRange(Excel.Range range)
        {
            if (range == null)
            {
                MessageBox.Show("Invalid range.", "Error");
                return;
            }

            RefreshInternal(null, range);
        }

        private void RefreshInternal(Excel.Worksheet targetSheet, Excel.Range targetRange)
        {
            // Store original calculation mode
            Excel.XlCalculation originalCalcMode = xlApp.Calculation;

            try
            {
                // Disable automatic calculation during refresh
                xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                // Step 1: Find all cells with #REFRESH or #N/A
                var cellsToRefresh = FindCellsNeedingRefresh(targetSheet, targetRange);

                if (cellsToRefresh.Count == 0)
                {
                    string scope = targetRange != null ? "in selection" :
                                   targetSheet != null ? $"on sheet '{targetSheet.Name}'" :
                                   "in workbook";
                    MessageBox.Show($"No cells need refreshing {scope}. All formulas are up to date!");
                    return;
                }

                //MessageBox.Show($"Found {cellsToRefresh.Count} cells needing refresh.\n\nAnalyzing query patterns...");

                // Step 2: Analyze and create pools
                var analysis = QueryPoolAnalyzer.AnalyzeQueries(cellsToRefresh, MIN_POOL_SIZE);

                int pooledQueries = analysis.Pools.Sum(p => p.Items.Count);
                int orphanQueries = analysis.Orphans.Count;

                //MessageBox.Show($"Pool Analysis:\n\n" +
                //    $"Pools created: {analysis.Pools.Count}\n" +
                //    $"Queries in pools: {pooledQueries}\n" +
                //    $"Orphan queries: {orphanQueries}\n\n" +
                //    $"Building DAX queries...");

                // Step 3: Build pooled queries and track which pools are in each query
                var queryBatches = BuildPooledQueriesWithTracking(analysis.Pools, MAX_QUERY_LENGTH);

                // DIAGNOSTIC: Show pool details
                var poolDetails = string.Join("\n\n", analysis.Pools.Select((p, i) =>
                    $"Pool {i + 1}:\n" +
                    $"  Varying: Param{p.VaryingParamIndex}\n" +
                    $"  Fixed: {string.Join(", ", p.FixedParameters.Select(kvp => $"Param{kvp.Key}='{kvp.Value}'"))}\n" +
                    $"  Values: {p.VaryingValues.Count} items\n" +
                    $"  Queries: {p.Items.Count} formulas"
                ));

                //MessageBox.Show($"Pool Details:\n\n{poolDetails}");

                //MessageBox.Show($"Built {queryBatches.Count} pooled DAX queries.\n\n" +
                //    $"Query lengths:\n{string.Join("\n", queryBatches.Select((q, i) => $"Query {i + 1}: {q.DaxQuery.Length:N0} chars, {q.Pools.Count} pools"))}");

                // Step 4: Copy first query to clipboard for testing (commented out)
                //if (queryBatches.Count > 0)
                //{
                //    string firstQuery = queryBatches[0].DaxQuery;
                //    Clipboard.SetText(firstQuery);
                //    MessageBox.Show($"First pooled query copied to clipboard!\n\n" +
                //        $"Length: {firstQuery.Length:N0} characters\n\n" +
                //        $"You can now paste this into DAX Studio to test.\n\n" +
                //        $"Preview (first 500 chars):\n{firstQuery.Substring(0, Math.Min(500, firstQuery.Length))}...");
                //}

                // Step 5: Execute pooled queries
                int totalProcessed = 0;
                int totalFailed = 0;

                for (int i = 0; i < queryBatches.Count; i++)
                {
                    var batch = queryBatches[i];

                    try
                    {
                        //MessageBox.Show($"Executing pooled query {i + 1}/{queryBatches.Count}...");

                        var results = ExecuteBatchQuery(batch.DaxQuery);

                        //MessageBox.Show($"Query returned {results.Count} results.\n\n" +
                        //    $"Sample keys:\n{string.Join("\n", results.Keys.Take(5))}");

                        // Get items from the pools in this batch
                        var itemsInThisBatch = batch.Pools.SelectMany(p => p.Items).ToList();

                        //MessageBox.Show($"Batch contains {batch.Pools.Count} pools:\n\n" +
                        //    string.Join("\n", batch.Pools.Select((p, idx) =>
                        //        $"Pool {idx + 1}: {p.Items.Count} items, " +
                        //        $"Sample key: {p.Items.FirstOrDefault()?.CacheKey ?? "N/A"}"
                        //    )));

                        var itemsToStore = PrepareResultsForStorage(results, itemsInThisBatch);

                        //MessageBox.Show($"Prepared {itemsToStore.Count} items for storage.\n\n" +
                        //    $"Sample:\n{string.Join("\n", itemsToStore.Take(3).Select(kvp => $"{kvp.Key} = {kvp.Value.result}"))}");

                        CacheManager.StoreBatch(workbook, itemsToStore);
                        totalProcessed += itemsToStore.Count;

                        //MessageBox.Show($"Pooled query {i + 1}/{queryBatches.Count} complete!\n" +
                        //    $"Stored {itemsToStore.Count} results");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Pooled query {i + 1}/{queryBatches.Count} FAILED:\n\n{ex.Message}\n\n{ex.StackTrace}");
                        totalFailed++;
                    }
                }

                // Step 6: Handle orphan queries with traditional batching
                if (analysis.Orphans.Count > 0)
                {
                    //MessageBox.Show($"Processing {analysis.Orphans.Count} orphan queries using traditional batching...");

                    var orphansByFunction = analysis.Orphans.GroupBy(o => o.Config.FunctionName);

                    foreach (var funcGroup in orphansByFunction)
                    {
                        var config = funcGroup.First().Config;
                        var batchItems = funcGroup.Select(item => new DAXQueryBuilder.BatchQueryItem
                        {
                            CacheKey = item.CacheKey,
                            Parameters = item.Parameters
                        }).ToList();

                        try
                        {
                            string orphanQuery = DAXQueryBuilder.BuildBatchQuery(config, batchItems);

                            //MessageBox.Show($"Orphan query for {funcGroup.Key}:\n\n{orphanQuery.Substring(0, Math.Min(1000, orphanQuery.Length))}...", "Debug Query");
                            Clipboard.SetText(orphanQuery);  // Copy to clipboard so you can test in DAX Studio


                            var results = ExecuteBatchQuery(orphanQuery);
                            var itemsToStore = PrepareResultsForStorage(results, analysis.Orphans);
                            CacheManager.StoreBatch(workbook, itemsToStore);
                            totalProcessed += itemsToStore.Count;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Orphan batch failed for {funcGroup.Key}:\n\n{ex.Message}");
                            totalFailed += funcGroup.Count();
                        }
                    }
                }

                // Step 7: Recalculate
                xlApp.Calculation = originalCalcMode;
                workbook.Application.CalculateFullRebuild();

                //MessageBox.Show($"Refresh complete!\n\n" +
                //    $"Processed: {totalProcessed}\n" +
                //    $"Failed: {totalFailed}\n" +
                //    $"Pools used: {analysis.Pools.Count}\n" +
                //    $"Orphans: {analysis.Orphans.Count}");

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during refresh:\n\n{ex.Message}\n\n{ex.StackTrace}");
            }
            finally
            {
                // Always restore calculation mode
                xlApp.Calculation = originalCalcMode;
            }
        }

        /// <summary>
        /// Build pooled queries with tracking of which pools are in each query
        /// </summary>
        private List<PoolQueryBatch> BuildPooledQueriesWithTracking(List<QueryPool> pools, int maxQueryLength)
        {
            var batches = new List<PoolQueryBatch>();
            var currentBatch = new PoolQueryBatch { Pools = new List<QueryPool>() };
            int currentLength = 0;

            foreach (var pool in pools)
            {
                string poolDax = QueryPoolAnalyzer.BuildSinglePoolDax(pool);
                int poolLength = poolDax.Length;

                int projectedLength = currentLength + poolLength + 20;

                if (currentBatch.Pools.Count > 0 && projectedLength > maxQueryLength)
                {
                    // Finalize current batch
                    currentBatch.DaxQuery = QueryPoolAnalyzer.WrapPoolsInUnionDax(
                        currentBatch.Pools.Select(p => QueryPoolAnalyzer.BuildSinglePoolDax(p)).ToList()
                    );
                    batches.Add(currentBatch);
                    currentBatch = new PoolQueryBatch { Pools = new List<QueryPool>() };
                    currentLength = 0;
                }

                currentBatch.Pools.Add(pool);
                currentLength += poolLength + 1;
            }

            if (currentBatch.Pools.Count > 0)
            {
                currentBatch.DaxQuery = QueryPoolAnalyzer.WrapPoolsInUnionDax(
                    currentBatch.Pools.Select(p => QueryPoolAnalyzer.BuildSinglePoolDax(p)).ToList()
                );
                batches.Add(currentBatch);
            }

            return batches;
        }

        /// <summary>
        /// Represents a batch of pools that are combined into a single DAX query
        /// </summary>
        private class PoolQueryBatch
        {
            public List<QueryPool> Pools { get; set; }
            public string DaxQuery { get; set; }
        }

        /// <summary>
        /// Prepare results dictionary for batch storage with signatures
        /// </summary>
        private Dictionary<string, (object result, string signature)> PrepareResultsForStorage(
            Dictionary<string, object> results,
            List<RefreshItem> allItems)
        {
            var itemsToStore = new Dictionary<string, (object result, string signature)>();

            // DIAGNOSTIC: Show what we're working with
            //MessageBox.Show($"PrepareResultsForStorage:\n\n" +
            //    $"Results count: {results.Count}\n" +
            //    $"AllItems count: {allItems.Count}\n\n" +
            //    $"Sample result keys:\n{string.Join("\n", results.Keys.Take(3))}\n\n" +
            //    $"Sample item keys:\n{string.Join("\n", allItems.Take(3).Select(i => i.CacheKey))}");

            foreach (var kvp in results)
            {
                string cacheKey = kvp.Key;
                object result = kvp.Value;

                // Skip invalid keys
                if (string.IsNullOrWhiteSpace(cacheKey) ||
                    cacheKey.StartsWith("[") ||
                    cacheKey.Equals("CacheKey", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Find matching items to build signature
                var matchingItems = allItems.Where(item => item.CacheKey == cacheKey).ToList();

                string signature;
                if (matchingItems.Any())
                {
                    // Build signature from first matching item
                    var firstMatch = matchingItems.First();
                    signature = BuildSignature(firstMatch);
                }
                else
                {
                    // No match found - build signature from cache key
                    // Cache key format: FUNCTIONNAME|param0|param1|param2|
                    signature = BuildSignatureFromCacheKey(cacheKey);
                }

                itemsToStore[cacheKey] = (result, signature);
            }

            return itemsToStore;
        }

        /// <summary>
        /// Build formula signature from cache key
        /// Format: FUNCTIONNAME|param0|param1|param2| -> =FUNCTIONNAME("param0","param1","param2")
        /// </summary>
        private string BuildSignatureFromCacheKey(string cacheKey)
        {
            try
            {
                // Split by pipe
                var parts = cacheKey.Split('|');

                if (parts.Length < 2) return $"={cacheKey}";

                string functionName = parts[0];

                // Build parameter list (skip function name and trailing empty)
                var parameters = parts.Skip(1).Where(p => !string.IsNullOrEmpty(p)).ToList();

                // Format each parameter
                var formattedParams = parameters.Select(p =>
                {
                    // If it looks like a number, don't quote it
                    if (double.TryParse(p, out _))
                    {
                        return p;
                    }
                    // Otherwise quote it
                    return $"\"{p}\"";
                }).ToList();

                return $"={functionName}({string.Join(",", formattedParams)})";
            }
            catch
            {
                return $"={cacheKey}";
            }
        }

        /// <summary>
        /// Build formula signature for cache storage
        /// </summary>
        private string BuildSignature(RefreshItem item)
        {
            var formattedParams = new List<string>();

            foreach (var param in item.Parameters)
            {
                if (string.IsNullOrEmpty(param))
                {
                    formattedParams.Add("\"\"");
                }
                else if (param.Contains(","))
                {
                    formattedParams.Add($"\"{param}\"");
                }
                else if (double.TryParse(param, out _))
                {
                    formattedParams.Add(param);
                }
                else
                {
                    formattedParams.Add($"\"{param}\"");
                }
            }

            return $"={item.Config.FunctionName}({string.Join(",", formattedParams)})";
        }

        /// <summary>
        /// Delete the cache worksheet to force all cells to refresh
        /// </summary>
        private void DeleteCacheWorksheet()
        {
            try
            {
                Excel.Worksheet cacheSheet = workbook.Worksheets["__CubeConnector_Cache__"];
                xlApp.DisplayAlerts = false;
                cacheSheet.Delete();
                xlApp.DisplayAlerts = true;
            }
            catch
            {
                // Cache worksheet doesn't exist, that's fine
            }
        }

        /// <summary>
        /// Find all cells that contain UDF formulas showing #REFRESH or #N/A
        /// </summary>
        /// <param name="targetSheet">Optional: Only search this specific sheet</param>
        /// <param name="targetRange">Optional: Only search this specific range</param>
        private List<RefreshItem> FindCellsNeedingRefresh(Excel.Worksheet targetSheet = null, Excel.Range targetRange = null)
        {
            var results = new List<RefreshItem>();

            // If a specific range is provided, only search that range
            if (targetRange != null)
            {
                return FindCellsInRange(targetRange);
            }

            // If a specific sheet is provided, only search that sheet
            if (targetSheet != null)
            {
                return FindCellsInSheet(targetSheet);
            }

            // Otherwise, search all visible sheets
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                {
                    continue;
                }

                results.AddRange(FindCellsInSheet(sheet));
            }

            return results;
        }

        /// <summary>
        /// Find cells needing refresh in a specific sheet
        /// </summary>
        private List<RefreshItem> FindCellsInSheet(Excel.Worksheet sheet)
        {
            var results = new List<RefreshItem>();
            Excel.Range usedRange = sheet.UsedRange;

            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    Excel.Range cell = usedRange.Cells[row, col];
                    var item = CheckCellForRefresh(cell);
                    if (item != null)
                    {
                        results.Add(item);
                    }
                }
            }

            return results;
        }

        /// <summary>
        /// Find cells needing refresh in a specific range
        /// </summary>
        private List<RefreshItem> FindCellsInRange(Excel.Range range)
        {
            var results = new List<RefreshItem>();

            foreach (Excel.Range cell in range.Cells)
            {
                var item = CheckCellForRefresh(cell);
                if (item != null)
                {
                    results.Add(item);
                }
            }

            return results;
        }

        /// <summary>
        /// Check if a single cell needs refresh
        /// </summary>
        private RefreshItem CheckCellForRefresh(Excel.Range cell)
        {
            if (!cell.HasFormula)
            {
                return null;
            }

            string formula = cell.Formula.ToString();

            // Extract the actual function name from the formula
            string functionName = ExtractFunctionName(formula);
            if (string.IsNullOrEmpty(functionName))
            {
                return null;
            }

            // Check if this function name matches one of our configured functions
            var allFunctionNames = ConfigurationStore.GetAllConfigs().Select(c => c.FunctionName).ToList();
            bool isOurFunction = allFunctionNames.Any(name => name.Equals(functionName, StringComparison.OrdinalIgnoreCase));

            if (!isOurFunction)
            {
                return null;
            }

            var value = cell.Value2;
            var text = cell.Text;
            string displayValue = value?.ToString() ?? text?.ToString() ?? "";

            if (displayValue.Contains("#REFRESH") || displayValue.Contains("#N/A"))
            {
                return ParseFormulaToRefreshItem(formula, cell);
            }

            return null;
        }

        private RefreshItem ParseFormulaToRefreshItem(string formula, Excel.Range cell)
        {
            try
            {
                string functionName = ExtractFunctionName(formula);
                if (string.IsNullOrEmpty(functionName)) return null;

                var config = ConfigurationStore.GetConfig(functionName);
                if (config == null) return null;

                var parameters = ExtractParametersFromCell(cell, config);
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

        private string ExtractFunctionName(string formula)
        {
            if (!formula.StartsWith("=")) return null;
            int parenIndex = formula.IndexOf('(');
            if (parenIndex < 0) return null;
            return formula.Substring(1, parenIndex - 1).Trim();
        }

        /// <summary>
        /// Normalize a parameter value to match CacheKey.NormalizeParameter logic
        /// </summary>
        private string NormalizeParameterValue(object value, ParameterConfig paramConfig = null)
        {
            if (value == null)
            {
                return "";
            }

            // Handle DateTime objects
            if (value is DateTime dt)
            {
                return dt.ToString("yyyy-MM-dd");
            }

            // Handle doubles that might be dates or years
            if (value is double dbl)
            {
                // Check if this is a date parameter
                if (paramConfig != null && paramConfig.DataType == "date")
                {
                    // Check if it's a year value (1900-2099)
                    if (dbl >= 1900 && dbl <= 2099 && Math.Abs(dbl - Math.Round(dbl)) < 0.0001)
                    {
                        int year = (int)Math.Round(dbl);

                        // Convert based on filter type
                        if (paramConfig.FilterType == FilterType.RangeStart)
                        {
                            return new DateTime(year, 1, 1).ToString("yyyy-MM-dd");
                        }
                        else if (paramConfig.FilterType == FilterType.RangeEnd)
                        {
                            return new DateTime(year, 12, 31).ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            // Default to first day of year if filter type not specified
                            return new DateTime(year, 1, 1).ToString("yyyy-MM-dd");
                        }
                    }
                    // Check if it's an Excel date serial
                    else if (dbl > 25569 && dbl < 73050)
                    {
                        DateTime date = DateTime.FromOADate(dbl);
                        return date.ToString("yyyy-MM-dd");
                    }
                }
                // Not a date parameter, or doesn't look like a date
                else if (dbl > 25569 && dbl < 73050)
                {
                    // Might be a date even if config doesn't say so
                    DateTime date = DateTime.FromOADate(dbl);
                    return date.ToString("yyyy-MM-dd");
                }
            }

            return value.ToString();
        }

        private string[] ExtractParametersFromCell(Excel.Range cell, UDFConfig config)
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

                for (int i = 0; i < argTokens.Count; i++)
                {
                    string cleanToken = argTokens[i].Trim();

                    // Get parameter config for this position
                    ParameterConfig paramConfig = null;
                    if (config?.Parameters != null && i < config.Parameters.Count)
                    {
                        paramConfig = config.Parameters[i];
                    }

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
                        // Normalize the number value (might be a year for date parameters)
                        parameters.Add(NormalizeParameterValue(numValue, paramConfig));
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
                                    if (v != null) values.Add(NormalizeParameterValue(v, paramConfig));
                                }
                                parameters.Add(string.Join(",", values));
                            }
                            else
                            {
                                parameters.Add(NormalizeParameterValue(value, paramConfig));
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

            // Dynamically pad to match the number of parameters in the config
            int expectedParamCount = config?.Parameters?.Count ?? 0;
            while (parameters.Count < expectedParamCount)
            {
                parameters.Add("");
            }

            return parameters.ToArray();
        }

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

        private Dictionary<string, object> ExecuteBatchQuery(string batchQuery)
        {
            var results = new Dictionary<string, object>();

            try
            {
                Excel.WorkbookConnection conn;
                try
                {
                    conn = workbook.Connections["CubeConnector"];
                }
                catch
                {
                    throw new Exception("Connection 'CubeConnector' not found.");
                }

                Excel.Worksheet querySheet;
                Excel.ListObject queryTable;

                try
                {
                    querySheet = workbook.Worksheets["__CubeConnector_Query__"];
                    queryTable = querySheet.ListObjects["CubeConnector_QueryTable"];
                }
                catch
                {
                    querySheet = workbook.Worksheets.Add();
                    querySheet.Name = "__CubeConnector_Query__";
                    querySheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;

                    Excel.OLEDBConnection oledbConn = conn.OLEDBConnection;

                    queryTable = querySheet.ListObjects.Add(
                        SourceType: Excel.XlListObjectSourceType.xlSrcExternal,
                        Source: conn,
                        LinkSource: true,
                        XlListObjectHasHeaders: Excel.XlYesNoGuess.xlNo,
                        Destination: querySheet.Range["A1"]
                    );

                    queryTable.Name = "CubeConnector_QueryTable";

                    Excel.QueryTable qt = queryTable.QueryTable;
                    qt.CommandType = oledbConn.CommandType;
                    qt.CommandText = "EVALUATE { 1 }";
                    qt.Refresh(BackgroundQuery: false);
                }

                Excel.QueryTable queryTableObj = queryTable.QueryTable;
                queryTableObj.CommandText = batchQuery;
                bool refreshed = queryTableObj.Refresh(BackgroundQuery: false);

                if (queryTableObj.ResultRange == null)
                {
                    throw new Exception("Batch query returned no results");
                }

                int rowCount = queryTableObj.ResultRange.Rows.Count;
                int colCount = queryTableObj.ResultRange.Columns.Count;

                if (rowCount > 0 && colCount >= 2)
                {
                    for (int row = 1; row <= rowCount; row++)
                    {
                        var cacheKeyCell = queryTableObj.ResultRange.Cells[row, 1].Value2;
                        var resultCell = queryTableObj.ResultRange.Cells[row, 2].Value2;

                        string cacheKey = cacheKeyCell?.ToString();

                        if (string.IsNullOrEmpty(cacheKey) || cacheKey.StartsWith("["))
                        {
                            continue;
                        }

                        results[cacheKey] = resultCell ?? "#NULL";
                    }
                }

                return results;
            }
            catch (Exception ex)
            {
                throw new Exception($"Batch query execution failed: {ex.Message}", ex);
            }
        }
    }

    public class RefreshItem
    {
        public string CacheKey { get; set; }
        public UDFConfig Config { get; set; }
        public string[] Parameters { get; set; }
        public string FunctionSignature { get; set; }
        public Excel.Range Cell { get; set; }
    }
}