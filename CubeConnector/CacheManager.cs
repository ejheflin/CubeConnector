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
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace CubeConnector
{
    /// <summary>
    /// Manages the cache table for UDF results
    /// </summary>
    public static class CacheManager
    {
        private const string CACHE_SHEET_NAME = "__CubeConnector_Cache__";
        private const string CACHE_TABLE_NAME = "CubeConnector_CacheTable";

        /// <summary>
        /// Look up a value from the cache by key
        /// </summary>
        public static object Lookup(string cacheKey)
        {
            try
            {
                var xlApp = (Excel.Application)ExcelDnaUtil.Application;
                var workbook = xlApp.ActiveWorkbook;

                // Get cache sheet
                Excel.Worksheet cacheSheet;
                try
                {
                    cacheSheet = workbook.Worksheets[CACHE_SHEET_NAME];
                }
                catch
                {
                    return "#REFRESH";
                }

                // Get cache table
                Excel.ListObject cacheTable;
                try
                {
                    cacheTable = cacheSheet.ListObjects[CACHE_TABLE_NAME];
                }
                catch
                {
                    return "#REFRESH";
                }

                // Get the data range
                if (cacheTable.ListRows.Count == 0)
                {
                    return "#REFRESH"; // Empty cache
                }

                // Normalize cache key by trimming trailing pipes for comparison
                string normalizedSearchKey = cacheKey.TrimEnd('|');

                // Search for the cache key
                Excel.Range dataRange = cacheTable.DataBodyRange;

                for (int row = 1; row <= dataRange.Rows.Count; row++)
                {
                    var keyCell = dataRange.Cells[row, 1].Value2;
                    if (keyCell != null)
                    {
                        string storedKey = keyCell.ToString().TrimEnd('|');
                        if (storedKey == normalizedSearchKey)
                        {
                            // Found it! Return the result (column 2)
                            var result = dataRange.Cells[row, 2].Value2;
                            return result ?? "#NULL";
                        }
                    }
                }

                // Not found in cache
                return "#REFRESH";
            }
            catch (Exception ex)
            {
                return "#REFRESH";
            }
        }

        /// <summary>
        /// Store a value in the cache (VSTO context - pass workbook)
        /// </summary>
        public static void Store(Excel.Workbook workbook, string cacheKey, object result, string functionSignature)
        {
            try
            {
                // Validate inputs - DON'T store if cache key is empty
                if (string.IsNullOrWhiteSpace(cacheKey))
                {
                    return; // Don't store empty keys!
                }

                // Check if cache sheet exists
                Excel.Worksheet cacheSheet = null;
                try
                {
                    cacheSheet = workbook.Worksheets[CACHE_SHEET_NAME];
                }
                catch
                {
                    throw new Exception($"Cache sheet '{CACHE_SHEET_NAME}' not found.");
                }

                Excel.ListObject cacheTable;
                try
                {
                    cacheTable = cacheSheet.ListObjects[CACHE_TABLE_NAME];
                }
                catch
                {
                    throw new Exception($"Cache table '{CACHE_TABLE_NAME}' not found.");
                }

                // Check if key already exists - update if so
                bool found = false;

                if (cacheTable.ListRows.Count > 0)
                {
                    Excel.Range dataRange = cacheTable.DataBodyRange;

                    if (dataRange != null)
                    {
                        for (int row = 1; row <= dataRange.Rows.Count; row++)
                        {
                            Excel.Range keyCell = (Excel.Range)dataRange.Cells[row, 1];
                            var keyValue = keyCell.Value2;

                            if (keyValue != null && keyValue.ToString() == cacheKey)
                            {
                                // Update existing row
                                ((Excel.Range)dataRange.Cells[row, 2]).Value2 = result;
                                ((Excel.Range)dataRange.Cells[row, 3]).Value2 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                ((Excel.Range)dataRange.Cells[row, 4]).Value2 = "'" + functionSignature;
                                found = true;
                                break;
                            }
                        }
                    }
                }

                if (!found)
                {
                    // Add new row
                    Excel.ListRow newRow = cacheTable.ListRows.Add();
                    Excel.Range newRowRange = newRow.Range;

                    ((Excel.Range)newRowRange.Cells[1, 1]).Value2 = cacheKey;
                    ((Excel.Range)newRowRange.Cells[1, 2]).Value2 = result;
                    ((Excel.Range)newRowRange.Cells[1, 3]).Value2 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    ((Excel.Range)newRowRange.Cells[1, 4]).Value2 = "'" + functionSignature;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error storing in cache:\n\n{ex.Message}\n\n{ex.StackTrace}");
                throw;
            }
        }

        /// <summary>
        /// Clear all cache entries (VSTO context - pass in workbook)
        /// </summary>
        public static void Clear(Excel.Workbook workbook)
        {
            try
            {
                Excel.Worksheet cacheSheet = workbook.Worksheets[CACHE_SHEET_NAME];
                Excel.ListObject cacheTable = cacheSheet.ListObjects[CACHE_TABLE_NAME];

                // Delete all rows
                if (cacheTable.ListRows.Count > 0)
                {
                    for (int i = cacheTable.ListRows.Count; i >= 1; i--)
                    {
                        cacheTable.ListRows[i].Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error clearing cache: {ex.Message}");
            }
        }

        /// <summary>
        /// Clear all cache entries (UDF context - uses ExcelDna)
        /// </summary>
        public static void Clear()
        {
            try
            {
                var xlApp = (Excel.Application)ExcelDnaUtil.Application;
                var workbook = xlApp.ActiveWorkbook;
                Clear(workbook);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error clearing cache: {ex.Message}");
            }
        }

        /// <summary>
        /// Store multiple values in the cache at once (more efficient, avoids row count issues)
        /// </summary>
        public static void StoreBatch(Excel.Workbook workbook, Dictionary<string, (object result, string signature)> items)
        {
            try
            {
                if (items == null || items.Count == 0)
                {
                    return;
                }

                // Get cache sheet and table
                Excel.Worksheet cacheSheet = workbook.Worksheets[CACHE_SHEET_NAME];
                Excel.ListObject cacheTable = cacheSheet.ListObjects[CACHE_TABLE_NAME];

                int initialRowCount = cacheTable.ListRows.Count;

                // Special handling for empty table
                bool tableWasEmpty = (initialRowCount == 0);

                // Build a list of items to add vs update
                var itemsToAdd = new List<KeyValuePair<string, (object result, string signature)>>();
                var existingKeys = new HashSet<string>();

                // First pass: identify existing keys (skip if table was empty)
                if (!tableWasEmpty && cacheTable.ListRows.Count > 0)
                {
                    Excel.Range dataRange = cacheTable.DataBodyRange;
                    if (dataRange != null)
                    {
                        for (int row = 1; row <= dataRange.Rows.Count; row++)
                        {
                            var keyValue = ((Excel.Range)dataRange.Cells[row, 1]).Value2?.ToString();
                            if (!string.IsNullOrEmpty(keyValue))
                            {
                                existingKeys.Add(keyValue);
                            }
                        }
                    }
                }

                // Second pass: update existing, collect new
                foreach (var kvp in items)
                {
                    string cacheKey = kvp.Key;
                    object result = kvp.Value.result;
                    string signature = kvp.Value.signature;

                    if (string.IsNullOrWhiteSpace(cacheKey))
                    {
                        continue; // Skip empty keys
                    }

                    if (existingKeys.Contains(cacheKey))
                    {
                        // Update existing row
                        UpdateExistingRow(cacheTable, cacheKey, result, signature);
                    }
                    else
                    {
                        // Queue for batch add
                        itemsToAdd.Add(kvp);
                    }
                }

                // Third pass: batch add all new rows
                int addedCount = 0;
                foreach (var kvp in itemsToAdd)
                {
                    string cacheKey = kvp.Key;
                    object result = kvp.Value.result;
                    string signature = kvp.Value.signature;

                    Excel.ListRow newRow = cacheTable.ListRows.Add();
                    Excel.Range newRowRange = newRow.Range;

                    ((Excel.Range)newRowRange.Cells[1, 1]).Value2 = cacheKey;
                    ((Excel.Range)newRowRange.Cells[1, 2]).Value2 = result;
                    ((Excel.Range)newRowRange.Cells[1, 3]).Value2 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    ((Excel.Range)newRowRange.Cells[1, 4]).Value2 = "'" + signature;

                    addedCount++;

                    // If this was the first row added to an empty table AND count jumped to 2, delete the blank row
                    if (tableWasEmpty && addedCount == 1 && cacheTable.ListRows.Count == 2)
                    {
                        // The blank row is usually row 1 (first row)
                        // Check both rows to find which one is blank
                        for (int checkRow = 1; checkRow <= 2; checkRow++)
                        {
                            try
                            {
                                Excel.ListRow row = cacheTable.ListRows[checkRow];
                                var keyValue = ((Excel.Range)row.Range.Cells[1, 1]).Value2;

                                if (keyValue == null || string.IsNullOrWhiteSpace(keyValue.ToString()))
                                {
                                    row.Delete();
                                    break; // Stop after deleting one blank row
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Windows.Forms.MessageBox.Show($"Error checking row {checkRow}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error in StoreBatch:\n\n{ex.Message}\n\n{ex.StackTrace}");
                throw;
            }
        }

        /// <summary>
        /// Update an existing row in the cache table
        /// </summary>
        private static void UpdateExistingRow(Excel.ListObject cacheTable, string cacheKey, object result, string signature)
        {
            if (cacheTable.ListRows.Count == 0) return;

            Excel.Range dataRange = cacheTable.DataBodyRange;
            if (dataRange == null) return;

            for (int row = 1; row <= dataRange.Rows.Count; row++)
            {
                var keyValue = ((Excel.Range)dataRange.Cells[row, 1]).Value2?.ToString();
                if (keyValue == cacheKey)
                {
                    ((Excel.Range)dataRange.Cells[row, 2]).Value2 = result;
                    ((Excel.Range)dataRange.Cells[row, 3]).Value2 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    ((Excel.Range)dataRange.Cells[row, 4]).Value2 = "'" + signature;
                    return;
                }
            }
        }
    }
}