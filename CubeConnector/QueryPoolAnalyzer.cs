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
using System.Text;

namespace CubeConnector
{
    /// <summary>
    /// Analyzes queries to find consolidation opportunities and builds optimized DAX pools
    /// </summary>
    public static class QueryPoolAnalyzer
    {
        /// <summary>
        /// Analyze queries and create pools where only 1 parameter varies
        /// </summary>
        public static PoolAnalysisResult AnalyzeQueries(List<RefreshItem> items, int minPoolSize = 3)
        {
            var result = new PoolAnalysisResult
            {
                Pools = new List<QueryPool>(),
                Orphans = new List<RefreshItem>()
            };

            if (items == null || items.Count == 0)
            {
                return result;
            }

            // Group by function first
            var byFunction = items.GroupBy(i => i.Config.FunctionName);

            foreach (var funcGroup in byFunction)
            {
                var config = funcGroup.First().Config;
                int paramCount = config.Parameters.Count;
                var unassigned = funcGroup.ToList();

                // Try each parameter position as the varying parameter
                for (int varyingParamIndex = 0; varyingParamIndex < paramCount; varyingParamIndex++)
                {
                    // Group by "signature" = all params except the varying one
                    var grouped = unassigned.GroupBy(item =>
                    {
                        var sig = new List<string>();
                        for (int i = 0; i < item.Parameters.Length; i++)
                        {
                            if (i != varyingParamIndex)
                            {
                                sig.Add(item.Parameters[i] ?? "");
                            }
                        }
                        return string.Join("||", sig); // Use || to avoid conflicts
                    });

                    foreach (var group in grouped)
                    {
                        if (group.Count() < minPoolSize)
                        {
                            continue; // Not worth consolidating
                        }

                        // Create a pool
                        var pool = new QueryPool
                        {
                            Config = config,
                            VaryingParamIndex = varyingParamIndex,
                            FixedParameters = new Dictionary<int, string>(),
                            VaryingValues = new List<string>(),
                            Items = group.ToList()
                        };

                        // Extract fixed parameters
                        var firstItem = group.First();
                        for (int i = 0; i < paramCount; i++)
                        {
                            if (i != varyingParamIndex)
                            {
                                pool.FixedParameters[i] = firstItem.Parameters[i] ?? "";
                            }
                        }

                        // Extract varying values
                        pool.VaryingValues = group
                            .Select(item => item.Parameters[varyingParamIndex] ?? "")
                            .Distinct()
                            .OrderBy(v => v)
                            .ToList();

                        result.Pools.Add(pool);

                        // Remove assigned items from unassigned list
                        foreach (var item in group)
                        {
                            unassigned.Remove(item);
                        }
                    }
                }

                // Remaining items are orphans
                result.Orphans.AddRange(unassigned);
            }

            return result;
        }

        /// <summary>
        /// Build a compact UNION query from pools, respecting batch size limit
        /// Returns multiple queries if needed to stay under character limit
        /// </summary>
        public static List<string> BuildPooledQueries(
            List<QueryPool> pools,
            int maxQueryLength = 30000)
        {
            var queries = new List<string>();
            var currentBatch = new List<QueryPool>();
            int currentLength = 0;

            foreach (var pool in pools)
            {
                string poolDax = BuildSinglePool(pool);
                int poolLength = poolDax.Length;

                // Check if adding this pool would exceed limit
                // Account for UNION wrapping: "EVALUATE UNION(" + pools + ")"
                int projectedLength = currentLength + poolLength + 20; // 20 for UNION overhead

                if (currentBatch.Count > 0 && projectedLength > maxQueryLength)
                {
                    // Finalize current batch
                    queries.Add(WrapPoolsInUnion(currentBatch));
                    currentBatch.Clear();
                    currentLength = 0;
                }

                currentBatch.Add(pool);
                currentLength += poolLength + 1; // +1 for comma separator
            }

            // Add remaining batch
            if (currentBatch.Count > 0)
            {
                queries.Add(WrapPoolsInUnion(currentBatch));
            }

            return queries;
        }

        /// <summary>
        /// Build DAX for a single pool - ultra compact (PUBLIC for external use)
        /// </summary>
        public static string BuildSinglePoolDax(QueryPool pool)
        {
            return BuildSinglePool(pool);
        }

        /// <summary>
        /// Wrap pool DAX strings in UNION (PUBLIC for external use)
        /// </summary>
        public static string WrapPoolsInUnionDax(List<string> poolDaxStrings)
        {
            if (poolDaxStrings.Count == 0) return "";
            if (poolDaxStrings.Count == 1) return $"EVALUATE {poolDaxStrings[0]}";
            return $"EVALUATE UNION({string.Join(",", poolDaxStrings)})";
        }

        /// <summary>
        /// Build DAX for a single pool - ultra compact
        /// </summary>
        private static string BuildSinglePool(QueryPool pool)
        {
            var config = pool.Config;
            int varyingIndex = pool.VaryingParamIndex;
            var varyingParam = config.Parameters[varyingIndex];

            // Build individual ROW statements for each varying value
            // This ensures we use the user's exact input values in cache keys (like orphan queries)
            var rowStatements = new List<string>();

            foreach (var varyingValue in pool.VaryingValues)
            {
                // Build the complete parameter array for this specific item
                var parameters = new string[config.Parameters.Count];
                for (int i = 0; i < config.Parameters.Count; i++)
                {
                    if (i == varyingIndex)
                    {
                        parameters[i] = varyingValue;
                    }
                    else if (pool.FixedParameters.ContainsKey(i))
                    {
                        parameters[i] = pool.FixedParameters[i];
                    }
                    else
                    {
                        parameters[i] = "";
                    }
                }

                // Build cache key from user's exact parameters (same as orphan queries)
                string cacheKey = CacheKey.BuildFromStrings(config.FunctionName, parameters);
                string escapedKey = cacheKey.Replace("\"", "\"\"");

                // Build filters for this specific set of parameters
                var filters = new List<string>();
                for (int i = 0; i < parameters.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(parameters[i])) continue;

                    var paramConfig = config.Parameters[i];
                    string filter = BuildFilterCompact(paramConfig, parameters[i]);
                    if (!string.IsNullOrEmpty(filter))
                    {
                        filters.Add(filter);
                    }
                }

                // Build CALCULATE expression
                string measureName = config.MeasureName.Trim();
                while (measureName.StartsWith("[") && measureName.EndsWith("]"))
                {
                    measureName = measureName.Substring(1, measureName.Length - 2).Trim();
                }
                measureName = $"[{measureName}]";

                string calculateExpr = filters.Count > 0
                    ? $"CALCULATE({measureName},{string.Join(",", filters)})"
                    : measureName;

                // Build ROW statement with literal cache key (like orphan queries)
                string rowStmt = $"ROW(\"CacheKey\",\"{escapedKey}\",\"Result\",{calculateExpr})";
                rowStatements.Add(rowStmt);
            }

            // UNION all rows together
            if (rowStatements.Count == 1)
            {
                return rowStatements[0];
            }
            else
            {
                return $"UNION({string.Join(",", rowStatements)})";
            }
        }

        /// <summary>
        /// Check if a value looks like a year (4-digit integer between 1900-2150)
        /// </summary>
        private static bool IsYearValue(string value)
        {
            if (int.TryParse(value, out int year))
            {
                return year >= 1900 && year <= 2150;
            }
            return false;
        }

        /// <summary>
        /// Build a compact filter expression for a parameter
        /// </summary>
        private static string BuildFilterCompact(ParameterConfig paramConfig, string paramValue)
        {
            string table = paramConfig.TableName;
            string field = paramConfig.FieldName;
            string fullField = $"{table}[{field}]";

            switch (paramConfig.FilterType)
            {
                case FilterType.List:
                    // Handle comma-separated values
                    var values = paramValue.Split(',')
                        .Select(v => v.Trim())
                        .Where(v => !string.IsNullOrEmpty(v))
                        .Select(v => FormatValueCompact(v, paramConfig.DataType));

                    if (!values.Any()) return "";

                    if (values.Count() == 1)
                    {
                        return $"{fullField}={values.First()}";
                    }
                    else
                    {
                        return $"{fullField} IN {{{string.Join(",", values)}}}";
                    }

                case FilterType.RangeStart:
                    // Check if it's a year value for date fields
                    if (paramConfig.DataType.ToLower() == "date" || paramConfig.DataType.ToLower() == "datetime")
                    {
                        if (IsYearValue(paramValue))
                        {
                            int year = int.Parse(paramValue);
                            return $"{fullField}>=DATE({year},1,1)";
                        }
                    }

                    var minValues = paramValue.Split(',').Where(v => !string.IsNullOrEmpty(v)).ToList();
                    if (!minValues.Any()) return "";
                    string minValue = FormatValueCompact(minValues.Min(), paramConfig.DataType);
                    return $"{fullField}>={minValue}";

                case FilterType.RangeEnd:
                    // Check if it's a year value for date fields
                    if (paramConfig.DataType.ToLower() == "date" || paramConfig.DataType.ToLower() == "datetime")
                    {
                        if (IsYearValue(paramValue))
                        {
                            int year = int.Parse(paramValue);
                            return $"{fullField}<=DATE({year},12,31)";
                        }
                    }

                    var maxValues = paramValue.Split(',').Where(v => !string.IsNullOrEmpty(v)).ToList();
                    if (!maxValues.Any()) return "";
                    string maxValue = FormatValueCompact(maxValues.Max(), paramConfig.DataType);
                    return $"{fullField}<={maxValue}";

                default:
                    return "";
            }
        }

        /// <summary>
        /// Format a value for DAX - ultra compact (no extra spaces)
        /// </summary>
        private static string FormatValueCompact(string value, string dataType)
        {
            switch (dataType.ToLower())
            {
                case "text":
                    string escaped = value.Replace("\"", "\"\"");
                    return $"\"{escaped}\"";

                case "number":
                case "integer":
                    return value.Trim();

                case "date":
                case "datetime":
                    // Try parsing as Excel date number first (e.g., "46022")
                    if (double.TryParse(value, out double excelDateNumber))
                    {
                        try
                        {
                            DateTime dt = DateTime.FromOADate(excelDateNumber);
                            return $"DATE({dt.Year},{dt.Month},{dt.Day})";
                        }
                        catch
                        {
                            // OADate conversion failed, fall through to string parsing
                        }
                    }

                    // Try parsing as date string (e.g., "2024-01-15")
                    if (DateTime.TryParse(value, out DateTime parsedDt))
                    {
                        return $"DATE({parsedDt.Year},{parsedDt.Month},{parsedDt.Day})";
                    }

                    // Fallback: return as-is (shouldn't happen if config is correct)
                    return value;

                default:
                    return $"\"{value}\"";
            }
        }

        /// <summary>
        /// Wrap multiple pools in a UNION - ultra compact
        /// </summary>
        private static string WrapPoolsInUnion(List<QueryPool> pools)
        {
            if (pools.Count == 0)
            {
                return "";
            }

            if (pools.Count == 1)
            {
                // Single pool - no UNION needed
                return $"EVALUATE {BuildSinglePool(pools[0])}";
            }

            // Multiple pools - wrap in UNION
            var poolStrings = pools.Select(p => BuildSinglePool(p));
            return $"EVALUATE UNION({string.Join(",", poolStrings)})";
        }
    }

    /// <summary>
    /// Result of pool analysis
    /// </summary>
    public class PoolAnalysisResult
    {
        public List<QueryPool> Pools { get; set; }
        public List<RefreshItem> Orphans { get; set; }
    }

    /// <summary>
    /// A pool of queries that differ only in one parameter
    /// </summary>
    public class QueryPool
    {
        public UDFConfig Config { get; set; }
        public int VaryingParamIndex { get; set; }
        public Dictionary<int, string> FixedParameters { get; set; }
        public List<string> VaryingValues { get; set; }
        public List<RefreshItem> Items { get; set; }
    }
}