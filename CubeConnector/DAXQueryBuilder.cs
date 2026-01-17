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

namespace CubeConnector
{
    /// <summary>
    /// Builds DAX queries from UDF configurations and parameters
    /// </summary>
    public static class DAXQueryBuilder
    {
        /// <summary>
        /// Build a CALCULATE query for a UDF with given parameters
        /// </summary>
        public static string BuildCalculateQuery(UDFConfig config, string[] parameters)
        {
            // Start with the measure
            string dax = $"EVALUATE {{ {config.MeasureName} ";

            // Build filters
            var filters = new List<string>();

            for (int i = 0; i < config.Parameters.Count && i < parameters.Length; i++)
            {
                var paramConfig = config.Parameters[i];
                var paramValue = parameters[i];

                // Skip empty/optional parameters
                if (string.IsNullOrWhiteSpace(paramValue))
                {
                    continue;
                }

                string filter = BuildFilter(paramConfig, paramValue);
                if (!string.IsNullOrEmpty(filter))
                {
                    filters.Add(filter);
                }
            }

            // Add filters to CALCULATE if any exist
            if (filters.Count > 0)
            {
                dax = $"EVALUATE {{ CALCULATE( {config.MeasureName}, {string.Join(", ", filters)} ) }}";
            }
            else
            {
                dax = $"EVALUATE {{ {config.MeasureName} }}";
            }

            return dax;
        }

        /// <summary>
        /// Build a filter expression for a single parameter
        /// </summary>
        private static string BuildFilter(ParameterConfig paramConfig, string paramValue)
        {
            string tableName = paramConfig.TableName;
            string fieldName = paramConfig.FieldName;
            string fullField = $"{tableName}[{fieldName}]";

            switch (paramConfig.FilterType)
            {
                case FilterType.List:
                    return BuildListFilter(fullField, paramValue, paramConfig.DataType);

                case FilterType.RangeStart:
                    return BuildRangeStartFilter(fullField, paramValue, paramConfig.DataType);

                case FilterType.RangeEnd:
                    return BuildRangeEndFilter(fullField, paramValue, paramConfig.DataType);

                default:
                    return "";
            }
        }

        /// <summary>
        /// Build IN filter for List type: Field IN { 'value1', 'value2' }
        /// </summary>
        private static string BuildListFilter(string field, string paramValue, string dataType)
        {
            // Handle comma-separated values (from ranges)
            var values = paramValue.Split(',').Select(v => v.Trim()).Where(v => !string.IsNullOrEmpty(v));

            if (!values.Any())
            {
                return "";
            }

            // Format values based on data type
            var formattedValues = values.Select(v => FormatValue(v, dataType));

            return $"{field} IN {{ {string.Join(", ", formattedValues)} }}";
        }

        /// <summary>
        /// Build >= filter for RangeStart: Field >= value
        /// </summary>
        private static string BuildRangeStartFilter(string field, string paramValue, string dataType)
        {
            // For ranges, take MIN value
            var values = paramValue.Split(',').Select(v => v.Trim()).Where(v => !string.IsNullOrEmpty(v)).ToList();
            if (!values.Any()) return "";

            string minValue = values.Min(); // This works for strings, numbers, and dates in yyyy-MM-dd format
            string formattedValue = FormatValue(minValue, dataType);

            return $"{field} >= {formattedValue}";
        }

        /// <summary>
        /// Build <= filter for RangeEnd: Field <= value
        /// </summary>
        private static string BuildRangeEndFilter(string field, string paramValue, string dataType)
        {
            // For ranges, take MAX value
            var values = paramValue.Split(',').Select(v => v.Trim()).Where(v => !string.IsNullOrEmpty(v)).ToList();
            if (!values.Any()) return "";

            string maxValue = values.Max();
            string formattedValue = FormatValue(maxValue, dataType);

            return $"{field} <= {formattedValue}";
        }

        /// <summary>
        /// Format a value for DAX based on data type
        /// </summary>
        private static string FormatValue(string value, string dataType)
        {
            switch (dataType.ToLower())
            {
                case "text":
                    // Escape single quotes by doubling them
                    string escaped = value.Replace("'", "''");
                    return $"\"{escaped}\"";  // Use double quotes for text in DAX

                case "number":
                case "integer":
                    // No quotes for numbers - just return the value
                    // Remove any whitespace
                    return value.Trim();

                case "date":
                case "datetime":
                    // Parse date and format as DATE(year, month, day)
                    if (DateTime.TryParse(value, out DateTime dt))
                    {
                        return $"DATE({dt.Year}, {dt.Month}, {dt.Day})";
                    }
                    return value; // Fallback

                default:
                    return $"\"{value}\""; // Default to text
            }
        }

        /// <summary>
        /// Build a batched UNION query for multiple parameter sets
        /// Returns a query that includes a key column to match results back
        /// </summary>
        public static string BuildBatchQuery(UDFConfig config, List<BatchQueryItem> items)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("No items to batch");
            }

            // Build queries for all items
            var queries = new List<string>();

            foreach (var item in items)
            {
                string queryPart = BuildCalculateQueryWithKey(config, item.Parameters, item.CacheKey);
                queries.Add(queryPart);
            }

            // If only one item, wrap in EVALUATE
            if (items.Count == 1)
            {
                return $"EVALUATE\n{queries[0]}";
            }

            // Multiple items: Combine with UNION
            string batchQuery = "EVALUATE\nUNION(\n    " + string.Join(",\n    ", queries) + "\n)";

            return batchQuery;
        }

        /// <summary>
        /// Build a CALCULATE query that includes a key column for result matching
        /// </summary>
        private static string BuildCalculateQueryWithKey(UDFConfig config, string[] parameters, string cacheKey)
        {
            // Build filters
            var filters = new List<string>();

            for (int i = 0; i < config.Parameters.Count && i < parameters.Length; i++)
            {
                var paramConfig = config.Parameters[i];
                var paramValue = parameters[i];

                // Skip empty/optional parameters
                if (string.IsNullOrWhiteSpace(paramValue))
                {
                    continue;
                }

                string filter = BuildFilter(paramConfig, paramValue);
                if (!string.IsNullOrEmpty(filter))
                {
                    filters.Add(filter);
                }
            }

            // Build the CALCULATE expression with filters
            string calculateExpr;
            if (filters.Count > 0)
            {
                calculateExpr = $"CALCULATE( {config.MeasureName}, {string.Join(", ", filters)} )";
            }
            else
            {
                calculateExpr = config.MeasureName;
            }

            // Escape the cache key for DAX (double quotes)
            string escapedKey = cacheKey.Replace("\"", "\"\"");

            // Return a table with CacheKey and Result columns
            string query = $"ROW( \"CacheKey\", \"{escapedKey}\", \"Result\", {calculateExpr} )";

            return query;
        }

        /// <summary>
        /// Item for batch query execution
        /// </summary>
        public class BatchQueryItem
        {
            public string CacheKey { get; set; }
            public string[] Parameters { get; set; }
        }
    }
}