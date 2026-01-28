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
namespace CubeConnector
{
    /// <summary>
    /// Handles building and parsing cache keys
    /// </summary>
    public static class CacheKey
    {
        private const string DELIMITER = "|";

        /// <summary>
        /// Build a cache key from function name and parameters
        /// </summary>
        public static string Build(string functionName, params object[] parameters)
        {
            // Get the function config to understand parameter types
            var config = ConfigurationStore.GetConfig(functionName);

            var parts = new List<string> { functionName };

            // Normalize all parameters
            var normalizedParams = new List<string>();
            for (int i = 0; i < parameters.Length; i++)
            {
                ParameterConfig paramConfig = null;
                if (config?.Parameters != null && i < config.Parameters.Count)
                {
                    paramConfig = config.Parameters[i];
                }

                normalizedParams.Add(NormalizeParameter(parameters[i], paramConfig));
            }

            // Find the last non-empty parameter to match DAX query behavior (skip trailing empty params)
            int lastNonEmptyIndex = -1;
            for (int i = normalizedParams.Count - 1; i >= 0; i--)
            {
                if (!string.IsNullOrEmpty(normalizedParams[i]))
                {
                    lastNonEmptyIndex = i;
                    break;
                }
            }

            // Only include parameters up to the last non-empty one
            for (int i = 0; i <= lastNonEmptyIndex; i++)
            {
                parts.Add(normalizedParams[i]);
            }

            return string.Join(DELIMITER, parts);
        }

        /// <summary>
        /// Build a cache key from function name and string parameters (VSTO context)
        /// </summary>
        public static string BuildFromStrings(string functionName, params string[] parameters)
        {
            // Get the function config to understand parameter types for normalization
            var config = ConfigurationStore.GetConfig(functionName);

            var parts = new List<string> { functionName };

            // Normalize all parameters
            var normalizedParams = new List<string>();
            for (int i = 0; i < parameters.Length; i++)
            {
                string param = parameters[i] ?? "";

                // Uppercase text parameters for case-insensitive matching
                if (config?.Parameters != null && i < config.Parameters.Count)
                {
                    var paramConfig = config.Parameters[i];
                    if (paramConfig.DataType.ToLower() == "text")
                    {
                        param = param.ToUpper();
                    }
                }

                normalizedParams.Add(param);
            }

            // Find the last non-empty parameter to match DAX query behavior (skip trailing empty params)
            int lastNonEmptyIndex = -1;
            for (int i = normalizedParams.Count - 1; i >= 0; i--)
            {
                if (!string.IsNullOrEmpty(normalizedParams[i]))
                {
                    lastNonEmptyIndex = i;
                    break;
                }
            }

            // Only include parameters up to the last non-empty one
            for (int i = 0; i <= lastNonEmptyIndex; i++)
            {
                parts.Add(normalizedParams[i]);
            }

            return string.Join(DELIMITER, parts);
        }

        /// <summary>
        /// Normalize a parameter value for cache key
        /// </summary>
        private static string NormalizeParameter(object param, ParameterConfig paramConfig)
        {
            // Handle Excel special values
            if (param == null || param is ExcelMissing || param is ExcelEmpty)
            {
                return "";
            }

            // Handle ranges (convert to comma-separated list)
            if (param is object[,] range)
            {
                return NormalizeRange(range, paramConfig);
            }

            if (param is object[] array)
            {
                return NormalizeArray(array, paramConfig);
            }

            // Handle dates
            if (param is DateTime dt)
            {
                return dt.ToString("yyyy-MM-dd");
            }

            // Handle doubles that might be dates or years
            if (param is double dbl)
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

            // Normalize text parameters to uppercase for case-insensitive matching
            string result = param.ToString();
            if (paramConfig != null && paramConfig.DataType.ToLower() == "text")
            {
                result = result.ToUpper();
            }
            return result;
        }

        /// <summary>
        /// Convert Excel range to normalized string
        /// </summary>
        private static string NormalizeRange(object[,] range, ParameterConfig paramConfig)
        {
            var values = new List<string>();

            for (int i = range.GetLowerBound(0); i <= range.GetUpperBound(0); i++)
            {
                for (int j = range.GetLowerBound(1); j <= range.GetUpperBound(1); j++)
                {
                    var val = range[i, j];
                    if (val != null && !(val is ExcelEmpty) && !(val is ExcelMissing))
                    {
                        // Recursively normalize each value in the range
                        values.Add(NormalizeParameter(val, paramConfig));
                    }
                }
            }

            return string.Join(",", values);
        }

        /// <summary>
        /// Convert array to normalized string
        /// </summary>
        private static string NormalizeArray(object[] array, ParameterConfig paramConfig)
        {
            var values = array
                .Where(v => v != null && !(v is ExcelEmpty) && !(v is ExcelMissing))
                .Select(v => NormalizeParameter(v, paramConfig));

            return string.Join(",", values);
        }

        /// <summary>
        /// Parse parameters from a cache key
        /// </summary>
        public static (string functionName, string[] parameters) Parse(string cacheKey)
        {
            var parts = cacheKey.Split(new[] { DELIMITER }, StringSplitOptions.None);
            var functionName = parts[0];
            var parameters = parts.Skip(1).ToArray();
            return (functionName, parameters);
        }
    }
}