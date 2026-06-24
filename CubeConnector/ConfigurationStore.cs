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
    /// Stores UDF configurations loaded from the per-user FunctionStore
    /// </summary>
    public static class ConfigurationStore
    {
        private static List<UDFConfig> _configs;

        public static List<UDFConfig> GetAllConfigs()
        {
            if (_configs == null)
            {
                InitializeConfigs();
            }
            return _configs;
        }

        public static UDFConfig GetConfig(string functionName)
        {
            return GetAllConfigs().FirstOrDefault(c =>
                c.FunctionName.Equals(functionName, StringComparison.OrdinalIgnoreCase));
        }

        private static void InitializeConfigs()
        {
            _configs = FunctionStore.GetAll();
            if (_configs == null || _configs.Count == 0)
                _configs = GetFallbackConfigs();
        }

        /// <summary>
        /// Fallback hardcoded configuration (used if per-user file not found or empty)
        /// </summary>
        private static List<UDFConfig> GetFallbackConfigs()
        {
            return new List<UDFConfig>
            {
                new UDFConfig
                {
                    FunctionName = "CC.AmtNet",
                    TenantId = "your-tenant-id-here",
                    DatasetId = "your-dataset-id-here",
                    MeasureName = "[AmtNet]",
                    Parameters = new List<ParameterConfig>
                    {
                        new ParameterConfig
                        {
                            Name = "accounts",
                            Position = 0,
                            TableName = "Account",
                            FieldName = "AccountID",
                            DataType = "text",
                            FilterType = FilterType.List,
                            IsOptional = true
                        },
                        new ParameterConfig
                        {
                            Name = "acctg_period_start",
                            Position = 1,
                            TableName = "AcctgPeriod",
                            FieldName = "Date",
                            DataType = "date",
                            FilterType = FilterType.RangeStart,
                            IsOptional = true
                        },
                        new ParameterConfig
                        {
                            Name = "acctg_period_end",
                            Position = 2,
                            TableName = "AcctgPeriod",
                            FieldName = "Date",
                            DataType = "date",
                            FilterType = FilterType.RangeEnd,
                            IsOptional = true
                        },
                        new ParameterConfig
                        {
                            Name = "cost_centers",
                            Position = 3,
                            TableName = "CostCenter",
                            FieldName = "CostCenterNumber",
                            DataType = "text",
                            FilterType = FilterType.List,
                            IsOptional = true
                        },
                        new ParameterConfig
                        {
                            Name = "afe",
                            Position = 4,
                            TableName = "Afe",
                            FieldName = "AfeNumber",
                            DataType = "text",
                            FilterType = FilterType.List,
                            IsOptional = true
                        }
                    }
                }
            };
        }
    }
}
