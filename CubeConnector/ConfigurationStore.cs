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
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace CubeConnector
{
    /// <summary>
    /// Stores UDF configurations loaded from JSON file
    /// </summary>
    public static class ConfigurationStore
    {
        private static List<UDFConfig> _configs;
        private const string CONFIG_FILE_NAME = "CubeConnectorConfig.json";

        // Data contract classes for JSON deserialization
        [DataContract]
        private class ConfigFileWrapper
        {
            [DataMember(Name = "functions")]
            public List<FunctionConfigJson> Functions { get; set; }
        }

        [DataContract]
        private class FunctionConfigJson
        {
            [DataMember(Name = "functionName")]
            public string FunctionName { get; set; }

            [DataMember(Name = "tenantId")]
            public string TenantId { get; set; }

            [DataMember(Name = "datasetPrefix")]
            public string DatasetPrefix { get; set; }

            [DataMember(Name = "datasetId")]
            public string DatasetId { get; set; }

            [DataMember(Name = "measureName")]
            public string MeasureName { get; set; }

            [DataMember(Name = "parameters")]
            public List<ParameterConfigJson> Parameters { get; set; }
        }

        [DataContract]
        private class ParameterConfigJson
        {
            [DataMember(Name = "name")]
            public string Name { get; set; }

            [DataMember(Name = "position")]
            public int Position { get; set; }

            [DataMember(Name = "tableName")]
            public string TableName { get; set; }

            [DataMember(Name = "fieldName")]
            public string FieldName { get; set; }

            [DataMember(Name = "dataType")]
            public string DataType { get; set; }

            [DataMember(Name = "filterType")]
            public string FilterType { get; set; }

            [DataMember(Name = "isOptional")]
            public bool IsOptional { get; set; }
        }

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
            // Try to load from JSON file first
            _configs = LoadFromJson();

            // If JSON load failed, fall back to hardcoded config
            if (_configs == null || _configs.Count == 0)
            {
                _configs = GetFallbackConfigs();
            }
        }

        /// <summary>
        /// Load configurations from JSON file
        /// </summary>
        private static List<UDFConfig> LoadFromJson()
        {
            try
            {
                // Get the XLL directory (where the add-in file is located, not the unpacked DLL location)
                string xllPath = ExcelDna.Integration.ExcelDnaUtil.XllPath;
                string addInDirectory = Path.GetDirectoryName(xllPath);
                string configPath = Path.Combine(addInDirectory, CONFIG_FILE_NAME);

                if (!File.Exists(configPath))
                {
                    return null;
                }

                // Deserialize using DataContractJsonSerializer
                var serializer = new DataContractJsonSerializer(typeof(ConfigFileWrapper));
                ConfigFileWrapper configWrapper;

                using (var stream = File.OpenRead(configPath))
                {
                    configWrapper = (ConfigFileWrapper)serializer.ReadObject(stream);
                }

                if (configWrapper?.Functions == null || configWrapper.Functions.Count == 0)
                {
                    return null;
                }

                // Convert from JSON contract classes to UDFConfig
                var configs = new List<UDFConfig>();
                foreach (var funcJson in configWrapper.Functions)
                {
                    var config = ConvertToUDFConfig(funcJson);
                    if (config != null)
                    {
                        configs.Add(config);
                    }
                }

                return configs;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Convert JSON config to UDFConfig
        /// </summary>
        private static UDFConfig ConvertToUDFConfig(FunctionConfigJson funcJson)
        {
            try
            {
                var config = new UDFConfig
                {
                    FunctionName = funcJson.FunctionName,
                    TenantId = funcJson.TenantId,
                    DatasetPrefix = funcJson.DatasetPrefix,
                    DatasetId = funcJson.DatasetId,
                    MeasureName = funcJson.MeasureName,
                    Parameters = new List<ParameterConfig>()
                };

                // Auto-prefix datasetId if it's just a GUID and a prefix is provided
                if (!string.IsNullOrEmpty(config.DatasetPrefix) && !string.IsNullOrEmpty(config.DatasetId))
                {
                    // Check if datasetId looks like a GUID (36 characters, correct format)
                    if (IsGuid(config.DatasetId))
                    {
                        // Only add prefix if it's not already there
                        if (!config.DatasetId.StartsWith(config.DatasetPrefix))
                        {
                            config.DatasetId = config.DatasetPrefix + config.DatasetId;
                        }
                    }
                }

                if (funcJson.Parameters != null)
                {
                    foreach (var paramJson in funcJson.Parameters)
                    {
                        var param = new ParameterConfig
                        {
                            Name = paramJson.Name,
                            Position = paramJson.Position,
                            TableName = paramJson.TableName,
                            FieldName = paramJson.FieldName,
                            DataType = paramJson.DataType ?? "text",
                            IsOptional = paramJson.IsOptional
                        };

                        // Parse FilterType enum
                        FilterType filterType;
                        if (!string.IsNullOrEmpty(paramJson.FilterType) &&
                            Enum.TryParse(paramJson.FilterType, true, out filterType))
                        {
                            param.FilterType = filterType;
                        }

                        config.Parameters.Add(param);
                    }
                }

                return config;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Check if a string is a valid GUID format
        /// </summary>
        private static bool IsGuid(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length != 36)
                return false;

            Guid result;
            return Guid.TryParse(value, out result);
        }

        /// <summary>
        /// Fallback hardcoded configuration (used if JSON file not found or invalid)
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