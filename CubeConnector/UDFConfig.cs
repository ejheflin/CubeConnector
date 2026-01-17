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

namespace CubeConnector
{
    /// <summary>
    /// Configuration for a single UDF
    /// </summary>
    public class UDFConfig
    {
        public string FunctionName { get; set; }
        public string TenantId { get; set; }
        public string DatasetPrefix { get; set; }
        public string DatasetId { get; set; }
        public string MeasureName { get; set; }
        public List<ParameterConfig> Parameters { get; set; }

        public UDFConfig()
        {
            Parameters = new List<ParameterConfig>();
        }
    }

    /// <summary>
    /// Configuration for a UDF parameter
    /// </summary>
    public class ParameterConfig
    {
        public string Name { get; set; }
        public int Position { get; set; }
        public string TableName { get; set; }
        public string FieldName { get; set; }
        public string DataType { get; set; } // "text", "number", "date", "datetime"
        public FilterType FilterType { get; set; }
        public bool IsOptional { get; set; }
    }

    public enum FilterType
    {
        List,
        RangeStart,
        RangeEnd
    }
}