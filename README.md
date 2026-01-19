![CubeConnector Banner](images/banner.png)

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.7.2-512BD4)](https://dotnet.microsoft.com/)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.9.0-green)](https://excel-dna.net/)
[![Power BI](https://img.shields.io/badge/Power%20BI-Pro%20|%20Premium%20|%20Fabric-F2C811)](https://powerbi.microsoft.com/)
[![Version](https://img.shields.io/badge/version-1.0.0-brightgreen)](https://github.com/[owner]/CubeConnector/releases)

A fast, flexible Excel add-in built on Excel-DNA that creates dynamic Excel functions (UDFs) to query Power BI datasets with intelligent caching and drillthrough capabilities.

## Overview

CubeConnector is a faster, easier, and more flexible alternative to Excel's built-in cube functions. Instead of manually writing complex formulas or relying on static pivot tables, you define custom Excel functions in a JSON configuration file that query Power BI measures directly. Each function can accept up to 15 user-defined parameters and supports model-defined drillthrough capabilities.
![Cube Connector Demo](images/CCdemo_small.gif)

## Security and Authentication

CubeConnector leverages **Microsoft's own "Analyze in Excel" connection** infrastructure for secure, seamless authentication:

- **Microsoft-Managed Authentication**: All authentication and token management is handled by Microsoft
- **No Credential Storage**: CubeConnector cannot access or store your credentials
- **Existing Security Respected**: All Row-Level Security (RLS) and model security permissions are fully respected
- **Zero Trust Required**: No admin rights, permissions, or delegations needed
- **User Control**: Cautious users can manually create their own connection strings

## How It Works

1. **Define Functions**: You define custom Excel functions in a JSON configuration file, with each function bound to a specific measure in your Power BI model
2. **Configure Parameters**: Each function accepts up to 15 parameters, each bound to a `Table[Field]` in your model for filtering
3. **Dynamic Registration**: Functions are automatically registered with Excel when the add-in loads, appearing like any other built-in Excel function
4. **Execute Queries**: When called, CubeConnector dynamically generates and executes DAX queries against your Power BI connection
5. **Intelligent Caching**: Results are cached for performance, with manual and automatic refresh options

## Key Features

### Intelligent Caching
- Query results are cached in a hidden worksheet table
- Dramatically improves performance for repeated queries
- Manual and automatic refresh options
- Cache invalidation strategies for data accuracy

### Drillthrough Capabilities
- **Drill to Details**: Right-click any cell to see the underlying detail records
- **Drill to Pivot**: Convert a query result into a filtered pivot table for further exploration
- Context menu integration for seamless workflows

### Flexible Parameter System
- Support for multiple filter types:
  - List filters (comma-separated values)
  - RangeStart or RangeEnd (start/end dates)
- Optional parameters
- Type-safe parameter handling (text, date, numeric)

### Excel Ribbon Integration
- Custom "CubeConnector" group in the Data tab
- Quick access to:
  - Refresh cache
  - Drill to details
  - Drill to pivot

## Use Cases

- **Financial Reporting**: Create reusable functions for budget analysis, variance reporting, and KPI tracking
- **Sales Analytics**: Build custom functions for revenue analysis, customer segmentation, and pipeline reporting
- **Operational Dashboards**: Query real-time operational metrics directly in Excel
- **Ad-hoc Analysis**: Quickly slice and dice Power BI data without leaving Excel
- **Data Distribution**: Share Excel templates with embedded Power BI queries

## System Requirements

- Microsoft Excel (Windows)
- .NET Framework 4.7.2 or higher
- Access to Power BI Pro, Premium, or Fabric workspace

## Getting Started

### Configuration

1. Download the latest release from the [Releases](../../releases) page
2. Extract the files to a local directory
3. Edit the `CubeConnectorConfig.json` file in the same directory as the `.xll` file
    - Paste in your tenantID
    - Paste in your datasetID
    - (optional) Define each param - these are used as filters for the function  
    - Note: if you already have an "analyze in excel" pivot, you can find your tenantID and datasetID by inspecting the connection:
![Finding Connection String IDs](images/connectionstring.png)

### Basic Structure

```json
{
  "functions": [
    {
      "functionName": "CC.AmtNet",
      "tenantId": "your-tenant-id-here",
      "datasetId": "your-dataset-id-here",
      "measureName": "[AmtNet]",
      "parameters": [ ... ]
    }
  ]
}
```

### Configuration Properties

| Property | Type | Description |
|----------|------|-------------|
| `functionName` | string | Name of the Excel function (e.g., "CC.SumSales", "MyFunction") |
| `tenantId` | string | Azure AD tenant ID |
| `datasetId` | string | Power BI dataset ID - can be full ID or just the GUID if using datasetPrefix |
| `measureName` | string | DAX measure name (e.g., "[Revenue]") |
| `parameters` | array | Array of parameter configurations |

### Parameter Configuration

| Property | Type | Description |
|----------|------|-------------|
| `name` | string | Parameter name (for documentation) |
| `position` | integer | Zero-based parameter position |
| `tableName` | string | Power BI table name for filtering |
| `fieldName` | string | Power BI field/column name |
| `dataType` | string | Data type: "text", "date", or "numeric" |
| `filterType` | string | Filter type: "List", "RangeStart", "RangeEnd"|
| `isOptional` | boolean | Whether parameter can be empty |

### Filter Types

**List**: Comma-separated values (e.g., "1000,2000,3000")
```json
{
  "filterType": "List",
  "dataType": "text"
}
```

**RangeStart/RangeEnd**: Date or numeric range boundaries
```json
{
  "filterType": "RangeStart",
  "dataType": "date"
}
```

5. Close out of all Excel instances and restart Excel

### Excel Installation

**Simple setup - no traditional install required:**

1. In Excel, go to File → Options → Add-ins → Manage Excel Add-ins → Browse
2. Select the `CubeConnector.xll` file from the downloaded directory
3. Click OK to enable the add-in

**Key Benefits:**
- ✅ No administrative rights required
- ✅ No permissions or delegations needed
- ✅ Just load the XLL - that's it!

### Basic Usage

Once configured, your custom functions appear in Excel's formula autocomplete:

```excel
=CC.MyMeasure(p1, p2, p3, ...)
```

Where:
- `p1` = Param1 as defined in your json (e.g. Account codes "1000,2000,3000")
- `p2` = Param2 as defined in your json (e.g. Start date "1/1/2026")
- `p3` = Param3 as defined in your json (e.g. End date "12/31/2026")

### Refreshing Data

**Option 1: Ribbon Button**
- Click the "Refresh" button in the CubeConnector group on the Data tab

**Option 2: Context Menu**
- Right-click any cell
- Select "CubeConnector - Refresh Cache"

**Option 3: Programmatic**
- Functions automatically refresh when cache is cleared

### Drillthrough

**Drill to Details:**
1. Click on a cell with a CubeConnector function result
2. Right-click → "CubeConnector - Drill to Details"
3. A new sheet opens with the underlying detail records

**Drill to Pivot:**
1. Click on a cell with a CubeConnector function result
2. Right-click → "CubeConnector - Drill to Pivot"
3. A new sheet opens with an interactive pivot table


## Advanced Topics

### Finding Your Power BI IDs

**Tenant ID:**
- Azure Portal → Azure Active Directory → Properties → Tenant ID

**Dataset ID:**
- Open dataset settings in Power BI
- Copy the GUID from the URL: `https://app.powerbi.com/groups/.../datasets/{dataset-id}/...`
- **Note**: Some Power BI deployments use prefixed dataset IDs (e.g., `sobe_wowvirtualserver-<guid>`). You can either:
  - Provide the full dataset ID: `"datasetId": "sobe_wowvirtualserver-cae9a534-453a-4513-b77d-cda5bfc91fd0"`
  - Or use the `datasetPrefix` field and provide just the GUID:
    ```json
    "datasetPrefix": "sobe_wowvirtualserver-",
    "datasetId": "cae9a534-453a-4513-b77d-cda5bfc91fd0"
    ```
  - The prefix will be automatically added if the `datasetId` is detected as a GUID

#### Alternative Method: From Existing Excel Connection

This method is particularly useful when you're setting up CubeConnector to match an existing Power BI connection.

### Cache Management

The cache is stored in a hidden worksheet named `__CubeConnector_Cache__` with a table named `CubeConnector_CacheTable`. The cache structure:

| Column | Purpose |
|--------|---------|
| CacheKey | Unique hash of function + parameters |
| Value | Query result |
| Timestamp | Last refresh time |

### Performance Optimization

1. **Use caching**: Don't disable cache unless absolutely necessary
2. **Limit parameter cardinality**: Fewer unique parameter combinations = better performance
3. **Optimize DAX measures**: Ensure your Power BI measures are optimized
4. **Batch refreshes**: Refresh cache in bulk rather than individual queries
5. **Use optional parameters**: Skip unnecessary filters to reduce query complexity

### Troubleshooting

**#REFRESH Error:**
- Cache needs to be refreshed
- Click "Refresh" in the data tab of the ribbon or context menu

**Authentication Errors:**
- Verify tenant ID is correct
- Ensure you're signed into Excel with the correct Azure AD account
- Check Power BI workspace permissions

**Function Not Found:**
- Verify `CubeConnectorConfig.json` is in the same directory as the `.xll` file
- Check JSON syntax is valid
- Restart Excel after configuration changes

**No Data Returned:**
- Verify workspace ID and dataset ID are correct
- Check Power BI dataset is published and accessible
- Ensure measure name exactly matches (case-sensitive)

## Project Structure

```
CubeConnector/
├── CacheKey.cs                    # Cache key generation and hashing
├── CacheManager.cs                # Cache table management
├── ConfigurationStore.cs          # JSON configuration loading
├── CubeConnectorFunctions.cs      # UDF entry points
├── CubeConnectorRibbon.cs         # Excel ribbon UI
├── DAXQueryBuilder.cs             # Dynamic DAX query generation
├── DrillthroughManager.cs         # Drill-to-details functionality
├── DynamicFunctionRegistration.cs # Function registration engine
├── PivotManager.cs                # Drill-to-pivot functionality
├── QueryPoolAnalyzer.cs           # Query optimization and analysis
├── RefreshManager.cs              # Cache refresh orchestration
├── UDFConfig.cs                   # Configuration models
├── CubeConnectorConfig.json       # User configuration file
└── packages.config                # NuGet dependencies
```

## Technologies Used

- **[Excel-DNA](https://excel-dna.net/)**: High-performance Excel add-in framework for .NET
- **.NET Framework 4.7.2**: Core runtime
- **Microsoft.AnalysisServices**: Power BI connectivity and DAX query execution
- **Excel Interop**: Excel object model integration

## License

### Open Source License

CubeConnector is licensed under the **GNU General Public License v3.0 (GPLv3)**.

This means you are free to:
- ✅ Use the software for any purpose
- ✅ Study and modify the source code
- ✅ Distribute copies of the software
- ✅ Distribute modified versions

**Under these conditions:**
- 📋 You must disclose the source code when you distribute the software
- 📋 You must license your modifications under GPLv3
- 📋 You must include the original copyright and license notices
- 📋 You must state significant changes made to the software

See the [LICENSE](LICENSE) file for the full license text.

### Enterprise Licensing

For organizations that require:
- Proprietary modifications without source code disclosure
- Commercial licensing without GPLv3 obligations
- Custom support and service level agreements
- White-label or OEM distribution rights

**Custom enterprise licenses are available.** Contact the project maintainers to discuss enterprise licensing options that fit your organization's needs.

### Third-Party Dependencies

CubeConnector depends on the following open-source libraries:

- **Excel-DNA** (v1.9.0) - Licensed under [zlib License](https://github.com/Excel-DNA/ExcelDna/blob/master/LICENSE.txt)
  - Permissive license compatible with GPLv3
  - No additional restrictions

All dependencies are compatible with the GPLv3 license.

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## Security

For security concerns or vulnerability reports, please see [SECURITY.md](SECURITY.md).

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history and release notes.

## Acknowledgments

- Built with [Excel-DNA](https://excel-dna.net/) by Govert van Drimmelen
- Icon: [3d icons created by Freepik - Flaticon](https://www.flaticon.com/free-icons/3d)

## Support

- **Issues**: [GitHub Issues](../../issues)
- **Discussions**: [GitHub Discussions](../../discussions)
- **Documentation**: [Wiki](../../wiki)

---

**Compatible with Power BI Pro, Premium, and Fabric workspaces.**
