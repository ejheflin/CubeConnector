# Changelog

All notable changes to CubeConnector will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Removed
- Retired the standalone web-based Config Editor (GitHub Pages `docs/index.html`). Configuration now lives in the in-Excel task pane (**Data ▸ CubeConnector ▸ Manage Formulas**), which authenticates, enumerates models/measures, and registers formulas directly.

### Added
- Initial public release
- Dynamic function registration from JSON configuration
- Intelligent caching system with hidden worksheet storage
- Drill-to-details functionality via context menu
- Drill-to-pivot functionality via context menu
- Custom Excel ribbon integration in Data tab
- Support for multiple filter types (List, RangeStart, RangeEnd, Single)
- Support for multiple data types (text, date, numeric)
- DAX query builder with dynamic CALCULATE generation
- Power BI XMLA endpoint connectivity
- Azure AD authentication support
- Cache refresh via ribbon button and context menu
- Query pool analyzer for optimization
- Comprehensive documentation and examples
- **Support for UDFs in arithmetic expressions** (e.g., `=MyFunc(1)+50`, `=MyFunc(1)*2`)
- **Support for multiple UDFs per cell** (e.g., `=MyFunc(1)+MyFunc(2)`) - each UDF cached separately
- **Force Refresh option** when no cells need refreshing - clears and rebuilds entire cache
- **Protection for drill features** - prevents drillthrough on cells with multiple UDFs
- **Enhanced error detection** - refresh finds cells with #VALUE! in addition to #REFRESH
- **COM Variant type conversion** for proper Excel marshaling and calculation support

### Changed
- UDF functions now use `IsMacroType=false` to enable arithmetic expression support
- Cache manager uses static Excel Application reference instead of per-call access
- Refresh manager detects and parses multiple UDF calls within a single formula
- Error display: standalone UDFs show #REFRESH, UDFs in expressions show #VALUE! (both resolved on refresh)

### Deprecated
- N/A (initial release)

### Removed
- N/A (initial release)

### Fixed
- **Critical**: UDFs now work correctly in arithmetic expressions without #VALUE! errors
- **Critical**: Multiple UDFs per cell are properly identified and cached during refresh
- Type marshaling issues that prevented calculated results in Excel formulas

### Security
- N/A (initial release)

## [1.0.0] - YYYY-MM-DD

### Added
- Core functionality for querying Power BI datasets from Excel
- Configuration-based UDF generation
- Cache management system
- Drillthrough capabilities
- Excel ribbon UI

---

## Version History Format

### [Version] - YYYY-MM-DD

#### Added
New features and capabilities

#### Changed
Changes to existing functionality

#### Deprecated
Features that will be removed in future versions

#### Removed
Features that have been removed

#### Fixed
Bug fixes

#### Security
Security-related changes and fixes

---

## Release Notes

### How to Read Version Numbers

Given a version number MAJOR.MINOR.PATCH (e.g., 1.2.3):

- **MAJOR** version changes when incompatible API changes are made
- **MINOR** version changes when functionality is added in a backwards-compatible manner
- **PATCH** version changes when backwards-compatible bug fixes are made

### Upgrade Guide

When upgrading between versions, check the relevant sections above for:
- **Breaking changes** (marked in Changed or Removed sections)
- **New features** (marked in Added section)
- **Deprecation warnings** (marked in Deprecated section)

Always back up your workbooks and configurations before upgrading.

---

[Unreleased]: https://github.com/[owner]/CubeConnector/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/[owner]/CubeConnector/releases/tag/v1.0.0
