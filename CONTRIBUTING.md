# Contributing to CubeConnector

Thank you for your interest in contributing to CubeConnector! This document provides guidelines and information for contributors.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [How to Contribute](#how-to-contribute)
- [Coding Standards](#coding-standards)
- [Testing Guidelines](#testing-guidelines)
- [Pull Request Process](#pull-request-process)
- [Issue Reporting](#issue-reporting)

## Code of Conduct

This project adheres to a code of professional conduct. By participating, you are expected to:

- Be respectful and inclusive
- Welcome newcomers and help them get started
- Focus on constructive criticism
- Accept feedback gracefully
- Prioritize the community's best interests

## Getting Started

### Prerequisites

Before you begin, ensure you have:

- Visual Studio 2019 or later
- .NET Framework 4.7.2 SDK or higher
- Microsoft Excel (Windows)
- Git for version control
- A Power BI Premium or Premium Per User workspace (for testing)

### Development Setup

1. **Fork and Clone the Repository**
   ```bash
   git fork https://github.com/[owner]/CubeConnector
   cd CubeConnector
   ```

2. **Restore NuGet Packages**
   ```bash
   nuget restore
   ```
   Or in Visual Studio: Right-click solution → Restore NuGet Packages

3. **Build the Project**
   ```bash
   msbuild CubeConnector.sln /p:Configuration=Debug
   ```
   Or in Visual Studio: Build → Build Solution (Ctrl+Shift+B)

4. **Create Test Configuration**
   - Copy `CubeConnectorConfig.example.json` to `CubeConnectorConfig.json`
   - Fill in your Power BI workspace and dataset details
   - **Never commit your personal configuration file**

5. **Load the Add-in in Excel**
   - Open Excel
   - File → Options → Add-ins → Manage Excel Add-ins → Browse
   - Navigate to `bin\Debug\CubeConnector.xll`
   - Click OK

## How to Contribute

### Types of Contributions

We welcome several types of contributions:

- **Bug Fixes**: Fix issues reported in GitHub Issues
- **Features**: Implement new functionality
- **Documentation**: Improve README, wiki, or code comments
- **Testing**: Add unit tests or integration tests
- **Performance**: Optimize existing code
- **Refactoring**: Improve code structure without changing behavior

### Finding Work

- Check the [Issues](../../issues) page for open tasks
- Look for issues labeled `good first issue` for beginner-friendly tasks
- Look for issues labeled `help wanted` for areas where we need assistance
- Propose new features in [Discussions](../../discussions) before implementing

### Before You Start

1. **Check for existing work**: Search issues and pull requests to avoid duplicate effort
2. **Discuss major changes**: Open an issue to discuss significant changes before coding
3. **Create a branch**: Never work directly on `main`
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/issue-number-description
   ```

## Coding Standards

### C# Style Guidelines

Follow these conventions for consistency:

#### Naming Conventions

- **PascalCase** for class names, method names, properties, and public fields
  ```csharp
  public class CacheManager
  public void RefreshCache()
  public string TenantId { get; set; }
  ```

- **camelCase** for local variables and parameters
  ```csharp
  string cacheKey = GenerateKey();
  public void ProcessData(string configPath, int maxRetries)
  ```

- **_camelCase** for private fields (optional)
  ```csharp
  private string _connectionString;
  ```

- **UPPER_CASE** for constants
  ```csharp
  private const string CACHE_SHEET_NAME = "__CubeConnector_Cache__";
  ```

#### Code Structure

- **One class per file** (except for small helper/model classes)
- **Organize using statements**: Remove unused, sort alphabetically
- **Use explicit access modifiers**: Always specify `public`, `private`, `internal`
- **Prefer readonly when possible**: Mark fields as `readonly` if they don't change

#### Comments and Documentation

- **Use XML documentation** for all public APIs
  ```csharp
  /// <summary>
  /// Retrieves a value from the cache using the specified key.
  /// </summary>
  /// <param name="cacheKey">The unique cache key.</param>
  /// <returns>The cached value or #REFRESH if not found.</returns>
  public static object Lookup(string cacheKey)
  ```

- **Explain "why" not "what"**: Code should be self-documenting; comments explain reasoning
  ```csharp
  // BAD: Increment counter
  counter++;

  // GOOD: Skip the header row
  counter++;
  ```

- **Document complex algorithms**: Add inline comments for non-obvious logic

#### Error Handling

- **Use specific exception types**: Don't catch `Exception` unless necessary
- **Provide context in error messages**: Include relevant details
  ```csharp
  throw new ArgumentException($"Invalid filter type '{filterType}' for parameter '{paramName}'");
  ```
- **Clean up resources**: Use `using` statements for IDisposable objects
- **Don't swallow exceptions**: Log or rethrow

### JSON Configuration

- Use consistent formatting (2-space indentation)
- Include comments when possible (use JSON5 or document separately)
- Validate against schema

## Testing Guidelines

### Manual Testing

Before submitting a pull request:

1. **Test the basic workflow**:
   - Load the add-in in Excel
   - Create a test function configuration
   - Call the function in a cell
   - Verify results are correct
   - Test cache refresh
   - Test drillthrough features

2. **Test edge cases**:
   - Empty parameters
   - Invalid parameters
   - Large datasets
   - Network errors
   - Authentication failures

3. **Test across Excel versions** (if possible):
   - Excel 2016
   - Excel 2019
   - Excel 365

### Automated Testing

While the project doesn't currently have a comprehensive test suite, we welcome contributions that add:

- Unit tests for core logic (DAXQueryBuilder, CacheManager, etc.)
- Integration tests for Power BI connectivity
- Mock-based tests for Excel interop

Use a testing framework like:
- xUnit
- NUnit
- MSTest

## Pull Request Process

### Before Submitting

1. **Update documentation**: If you changed functionality, update README.md
2. **Test thoroughly**: Verify your changes work as expected
3. **Follow coding standards**: Ensure code matches project style
4. **Keep commits clean**: Use clear, descriptive commit messages
   ```
   Add support for numeric range filters

   - Extended DAXQueryBuilder to handle numeric ranges
   - Added unit tests for numeric filter generation
   - Updated configuration documentation
   ```

5. **Update CHANGELOG.md**: Add your changes under "Unreleased"

### Submitting a Pull Request

1. **Push your branch**:
   ```bash
   git push origin feature/your-feature-name
   ```

2. **Create pull request on GitHub**:
   - Provide a clear title and description
   - Reference any related issues (e.g., "Fixes #123")
   - Describe what changed and why
   - Include screenshots/examples if applicable

3. **Pull request template**:
   ```markdown
   ## Description
   Brief description of changes

   ## Related Issues
   Fixes #123

   ## Changes Made
   - Added feature X
   - Fixed bug Y
   - Updated documentation Z

   ## Testing Done
   - Tested with Excel 2019
   - Verified cache refresh works
   - Tested with 10,000+ row datasets

   ## Screenshots (if applicable)
   [Attach images]

   ## Checklist
   - [ ] Code follows project style guidelines
   - [ ] Documentation updated
   - [ ] Manual testing completed
   - [ ] CHANGELOG.md updated
   ```

### Review Process

1. **Automated checks**: CI/CD may run builds and tests (if configured)
2. **Code review**: Maintainers will review your code
3. **Feedback**: Address review comments by pushing new commits
4. **Approval**: Once approved, a maintainer will merge

### After Merge

- Your contribution will appear in the next release
- You'll be credited in the release notes
- Thank you for contributing!

## Issue Reporting

### Before Creating an Issue

- **Search existing issues**: Your issue may already be reported
- **Try the latest version**: The bug may already be fixed
- **Verify it's reproducible**: Ensure the issue is consistent

### Creating a Bug Report

Include:

- **Title**: Clear, concise description
- **Environment**:
  - Excel version
  - Windows version
  - CubeConnector version
  - .NET Framework version
- **Steps to reproduce**:
  1. Step one
  2. Step two
  3. Step three
- **Expected behavior**: What should happen
- **Actual behavior**: What actually happens
- **Screenshots/Logs**: If applicable
- **Configuration**: Sanitized JSON config (remove sensitive IDs)

### Feature Requests

Include:

- **Problem statement**: What problem does this solve?
- **Proposed solution**: How would you implement it?
- **Alternatives considered**: What other approaches did you think about?
- **Use cases**: Who would benefit from this feature?

## Development Tips

### Debugging the Add-in

1. **Attach debugger to Excel**:
   - In Visual Studio: Debug → Attach to Process → Excel.exe
   - Set breakpoints in your code
   - Trigger the functionality in Excel

2. **Use MessageBox for quick debugging**:
   ```csharp
   System.Windows.Forms.MessageBox.Show($"Value: {variable}");
   ```

3. **Enable Excel-DNA diagnostic logging**:
   - Add `<ExcelDna Diagnostic="true" />` to .dna file

### Working with Excel Interop

- **Always release COM objects**: Use `Marshal.ReleaseComObject()` to prevent Excel processes from hanging
- **Be careful with range indexing**: Excel ranges are 1-based, not 0-based
- **Test with calculation mode**: Test with automatic and manual calculation modes

### Power BI Connection Tips

- **Use premium workspace**: XMLA endpoints require Premium capacity
- **Check firewall**: Ensure outbound connections to `*.powerbi.com` and `*.analysis.windows.net`
- **Refresh authentication**: Tokens expire; test with long-running sessions

## Questions?

- **General questions**: Use [GitHub Discussions](../../discussions)
- **Bugs**: Create an [Issue](../../issues)
- **Security concerns**: See [SECURITY.md](SECURITY.md)

## License

By contributing, you agree that your contributions will be licensed under the same license as the project.

---

Thank you for contributing to CubeConnector!
