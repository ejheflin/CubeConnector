# Security Policy

## Supported Versions

We release patches for security vulnerabilities for the following versions:

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |
| < 1.0   | :x:                |

## Reporting a Vulnerability

We take the security of CubeConnector seriously. If you have discovered a security vulnerability, we appreciate your help in disclosing it to us in a responsible manner.

### Please Do Not

- **Do not** open a public GitHub issue for security vulnerabilities
- **Do not** disclose the vulnerability publicly until it has been addressed
- **Do not** exploit the vulnerability beyond what is necessary to demonstrate the issue

### How to Report

Please report security vulnerabilities by emailing the project maintainer(s) directly or using GitHub's private security advisory feature:

1. **GitHub Security Advisories** (Preferred):
   - Navigate to the repository's Security tab
   - Click "Report a vulnerability"
   - Fill out the vulnerability details
   - Submit the report

2. **Direct Email**:
   - Send details to: [maintainer-email@example.com]
   - Use subject line: "CubeConnector Security Vulnerability"

### What to Include

Please include the following information in your report:

- **Description**: Clear description of the vulnerability
- **Impact**: What an attacker could achieve
- **Steps to reproduce**: Detailed steps to reproduce the issue
- **Affected versions**: Which versions are affected
- **Proof of concept**: Code or commands demonstrating the vulnerability (if applicable)
- **Suggested fix**: If you have ideas on how to address the issue
- **Your contact information**: So we can follow up with questions

### Example Report

```
Subject: CubeConnector Security Vulnerability - DAX Injection

Description:
The DAX query builder does not properly sanitize user input, allowing
for DAX injection attacks through the List filter type.

Impact:
An attacker could execute arbitrary DAX queries against the Power BI
dataset, potentially accessing unauthorized data or causing performance
issues.

Steps to Reproduce:
1. Create a function with a List filter parameter
2. Input: '); EVALUATE Users; //
3. The resulting DAX query includes the injected EVALUATE statement

Affected Versions:
1.0.0 and earlier

Proof of Concept:
[Attach code or screenshots]

Suggested Fix:
Implement parameterized queries or proper input validation/escaping
before constructing DAX queries.

Contact:
John Doe - john@example.com
```

## Response Timeline

- **Acknowledgment**: We will acknowledge receipt of your vulnerability report within 48 hours
- **Initial assessment**: We will provide an initial assessment within 5 business days
- **Status updates**: We will keep you informed of our progress
- **Resolution**: We aim to address critical vulnerabilities within 30 days
- **Public disclosure**: Once a fix is available, we will coordinate public disclosure with you

## Security Best Practices for Users

### Configuration Security

1. **Protect Configuration Files**:
   - Never commit `CubeConnectorConfig.json` with real credentials to public repositories
   - Use file system permissions to restrict access to configuration files
   - Store configuration files in secure locations

2. **Power BI Security**:
   - Use principle of least privilege for Power BI workspace access
   - Regularly audit workspace membership
   - Enable Multi-Factor Authentication (MFA) for all Power BI accounts
   - Review dataset permissions regularly

3. **Azure AD Security**:
   - Use dedicated service accounts for automated scenarios
   - Implement Conditional Access policies
   - Monitor sign-in logs for suspicious activity
   - Rotate credentials regularly

### Excel Workbook Security

1. **Protect Workbooks**:
   - Use Excel's workbook protection features
   - Distribute workbooks via secure channels
   - Be cautious when opening workbooks from untrusted sources

2. **Macro Security**:
   - Keep macro security settings at recommended levels
   - Only enable the CubeConnector add-in from trusted sources
   - Verify add-in signatures (when available)

3. **Data Handling**:
   - Be aware that cache data is stored in a hidden worksheet
   - Sensitive data may remain in workbook even after formulas are removed
   - Clear cache before sharing workbooks containing sensitive data

### Network Security

1. **Firewall Configuration**:
   - Ensure your firewall allows outbound connections to:
     - `*.powerbi.com`
     - `*.analysis.windows.net`
     - `login.microsoftonline.com`

2. **TLS/SSL**:
   - Connections to Power BI use TLS encryption
   - Ensure your system supports modern TLS versions (1.2+)

3. **Proxy Considerations**:
   - Configure Excel to use your corporate proxy if required
   - SSL inspection may interfere with authentication

## Known Security Considerations

### DAX Query Execution

- The add-in executes DAX queries against your Power BI datasets
- Queries are constructed dynamically based on user input
- Input validation is performed to prevent injection attacks
- Users should still audit queries for sensitive data access

### Caching Behavior

- Query results are cached in a hidden Excel worksheet
- Cache is not encrypted within the workbook
- Sensitive data persists in cache until explicitly refreshed or cleared
- Consider data sensitivity when using caching features

### Authentication Token Storage

- Azure AD tokens are managed by Excel's authentication framework
- Tokens are stored in Windows Credential Manager
- Tokens have a limited lifetime and are automatically refreshed
- Users should sign out when finished on shared computers

### Code Execution

- The add-in runs with the same permissions as Excel
- Add-in code has access to the Excel object model
- Excel macros may interact with the add-in
- Only install add-ins from trusted sources

## Security Hardening Recommendations

### For Individual Users

1. Keep Excel and Windows updated with latest security patches
2. Use strong, unique passwords for Power BI accounts
3. Enable Multi-Factor Authentication on your Azure AD account
4. Review Power BI workspace access permissions regularly
5. Clear cache before sharing workbooks externally
6. Disable the add-in when not actively using it

### For Organizations

1. **Access Controls**:
   - Implement Role-Based Access Control (RBAC) in Power BI
   - Use Row-Level Security (RLS) in datasets
   - Restrict workspace access to authorized users only

2. **Monitoring**:
   - Enable audit logging in Power BI
   - Monitor for unusual query patterns
   - Track add-in usage across organization
   - Alert on authentication failures

3. **Deployment**:
   - Test add-in in isolated environment first
   - Use code signing for internal distributions
   - Implement application whitelisting
   - Consider using App-V or similar containerization

4. **Data Governance**:
   - Classify data sensitivity levels
   - Apply appropriate controls based on classification
   - Implement Data Loss Prevention (DLP) policies
   - Regular security audits

## Vulnerability Disclosure Policy

Once a security vulnerability has been fixed:

1. We will release a patched version
2. We will publish a security advisory with:
   - Description of the vulnerability
   - Affected versions
   - Remediation steps
   - Credit to the reporter (if desired)
3. We will update this SECURITY.md file if needed
4. We will notify users through:
   - GitHub releases
   - Repository README
   - Discussion forums (if applicable)

## Security Hall of Fame

We recognize and thank the following security researchers who have responsibly disclosed vulnerabilities:

- *No vulnerabilities reported yet*

If you report a valid security vulnerability, you may be listed here (with your permission).

## Contact

For security-related questions that are not vulnerability reports:
- Open a [GitHub Discussion](../../discussions)
- Tag with "security" label

For urgent security matters:
- Use the vulnerability reporting process above

## Additional Resources

- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
- [Microsoft Security Response Center](https://www.microsoft.com/en-us/msrc)
- [Power BI Security Whitepaper](https://docs.microsoft.com/power-bi/guidance/whitepaper-powerbi-security)
- [Azure AD Security](https://docs.microsoft.com/azure/active-directory/fundamentals/security-operations-introduction)

---

**Last Updated**: 2026-01-16
