# AD2Markdown

A set of vibe-coded PowerShell scripts that collect data from an Active Directory environment and output an unordered list in Markdown format that can be used with md2ADUC.


### Export-ADStructureToMarkdown.ps1 (Full-Featured)
The comprehensive script with extensive options:
- Automatic domain detection or specify custom SearchBase
- Filtering options:
  -ExcludeSystemContainers - Skip System, Builtin, etc.
  -IncludeDisabled - Include/exclude disabled accounts
  -IncludeContacts - Include contact objects
  -ExcludeServiceAccounts - Exclude managed service accounts
- Depth control with -MaxDepth parameter
- Verbose logging for troubleshooting
- Properly formats object types with [user], [computer], [group] markers

Example usage:
```powershell
# Export entire domain, excluding system containers
.\Export-ADStructureToMarkdown.ps1 -ExcludeSystemContainers -OutputFile "domain.md"

# Export specific OU
.\Export-ADStructureToMarkdown.ps1 -SearchBase "OU=Corporate,DC=contoso,DC=com" -OutputFile "corporate.md"
```

### Export-ADStructure-Simple.ps1 (Lightweight)
Quick and easy script for rapid exports:
- Minimal parameters
- Automatic system container filtering
- Fast execution
- Perfect for quick documentation

Example usage:
```powershell
.\Export-ADStructure-Simple.ps1 -OutFile "quick-export.md"
```

### Convert-ADCsvToMarkdown.ps1 (CSV Converter)
For offline processing or when you don't have AD access:
- Processes CSV exports from AD
- Rebuilds hierarchy from Distinguished Names
- Maintains object relationships

Example usage:
```powershell
# First export AD to CSV
Get-ADObject -Filter * -Properties Name,ObjectClass,Enabled | Export-Csv "AD.csv"

# Then convert
.\Convert-ADCsvToMarkdown.ps1 -CsvFile "AD.csv" -OutputFile "structure.md"
```

## Complete Workflow:
1. Run PowerShell script on domain-joined machine:
```powershell
.\Export-ADStructureToMarkdown.ps1 -OutputFile "MyDomain.md"
```
2. Open the HTML visualizer and load the markdown file
3. Export as PNG for PowerPoint presentations

Key Features:
- Intelligent Object Recognition: Automatically identifies users, computers, groups, OUs, and other AD objects
- Clean Formatting: Removes computer name trailing $, organizes objects by type
- Recursive Traversal: Handles nested OUs to any depth
- Error Handling: Continues on access denied errors, logs issues
- Performance Optimized: Uses efficient LDAP queries

The scripts handle real-world AD complexities like:
- Deeply nested OU structures
- Mixed object types in containers
- System containers and built-in objects
- Disabled accounts
- Service accounts (MSAs/gMSAs)
- Various container types beyond just OUs