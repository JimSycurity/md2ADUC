# AD Structure Visualization Toolkit

A complete solution for visualizing and documenting Active Directory structures using markdown and an HTML-based ADUC simulator.

## Components

### 1. **aduc-simulator.html** - Web-based ADUC Interface Simulator
- Converts markdown lists to realistic Active Directory Users and Computers interface
- Exports high-resolution PNG images for PowerPoint presentations
- Save/load markdown files for reusable AD structure templates
- No installation required - runs entirely in browser

### 2. **Export-ADStructureToMarkdown.ps1** - Full-Featured AD Export Script
- Exports live AD structure to markdown format
- Comprehensive filtering options (disabled accounts, system containers, etc.)
- Service accounts (MSAs/gMSAs) included by default
- Configurable depth and scope
- Detailed logging and progress tracking

### 3. **Export-ADStructure-Simple.ps1** - Lightweight AD Export Script  
- Quick and simple AD structure export
- Minimal parameters for ease of use
- Perfect for quick documentation tasks

### 4. **Convert-ADCsvToMarkdown.ps1** - CSV to Markdown Converter
- Converts CSV exports to markdown format
- Useful for offline processing or archived data
- Works with standard AD export formats

## Quick Start Guide

### Visualizing an Existing AD Structure

1. **Export from Active Directory:**
   ```powershell
   # Full export with all options
   .\Export-ADStructureToMarkdown.ps1 -OutputFile "mydomain.md" -ExcludeSystemContainers
   
   # Quick export
   .\Export-ADStructure-Simple.ps1 -OutFile "structure.md"
   
   # Export specific OU
   .\Export-ADStructureToMarkdown.ps1 -SearchBase "OU=Corporate,DC=contoso,DC=com" -OutputFile "corporate.md"
   ```

2. **Open the Visualizer:**
   - Open `aduc-simulator.html` in any modern browser
   - Click "Open Markdown" and select your exported `.md` file
   - The structure will automatically render

3. **Export for Presentation:**
   - Click "Export as Image" to save as PNG
   - Insert the PNG into PowerPoint or documentation

### Creating Documentation from Scratch

1. **Open the Visualizer:**
   - Open `aduc-simulator.html` in your browser

2. **Write Your Structure:**
   ```markdown
   - contoso.com
     - Domain Controllers
       - DC01 [computer]
       - DC02 [computer]
     - Corporate
       - Users
         - John Doe [user]
       - Computers
         - DESKTOP-01 [computer]
   ```

3. **Save and Export:**
   - Click "Save Markdown" to save your structure
   - Click "Export as Image" for PowerPoint

## Markdown Format Reference

### Basic Structure
- Use `-` or `*` for list items
- Indent with 2 spaces for each level
- Add object types in brackets: `[type]`

### Supported Object Types
| Type | Marker | Icon |
|------|--------|------|
| Organizational Unit | (none) | 📁 |
| User | `[user]` | 👤 |
| Computer | `[computer]` | 💻 |
| Group | `[group]` | 👥 |
| Contact | `[contact]` | 📇 |
| Printer | `[printer]` | 🖨️ |
| Share | `[share]` | 📤 |
| Policy | `[policy]` | 📋 |
| Container | `[container]` | 📦 |

### Example Structure
```markdown
- contoso.com
  - Domain Controllers
    - DC01 [computer]
    - DC02 [computer]
  - Corporate
    - Finance
      - Accounting
        - Bob Smith [user]
        - AccountingTeam [group]
      - Payroll
        - PayrollPrinter [printer]
    - IT Department
      - Servers
        - WEB-SRV01 [computer]
        - SQL-SRV01 [computer]
      - Service Accounts
        - svc_backup [user]
        - svc_sql [user]
```

## PowerShell Script Examples

### Default Behavior
By default, the scripts include:
- All enabled user and computer accounts
- All groups and OUs
- **Service accounts (MSAs and gMSAs)** - these are critical infrastructure components
- The "Managed Service Accounts" container when present

The scripts exclude by default:
- Disabled accounts (use `-IncludeDisabled` to include)
- Contact objects (use `-IncludeContacts` to include)
- System containers like Builtin, System (use `-ExcludeSystemContainers` to exclude)

### Export Entire Domain
```powershell
.\Export-ADStructureToMarkdown.ps1 -OutputFile "FullDomain.md"
```

### Export with Filters
```powershell
# Exclude disabled accounts and system containers
.\Export-ADStructureToMarkdown.ps1 `
    -ExcludeSystemContainers `
    -OutputFile "CleanStructure.md"

# Include everything (disabled, contacts, but exclude service accounts)
.\Export-ADStructureToMarkdown.ps1 `
    -IncludeDisabled `
    -IncludeContacts `
    -ExcludeServiceAccounts `
    -OutputFile "NoServiceAccounts.md"

# Standard export (includes service accounts by default)
.\Export-ADStructureToMarkdown.ps1 `
    -OutputFile "Complete.md"
```

### Export Specific OU with Limited Depth
```powershell
.\Export-ADStructureToMarkdown.ps1 `
    -SearchBase "OU=BranchOffices,DC=contoso,DC=com" `
    -MaxDepth 3 `
    -OutputFile "BranchOffices.md"
```

### Working with CSV Exports
```powershell
# First, export AD to CSV
Get-ADObject -Filter * -Properties Name,ObjectClass,Enabled |
    Export-Csv -Path "ADExport.csv" -NoTypeInformation

# Then convert to markdown
.\Convert-ADCsvToMarkdown.ps1 -CsvFile "ADExport.csv" -OutputFile "FromCSV.md"
```

## Use Cases

### 1. Documentation
- Create visual AD structure documentation
- Document planned AD changes before implementation
- Archive AD structure snapshots for compliance

### 2. Planning & Design
- Design new AD structures before deployment
- Plan OU reorganizations
- Visualize proposed changes for approval

### 3. Training & Education
- Create training materials with realistic AD interfaces
- Build lab environment documentation
- Demonstrate AD concepts visually

### 4. Presentations
- Executive briefings on AD structure
- Security assessments and recommendations
- Migration planning presentations

### 5. Troubleshooting
- Visualize complex AD hierarchies
- Document problematic areas
- Compare before/after states

## Tips & Best Practices

1. **Organize by Function**: Group related objects together (all users in a Users OU, all computers in Computers OU)

2. **Use Descriptive Names**: Make OU names clear and purposeful

3. **Limit Depth**: Try to keep structures under 5-6 levels deep for better manageability

4. **Save Templates**: Create and save common structure templates for reuse

5. **Version Control**: Save markdown files in Git for tracking changes over time

6. **Regular Exports**: Schedule regular exports for documentation updates

## Troubleshooting

### PowerShell Scripts

**Issue**: "Active Directory module not found"
- Install RSAT tools or run on a domain controller
- For Windows 10/11: `Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0`

**Issue**: "Access Denied" errors
- Ensure you have appropriate AD read permissions
- Run PowerShell as administrator if needed

**Issue**: "Maximum depth reached" warnings
- Increase `-MaxDepth` parameter
- Check for circular references in AD

### HTML Visualizer

**Issue**: Export button doesn't work
- Ensure you've clicked "Render Tree" first
- Check browser console for errors
- Try a different browser (Chrome/Edge recommended)

**Issue**: Icons not displaying
- Ensure your browser supports Unicode emoji
- Try refreshing the page

## Security Considerations

1. **Sensitive Information**: Be aware that AD structures can reveal organizational information
2. **File Storage**: Store exported markdown files securely
3. **Sharing**: Review exports before sharing to ensure no sensitive data is included
4. **Permissions**: PowerShell scripts only require read access to AD

## Requirements

### For PowerShell Scripts
- Windows PowerShell 5.1 or PowerShell 7+
- Active Directory PowerShell module (RSAT)
- Domain-joined computer or appropriate credentials
- Read access to target AD objects

### For HTML Visualizer
- Modern web browser (Chrome, Edge, Firefox, Safari)
- JavaScript enabled
- No installation or internet connection required

## Support & Updates

This toolkit is provided as-is for AD documentation and visualization purposes. Feel free to modify the scripts and HTML to suit your specific needs.

## License

Free to use and modify for personal and commercial use.

---

*Created for AD administrators and architects who need professional visualization tools for documentation and presentations.*