# AD Export Scripts - Troubleshooting Guide

## Common Issues and Solutions

### ✅ FIXED: "Key cannot be null" error in CSV conversion

**Cause:** The CSV contained rows with null or empty DistinguishedName values, or the DN format was invalid.

**Solution:** This has been fixed with proper validation. The script now:
- Validates DNs before processing
- Skips rows with empty or null DNs
- Shows diagnostic information about the CSV structure
- Handles malformed DN components gracefully

**To test the fix:**
```powershell
.\Test-CSVConversion.ps1
```

**To properly export AD to CSV:**
```powershell
# Use the provided export script for best results
.\Export-ADToCSV.ps1 -OutputFile "ADObjects.csv"

# Then convert
.\Convert-ADCsvToMarkdown.ps1 -CsvFile "ADObjects.csv" -OutputFile "structure.md"
```

**CSV Requirements:**
- Must have `DistinguishedName` (or `DN`) column
- Must have `Name` column
- Should have `ObjectClass` (or `Type`) column
- DNs must be properly formatted: `CN=Name,OU=Container,DC=domain,DC=com`

---

### ✅ FIXED: "Invalid path warnings" in CSV conversion

**Cause:** The DN parser wasn't handling domain components (DC=) correctly, causing empty paths for domain roots and containers.

**Solution:** This has been fixed. The script now:
- Properly parses DC components and creates domain root entries
- Combines multiple DC parts into a single domain name (e.g., DC=corp,DC=lab becomes "corp.lab")
- Handles all component types: DC (domain), OU (organizational unit), CN (container/object)
- Provides detailed warnings with the actual DN when issues occur

**To test the DN parsing:**
```powershell
# Run the DN parsing test
.\Test-DNParsing.ps1

# Convert with verbose output to see what's being processed
.\Convert-ADCsvToMarkdown.ps1 -CsvFile "ADObjects.csv" -OutputFile "structure.md" -Verbose
```

**Expected behavior:**
- Domain roots (DC=domain,DC=com) appear as single entries like "domain.com"
- All containers and OUs maintain proper hierarchy
- Service accounts and system containers are properly nested

---

### ✅ FIXED: "A parameter with the name 'Verbose' was defined multiple times"

**Cause:** The script previously defined `-Verbose` as a custom parameter, but this is already a built-in PowerShell common parameter when using `[CmdletBinding()]`.

**Solution:** This has been fixed. The script now uses the built-in `-Verbose` parameter correctly.

**Usage:**
```powershell
# Use -Verbose to see detailed processing information
.\Export-ADStructureToMarkdown.ps1 -Verbose -OutputFile "output.md"
```

---

### Issue: "Active Directory module is not installed"

**Solution for Windows 10/11:**
```powershell
# Run as Administrator
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

**Solution for Windows Server:**
```powershell
# Run as Administrator
Install-WindowsFeature -Name RSAT-AD-PowerShell
```

---

### Issue: "Cannot find object with DN"

**Causes:**
1. Invalid Distinguished Name format
2. Object doesn't exist
3. No permissions to read object

**Solution:**
```powershell
# Test if you can access the domain
Get-ADDomain

# Test specific SearchBase
Get-ADObject -Identity "OU=YourOU,DC=domain,DC=com"

# If access denied, check permissions
whoami /groups
```

---

### Issue: "Access is denied" errors

**Solution:**
```powershell
# Option 1: Run as a user with appropriate permissions
runas /user:DOMAIN\AdminAccount powershell.exe

# Option 2: Use credentials
$cred = Get-Credential
Get-ADObject -Filter * -Credential $cred
```

---

### Issue: Script runs but no output

**Common causes:**
1. Empty OUs (use `-IncludeEmptyOUs`)
2. All objects are disabled (use `-IncludeDisabled`)
3. SearchBase has no child objects

**Debug steps:**
```powershell
# Run with verbose to see what's happening
.\Export-ADStructureToMarkdown.ps1 -Verbose

# Test with minimal depth first
.\Export-ADStructureToMarkdown.ps1 -MaxDepth 1

# Check if SearchBase has objects
Get-ADObject -SearchBase "YourDN" -SearchScope OneLevel -Filter *
```

---

### Issue: "The term 'Get-ADObject' is not recognized"

**Solution:**
```powershell
# Import the module explicitly
Import-Module ActiveDirectory

# If still not working, verify RSAT installation
Get-WindowsCapability -Name RSAT* -Online
```

---

### Issue: Exported markdown is too large

**Solutions:**
```powershell
# Limit depth
.\Export-ADStructureToMarkdown.ps1 -MaxDepth 3 -OutputFile "limited.md"

# Export specific OU only
.\Export-ADStructureToMarkdown.ps1 -SearchBase "OU=Specific,DC=domain,DC=com"

# Exclude disabled and system objects
.\Export-ADStructureToMarkdown.ps1 -ExcludeSystemContainers -OutputFile "clean.md"
```

---

### Issue: Service accounts not showing

**Check:**
```powershell
# Service accounts are included by default now
# Make sure you're NOT using -ExcludeServiceAccounts

# List service accounts directly
Get-ADObject -Filter {objectClass -eq "msDS-ManagedServiceAccount"} 
Get-ADObject -Filter {objectClass -eq "msDS-GroupManagedServiceAccount"}

# Check the Managed Service Accounts container
Get-ADObject -SearchBase "CN=Managed Service Accounts,DC=domain,DC=com" -Filter *
```

---

### Issue: Computer names show with $ at the end

**Status:** This is fixed in the current version. Computer names are automatically cleaned.

If still seeing $:
```powershell
# Check script version - should be 1.0 or later
# Re-download the latest version of the script
```

---

## Performance Tips

### For Large Domains

```powershell
# 1. Start with shallow depth
.\Export-ADStructureToMarkdown.ps1 -MaxDepth 2 -OutputFile "test.md"

# 2. Export specific OUs separately
$ous = @("OU=Corp,DC=domain,DC=com", "OU=Branch,DC=domain,DC=com")
foreach ($ou in $ous) {
    $name = ($ou -split ',')[0] -replace 'OU=',''
    .\Export-ADStructureToMarkdown.ps1 -SearchBase $ou -OutputFile "$name.md"
}

# 3. Exclude unnecessary objects
.\Export-ADStructureToMarkdown.ps1 `
    -ExcludeSystemContainers `
    -OutputFile "optimized.md"
```

---

## Testing Your Setup

Run the test script to verify everything works:
```powershell
.\Test-ADExportScripts.ps1
```

This will:
1. Check if AD module is installed
2. Show example commands
3. Optionally run a quick test export

---

## Getting Help

All scripts include built-in help:
```powershell
Get-Help .\Export-ADStructureToMarkdown.ps1 -Full
Get-Help .\Export-ADStructure-Simple.ps1 -Examples
Get-Help .\Convert-ADCsvToMarkdown.ps1 -Parameter OutputFile
```

---

## Quick Validation

Test if everything is working:
```powershell
# Should return your domain info
(Get-ADDomain).DNSRoot

# Should list some objects
Get-ADObject -Filter * -ResultSetSize 5

# Should not throw any errors
.\Export-ADStructureToMarkdown.ps1 -MaxDepth 1
```