<#
.SYNOPSIS
    Converts CSV export of AD objects to markdown format for AD Structure Visualizer.

.DESCRIPTION
    This script processes a CSV file containing AD object information and converts it
    to the markdown format used by the AD Structure Visualizer. Useful when you have
    CSV exports from AD or when you don't have direct AD access.

.PARAMETER CsvFile
    Path to the CSV file containing AD object data.
    Required columns: DistinguishedName, Name, ObjectClass

.PARAMETER OutputFile
    Path to save the markdown file.

.PARAMETER IncludeDisabled
    Include objects marked as disabled in the Enabled column.

.EXAMPLE
    .\Convert-ADCsvToMarkdown.ps1 -CsvFile "ADExport.csv" -OutputFile "structure.md"

.EXAMPLE
    # First export from AD to CSV:
    Get-ADObject -Filter * -Properties objectClass, Enabled | 
        Export-Csv -Path "ADObjects.csv" -NoTypeInformation
    
    # Then convert:
    .\Convert-ADCsvToMarkdown.ps1 -CsvFile "ADObjects.csv" -OutputFile "structure.md"

.NOTES
    The CSV should have at minimum these columns:
    - DistinguishedName (or DN)
    - Name
    - ObjectClass (or Type)
    
    Optional columns:
    - Enabled (for filtering disabled objects)
    - Description (can be included as comments)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvFile,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputFile,
    
    [switch]$IncludeDisabled
)

# Verify CSV exists
if (-not (Test-Path $CsvFile)) {
    Write-Error "CSV file not found: $CsvFile"
    exit 1
}

# Import CSV
Write-Host "Importing CSV file..." -ForegroundColor Cyan
$adObjects = Import-Csv -Path $CsvFile

# Show CSV info for debugging
Write-Host "CSV contains $($adObjects.Count) objects" -ForegroundColor Yellow
if ($adObjects.Count -gt 0) {
    Write-Host "Available columns: $($adObjects[0].PSObject.Properties.Name -join ', ')" -ForegroundColor Gray
    
    # Show sample of first object
    Write-Host "`nSample object:" -ForegroundColor Gray
    $adObjects[0] | Format-List | Out-String | Write-Host
}

# Verify required columns
$requiredColumns = @('Name')
$hasDistinguishedName = $false
$hasDN = $false

if ($adObjects[0].PSObject.Properties['DistinguishedName']) {
    $hasDistinguishedName = $true
} elseif ($adObjects[0].PSObject.Properties['DN']) {
    $hasDN = $true
} else {
    Write-Error "CSV must contain either 'DistinguishedName' or 'DN' column"
    exit 1
}

# Function to get DN from object
function Get-ObjectDN($obj) {
    if ($hasDistinguishedName) {
        return $obj.DistinguishedName
    } elseif ($hasDN) {
        return $obj.DN
    }
    return $null
}

# Function to get object class
function Get-ObjectClass($obj) {
    $rawClass = $null
    
    # Get the raw class value
    if ($obj.PSObject.Properties['ObjectClass']) {
        $rawClass = $obj.ObjectClass
    } elseif ($obj.PSObject.Properties['Type']) {
        $rawClass = $obj.Type
    } elseif ($obj.PSObject.Properties['objectCategory']) {
        # Try to infer from objectCategory
        if ($obj.objectCategory -like '*Person*') { return 'user' }
        if ($obj.objectCategory -like '*Computer*') { return 'computer' }
        if ($obj.objectCategory -like '*Group*') { return 'group' }
        if ($obj.objectCategory -like '*Organizational-Unit*') { return 'organizationalUnit' }
        return 'unknown'
    }
    
    if ($null -eq $rawClass) {
        return 'unknown'
    }

    if ($rawClass -isnot [string]) {
        $rawClass = [string]$rawClass
    }

    $rawClass = $rawClass.Trim()

    if ([string]::IsNullOrWhiteSpace($rawClass)) {
        return 'unknown'
    }

    # Normalize the object class
    switch -Wildcard ($rawClass) {
        'user' { return 'user' }
        'computer' { return 'computer' }
        'group' { return 'group' }
        'contact' { return 'contact' }
        'organizationalUnit' { return 'organizationalUnit' }
        'container' { return 'container' }
        'domainDNS' { return 'domainDNS' }
        'domain' { return 'domainDNS' }
        'printQueue' { return 'printer' }
        'volume' { return 'share' }
        'groupPolicyContainer' { return 'policy' }
        'msDS-ManagedServiceAccount' { return 'computer' }
        'msDS-GroupManagedServiceAccount' { return 'computer' }
        'msExchSystemObjects*' { return 'container' }
        'msImaging-PSPs' { return 'container' }
        'rpcContainer' { return 'container' }
        'msDS-Device*' { return 'computer' }
        '*ServiceAccount*' { return 'computer' }
        default { 
            # Default based on common patterns
            if ($rawClass -like '*computer*') { return 'computer' }
            if ($rawClass -like '*user*') { return 'user' }
            if ($rawClass -like '*group*') { return 'group' }
            return $rawClass 
        }
    }
}

# Function to parse DN into components
function Parse-DN {
    param([string]$DN)
    
    $components = @()
    
    if ([string]::IsNullOrWhiteSpace($DN)) {
        return $components
    }
    
    $parts = $DN -split ',(?=\w+=)'
    
    if ($parts.Count -eq 0) {
        return $components
    }
    
    # Identify trailing DC components (domain root) and leave other DC components alone
    $domainStartIndex = $parts.Count
    for ($i = $parts.Count - 1; $i -ge 0; $i--) {
        $part = $parts[$i].Trim()
        if ($part -match '^DC=(.+)$') {
            $domainStartIndex = $i
            continue
        }
        break
    }
    
    if ($domainStartIndex -lt $parts.Count) {
        $domainNameParts = @()
        for ($j = $domainStartIndex; $j -lt $parts.Count; $j++) {
            $dcPart = $parts[$j].Trim()
            if ($dcPart -match '^DC=(.+)$') {
                $value = $Matches[1].Trim()
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $domainNameParts += $value
                }
            }
        }
        
        if ($domainNameParts.Count -gt 0) {
            $domainName = $domainNameParts -join '.'
            $components += @{
                Type = 'DC'
                Name = $domainName
                Level = 0
                IsDomainRoot = $true
            }
        }
    }
    
    $hierarchyParts = @()
    for ($k = 0; $k -lt $domainStartIndex; $k++) {
        $hierarchyParts += $parts[$k]
    }
    
    $level = if ($components.Count -gt 0) { 1 } else { 0 }
    for ($i = $hierarchyParts.Count - 1; $i -ge 0; $i--) {
        $part = $hierarchyParts[$i]
        
        if ([string]::IsNullOrWhiteSpace($part)) {
            continue
        }
        
        if ($part -match '^([A-Za-z]+)=(.+)$') {
            $prefix = $Matches[1].ToUpperInvariant()
            $name = $Matches[2].Trim()
            
            if ([string]::IsNullOrWhiteSpace($name)) {
                continue
            }
            
            $components += @{
                Type = $prefix
                Name = $name
                Level = $level
                IsDomainRoot = $false
            }
            $level++
        }
    }
    
    # Always return an array so callers can rely on indexing/count semantics
    return ,$components
}

# Function to get type marker
function Get-TypeMarker($objectClass) {
    # Only these specific types get markers in the visualization tool
    # OUs and containers don't get markers
    switch -Wildcard ($objectClass) {
        'user' { return ' [user]' }
        'computer' { return ' [computer]' }
        'group' { return ' [group]' }
        'contact' { return ' [contact]' }
        'printer' { return ' [printer]' }
        'share' { return ' [share]' }
        'policy' { return ' [policy]' }
        # These don't get type markers in the output
        'organizationalUnit' { return '' }
        'container' { return '' }
        'domain' { return '' }
        'domainDNS' { return '' }
        '*ServiceAccount*' { return ' [computer]' }  # Managed service accounts show as computers
        default { return '' }  # Unknown types get no marker
    }
}

# Build tree structure
Write-Host "Building tree structure..." -ForegroundColor Cyan
$tree = @{}
$processedPaths = @{}
$normalizedTypeCounts = @{}

foreach ($obj in $adObjects) {
    # Skip disabled if requested
    if (-not $IncludeDisabled -and $obj.PSObject.Properties['Enabled']) {
        if ($obj.Enabled -eq 'False' -or $obj.Enabled -eq $false) {
            continue
        }
    }
    
    $dn = Get-ObjectDN $obj
    
    # Skip if DN is null or empty
    if ([string]::IsNullOrWhiteSpace($dn)) {
        Write-Warning "Skipping object with empty DN: $($obj.Name)"
        continue
    }
    
    Write-Verbose "Processing: $dn"
    
    $objectClass = Get-ObjectClass $obj
    
    if ([string]::IsNullOrWhiteSpace($objectClass)) {
        $objectClass = 'unknown'
    }

    if (-not $normalizedTypeCounts.ContainsKey($objectClass)) {
        $normalizedTypeCounts[$objectClass] = 0
    }
    $normalizedTypeCounts[$objectClass]++

    $components = Parse-DN -DN $dn
    
    # Skip if no valid components found
    if ($components.Count -eq 0) {
        Write-Warning "No valid components found in DN: $dn"
        continue
    }
    
    # Build path
    $currentPath = ""
    for ($i = 0; $i -lt $components.Count; $i++) {
        $comp = $components[$i]
        
        # Validate component has a name
        if ([string]::IsNullOrWhiteSpace($comp.Name)) {
            Write-Warning "Empty component name found at level $i in DN: $dn"
            continue
        }
        
        if ($i -eq 0) {
            $currentPath = $comp.Name
        } else {
            $currentPath = "$currentPath/$($comp.Name)"
        }
        
        # Validate currentPath is not null
        if ([string]::IsNullOrWhiteSpace($currentPath)) {
            Write-Warning "Invalid path generated for DN: $dn (Component: $($comp.Name))"
            continue
        }
        
        if (-not $processedPaths.ContainsKey($currentPath)) {
            # Determine type based on component type and position
            $nodeType = if ($i -eq $components.Count - 1) { 
                # This is the leaf node, use the object's class
                $objectClass 
            } elseif ($comp.Type -eq 'DC' -and $comp.IsDomainRoot) {
                # Domain component
                'domain'
            } elseif ($comp.Type -eq 'DC') {
                # Non-domain DC entries (e.g., dnsZone/dnsNode) are containers
                'container'
            } elseif ($comp.Type -eq 'OU') {
                # Organizational Unit
                'organizationalUnit'
            } elseif ($comp.Type -eq 'CN') {
                # Container (for non-leaf CN entries)
                'container'
            } else {
                'organizationalUnit'
            }
            
            $processedPaths[$currentPath] = @{
                Name = $comp.Name
                Level = $i
                Children = @()
                Type = $nodeType
                FullPath = $currentPath
            }
            
            # Add to parent if exists
            if ($i -gt 0) {
                $parentPath = $currentPath.Substring(0, $currentPath.LastIndexOf('/'))
                if ($processedPaths.ContainsKey($parentPath)) {
                    $processedPaths[$parentPath].Children += $currentPath
                }
            }
        }
    }
}

# Function to output tree recursively
function Write-Tree {
    param(
        [string]$Path,
        [int]$Level = 0
    )
    
    if (-not $processedPaths.ContainsKey($Path)) { return }
    
    $node = $processedPaths[$Path]
    $indent = '  ' * $Level
    $typeMarker = Get-TypeMarker $node.Type
    
    # Clean display names
    $displayName = $node.Name
    
    # Remove $ from computer names
    if ($node.Type -eq 'computer' -and $displayName -like '*$') {
        $displayName = $displayName -replace '\$$', ''
    }
    
    $output = "$indent- $displayName$typeMarker"
    
    # Sort children
    $children = $node.Children | Sort-Object
    
    foreach ($child in $children) {
        $output += "`n" + (Write-Tree -Path $child -Level ($Level + 1))
    }
    
    return $output
}

# Find root nodes (those without parents in our structure)
$rootNodes = @()
foreach ($path in $processedPaths.Keys) {
    if ($processedPaths[$path].Level -eq 0) {
        $rootNodes += $path
    }
}

# Generate markdown
Write-Host "Generating markdown..." -ForegroundColor Cyan
$markdown = @()
$markdown += "# Active Directory Structure"
$markdown += ""
$markdown += "Converted from: $((Get-Item $CsvFile).Name)"
$markdown += "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$markdown += "Total Objects: $($processedPaths.Count)"
$markdown += ""

foreach ($root in $rootNodes | Sort-Object) {
    $markdown += Write-Tree -Path $root -Level 0
}

$output = $markdown -join "`n"

# Save or output
if ($OutputFile) {
    $output | Set-Content -Path $OutputFile -Encoding UTF8
    Write-Host "Markdown saved to: $OutputFile" -ForegroundColor Green
    Write-Host "Total objects processed: $($processedPaths.Count)" -ForegroundColor Cyan
    
    # Show summary
    Write-Host "`nObject type summary (normalized from CSV):" -ForegroundColor Yellow
    
    $supportedTypes = @('user', 'computer', 'group', 'contact', 'printer', 'share', 'policy', 
                       'container', 'organizationalUnit', 'domain', 'domainDNS', 'dnsNode', 'dnsZone')
    $unsupportedTypes = @()
    
    if ($normalizedTypeCounts.Count -eq 0) {
        Write-Host "  No objects were processed from the CSV." -ForegroundColor Gray
    } else {
        $sortedTypes = $normalizedTypeCounts.GetEnumerator() | Sort-Object Value -Descending
        foreach ($type in $sortedTypes) {
            $marker = Get-TypeMarker $type.Key
            $display = if ($marker) { "$($type.Key)$marker" } else { $type.Key }
            Write-Host ("  {0}: {1}" -f $display, $type.Value) -ForegroundColor Gray
            
            # Check for unsupported types
            if ($type.Key -notin $supportedTypes -and -not [string]::IsNullOrWhiteSpace($type.Key)) {
                $unsupportedTypes += $type.Key
            }
        }
        
        $uniqueUnsupported = $unsupportedTypes | Where-Object { $_ } | Sort-Object -Unique
        if ($uniqueUnsupported.Count -gt 0) {
            Write-Warning "`nFound unsupported object types that may not display correctly:"
            foreach ($unsupType in $uniqueUnsupported) {
                Write-Warning "  - $unsupType"
            }
            Write-Host "Run Test-ObjectClassMapping.ps1 for diagnostics" -ForegroundColor Yellow
        } else {
            Write-Host "`nAll object classes mapped to known markers." -ForegroundColor Green
        }
    }
} else {
    Write-Output $output
}

Write-Host "`nConversion complete!" -ForegroundColor Green
Write-Host "Note: If you see warnings, check that your CSV has valid Distinguished Names." -ForegroundColor Yellow

# Provide sample export command
Write-Host "`nTip: To export AD objects to CSV for this script, use:" -ForegroundColor Yellow
Write-Host '  Get-ADObject -Filter * -Properties Name,ObjectClass,Enabled |' -ForegroundColor Gray
Write-Host '    Export-Csv -Path "ADExport.csv" -NoTypeInformation' -ForegroundColor Gray
