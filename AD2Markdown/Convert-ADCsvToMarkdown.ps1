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
    if ($obj.PSObject.Properties['ObjectClass']) {
        return $obj.ObjectClass
    } elseif ($obj.PSObject.Properties['Type']) {
        return $obj.Type
    } elseif ($obj.PSObject.Properties['objectCategory']) {
        # Try to infer from objectCategory
        if ($obj.objectCategory -like '*Person*') { return 'user' }
        if ($obj.objectCategory -like '*Computer*') { return 'computer' }
        if ($obj.objectCategory -like '*Group*') { return 'group' }
        if ($obj.objectCategory -like '*Organizational-Unit*') { return 'organizationalUnit' }
    }
    return 'unknown'
}

# Function to parse DN into components
function Parse-DN {
    param([string]$DN)
    
    $components = @()
    $parts = $DN -split ',(?=\w+=)'
    
    for ($i = $parts.Count - 1; $i -ge 0; $i--) {
        if ($parts[$i] -match '^(CN|OU)=(.+)$') {
            $type = $Matches[1]
            $name = $Matches[2]
            $components += @{
                Type = $type
                Name = $name
                Level = $parts.Count - $i - 1
            }
        }
    }
    
    return $components
}

# Function to get type marker
function Get-TypeMarker($objectClass) {
    switch -Wildcard ($objectClass) {
        'user' { return ' [user]' }
        'computer' { return ' [computer]' }
        'group' { return ' [group]' }
        'contact' { return ' [contact]' }
        'organizationalUnit' { return '' }
        'container' { return '' }
        '*ServiceAccount' { return ' [user]' }
        default { return '' }
    }
}

# Build tree structure
Write-Host "Building tree structure..." -ForegroundColor Cyan
$tree = @{}
$processedPaths = @{}

foreach ($obj in $adObjects) {
    # Skip disabled if requested
    if (-not $IncludeDisabled -and $obj.PSObject.Properties['Enabled']) {
        if ($obj.Enabled -eq 'False' -or $obj.Enabled -eq $false) {
            continue
        }
    }
    
    $dn = Get-ObjectDN $obj
    if (-not $dn) { continue }
    
    $objectClass = Get-ObjectClass $obj
    $components = Parse-DN -DN $dn
    
    # Build path
    $currentPath = ""
    for ($i = 0; $i -lt $components.Count; $i++) {
        $comp = $components[$i]
        
        if ($i -eq 0) {
            $currentPath = $comp.Name
        } else {
            $currentPath = "$currentPath/$($comp.Name)"
        }
        
        if (-not $processedPaths.ContainsKey($currentPath)) {
            $processedPaths[$currentPath] = @{
                Name = $comp.Name
                Level = $i
                Children = @()
                Type = if ($i -eq $components.Count - 1) { $objectClass } else { 'organizationalUnit' }
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
    
    # Clean computer names (remove $)
    $displayName = $node.Name -replace '\$$', ''
    
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
} else {
    Write-Output $output
}

Write-Host "Conversion complete!" -ForegroundColor Green

# Provide sample export command
Write-Host "`nTip: To export AD objects to CSV for this script, use:" -ForegroundColor Yellow
Write-Host '  Get-ADObject -Filter * -Properties Name,ObjectClass,Enabled |' -ForegroundColor Gray
Write-Host '    Export-Csv -Path "ADExport.csv" -NoTypeInformation' -ForegroundColor Gray