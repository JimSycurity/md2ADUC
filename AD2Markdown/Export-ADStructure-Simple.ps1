<#
.SYNOPSIS
    Simple AD structure export to markdown for AD Structure Visualizer.

.DESCRIPTION
    Lightweight script to quickly export AD structure to markdown format.
    Focuses on common object types and standard OU structures.

.PARAMETER Path
    OU path to export. Defaults to domain root.

.PARAMETER OutFile
    Output file path. Defaults to console output.

.PARAMETER Depth
    Maximum depth to traverse. Default is 5.

.EXAMPLE
    .\Export-ADStructure-Simple.ps1
    
.EXAMPLE
    .\Export-ADStructure-Simple.ps1 -Path "OU=Corporate,DC=contoso,DC=com" -OutFile "corporate.md"
#>

param(
    [string]$Path,
    [string]$OutFile,
    [int]$Depth = 5
)

# Import AD module
Import-Module ActiveDirectory -ErrorAction Stop

# Get domain DN if not specified
if (-not $Path) {
    $Path = (Get-ADDomain).DistinguishedName
}

# Function to get object type marker
function Get-TypeMarker($objectClass) {
    switch ($objectClass) {
        'user' { return ' [user]' }
        'computer' { return ' [computer]' }
        'group' { return ' [group]' }
        'organizationalUnit' { return '' }
        'container' { return '' }
        default { return '' }
    }
}

# Recursive function to build tree
function Get-ADTree {
    param(
        [string]$DN,
        [int]$Level = 0,
        [int]$MaxLevel = 5
    )
    
    if ($Level -ge $MaxLevel) { return }
    
    $indent = '  ' * $Level
    $output = @()
    
    # Get child objects
    $children = Get-ADObject -SearchBase $DN -SearchScope OneLevel `
                            -Filter * -Properties objectClass |
                Sort-Object objectClass, Name
    
    foreach ($child in $children) {
        # Skip some system containers (but keep service accounts and their container)
        if ($child.Name -in @('System', 'Builtin', 'ForeignSecurityPrincipals', 
                             'Program Data', 'NTDS Quotas')) {
            continue
        }
        
        # Format name (remove $ from computer names)
        $name = $child.Name -replace '\$$', ''
        $type = Get-TypeMarker $child.objectClass
        
        $output += "$indent- $name$type"
        
        # Recurse for containers and OUs
        if ($child.objectClass -in @('organizationalUnit', 'container')) {
            $subItems = Get-ADTree -DN $child.DistinguishedName `
                                  -Level ($Level + 1) `
                                  -MaxLevel $MaxLevel
            $output += $subItems
        }
    }
    
    return $output
}

# Get root name
$rootObj = Get-ADObject -Identity $Path
$rootName = $rootObj.Name

# Build structure
Write-Host "Exporting AD structure from: $Path" -ForegroundColor Cyan
$markdown = @("- $rootName")
$markdown += Get-ADTree -DN $Path -Level 1 -MaxLevel $Depth

# Output
$result = $markdown -join "`n"

if ($OutFile) {
    $result | Set-Content -Path $OutFile -Encoding UTF8
    Write-Host "Exported to: $OutFile" -ForegroundColor Green
} else {
    Write-Output $result
}

Write-Host "Export complete! Total objects: $($markdown.Count)" -ForegroundColor Green