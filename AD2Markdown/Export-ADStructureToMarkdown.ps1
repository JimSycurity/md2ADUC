<#
.SYNOPSIS
    Exports Active Directory structure to markdown format compatible with AD Structure Visualizer.

.DESCRIPTION
    This script recursively collects AD objects from a specified naming context and exports them
    as a markdown unordered list with proper indentation and object type annotations.
    The output is formatted for use with the AD Structure Visualizer HTML tool.

.PARAMETER SearchBase
    The Distinguished Name of the container to start the export from.
    Defaults to the current domain's distinguished name.

.PARAMETER OutputFile
    Path to save the markdown file. If not specified, outputs to console.

.PARAMETER IncludeDisabled
    Include disabled user and computer accounts in the export.

.PARAMETER IncludeEmptyOUs
    Include Organizational Units that contain no objects.

.PARAMETER MaxDepth
    Maximum depth to traverse. Default is 10 levels.

.PARAMETER ExcludeSystemContainers
    Exclude well-known system containers like 'CN=System', 'CN=Builtin', etc.

.PARAMETER IncludeContacts
    Include contact objects in the export.

.PARAMETER ExcludeServiceAccounts
    Exclude managed service accounts and group managed service accounts from the export.
    By default, service accounts are included.

.EXAMPLE
    .\Export-ADStructureToMarkdown.ps1
    Exports the entire domain structure to console output.

.EXAMPLE
    .\Export-ADStructureToMarkdown.ps1 -SearchBase "OU=Corporate,DC=contoso,DC=com" -OutputFile "C:\temp\corporate-structure.md"
    Exports only the Corporate OU structure to a file.

.EXAMPLE
    .\Export-ADStructureToMarkdown.ps1 -ExcludeSystemContainers -ExcludeServiceAccounts -OutputFile "AD-Structure.md"
    Exports domain structure excluding system containers and service accounts to a markdown file.

.NOTES
    Author: Security Consultant
    Requires: Active Directory PowerShell Module
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$SearchBase,
    
    [Parameter(Position = 1)]
    [string]$OutputFile,
    
    [switch]$IncludeDisabled,
    
    [switch]$IncludeEmptyOUs,
    
    [int]$MaxDepth = 10,
    
    [switch]$ExcludeSystemContainers,
    
    [switch]$IncludeContacts,
    
    [switch]$ExcludeServiceAccounts,
    
    [switch]$Verbose
)

#region Functions

function Write-VerboseLog {
    param([string]$Message)
    if ($Verbose) {
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor Gray
    }
}

function Get-ObjectType {
    <#
    .SYNOPSIS
        Determines the object type for markdown annotation.
    #>
    param(
        [Microsoft.ActiveDirectory.Management.ADObject]$ADObject
    )
    
    switch ($ADObject.objectClass) {
        'user' {
            # Check if it's a service account
            if ($ADObject.samAccountName -like 'svc_*' -or 
                $ADObject.samAccountName -like '*service*' -or
                $ADObject.objectClass -contains 'msDS-ManagedServiceAccount' -or
                $ADObject.objectClass -contains 'msDS-GroupManagedServiceAccount') {
                return 'user'  # Still mark as user but could add logic for different handling
            }
            return 'user'
        }
        'computer' { return 'computer' }
        'group' { return 'group' }
        'contact' { return 'contact' }
        'printQueue' { return 'printer' }
        'volume' { return 'share' }
        'groupPolicyContainer' { return 'policy' }
        'container' { return 'container' }
        'organizationalUnit' { return $null }  # OUs don't get a type marker
        'msDS-ManagedServiceAccount' { return 'user' }
        'msDS-GroupManagedServiceAccount' { return 'user' }
        default { 
            # For unknown types, check if it's a container-like object
            if ($ADObject.objectClass -contains 'container' -or 
                $ADObject.objectClass -contains 'organizationalUnit') {
                return $null
            }
            return 'default'
        }
    }
}

function Test-ShouldIncludeObject {
    <#
    .SYNOPSIS
        Determines if an object should be included in the export.
    #>
    param(
        [Microsoft.ActiveDirectory.Management.ADObject]$ADObject
    )
    
    # Check if disabled and we're excluding disabled
    if (-not $IncludeDisabled) {
        if ($ADObject.objectClass -eq 'user' -or $ADObject.objectClass -eq 'computer') {
            try {
                $userAccountControl = $ADObject.userAccountControl
                if ($userAccountControl -band 0x2) {  # ACCOUNTDISABLE flag
                    Write-VerboseLog "Excluding disabled object: $($ADObject.Name)"
                    return $false
                }
            }
            catch {
                # If we can't determine, include it
            }
        }
    }
    
    # Check contacts
    if ($ADObject.objectClass -eq 'contact' -and -not $IncludeContacts) {
        Write-VerboseLog "Excluding contact: $($ADObject.Name)"
        return $false
    }
    
    # Check service accounts
    if ($ExcludeServiceAccounts) {
        if ($ADObject.objectClass -eq 'msDS-ManagedServiceAccount' -or 
            $ADObject.objectClass -eq 'msDS-GroupManagedServiceAccount') {
            Write-VerboseLog "Excluding service account: $($ADObject.Name)"
            return $false
        }
    }
    
    return $true
}

function Get-ADStructure {
    <#
    .SYNOPSIS
        Recursively retrieves AD structure and formats as markdown.
    #>
    param(
        [string]$BaseDN,
        [int]$IndentLevel = 0,
        [int]$CurrentDepth = 0
    )
    
    if ($CurrentDepth -ge $MaxDepth) {
        Write-VerboseLog "Maximum depth reached at $BaseDN"
        return
    }
    
    $indent = "  " * $IndentLevel  # 2 spaces per level
    $output = @()
    
    try {
        # Get all child objects
        Write-VerboseLog "Processing: $BaseDN (Depth: $CurrentDepth)"
        
        $ldapFilter = "(|(objectClass=organizationalUnit)(objectClass=container)(objectClass=user)" +
                      "(objectClass=computer)(objectClass=group)(objectClass=contact)" +
                      "(objectClass=printQueue)(objectClass=volume)(objectClass=groupPolicyContainer)" +
                      "(objectClass=msDS-ManagedServiceAccount)(objectClass=msDS-GroupManagedServiceAccount))"
        
        $childObjects = Get-ADObject -SearchBase $BaseDN -SearchScope OneLevel -LDAPFilter $ldapFilter `
                                     -Properties objectClass, userAccountControl, samAccountName -ErrorAction SilentlyContinue
        
        if ($null -eq $childObjects -or $childObjects.Count -eq 0) {
            Write-VerboseLog "No child objects found in $BaseDN"
            if (-not $IncludeEmptyOUs) {
                return
            }
        }
        
        # Group objects by type for better organization
        $ous = $childObjects | Where-Object { $_.objectClass -eq 'organizationalUnit' } | Sort-Object Name
        $containers = $childObjects | Where-Object { $_.objectClass -eq 'container' -and $_.objectClass -ne 'organizationalUnit' } | Sort-Object Name
        $users = $childObjects | Where-Object { $_.objectClass -eq 'user' -and $_.objectClass -notlike '*ServiceAccount' } | Sort-Object Name
        $computers = $childObjects | Where-Object { $_.objectClass -eq 'computer' } | Sort-Object Name
        $groups = $childObjects | Where-Object { $_.objectClass -eq 'group' } | Sort-Object Name
        $other = $childObjects | Where-Object { 
            $_.objectClass -ne 'organizationalUnit' -and 
            $_.objectClass -ne 'container' -and 
            $_.objectClass -ne 'user' -and 
            $_.objectClass -ne 'computer' -and 
            $_.objectClass -ne 'group' 
        } | Sort-Object Name
        
        # Process OUs first
        foreach ($ou in $ous) {
            if (-not (Test-ShouldIncludeObject -ADObject $ou)) { continue }
            
            $output += "$indent- $($ou.Name)"
            
            # Recursively process this OU
            $subItems = Get-ADStructure -BaseDN $ou.DistinguishedName `
                                       -IndentLevel ($IndentLevel + 1) `
                                       -CurrentDepth ($CurrentDepth + 1)
            if ($subItems) {
                $output += $subItems
            }
        }
        
        # Process containers
        foreach ($container in $containers) {
            if (-not (Test-ShouldIncludeObject -ADObject $container)) { continue }
            
            # Skip system containers if requested
            if ($ExcludeSystemContainers) {
                $systemContainers = @('System', 'Builtin', 'ForeignSecurityPrincipals', 
                                     'Program Data', 'Microsoft Exchange Security Groups', 
                                     'NTDS Quotas', 'TPM Devices', 'Keys', 'Schema', 'Configuration')
                
                # Keep Managed Service Accounts container if we're including service accounts
                if (-not $ExcludeServiceAccounts) {
                    # Don't skip Managed Service Accounts container
                } else {
                    $systemContainers += 'Managed Service Accounts'
                }
                
                if ($container.Name -in $systemContainers) {
                    Write-VerboseLog "Excluding system container: $($container.Name)"
                    continue
                }
            }
            
            $type = Get-ObjectType -ADObject $container
            $typeMarker = if ($type) { " [$type]" } else { "" }
            $output += "$indent- $($container.Name)$typeMarker"
            
            # Recursively process this container
            $subItems = Get-ADStructure -BaseDN $container.DistinguishedName `
                                       -IndentLevel ($IndentLevel + 1) `
                                       -CurrentDepth ($CurrentDepth + 1)
            if ($subItems) {
                $output += $subItems
            }
        }
        
        # Process users
        foreach ($user in $users) {
            if (-not (Test-ShouldIncludeObject -ADObject $user)) { continue }
            
            $type = Get-ObjectType -ADObject $user
            $typeMarker = if ($type) { " [$type]" } else { "" }
            $displayName = if ($user.Name) { $user.Name } else { $user.samAccountName }
            $output += "$indent- $displayName$typeMarker"
        }
        
        # Process computers
        foreach ($computer in $computers) {
            if (-not (Test-ShouldIncludeObject -ADObject $computer)) { continue }
            
            $type = Get-ObjectType -ADObject $computer
            $typeMarker = if ($type) { " [$type]" } else { "" }
            $computerName = $computer.Name -replace '\$$', ''  # Remove trailing $
            $output += "$indent- $computerName$typeMarker"
        }
        
        # Process groups
        foreach ($group in $groups) {
            if (-not (Test-ShouldIncludeObject -ADObject $group)) { continue }
            
            $type = Get-ObjectType -ADObject $group
            $typeMarker = if ($type) { " [$type]" } else { "" }
            $output += "$indent- $($group.Name)$typeMarker"
        }
        
        # Process other objects
        foreach ($obj in $other) {
            if (-not (Test-ShouldIncludeObject -ADObject $obj)) { continue }
            
            $type = Get-ObjectType -ADObject $obj
            $typeMarker = if ($type) { " [$type]" } else { "" }
            $output += "$indent- $($obj.Name)$typeMarker"
        }
        
    }
    catch {
        Write-Warning "Error processing $BaseDN : $_"
    }
    
    return $output
}

#endregion Functions

#region Main Script

# Check for Active Directory module
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Error "Active Directory PowerShell module is not installed. Please install RSAT or run this on a domain controller."
    exit 1
}

# Import AD module
Import-Module ActiveDirectory -ErrorAction Stop

# Get default search base if not specified
if (-not $SearchBase) {
    try {
        $SearchBase = (Get-ADDomain).DistinguishedName
        Write-Host "Using domain root as search base: $SearchBase" -ForegroundColor Cyan
    }
    catch {
        Write-Error "Could not determine domain. Please specify -SearchBase parameter."
        exit 1
    }
}

# Validate search base exists
try {
    $rootObject = Get-ADObject -Identity $SearchBase -ErrorAction Stop
}
catch {
    Write-Error "Cannot find object with DN: $SearchBase"
    exit 1
}

Write-Host "Starting AD structure export..." -ForegroundColor Green
Write-Host "Search Base: $SearchBase" -ForegroundColor Cyan

# Build the markdown structure
$markdownLines = @()

# Add header
$markdownLines += "# Active Directory Structure Export"
$markdownLines += ""
$markdownLines += "Exported on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$markdownLines += "Domain: $((Get-ADDomain).DNSRoot)"
$markdownLines += "Search Base: $SearchBase"
$markdownLines += ""

# Get the root object name
$rootName = if ($rootObject.Name) { $rootObject.Name } else { $rootObject.DistinguishedName }

# Add the root
$markdownLines += "- $rootName"

# Get all child objects recursively
$structure = Get-ADStructure -BaseDN $SearchBase -IndentLevel 1 -CurrentDepth 0

if ($structure) {
    $markdownLines += $structure
}

# Output results
$finalOutput = $markdownLines -join "`n"

if ($OutputFile) {
    try {
        $finalOutput | Set-Content -Path $OutputFile -Encoding UTF8
        Write-Host "`nMarkdown file saved to: $OutputFile" -ForegroundColor Green
        Write-Host "Total lines exported: $($markdownLines.Count)" -ForegroundColor Cyan
        
        # Show file size
        $fileInfo = Get-Item $OutputFile
        Write-Host "File size: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Cyan
    }
    catch {
        Write-Error "Failed to save file: $_"
    }
}
else {
    # Output to console
    Write-Output $finalOutput
}

Write-Host "`nExport complete!" -ForegroundColor Green

#endregion Main Script